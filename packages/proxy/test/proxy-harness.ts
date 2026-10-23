import {
  createApp,
  createEvent,
  defineEventHandler,
  getRouterParam,
  setResponseStatus,
  toNodeListener,
  type H3Event,
} from "h3";
import { createServer } from "node:http";
import { readdirSync } from "node:fs";
import { join } from "node:path";
import { pathToFileURL } from "node:url";

export type NitroHandler = (event: H3Event, matched: RouteMatch) => unknown;

export interface RouteMatch {
  route: string;
  params: Record<string, string>;
  method: string;
}

const ROUTE_FILE = /\.(get|post|put|patch|delete|options|head)\.[cm]?[jt]s$/;

const PARAM_SEGMENT = /^\[(\.\.\.)?(.+)\]$/;

const PARAM_PLACEHOLDER = /^:(.+)$/;

const MIDDLEWARE_CORS = "middleware/cors.ts";

const AUTH_PLUGIN = "plugins/auth.ts";

export const toRoutePath = (file: string): string =>
  `/${file.replace(/^routes\//, "")}`
    .replace(ROUTE_FILE, "")
    .replace(/\/index$/, "")
    .split("/")
    .map((segment) => segment.replace(PARAM_SEGMENT, ":$2"))
    .join("/");

export const toRouteMethod = (file: string): string => file.match(ROUTE_FILE)![1].toUpperCase();

export const toNitroPath = (file: string): string => toRoutePath(file).split(":")[0];

export function routeFiles(root: string): string[] {
  const found: string[] = [];
  const walk = (dir: string) => {
    for (const entry of readdirSync(join(root, dir), { withFileTypes: true })) {
      const rel = `${dir}/${entry.name}`;
      if (entry.isDirectory()) walk(rel);
      else if (ROUTE_FILE.test(entry.name)) found.push(rel);
    }
  };
  walk("routes");
  return found;
}

function segments(path: string): string[] {
  return path.split("/").filter(Boolean);
}

export function matchRoute(route: string, method: string, path: string, requestMethod: string): RouteMatch | null {
  if (method !== requestMethod.toUpperCase()) return null;
  const wanted = segments(route);
  const given = segments(path.split("?")[0]);
  const params: Record<string, string> = {};
  for (let i = 0; i < wanted.length; i++) {
    const segment = wanted[i];
    const param = segment.match(PARAM_SEGMENT);
    const placeholder = segment.match(PARAM_PLACEHOLDER);
    if (param?.[1]) {
      params[param[2]] = given.slice(i).join("/");
      return { route, params, method };
    }
    const name = param?.[2] ?? placeholder?.[1];
    if (name) {
      if (given[i] === undefined) return null;
      params[name] = given[i];
      continue;
    }
    if (segment !== given[i]) return null;
  }
  return wanted.length === given.length ? { route, params, method } : null;
}

export interface MountedProxy {
  origin: string;
  request(path: string, init?: RequestInit): Promise<Response>;
  boot(): Promise<void>;
  close(): Promise<void>;
}

async function load(root: string, file: string): Promise<Record<string, unknown>> {
  return import(pathToFileURL(join(root, file)).href);
}

function handlerOf(mod: Record<string, unknown>, file: string): (event: H3Event) => unknown {
  const handler = mod.default;
  if (typeof handler !== "function") throw new Error(`${file} has no default export`);
  return handler as (event: H3Event) => unknown;
}

export async function startProxy(root: string): Promise<MountedProxy> {
  const routes: Array<{ path: string; method: string; handler: (event: H3Event) => unknown }> = [];
  for (const file of routeFiles(root)) {
    routes.push({ path: toRoutePath(file), method: toRouteMethod(file), handler: handlerOf(await load(root, file), file) });
  }
  const cors = handlerOf(await load(root, MIDDLEWARE_CORS), MIDDLEWARE_CORS);
  const boot = handlerOf(await load(root, AUTH_PLUGIN), AUTH_PLUGIN) as (nitro: unknown) => Promise<void>;

  const listener = toNodeListener(
    createApp().use(
      defineEventHandler(async (event) => {
        if (await cors(event)) return;
        const match = routes
          .map((route) => ({ route, match: matchRoute(route.path, route.method, event.path, event.method) }))
          .find((candidate) => candidate.match);
        if (!match?.match) return;
        event.context.params = match.match.params;
        event.context.matchedRoute = match.match;
        return match.route.handler(event);
      }),
    ),
  );

  const server = createServer((req, res) => {
    listener(req, res);
  });
  await new Promise<void>((resolve, reject) => {
    server.once("error", reject);
    server.listen(0, "127.0.0.1", () => resolve());
  });
  const address = server.address();
  const origin = `http://127.0.0.1:${typeof address === "object" && address ? address.port : 0}`;
  return {
    origin,
    request: (path, init) => fetch(`${origin}${path}`, init),
    boot: () => boot({}),
    close: () => new Promise<void>((done, fail) => server.close((err) => (err ? fail(err) : done()))),
  };
}

export { createEvent, getRouterParam, setResponseStatus };
