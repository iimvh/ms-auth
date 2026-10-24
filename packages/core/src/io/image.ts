
import { getToken, getImageArtifactToken } from "../adapters/msal/auth.js";
import { CopilotSession } from "../protocol/session.js";
import { selectCredential, turnContextFor } from "../app/credential.js";
import { createLogger, trunc } from "./log.js";
import { awaitTurnBudget } from "../budget/turn-budget.js";
import type { CapturedImage } from "../routing/copilot.js";

const log = createLogger("image");

export type ImageGenFailureReason = "quota_exceeded" | "capacity" | "content_filtered" | "no_image";

export class ImageGenerationError extends Error {
  constructor(public reason: ImageGenFailureReason, message: string) {
    super(message);
    this.name = "ImageGenerationError";
  }
}

export function classifyImageFailure(text: string): ImageGenFailureReason {
  const t = (text ?? "").toLowerCase();
  if (!t.trim()) return "no_image";
  if (/can.?t generate any more images|no more images (today|right now)|reached (your|the)[^.]*image|image (generation )?(limit|quota)|try again tomorrow/.test(t)) return "quota_exceeded";
  if (/trouble (creating|generating|making)[^.]*image|couldn.?t (create|generate)[^.]*image|try again (later|in a (bit|moment|few))/.test(t)) return "capacity";
  if (/can.?t (create|generate|make) (that|this|an? )[^.]*image|unable to (create|generate)[^.]*image|against[^.]*(policy|guideline)|can.?t help (with|create) that/.test(t)) return "content_filtered";
  return "no_image";
}

export type ImageOrientation = "landscape" | "portrait" | "square";

export type ImageStyle = "natural" | "icon" | "story" | "designer";

export interface GeneratedImage {
  url: string;
  contentType: string;
  data: Buffer;
  base64: string;
  size?: string;
  orientation?: string;
}

export interface GenerateImageOptions {
  model?: string;
  token?: string;
  artifactToken?: string;
  signal?: AbortSignal;
  urlsOnly?: boolean;
  orientation?: ImageOrientation;
  style?: ImageStyle;
}

export function buildImagePrompt(prompt: string, opts: { orientation?: ImageOrientation; style?: ImageStyle } = {}): string {
  const directives: string[] = [];
  switch (opts.style) {
    case "icon": directives.push("Render it as a clean app icon."); break;
    case "story": directives.push("Render it as a multi-panel illustrated story."); break;
    case "designer": directives.push("Render it as a polished graphic-design composition."); break;
    case "natural": case undefined: break;
  }
  switch (opts.orientation) {
    case "landscape": directives.push("Use a landscape (wide, 16:9) orientation."); break;
    case "portrait": directives.push("Use a portrait (tall, 9:16) orientation."); break;
    case "square": directives.push("Use a square (1:1) orientation."); break;
    case undefined: break;
  }
  return directives.length ? `${prompt.trim()}\n\n${directives.join(" ")}` : prompt;
}

export async function fetchImageBytes(
  url: string,
  artifactToken: string,
): Promise<{ data: Buffer; contentType: string }> {
  const res = await fetch(url, { headers: { Authorization: `Bearer ${artifactToken}` } });
  if (!res.ok) {
    throw new Error(`Image artifact fetch failed: ${res.status} ${res.statusText} for ${url.slice(0, 120)}`);
  }
  const contentType = res.headers.get("content-type") ?? "image/png";
  const data = Buffer.from(await res.arrayBuffer());
  return { data, contentType };
}

export async function generateImage(
  prompt: string,
  opts: GenerateImageOptions = {},
): Promise<GeneratedImage[]> {
  const token = opts.token ?? (await getToken());
  const model = opts.model ?? "m365-copilot";
  const effectivePrompt = buildImagePrompt(prompt, { orientation: opts.orientation, style: opts.style });

  const selection = await selectCredential(null);
  const session = new CopilotSession({
    model,
    turnContext: (plan) => turnContextFor(selection, { ...plan, images: true }),
  });
  await awaitTurnBudget({ signal: opts.signal });
  const stream = await session.chat(token, effectivePrompt, model, opts.signal, { generateImages: true });

  for await (const _ of stream) {
  }

  const captured: CapturedImage[] = stream.images;
  if (captured.length === 0) {
    const reason = classifyImageFailure(stream.fullText ?? "");
    if (reason !== "no_image") {
      const msg = stream.fullText?.trim() || `Image generation failed: ${reason}`;
      log.info(`Image gen failed (${reason}) for ${JSON.stringify(prompt.slice(0, 80))}: ${trunc(msg, 120)}`);
      throw new ImageGenerationError(reason, msg);
    }
    log.info(`No image generated for prompt ${JSON.stringify(prompt.slice(0, 80))} (model did not draw)`);
    return [];
  }
  log.info(`Captured ${captured.length} image(s) for prompt ${JSON.stringify(prompt.slice(0, 80))}`);

  const urls = captured
    .map((c) => ({ url: c.referenceUrls[0], meta: c }))
    .filter((x): x is { url: string; meta: CapturedImage } => !!x.url);

  if (opts.urlsOnly) {
    return urls.map(({ url, meta }) => ({
      url, contentType: "image/png", data: Buffer.alloc(0), base64: "",
      size: meta.size, orientation: meta.orientation,
    }));
  }

  const artifactToken = opts.artifactToken ?? (await getImageArtifactToken());
  if (!artifactToken) {
    throw new Error("Could not acquire designerappservice token to fetch image bytes");
  }

  const out: GeneratedImage[] = [];
  for (const { url, meta } of urls) {
    const { data, contentType } = await fetchImageBytes(url, artifactToken);
    out.push({ url, contentType, data, base64: data.toString("base64"), size: meta.size, orientation: meta.orientation });
  }
  return out;
}
