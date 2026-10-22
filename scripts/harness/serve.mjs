// In-process proxy server for the harness — wraps createApp() over node:http so the
// harness needs no separate Nitro build. Auth uses the repo's cached M365 login.
// Usage: node scripts/harness/serve.mjs [port]   (default 4142)
import { createServer } from "node:http";
import { createApp } from "../../packages/proxy-lib/dist/index.mjs";
import { getToken } from "../../packages/core/dist/index.mjs";

const PORT = Number(process.argv[2] || process.env.PORT || 4142);
await getToken(); // warm the token cache before serving
const app = createApp({});

createServer(async (req, res) => {
  try {
    const chunks = [];
    for await (const c of req) chunks.push(c);
    const hasBody = req.method !== "GET" && req.method !== "HEAD";
    const request = new Request(`http://localhost:${PORT}${req.url}`, {
      method: req.method,
      headers: req.headers,
      body: hasBody && chunks.length ? Buffer.concat(chunks) : undefined,
    });
    const response = await app.fetch(request);
    res.writeHead(response.status, Object.fromEntries(response.headers));
    res.end(Buffer.from(await response.arrayBuffer()));
  } catch (e) {
    res.writeHead(500, { "content-type": "application/json" });
    res.end(JSON.stringify({ error: { message: String(e?.message || e) } }));
  }
}).listen(PORT, () => console.log(`[serve] harness proxy on http://localhost:${PORT}`));
