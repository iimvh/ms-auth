import { readFileSync } from "node:fs";
import { z } from "zod/v4";

/** One line of a `M365_DUMP_FRAMES=1` turn dump: the shape `dumpFrame` in
 *  `session.ts` appends to `<FRAME_DIR>/<requestId>.ndjson`, plus the optional
 *  `synthetic` note a hand-written fixture carries on every line. */
export const FrameDumpLine = z.object({
  t: z.number(),
  dir: z.enum(["recv", "send"]),
  frame: z.unknown(),
  synthetic: z.string().optional(),
});

export type FrameDumpLine = z.infer<typeof FrameDumpLine>;

export class FrameFixtureError extends Error {}

export interface ReplayMessage {
  /** Milliseconds after the dump's first frame. */
  at: number;
  /** The 1-based dump line this message came from. */
  line: number;
  data: Buffer;
}

export interface ReplayFixture {
  /** Every parsed dump line, in dump order — both directions. */
  lines: FrameDumpLine[];
  /** The frames the proxy RECEIVED, one SignalR message each, in the order the
   *  live turn saw them. This is what a replay feeds the socket. */
  received: ReplayMessage[];
  /** The note every line carries in its `synthetic` field, or null on a
   *  recorded dump. */
  synthetic: string | null;
  /** The dump the fixture was loaded from. */
  file: string;
}

function dumpLines(text: string, file: string): FrameDumpLine[] {
  const lines: FrameDumpLine[] = [];
  for (const [index, raw] of text.split(/\r?\n/).entries()) {
    if (raw.trim().length === 0) continue;
    let parsed: unknown;
    try {
      parsed = JSON.parse(raw);
    } catch {
      throw new FrameFixtureError(`${file}:${index + 1}: not a JSON frame-dump line`);
    }
    const line = FrameDumpLine.safeParse(parsed);
    if (!line.success) throw new FrameFixtureError(`${file}:${index + 1}: not a frame-dump line (${line.error.issues.map((issue) => issue.path.join(".") || "line").join(", ")})`);
    lines.push(line.data);
  }
  return lines;
}

function syntheticNote(lines: FrameDumpLine[], file: string): string | null {
  const noted = lines.filter((line) => line.synthetic);
  if (noted.length === 0) return null;
  if (noted.length === lines.length) return noted[0]!.synthetic!;
  throw new FrameFixtureError(`${file}: ${noted.length} of ${lines.length} lines carry a synthetic note — a fixture is all synthetic or all recorded`);
}

/** Load a turn dump and rebuild what the socket delivered during that turn.
 *  A dump writes one line per received frame, and the live socket delivered
 *  each of those as one message; the dump's relative times are kept so a
 *  replay can pace the turn the way the recording arrived. */
export function loadFrameFixture(text: string, file = "<fixture>"): ReplayFixture {
  const lines = dumpLines(text, file);
  const first = lines[0];
  if (!first) throw new FrameFixtureError(`${file}: the dump is empty`);
  const received: ReplayMessage[] = [];
  for (const [index, line] of lines.entries()) {
    if (line.dir !== "recv") continue;
    received.push({ at: line.t - first.t, line: index + 1, data: Buffer.from(JSON.stringify(line.frame) + "\x1E") });
  }
  return { lines, received, synthetic: syntheticNote(lines, file), file };
}

export function readFrameFixture(path: string): ReplayFixture {
  return loadFrameFixture(readFileSync(path, "utf8"), path);
}
