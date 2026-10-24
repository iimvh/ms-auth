import { readFileSync } from "node:fs";
import { z } from "zod/v4";

export const FrameDumpLine = z.object({
  t: z.number(),
  dir: z.enum(["recv", "send"]),
  frame: z.unknown(),
  synthetic: z.string().optional(),
});

export type FrameDumpLine = z.infer<typeof FrameDumpLine>;

export class FrameFixtureError extends Error {}

export interface ReplayMessage {
  at: number;
  line: number;
  data: Buffer;
}

export interface ReplayFixture {
  lines: FrameDumpLine[];
  received: ReplayMessage[];
  synthetic: string | null;
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
