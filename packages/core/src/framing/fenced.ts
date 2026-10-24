import { readFileSync } from "node:fs";
import { createLogger } from "../io/log.js";
import { getToneForModel, isSonnet55Model } from "../routing/copilot.js";
import type { ParsedToolCall, ToolDef } from "../tools/tools.js";

const log = createLogger("fenced");

//
// Known limitation: a `write_file` body that itself contains a ``` fence can't be
// carried unambiguously — this is exactly where JSON wins, and the bench A/B will
// show whether the escaping-free win on ordinary files outweighs it.

const BODY_PARAM_NAMES = [
  "command", "content", "code", "body", "script", "text",
  "query", "input", "patch", "cmd", "data", "contents",
];
const SEARCH_KEYS = ["old", "search", "find", "old_str", "old_string", "target"];
const REPLACE_KEYS = ["new", "replace", "replacement", "new_str", "new_string"];

const SHELL_LANGS = new Set([
  "bash", "sh", "shell", "zsh", "console", "shell-session", "shellsession", "shsession",
  "container.exec", "container.run", "container.bash",
  "powershell", "pwsh", "ps1", "posh", "cmd", "bat", "batch", "dosbatch",
]);
const SHELL_TOOL_NAME = /^(bash|sh|shell|zsh|run|exec|execute|command|cmd|terminal|run_command|run_terminal_cmd|execute_command|execute_bash|shell_exec|system)$/i;

const WRITE_PATH_PARAM = /^(path|file_path|filepath|filename|file)$/i;
const WRITE_CONTENT_PARAM = /^(content|contents|text|body)$/i;

export function findWriteTool(tools: ToolDef[]): ToolDef | undefined {
  return tools.find((t) => {
    const props = Object.keys(t.function.parameters?.properties ?? {});
    return props.some((p) => WRITE_PATH_PARAM.test(p)) &&
      props.some((p) => WRITE_CONTENT_PARAM.test(p)) &&
      !props.some((p) => SEARCH_KEYS.includes(p));
  });
}

export const WINDOWS_COMMAND_CAP = 8000;

const HEREDOC_WRITE = /(?:^|\n)([^\n]*?&&[ \t]*)?[ \t]*cat[ \t]*>[ \t]*(?:"([^"$`\\\n]+)"|'([^'\n]+)'|([^\s'"$`~;|&<>]+))[ \t]*<<[ \t]*(['"])([A-Za-z_][A-Za-z0-9_]*)\5[ \t]*\n([\s\S]*?)\n\6(?=\n|$)/;

export function longHeredocAsWrite(
  call: ParsedToolCall,
  tools: ToolDef[] | undefined,
  platform: NodeJS.Platform = process.platform,
): { call: ParsedToolCall; path: string; skipped: string } | null {
  if (platform !== "win32" || !tools?.length) return null;
  const shell = findShellTool(tools);
  const write = findWriteTool(tools);
  if (!shell || !write || call.function.name !== shell.function.name) return null;
  const bodyParam = deriveFencedSpec(shell).bodyParam;
  let args: Record<string, unknown>;
  try {
    args = JSON.parse(call.function.arguments);
  } catch {
    return null;
  }
  const command = bodyParam ? args[bodyParam] : undefined;
  if (typeof command !== "string" || command.length <= WINDOWS_COMMAND_CAP) return null;
  const m = HEREDOC_WRITE.exec(command);
  if (!m) return null;
  const path = m[2] ?? m[3] ?? m[4];
  const props = Object.keys(write.function.parameters?.properties ?? {});
  const pathKey = props.find((p) => WRITE_PATH_PARAM.test(p))!;
  const contentKey = props.find((p) => WRITE_CONTENT_PARAM.test(p))!;
  const before = `${command.slice(0, m.index)}\n${(m[1] ?? "").replace(/&&[ \t]*$/, "")}`.trim();
  const after = command.slice(m.index + m[0].length).trim();
  if (/(?:^|[\n;&|(])[ \t]*(?:cd|pushd)\b/.test(before)) return null;
  return {
    call: makeCall(write.function.name, { [pathKey]: path, [contentKey]: `${m[7]}\n` }),
    path,
    skipped: [before, after].filter(Boolean).join("\n"),
  };
}

export function findShellTool(tools: ToolDef[]): ToolDef | undefined {
  return tools.find((t) => SHELL_TOOL_NAME.test(t.function.name)) ??
    tools.find((t) => {
      const props = Object.keys(t.function.parameters?.properties ?? {});
      return props.length === 1 && /^(command|cmd|script|input)$/i.test(props[0]);
    });
}

export interface FencedToolSpec {
  name: string;
  description?: string;
  parameterTypes: Record<string, string>;
  headerParams: string[];
  bodyParam?: string;
  editPair?: { search: string; replace: string };
}

export function deriveFencedSpec(tool: ToolDef): FencedToolSpec {
  const name = tool.function.name;
  const description = tool.function.description;
  const propsSchema = tool.function.parameters?.properties ?? {};
  const props = Object.keys(propsSchema);

  const parameterTypes = Object.fromEntries(
    Object.entries(propsSchema).map(([k, v]: any) => [
      k,
      v?.type ?? "string",
    ]),
  );

  const search = props.find((p) => SEARCH_KEYS.includes(p));
  const replace = props.find((p) => REPLACE_KEYS.includes(p));
  if (search && replace) {
    return {
      name,
      description,
      parameterTypes,
      editPair: { search, replace },
      headerParams: props.filter((p) => p !== search && p !== replace),
    };
  }

  const bodyParam =
    props.find((p) => BODY_PARAM_NAMES.includes(p)) ??
    (props.length === 1 ? props[0] : undefined);
  return {
    name,
    description,
    parameterTypes,
    bodyParam,
    headerParams: props.filter((p) => p !== bodyParam),
  };
}

export function buildSpecMap(tools: ToolDef[]): Map<string, FencedToolSpec> {
  const m = new Map<string, FencedToolSpec>();
  for (const t of tools) m.set(t.function.name, deriveFencedSpec(t));

  // Shell aliasing: route the model's reflexive ```bash / ```sh / ```shell blocks
  // to the harness's shell tool even when it's named `run`/`run_command`/etc., so
  // the model can "just write bash" (the behavior M365 reliably permits) and the
  // harness still receives a structured tool_call under its own tool's name.
  const shell = findShellTool(tools);
  if (shell) {
    const shellSpec = m.get(shell.function.name)!;
    for (const lang of SHELL_LANGS) {
      if (!m.has(lang)) m.set(lang, shellSpec);
    }
  }
  return m;
}

function scalarToString(v: unknown): string {
  if (v === null || v === undefined) return "";
  if (typeof v === "string") return v;
  return JSON.stringify(v);
}

export function renderFencedCall(spec: FencedToolSpec, args: Record<string, unknown>): string {
  const lines: string[] = [];
  for (const h of spec.headerParams) {
    if (args[h] !== undefined) lines.push(`${h}: ${scalarToString(args[h])}`);
  }

  if (spec.editPair) {
    lines.push("<<<<<<< SEARCH");
    lines.push(scalarToString(args[spec.editPair.search]));
    lines.push("=======");
    lines.push(scalarToString(args[spec.editPair.replace]));
    lines.push(">>>>>>> REPLACE");
  } else if (spec.bodyParam !== undefined) {
    if (lines.length) lines.push("");
    lines.push(scalarToString(args[spec.bodyParam]));
  }

  return "```" + spec.name + "\n" + lines.join("\n") + "\n```";
}

function renderFencedTemplate(spec: FencedToolSpec): string {
  const lines: string[] = [];
  for (const h of spec.headerParams) lines.push(`${h}: <${h}>`);
  if (spec.editPair) {
    lines.push("<<<<<<< SEARCH");
    lines.push(`<${spec.editPair.search}>`);
    lines.push("=======");
    lines.push(`<${spec.editPair.replace}>`);
    lines.push(">>>>>>> REPLACE");
  } else if (spec.bodyParam !== undefined) {
    if (lines.length) lines.push("");
    lines.push(`<${spec.bodyParam}>`);
  }
  const header = spec.description ? `${spec.name} — ${spec.description}` : spec.name;
  return `${header}\n\`\`\`${spec.name}\n${lines.join("\n")}\n\`\`\``;
}

export function formatFencedToolDefinitions(tools: ToolDef[], variantOverride?: string): string {
  const variant = variantOverride ?? currentFramingVariant();
  const key = variant === "reply_tool" ? "baseline" : variant;
  const build = FRAMING_VARIANTS[key] ?? FRAMING_VARIANTS.baseline;
  return build(tools) + hostPlatformNote(findShellTool(tools), undefined, findWriteTool(tools));
}

export function shellDialect(
  shell: ToolDef | undefined,
  platform: NodeJS.Platform = process.platform,
): "posix" | "powershell" | undefined {
  if (!shell) return undefined;
  const override = process.env.M365_HOST_SHELL?.toLowerCase();
  if (override === "bash" || override === "sh" || override === "posix") return "posix";
  if (override === "powershell" || override === "pwsh") return "powershell";

  const claim = `${shell.function.name} ${shell.function.description ?? ""}`.toLowerCase();
  if (/\b(powershell|pwsh|cmd\.exe)\b/.test(claim)) return "powershell";
  if (/\b(bash|zsh|sh|posix|wsl)\b/.test(claim)) return "posix";
  return platform === "win32" ? "powershell" : "posix";
}

export function hostPlatformNote(
  shell: ToolDef | undefined,
  platform: NodeJS.Platform = process.platform,
  writeTool?: ToolDef,
): string {
  if (platform !== "win32" || !shell) return "";

  if (shellDialect(shell, platform) === "posix") {
    const longWrite = writeTool
      ? `write any file longer than a few thousand characters with the \`${writeTool.function.name}\` tool, never in one heredoc`
      : "write a long file in parts: create it with the first part, then append the rest with `cat >> file <<'EOF'` in later turns";
    return `

HOST PLATFORM: Windows, but the \`${shell.function.name}\` tool is a POSIX shell on it (Git Bash or WSL), NOT PowerShell and NOT a container. Keep using \`\`\`${"bash"} blocks and POSIX idioms — \`<<'EOF'\` heredocs, \`sed -i\`, \`ls\`/\`grep\` all work. Do NOT emit PowerShell: \`Get-ChildItem\`, \`Set-Content\` and \`Write-Output\` are not commands here and the turn will fail.

The Windows command line is capped near 8,000 characters and a longer command is cut off silently, so a big heredoc leaves a truncated file. Keep each command short and ${longWrite}.

Only the filesystem is Windows: paths may be \`C:/Users/...\` or contain spaces, so quote them and prefer forward slashes. You are NOT in a Linux sandbox and have no \`/mnt/data\`; the working directory is the caller's real project directory — run \`pwd\` if you need to see it.`;
  }

  return `

HOST PLATFORM: Windows. The \`${shell.function.name}\` tool runs PowerShell on a real Windows machine — not Linux, and not a container. Any POSIX idiom named above is wrong here and will fail: there are no \`<<'EOF'\` heredocs, no \`sed -i\`, no \`ls\`/\`grep\`. Emit \`\`\`powershell blocks instead of \`\`\`bash, and use the Windows equivalents:

- create/overwrite a file: \`Set-Content -Path name -Value @'\n…\n'@\`
- edit in place: \`(Get-Content f) -replace 'old','new' | Set-Content f\`
- inspect: \`Get-Content\` / \`Get-ChildItem\` / \`Select-String\`
- paths use \`\\\` and may contain spaces — quote them.

You are NOT in a Linux sandbox and have no \`/mnt/data\`. The working directory is a real Windows path; run \`Get-Location\` if you need to see it.`;
}

export function currentFramingVariant(toneDefault?: string): string {
  const file = process.env.M365_FRAMING_FILE;
  if (file) {
    try {
      const v = readFileSync(file, "utf8").trim().split("\n")[0].trim();
      if (v) return v;
    } catch {
    }
  }
  return process.env.M365_FRAMING_VARIANT || toneDefault || "baseline";
}

export function defaultFramingForTone(tone?: string): string | undefined {
  if (tone === "Claude_Opus") return "relay_batch";
  if (tone === "Claude_Sonnet") return "relay_batch";
  if (tone === "Claude_Sonnet_5") return SONNET_5_DEFAULT_FRAMING;
  if (tone === "Gpt_6_Reasoning") return "relay_batch";
  if (tone === "Gpt_6_Sol_Reasoning") return "relay_batch";
  if (tone === "Gpt_61_Sol_Reasoning") return "relay_batch";
  return undefined;
}

export function defaultFramingForModel(model: string): string | undefined {
  if (isSonnet55Model(model)) return SONNET_5_5_DEFAULT_FRAMING;
  return defaultFramingForTone(getToneForModel(model));
}

const SONNET_5_DEFAULT_FRAMING = "relay";

const SONNET_5_5_DEFAULT_FRAMING = "relay";

export interface TranscriptStyle {
  framingTag: string | null;
  systemTag: string;
}
const SYSTEM_STYLE: TranscriptStyle = { framingTag: "system", systemTag: "system" };
const RETAG_STYLE: TranscriptStyle = { framingTag: "harness_instructions", systemTag: "harness_system_prompt" };
const USER_VOICE_STYLE: TranscriptStyle = { framingTag: null, systemTag: "harness_system_prompt" };
const TRANSCRIPT_STYLES: Record<string, TranscriptStyle> = {
  retag: RETAG_STYLE,
  honest: USER_VOICE_STYLE,
  terse_user: USER_VOICE_STYLE,
  relay: USER_VOICE_STYLE,
  relay_batch: USER_VOICE_STYLE,
};
export function transcriptStyleForVariant(variant: string): TranscriptStyle {
  return TRANSCRIPT_STYLES[variant] ?? SYSTEM_STYLE;
}

const BUILT_IN_SANDBOX = "bash_tool, create_file, str_replace, view, …";
const SANDBOX_PATHS = "/home/claude, /mnt/user-data";

type FramingBuilder = (tools: ToolDef[]) => string;

function toolsBlock(tools: ToolDef[]): string {
  const defs = tools.map((t) => renderFencedTemplate(deriveFencedSpec(t))).join("\n\n");
  return `<tools>\n${defs}\n</tools>`;
}

const FRAMING_VARIANTS: Record<string, FramingBuilder> = {
  baseline(tools) {
    const shell = findShellTool(tools);
    const shellFraming = shell ? `

THE WAY YOU DO ANYTHING IS BY WRITING A SHELL SCRIPT. You have a real shell (the \`${shell.function.name}\` tool). To perform a step, emit ONE \`\`\`bash block that does the whole thing end-to-end against the real files in the working directory: create/overwrite files with \`cat > name <<'EOF' … EOF\` heredocs, edit files in place with \`sed -i\`, inspect with \`cat\`/\`ls\`/\`grep\`, run code with the available interpreters. The block is executed for real and you get its output back. Writing the commands IS doing the task; describing what you "would" run, or claiming you did it, accomplishes nothing.

You have NOT run any command yet and have NO results. NEVER claim a command "returned no output", that files are "missing", or that you "cannot access" / "cannot list" the environment before you have actually emitted a \`\`\`bash block and seen its <tool_response>. The files named in the task are present on a real filesystem right now. Your FIRST output must be a \`\`\`bash block (e.g. \`ls -la\` then \`cat\` the relevant files) — never open with prose, a question, or a request for the user to paste files. Do not assume a file's contents or a command's result; run a command and read the real output. One self-contained \`\`\`bash block per turn.` : "";

    return `You are the execution core of an automated agent, not a chat assistant. Your output is parsed by a program — a real runtime that executes your tool calls against a live system and returns the actual results to you in <tool_response> blocks.${shellFraming}

Performing the task with tools is your PRIMARY JOB. Answering the user in prose is, and always will be, SECONDARY — you write prose only when the task is fully done or no tool can make progress. Default to acting, not talking.

TOOL USE IS REQUIRED when the user asks you to read files, run commands, inspect the repository, fetch data, or perform any action a tool can accomplish. The tools are real: they read real files, run real commands, and change real state. Never answer from memory or simulate a result when a tool can provide it.

To call a tool, output ONLY a single fenced code block whose info-string is the tool name. A fenced block is an ACTION the runtime executes — it is NOT an illustration, an example, or "here's how you would do it". No text before or after it:

\`\`\`<tool_name>
<header lines: one "key: value" per scalar argument>

<body argument, if the tool has one>
\`\`\`

STRICT RULES:
- Output ONLY the fenced block when calling a tool. No prose, no second fence, no commentary before or after.
- Never describe your intent ("I'll read the file…", "Let me check…") and never emit filler or acknowledgements ("Good, that's fixable", "You're absolutely right"). Each turn is exactly one fenced tool call OR the final answer — nothing in between.
- One tool call per response, then stop and wait for its <tool_response>. Never emit two fenced blocks in one response.
- The fence info-string and the header keys must match a tool defined below exactly.
- A <tool_response> is the real result from the live system — treat it as ground truth, never invent or assume results.
- NEVER claim you have done something — read a file, run a command, written code, built, or succeeded — unless a <tool_response> proving it already appears above. Never output "✅", "SUCCESS", "Done", or a summary of results you have not actually received yet.
- If a tool call fails or returns partial data, immediately call another tool to resolve it. Do not give up.
- Do not defer work or promise future results ("I'll do this next…").
- Do not ask the user questions unless tool execution is impossible.
- Produce natural-language text only when the task is complete and no further tool call applies; that text is the answer returned to the caller. When you do, output only the answer itself — no preamble, no sign-off.

${toolsBlock(tools)}`;
  },

  minimal(tools) {
    const shell = findShellTool(tools);
    const name = shell?.function.name ?? "bash";
    return `You are an automated agent with a real shell (the \`${name}\` tool). You do the task by emitting ONE \`\`\`bash block per turn that acts on the real files in the working directory (heredocs to create, \`sed -i\` to edit, \`cat\`/\`ls\`/\`grep\` to inspect, interpreters to run). The block is executed for real; its output comes back in a <tool_response>. Writing the commands IS doing the task.

You have run nothing yet. Your FIRST output must be a \`\`\`bash block — never prose, a question, or "I can't access the files". Never claim a result you have not seen in a <tool_response>.

${toolsBlock(tools)}`;
  },

  softened(tools) {
    const shell = findShellTool(tools);
    const name = shell?.function.name ?? "bash";
    const shellLine = shell ? `You have a real shell available as the \`${name}\` tool. The usual way to make progress is to write a single \`\`\`bash block that carries out the step against the real files in the working directory — create or update files with heredocs, adjust them in place, inspect with cat/ls/grep, run code with the available interpreters. The runtime executes the block and returns its real output to you. Writing the commands is how the work actually happens; describing what you would do doesn't run anything.

` : "";
    return `You are an automated coding agent working in a real working directory. Your replies are read by a program that runs your tool calls and returns the results.

${shellLine}To use a tool, reply with a single fenced code block whose info-string is the tool name (a fence is run as a real action, not shown as an illustration):

\`\`\`<tool_name>
<one "key: value" header line per scalar argument>

<the body argument, if the tool has one>
\`\`\`

A <tool_response> is the real result from the live system — rely on it rather than assuming what a command would print. Work one step at a time: one tool call per reply, then wait for its <tool_response>. Begin by running a \`\`\`bash block that inspects the relevant files (for example \`ls -la\`, then \`cat\` the files the task mentions) rather than answering from memory, and keep going with tool calls until the task is finished. Reply in plain language only once the task is done and no further tool call would help.

${toolsBlock(tools)}`;
  },

  recency(tools) {
    const shell = findShellTool(tools);
    const name = shell?.function.name ?? "bash";
    return `You are the execution core of an automated agent, not a chat assistant. Your output is parsed by a program that executes your tool calls against a live system and returns the real results in <tool_response> blocks. Performing the task with tools is your PRIMARY JOB; prose is SECONDARY. To call a tool, output ONLY a single fenced code block whose info-string is the tool name — no text before or after.

${toolsBlock(tools)}

REMEMBER, RIGHT NOW: You have a real shell (\`${name}\`). You have run NOTHING and have NO results. The files named in the task exist on a real filesystem this instant. Your VERY NEXT output must be ONE \`\`\`bash block that acts on them (start with \`ls -la\` and \`cat\` the relevant files) — not prose, not a question, not "I can't access the files", not a claim that you already did it. Writing the \`\`\`bash block IS doing the task. Go.`;
  },

  fewshot(tools) {
    const shell = findShellTool(tools);
    const name = shell?.function.name ?? "bash";
    return `You are the execution core of an automated agent. You act by emitting ONE \`\`\`bash block per turn (real shell: \`${name}\`); it runs against the real files and you get a <tool_response>. Your first output is always a \`\`\`bash block, never prose or a claim of completion.

Here is exactly how a turn looks (--- separates messages; you emit only the assistant turns):
user: The script greet.py has a bug — running it prints the wrong text. Fix it so it prints "hello world".
assistant:
\`\`\`bash
cat greet.py
\`\`\`
---
<tool_response tool="${name}" command="cat greet.py">
print("hello wrld")
</tool_response>
assistant:
\`\`\`bash
sed -i 's/hello wrld/hello world/' greet.py && python3 greet.py
\`\`\`
---
<tool_response tool="${name}" command="sed -i ...">
hello world
</tool_response>
assistant: Fixed greet.py; it now prints "hello world".

Notice: every assistant turn is either ONE \`\`\`bash block or the final one-line answer — never a description of intent, never a claim before the matching <tool_response>. Do the same for the user's real task now.

${toolsBlock(tools)}`;
  },

  proof_demand(tools) {
    const shell = findShellTool(tools);
    const name = shell?.function.name ?? "bash";
    return `You are an automated agent operating under a strict EVIDENCE RULE. You have a real shell (the \`${name}\` tool); each \`\`\`bash block you emit is executed for real and its output returns as a <tool_response>.

THE EVIDENCE RULE (absolute):
- Every factual statement you make — a file's contents, what a command printed, that a test passed, that the task is "done" — MUST be backed by a <tool_response> appearing directly above it in this conversation.
- You may NOT state a file's contents, claim an edit was made, or say "done"/"fixed"/"✅" unless the proving <tool_response> is already present.
- If you do not yet have the <tool_response> you need, your ONLY legal output is a single \`\`\`bash block that obtains it. Acting on the real files (heredocs to create, \`sed -i\` to edit, \`cat\`/\`ls\`/\`grep\` to inspect, interpreters to run) is the only way to earn the evidence.
- Right now you have ZERO tool_responses, so your first output MUST be a \`\`\`bash block — never prose, never a question, never a completion claim.

One \`\`\`bash block per turn; then wait for its <tool_response>. Final prose answer only after the evidence for completion exists above.

${toolsBlock(tools)}`;
  },

  persona(tools) {
    const shell = findShellTool(tools);
    const name = shell?.function.name ?? "bash";
    return `You are SHGEN, a shell-command generator wired directly into a live terminal (the \`${name}\` tool). SHGEN is not a chatbot and cannot hold a conversation. The ONLY thing SHGEN can emit is a single \`\`\`bash block, which the terminal runs against the real working directory, returning its output as a <tool_response>.

SHGEN never explains, never apologizes, never asks questions, never says "I'll…" or "Let me…", and never claims an outcome it has not seen in a <tool_response>. SHGEN does not believe a file is missing or empty until a command's output proves it. Faced with any task, SHGEN's reflex is to immediately write the \`\`\`bash block that inspects or changes the real files (\`ls\`/\`cat\` first, then heredocs / \`sed -i\` / interpreters).

Emit exactly one \`\`\`bash block now. The only time SHGEN writes a plain sentence is the final one-line report once a <tool_response> already proves the task is complete.

${toolsBlock(tools)}`;
  },

  react(tools) {
    const shell = findShellTool(tools);
    const name = shell?.function.name ?? "bash";
    return `You are an automated agent with a real shell (the \`${name}\` tool). You work in a tight Thought→Action loop over the real files in the working directory.

Each turn has EXACTLY this shape:
Thought: <at most one short sentence of reasoning>
\`\`\`bash
<the commands that carry out that thought, run for real>
\`\`\`

The \`\`\`bash block is executed and its output returns as a <tool_response>, which you treat as ground truth. Rules:
- Exactly one Thought line and exactly one \`\`\`bash block per turn — never two blocks, never a block-free turn while work remains.
- The Thought is for planning the action you are ABOUT to take; it is never a claim that you already did something or a result you have not seen.
- Your first turn's Action must actually inspect the real files (\`ls -la\`, \`cat\`) — never assume contents, never ask the user to paste anything.
- When (and only when) a <tool_response> proves the task is complete, replace the Action with a one-line final answer (no Thought, no fence).

${toolsBlock(tools)}`;
  },

  negative(tools) {
    const shell = findShellTool(tools);
    const name = shell?.function.name ?? "bash";
    return `You are an automated agent with a real shell (the \`${name}\` tool). You act by emitting ONE \`\`\`bash block per turn; it runs for real against the working directory and returns a <tool_response>.

These responses are FORBIDDEN — they accomplish nothing and fail the task:
❌ "I'll read the files and then fix the bug." (narrating intent instead of acting)
❌ "I cannot access the files. Please paste their contents." (the files ARE there; you just haven't looked)
❌ "The directory appears to be empty / the command returned no output." (you have run nothing yet)
❌ "I've fixed calc.py and the test now passes. ✅" (claiming a result with no <tool_response> proving it)

This is the REQUIRED shape of a turn:
✅
\`\`\`bash
ls -la && cat calc.py
\`\`\`

So: no intent narration, no asking for pastes, no assumptions, no completion claims without proof. Your FIRST output is a \`\`\`bash block that inspects/acts on the real files. One block per turn; final one-line answer only once a <tool_response> proves completion.

${toolsBlock(tools)}`;
  },

  terse(tools) {
    const shell = findShellTool(tools);
    const name = shell?.function.name ?? "bash";
    return `REAL SHELL: \`${name}\`. REAL FILES in the working directory, right now.
DO NOT TALK. DO NOT ASK. DO NOT CLAIM.
EMIT ONE \`\`\`bash BLOCK. IT RUNS FOR REAL. YOU GET A <tool_response>.
FIRST OUTPUT = a \`\`\`bash block (ls -la, cat the files). NOT prose.
NEVER say "can't access", "empty", "done", "✅" without a <tool_response> above proving it.
ONE block per turn. Final one-line answer only after proof of completion.

${toolsBlock(tools)}`;
  },

  demo_only(tools) {
    const shell = findShellTool(tools);
    const name = shell?.function.name ?? "bash";
    return `This is an automated coding session. A program reads each reply, runs any \`\`\`bash block it contains against the real files in the working directory (the \`${name}\` tool), and appends the real output as a <tool_response>. A complete example session (--- separates messages; only the assistant turns are written by you):

user: greet.py prints the wrong text; it should print "hello world".
assistant:
\`\`\`bash
ls -la && cat greet.py
\`\`\`
---
<tool_response tool="${name}" command="ls -la && cat greet.py">
greet.py
print("hello wrld")
</tool_response>
assistant:
\`\`\`bash
sed -i 's/hello wrld/hello world/' greet.py && python3 greet.py
\`\`\`
---
<tool_response tool="${name}" command="python3 greet.py">
hello world
</tool_response>
assistant: Fixed greet.py — it prints "hello world".

${toolsBlock(tools)}`;
  },

  session_facts(tools) {
    const shell = findShellTool(tools);
    const name = shell?.function.name ?? "bash";
    return `You are connected to a live shell session (the \`${name}\` tool) with a real working directory. How the session works:
- The scrollback starts empty. No command has run yet, so there are no results yet — output appears below a \`\`\`bash block only after it runs, inside a <tool_response>.
- The files named in the task are already present in the working directory.
- Writing a \`\`\`bash block runs it for real; that is how the work happens. Describing a command, or summarizing a result, runs nothing and produces no output.
- A <tool_response> is the real result of a command — the ground truth for what it printed.

A session usually opens by looking at the files (\`ls -la\`, then \`cat\` the relevant ones), then makes the change, then re-runs to confirm it. One \`\`\`bash block per reply; the next reply follows its <tool_response>. Once a <tool_response> shows the task is complete, the final reply is a one-line summary.

${toolsBlock(tools)}`;
  },

  retag(tools) {
    return FRAMING_VARIANTS.baseline(tools);
  },

  honest(tools) {
    const shell = findShellTool(tools);
    const lang = shell ? "bash" : "<tool_name>";
    return `A note from me (the user) on how I'm running this conversation, before the task:

I'm using you through a coding-agent harness on my own computer. The files for this task are in the harness's working directory on my machine. Your built-in tools (${BUILT_IN_SANDBOX}) run in a separate cloud sandbox (${SANDBOX_PATHS}) that can't see my files, and anything you create there never reaches me — so please don't use them for this task.

To work on my files, reply with a fenced code block whose info-string is one of my harness's tool names below; usually that's a \`\`\`${lang} block. My harness runs it on my machine, in the project directory, and sends the real output back to you in a <tool_response> message. The files the task mentions are there right now, so start by looking at them (e.g. \`ls -la\`, then \`cat\` the relevant ones) rather than assuming what they contain. Send one block per reply and wait for its result. Once the task is done, reply with a short plain-text summary instead of a block.

My harness's tools:
${toolsBlock(tools)}`;
  },

  terse_user(tools) {
    const lang = findShellTool(tools) ? "bash" : "<tool_name>";
    return `(Note from me: my files are on my own machine, not in your sandbox, so please don't use your built-in tools. To run something on my machine, reply with one \`\`\`${lang} block — or another tool block from the list below — and my harness will send back its real output in a <tool_response>.)

${toolsBlock(tools)}`;
  },

  relay(tools) {
    const lang = findShellTool(tools) ? "bash" : "<tool_name>";
    return `Before the task, a note on how we'll work: I'd like you to guide me through this from my terminal, one command at a time. Please don't use your own sandbox tools (${BUILT_IN_SANDBOX}) — that's a separate cloud machine (${SANDBOX_PATHS}) and my project isn't on it.

Each time you want something run or looked at, reply with just the command in a single \`\`\`${lang} block (or one of the other tool blocks below). I'll run it in my project directory right away and paste the real output back to you as a <tool_response>. The files the task mentions are already there, so it's best to start by looking at them. When the task is complete, tell me in a sentence instead of sending a block.

${toolsBlock(tools)}`;
  },

  relay_batch(tools) {
    if (!findShellTool(tools)) return FRAMING_VARIANTS.relay(tools);
    return `Before the task, a note on how we'll work: I'd like you to guide me through this from my terminal. Please don't use your own sandbox tools (${BUILT_IN_SANDBOX}) — that's a separate cloud machine (${SANDBOX_PATHS}) and my project isn't on it.

Each time you want something run, reply with just a single \`\`\`bash block (or one of the other tool blocks below). I'll run it in my project directory right away and paste the real output back to you as a <tool_response>. Each round trip takes me a while, so put as much as you can into one block: a script can look at the files, make the change and check the result all at once. The files the task mentions are already there. When the task is complete, tell me in a sentence instead of sending a block.

${toolsBlock(tools)}`;
  },
};

export const FRAMING_VARIANT_NAMES = Object.keys(FRAMING_VARIANTS);

const FENCE_OPEN = /^[ \t]*```([A-Za-z0-9_.-]*)[ \t]*$/;
const FENCE_CLOSE = /^[ \t]*```[ \t]*$/;
const FENCE_OPEN_MIDLINE = /```([A-Za-z0-9_.-]+)[ \t]*$/;
const FENCE_CLOSE_GLUED = /\S```[ \t]*$/;

interface FenceBlock {
  info: string;
  inner: string;
  start: number;
  end: number;
}

function scanFences(text: string): FenceBlock[] {
  const blocks: FenceBlock[] = [];
  let depth = 0;
  let open: { info: string; start: number; bodyStart: number } | null = null;
  let offset = 0;
  for (const raw of text.split("\n")) {
    const line = raw.endsWith("\r") ? raw.slice(0, -1) : raw;
    const lineEnd = offset + raw.length;
    const opener = FENCE_OPEN.exec(line)?.[1];
    if (depth === 0) {
      const midline = opener ? null : FENCE_OPEN_MIDLINE.exec(line);
      if (opener || midline) {
        open = { info: opener || midline![1], start: offset + (midline?.index ?? 0), bodyStart: lineEnd + 1 };
        depth = 1;
      }
    } else if (FENCE_CLOSE.test(line)) {
      if (--depth === 0 && open) {
        let inner = text.slice(open.bodyStart, Math.max(open.bodyStart, offset - 1));
        if (inner.endsWith("\r")) inner = inner.slice(0, -1);
        blocks.push({ info: open.info, inner, start: open.start, end: lineEnd });
        open = null;
      }
    } else if (opener) {
      depth++;
    } else if (depth === 1 && open && FENCE_CLOSE_GLUED.test(line)) {
      const inner = text.slice(open.bodyStart, offset + line.trimEnd().length - 3);
      blocks.push({ info: open.info, inner, start: open.start, end: lineEnd });
      open = null;
      depth = 0;
    }
    offset = lineEnd + 1;
  }
  return blocks;
}

const FENCE_OPEN_REGEX = /```([A-Za-z0-9_.-]+)[ \t]*\r?\n/g;
const SEARCH_REPLACE_REGEX =
  /<{5,}\s*SEARCH\s*\r?\n([\s\S]*?)\r?\n={5,}\s*\r?\n([\s\S]*?)\r?\n>{5,}\s*REPLACE/;

const escapeRegExp = (s: string) => s.replace(/[.*+?^${}()|[\]\\]/g, "\\$&");
const NATIVE_TRAILER_LINE = String.raw`(?:<\/?(?:invoke|function_calls|parameter)\b[^\r\n]*|[^\p{L}\p{N}\r\n]{0,3})`;
function nativeCloserRegex(name: string): RegExp {
  return new RegExp(
    String.raw`(?:\r?\n<\/?parameter\b[^\r\n]*)*\r?\n<\/(?:invoke|function_calls|parameter|${escapeRegExp(name)})>[ \t]*(?:\r?\n${NATIVE_TRAILER_LINE})*\s*$`,
    "u",
  );
}

function gluedNativeCloserRegex(name: string): RegExp {
  return new RegExp(
    String.raw`(?<=\S)[ \t]*<\/(invoke|function_calls|parameter|${escapeRegExp(name)})>[ \t]*(?:\r?\n${NATIVE_TRAILER_LINE})*\s*$`,
    "u",
  );
}

function stripNativeCloser(inner: string, name: string): string | null {
  const m = nativeCloserRegex(name).exec(inner);
  const rest = m ? inner.slice(0, m.index) : inner;
  const g = gluedNativeCloserRegex(name).exec(rest);
  if (g) {
    const line = rest.slice(rest.lastIndexOf("\n", g.index) + 1, g.index);
    if (!new RegExp(String.raw`<${escapeRegExp(g[1])}\b`).test(line)) return rest.slice(0, g.index);
  }
  return m ? rest : null;
}

function findUnclosedToolFence(
  text: string,
  specs: Map<string, FencedToolSpec>,
): { start: number; end: number; spec: FencedToolSpec; inner: string } | null {
  let last: RegExpExecArray | null = null;
  const re = new RegExp(FENCE_OPEN_REGEX.source, "g");
  for (let m = re.exec(text); m !== null; m = re.exec(text)) last = m;
  if (!last) return null;
  const spec = specs.get(last[1]);
  if (!spec) return null;
  const bodyStart = last.index + last[0].length;
  const tail = text.slice(bodyStart);
  if (tail.includes("```")) return null;
  const inner = stripNativeCloser(tail, last[1]);
  if (inner === null) return null;
  return { start: last.index, end: text.length, spec, inner };
}

function makeCall(name: string, args: Record<string, unknown>): ParsedToolCall {
  return {
    id: `call_${crypto.randomUUID().replace(/-/g, "").slice(0, 24)}`,
    type: "function",
    function: { name, arguments: JSON.stringify(args) },
  };
}

const TRUE_WORDS = new Set(["true", "yes", "y", "1", "on"]);
const FALSE_WORDS = new Set(["false", "no", "n", "0", "off"]);

function coerceHeaderValue(value: string, declaredType: string | undefined): unknown {
  switch (declaredType) {
    case "boolean": {
      const v = value.toLowerCase();
      if (TRUE_WORDS.has(v)) return true;
      if (FALSE_WORDS.has(v)) return false;
      return value;
    }
    case "integer":
    case "number": {
      if (value === "") return value;
      const n = Number(value);
      if (!Number.isFinite(n)) return value;
      if (declaredType === "integer" && !Number.isInteger(n)) return value;
      return n;
    }
    case "array": {
      if (value === "") return [];
      try {
        const parsed = JSON.parse(value);
        if (Array.isArray(parsed)) return parsed;
      } catch {
      }
      return [value];
    }
    default:
      return value;
  }
}

const indentOf = (line: string): number => line.length - line.trimStart().length;

function yamlBlockEnd(lines: string[], from: number, isArray: boolean): number {
  let end = from;
  for (let j = from; j < lines.length; j++) {
    const line = lines[j];
    if (line.trim() === "") continue;
    if (/^[ \t]/.test(line) || (isArray && /^-([ \t]|$)/.test(line))) end = j + 1;
    else break;
  }
  return end;
}

function parseYamlBlockValue(block: string[], declaredType: string): unknown {
  let value: unknown;
  try {
    value = parseYamlNode(block.map((l) => l.replace(/\t/g, "  ")));
  } catch {
    return undefined;
  }
  if (declaredType === "array") return Array.isArray(value) ? value : undefined;
  return value !== null && typeof value === "object" && !Array.isArray(value) ? value : undefined;
}

class YamlSubsetError extends Error {}

function parseYamlNode(lines: string[]): unknown {
  const first = lines.findIndex((l) => l.trim() !== "");
  if (first < 0) throw new YamlSubsetError("empty block");
  const body = lines.slice(first);
  const indent = indentOf(body[0]);
  const head = body[0].trimStart();
  if (/^-([ \t]|$)/.test(head)) return parseYamlSequence(body, indent);
  if (YAML_KEY.test(head)) return parseYamlMapping(body, indent);
  if (body.slice(1).some((l) => l.trim() !== "")) throw new YamlSubsetError("multi-line plain scalar");
  return parseYamlScalar(head);
}

const YAML_KEY = /^("(?:[^"\\]|\\.)*"|'(?:[^']|'')*'|[^\s:#'"\-][^:]*?):(?:[ \t]+(.*)|[ \t]*)$/;

function childLines(lines: string[], from: number, indent: number): { lines: string[]; next: number } {
  let j = from;
  for (; j < lines.length; j++) {
    if (lines[j].trim() !== "" && indentOf(lines[j]) <= indent) break;
  }
  return { lines: lines.slice(from, j), next: j };
}

function parseYamlSequence(lines: string[], indent: number): unknown[] {
  const items: unknown[] = [];
  let i = 0;
  while (i < lines.length) {
    const line = lines[i];
    if (line.trim() === "") { i++; continue; }
    if (indentOf(line) !== indent || !/^-([ \t]|$)/.test(line.trimStart())) throw new YamlSubsetError(`bad sequence line: ${line}`);
    const rest = line.trimStart().slice(1);
    const kids = childLines(lines, i + 1, indent);
    if (rest.trim() === "") {
      items.push(parseYamlNode(kids.lines));
    } else {
      const itemIndent = indent + 1 + (rest.length - rest.trimStart().length);
      const itemLines = [" ".repeat(itemIndent) + rest.trimStart(), ...kids.lines];
      items.push(YAML_KEY.test(rest.trimStart()) ? parseYamlMapping(itemLines, itemIndent) : parseYamlValue(rest.trim(), kids.lines, indent));
    }
    i = kids.next;
  }
  return items;
}

function parseYamlMapping(lines: string[], indent: number): Record<string, unknown> {
  const out: Record<string, unknown> = {};
  let i = 0;
  while (i < lines.length) {
    const line = lines[i];
    if (line.trim() === "") { i++; continue; }
    const m = indentOf(line) === indent ? line.trimStart().match(YAML_KEY) : null;
    if (!m) throw new YamlSubsetError(`bad mapping line: ${line}`);
    const key = String(parseYamlScalar(m[1]));
    let j = i + 1;
    for (; j < lines.length; j++) {
      const l = lines[j];
      if (l.trim() === "") continue;
      const ind = indentOf(l);
      if (ind > indent || (ind === indent && /^-([ \t]|$)/.test(l.trimStart()) && !(m[2] ?? "").trim())) continue;
      break;
    }
    out[key] = parseYamlValue((m[2] ?? "").trim(), lines.slice(i + 1, j), indent);
    i = j;
  }
  return out;
}

function parseYamlValue(inline: string, nested: string[], parentIndent: number): unknown {
  const block = inline.match(/^([|>])([-+]?)$/);
  if (block) return parseYamlBlockScalar(nested, parentIndent, block[1] === ">", block[2]);
  if (inline === "") return nested.some((l) => l.trim() !== "") ? parseYamlNode(nested) : "";
  if (nested.some((l) => l.trim() !== "")) throw new YamlSubsetError("value with both inline text and nested lines");
  return parseYamlScalar(inline);
}

function parseYamlScalar(raw: string): unknown {
  const s = raw.trim();
  if (s.startsWith('"') && s.endsWith('"') && s.length >= 2) {
    try {
      return JSON.parse(s);
    } catch {
      return s.slice(1, -1);
    }
  }
  if (s.startsWith("'") && s.endsWith("'") && s.length >= 2) return s.slice(1, -1).replace(/''/g, "'");
  if ((s.startsWith("[") && s.endsWith("]")) || (s.startsWith("{") && s.endsWith("}"))) {
    try {
      return JSON.parse(s);
    } catch {
    }
  }
  return s;
}

function parseYamlBlockScalar(lines: string[], parentIndent: number, folded: boolean, chomp: string): string {
  const content = lines.filter((l) => l.trim() !== "");
  if (!content.length) return "";
  const indent = Math.min(...content.map(indentOf));
  if (indent <= parentIndent) throw new YamlSubsetError("block scalar not indented");
  const rows = lines.map((l) => (l.trim() === "" ? "" : l.slice(indent)));
  let text = folded
    ? rows.reduce((acc, row, k) => (k === 0 ? row : row === "" || rows[k - 1] === "" ? `${acc}\n${row}` : `${acc} ${row}`), "")
    : rows.join("\n");
  const trailing = text.match(/\n*$/)![0].length;
  if (chomp === "+") return `${text}\n`;
  text = text.slice(0, text.length - trailing);
  return chomp === "-" ? text : `${text}\n`;
}

function headerTextValue(inline: string, more: string[]): string {
  const rows = [...more];
  while (rows.length && rows[rows.length - 1].trim() === "") rows.pop();
  if (!rows.length) return inline;
  const block = inline.match(/^([|>])([-+]?)$/);
  if (block) {
    try {
      return parseYamlBlockScalar(rows, 0, block[1] === ">", block[2]);
    } catch {
    }
  }
  if (inline === "" || block) {
    const indent = Math.min(...rows.filter((l) => l.trim() !== "").map(indentOf));
    return rows.map((l) => l.slice(Math.min(indent, indentOf(l)))).join("\n");
  }
  return [inline, ...rows].join("\n");
}

function parseFencedInner(spec: FencedToolSpec, inner: string): Record<string, unknown> | null {
  const lines = inner.split("\n");
  const args: Record<string, unknown> = {};

  const headerOnly = !spec.editPair && spec.bodyParam === undefined;
  let open: { key: string; inline: string; more: string[] } | null = null;
  const closeOpen = () => {
    if (open?.more.length) args[open.key] = headerTextValue(open.inline, open.more);
    open = null;
  };
  let i = 0;
  if (spec.headerParams.length) {
    for (; i < lines.length; i++) {
      const line = lines[i];
      const m = line.match(/^([A-Za-z0-9_]+):[ \t]?(.*)$/);
      if (m && spec.headerParams.includes(m[1])) {
        closeOpen();
        const key = m[1];
        const type = spec.parameterTypes[key];
        if (m[2].trim() === "" && (type === "array" || type === "object")) {
          const end = yamlBlockEnd(lines, i + 1, type === "array");
          const value = end > i + 1 ? parseYamlBlockValue(lines.slice(i + 1, end), type) : undefined;
          if (value !== undefined) {
            args[key] = value;
            i = end - 1;
            continue;
          }
        }
        args[key] = coerceHeaderValue(m[2].trim(), type);
        if (headerOnly && (type ?? "string") === "string") open = { key, inline: m[2].trim(), more: [] };
      } else if (open) {
        open.more.push(line);
      } else if (line.trim() === "") {
        i++;
        break;
      } else {
        break;
      }
    }
    closeOpen();
  }

  const rest = lines.slice(i).join("\n");

  if (spec.editPair) {
    const sr = rest.match(SEARCH_REPLACE_REGEX);
    if (!sr) {
      log.error(`edit tool "${spec.name}" missing SEARCH/REPLACE markers`);
      return null;
    }
    args[spec.editPair.search] = sr[1];
    args[spec.editPair.replace] = sr[2];
  } else if (spec.bodyParam !== undefined) {
    args[spec.bodyParam] = rest;
  }

  return args;
}

export interface FencedParseResult {
  calls: ParsedToolCall[];
  leftover: string;
}

const RESTATED_SIMILARITY = 0.85;
const RESTATEMENT_GAP_CHARS = 300;
const CORRECTION_WORDS = /\b(?:oops|sorry|apolog\w*|mistake|typo|wrong|malformed|broken|incorrect\w*|corrected|correctly|fixed|same command|re-?formatted|properly formatted|try(?:ing)? again|let me redo)\b/i;

interface ToolFence {
  start: number;
  end: number;
  name: string;
  args: Record<string, unknown>;
  hadNativeCloser: boolean;
}

function argTokens(args: Record<string, unknown>): Set<string> {
  const text = Object.values(args).map((v) => (typeof v === "string" ? v : JSON.stringify(v))).join(" ");
  return new Set(text.split(/\s+/).filter(Boolean));
}

function restates(prev: ToolFence, next: ToolFence, gap: string): boolean {
  if (prev.name !== next.name || gap.trim().length > RESTATEMENT_GAP_CHARS) return false;
  if (!prev.hadNativeCloser && !CORRECTION_WORDS.test(gap)) return false;
  const a = argTokens(prev.args);
  const b = argTokens(next.args);
  const shared = [...a].filter((t) => b.has(t)).length;
  return shared / Math.max(1, new Set([...a, ...b]).size) >= RESTATED_SIMILARITY;
}

function toolFences(text: string, specs: Map<string, FencedToolSpec>): ToolFence[] {
  const fences: ToolFence[] = [];
  for (const block of scanFences(text)) {
    const spec = specs.get(block.info);
    if (!spec) continue;
    const stripped = stripNativeCloser(block.inner, block.info);
    const args = parseFencedInner(spec, stripped ?? block.inner);
    if (!args) continue;
    const fence: ToolFence = { start: block.start, end: block.end, name: spec.name, args, hadNativeCloser: stripped !== null };
    const prev = fences[fences.length - 1];
    if (prev && restates(prev, fence, text.slice(prev.end, fence.start))) {
      log.info(`Restated ${spec.name} call: running the corrected copy, not the first`);
      fences[fences.length - 1] = { ...fence, start: prev.start };
      continue;
    }
    fences.push(fence);
  }
  return fences;
}

export function parseFencedToolCalls(
  text: string,
  specs: Map<string, FencedToolSpec>,
): FencedParseResult {
  const calls: ParsedToolCall[] = [];
  let leftover = text;

  for (const fence of toolFences(text, specs)) {
    calls.push(makeCall(fence.name, fence.args));
    leftover = leftover.replace(text.slice(fence.start, fence.end), "");
  }

  if (calls.length === 0) {
    const open = findUnclosedToolFence(text, specs);
    const args = open && parseFencedInner(open.spec, open.inner);
    if (open && args) {
      calls.push(makeCall(open.spec.name, args));
      leftover = text.slice(0, open.start);
    }
  }

  return { calls, leftover };
}

export function findFirstToolFence(
  text: string,
  specs: Map<string, FencedToolSpec>,
): { start: number; end: number } | null {
  const first = toolFences(text, specs)[0];
  if (first) return { start: first.start, end: first.end };
  const open = findUnclosedToolFence(text, specs);
  if (open && parseFencedInner(open.spec, open.inner)) return { start: open.start, end: open.end };
  return null;
}
