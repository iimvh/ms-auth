const { EventEmitter } = await import("node:events");

const RS = "\x1E";

export const recording = {
  url: "",
  options: undefined as unknown,
  sent: [] as string[],
  emit: (_event: string, _data: Buffer) => {},
};

export class RecordingWebSocket extends EventEmitter {
  static OPEN = 1;
  static CLOSED = 3;

  readyState = RecordingWebSocket.OPEN;

  constructor(url: string, options?: unknown) {
    super();
    recording.url = url;
    recording.options = options;
    recording.sent = [];
    recording.emit = (event, data) => this.emit(event, data);
  }

  send(data: string): void {
    recording.sent.push(data);
  }

  close(): void {
    this.readyState = RecordingWebSocket.CLOSED;
    setImmediate(() => this.emit("close"));
  }
}

export function handshake(): Buffer {
  return Buffer.from(JSON.stringify({ protocol: "json", version: 1 }) + RS);
}

export function botText(text: string): Buffer {
  return Buffer.from(JSON.stringify({ type: 1, target: "update", arguments: [{ messages: [{ author: "bot", text }] }] }) + RS);
}

export function replayFeed(frames: ReadonlyArray<{ data: Buffer }>): () => void {
  return () => {
    for (const frame of frames) recording.emit("message", frame.data);
  };
}
