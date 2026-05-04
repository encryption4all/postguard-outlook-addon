// Chunking protocol for Office.context.ui.messageChild / messageParent.
// The dialog API caps each message at ~32KB; chunking lets us pass
// arbitrary-sized payloads (encrypt requests with attachment bytes,
// encrypt results with multi-MB ciphertext) by splitting into chunks
// and reassembling on the other side.
//
// Each chunk is a JSON envelope: { type: "chunk", id, index, total, data }.
// The receiver collects chunks by id and parses the joined string back
// into the original payload once all are present.

const CHUNK_SIZE = 24 * 1024; // ~24K chars per chunk, leaves headroom for JSON envelope

export interface ChunkMessage {
  type: "chunk";
  id: string;
  index: number;
  total: number;
  data: string;
}

export function isChunkMessage(msg: { type?: unknown } | null | undefined): msg is ChunkMessage {
  return !!msg && (msg as { type?: unknown }).type === "chunk";
}

export function chunkPayload(payload: unknown): ChunkMessage[] {
  const json = JSON.stringify(payload);
  const id = `${Date.now().toString(36)}${Math.random().toString(36).slice(2, 8)}`;
  const total = Math.max(1, Math.ceil(json.length / CHUNK_SIZE));
  const chunks: ChunkMessage[] = [];
  for (let i = 0; i < total; i++) {
    chunks.push({
      type: "chunk",
      id,
      index: i,
      total,
      data: json.slice(i * CHUNK_SIZE, (i + 1) * CHUNK_SIZE),
    });
  }
  return chunks;
}

interface PendingPayload {
  chunks: (string | undefined)[];
  received: number;
  total: number;
}

export class ChunkAssembler {
  private buffers = new Map<string, PendingPayload>();

  // Returns the reassembled payload when the final chunk for an id
  // arrives; returns null while a payload is still in flight.
  ingest(msg: ChunkMessage): unknown | null {
    let buf = this.buffers.get(msg.id);
    if (!buf) {
      buf = { chunks: new Array(msg.total), received: 0, total: msg.total };
      this.buffers.set(msg.id, buf);
    }
    if (buf.chunks[msg.index] !== undefined) {
      return null; // duplicate chunk, ignore
    }
    buf.chunks[msg.index] = msg.data;
    buf.received++;
    if (buf.received < buf.total) return null;
    this.buffers.delete(msg.id);
    try {
      return JSON.parse(buf.chunks.join("")) as unknown;
    } catch {
      return null;
    }
  }
}
