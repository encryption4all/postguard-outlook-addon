// Thin Microsoft Graph wrapper for the housekeeping operations PostGuard
// needs: creating folders, importing raw MIME messages, deleting received
// messages. Encryption itself never touches Graph.

import { getGraphToken } from "./auth";

const GRAPH_BASE = "https://graph.microsoft.com/v1.0";

async function graphFetch(path: string, init: RequestInit = {}): Promise<Response> {
  const token = await getGraphToken(false);
  if (!token) throw new Error("Graph token unavailable");
  const headers = new Headers(init.headers);
  headers.set("Authorization", `Bearer ${token}`);
  if (!headers.has("Accept")) headers.set("Accept", "application/json");
  return fetch(`${GRAPH_BASE}${path}`, { ...init, headers });
}

export async function isGraphAvailable(): Promise<boolean> {
  const t = await getGraphToken(false);
  return !!t;
}

export interface GraphFolder {
  id: string;
  displayName: string;
}

export async function findFolder(displayName: string): Promise<GraphFolder | null> {
  const url = `/me/mailFolders?$filter=displayName eq '${encodeURIComponent(displayName)}'&$top=1`;
  const res = await graphFetch(url);
  if (!res.ok) throw new Error(`findFolder failed: ${res.status}`);
  const body = (await res.json()) as { value: GraphFolder[] };
  return body.value[0] ?? null;
}

export async function createFolder(displayName: string): Promise<GraphFolder> {
  const res = await graphFetch(`/me/mailFolders`, {
    method: "POST",
    headers: { "Content-Type": "application/json" },
    body: JSON.stringify({ displayName }),
  });
  if (!res.ok) throw new Error(`createFolder failed: ${res.status}`);
  return (await res.json()) as GraphFolder;
}

export async function getOrCreateFolder(displayName: string): Promise<GraphFolder> {
  const existing = await findFolder(displayName);
  if (existing) return existing;
  return createFolder(displayName);
}

// Imports a raw RFC 5322 MIME message into a folder. Graph accepts the
// MIME body via Content-Type: text/plain on /me/mailFolders/{id}/messages.
export async function importMimeIntoFolder(
  folderId: string,
  mime: Uint8Array
): Promise<{ id: string }> {
  // Graph wants base64 of the MIME content for the messages endpoint.
  const b64 = uint8ArrayToBase64(mime);
  const res = await graphFetch(`/me/mailFolders/${folderId}/messages`, {
    method: "POST",
    headers: { "Content-Type": "text/plain" },
    body: b64,
  });
  if (!res.ok) throw new Error(`importMime failed: ${res.status}`);
  return (await res.json()) as { id: string };
}

export async function deleteMessage(itemId: string): Promise<void> {
  const res = await graphFetch(`/me/messages/${encodeURIComponent(itemId)}`, {
    method: "DELETE",
  });
  if (!res.ok && res.status !== 404) {
    throw new Error(`deleteMessage failed: ${res.status}`);
  }
}

export async function getRawMime(itemId: string): Promise<Uint8Array> {
  const res = await graphFetch(`/me/messages/${encodeURIComponent(itemId)}/$value`);
  if (!res.ok) throw new Error(`getRawMime failed: ${res.status}`);
  const buf = await res.arrayBuffer();
  return new Uint8Array(buf);
}

function uint8ArrayToBase64(bytes: Uint8Array): string {
  let bin = "";
  const chunk = 8192;
  for (let i = 0; i < bytes.length; i += chunk) {
    bin += String.fromCharCode(...bytes.subarray(i, i + chunk));
  }
  return btoa(bin);
}
