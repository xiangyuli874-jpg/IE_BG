import { getDeployStore, getStore } from "@netlify/blobs";
import { createSession, isValidSessionId, normalizeSession } from "../../../shared/sessionCore";
import type { WorktimeSession } from "../../../src/types";

const STORE_NAME = "worktime-sessions";
export { createSession, normalizeSession };

function sessionKey(id: string) {
  return `sessions/${id}.json`;
}

function currentDeployContext() {
  const globalWithNetlify = globalThis as typeof globalThis & {
    Netlify?: { context?: { deploy?: { context?: string } } };
  };
  return globalWithNetlify.Netlify?.context?.deploy?.context;
}

function getSessionStore() {
  const options = { name: STORE_NAME, consistency: "strong" as const };
  return currentDeployContext() === "production" ? getStore(options) : getDeployStore(options);
}

export async function saveSession(session: WorktimeSession) {
  const normalized = normalizeSession(session);
  await getSessionStore().setJSON(sessionKey(normalized.id), normalized, {
    metadata: {
      updatedAt: normalized.updatedAt,
      lineId: normalized.lineId,
      modelId: normalized.modelId,
      modelName: normalized.modelName ?? ""
    }
  });
  return normalized;
}

export async function loadSession(id: string) {
  if (!id || !isValidSessionId(id)) {
    return null;
  }
  const session = await getSessionStore().get(sessionKey(id), { type: "json" });
  return session ? normalizeSession(session as WorktimeSession) : null;
}
