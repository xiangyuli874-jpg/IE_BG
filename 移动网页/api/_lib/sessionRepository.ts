import { isValidSessionId, normalizeSession } from "../../shared/sessionCore.js";
import type { WorktimeSession } from "../../src/types.js";
import { getRedis } from "./redis.js";

const SESSION_KEY_PREFIX = "worktime:session:";

export interface SessionRepositoryDeps {
  load: (id: string) => Promise<WorktimeSession | null>;
  save: (session: WorktimeSession) => Promise<WorktimeSession>;
  loadLegacy: (id: string) => Promise<WorktimeSession | null>;
}

function sessionKey(id: string) {
  return `${SESSION_KEY_PREFIX}${id}`;
}

export async function loadStoredSession(id: string) {
  if (!isValidSessionId(id)) return null;
  const session = await getRedis().get<WorktimeSession>(sessionKey(id));
  return session ? normalizeSession(session) : null;
}

export async function saveStoredSession(session: WorktimeSession) {
  const normalized = normalizeSession(session);
  await getRedis().set(sessionKey(normalized.id), normalized);
  return normalized;
}

export async function loadLegacySession(id: string) {
  const baseUrl = process.env.NETLIFY_MIGRATION_BASE_URL;
  const accessCode = process.env.NETLIFY_MIGRATION_ACCESS_CODE;
  if (!baseUrl || !accessCode || !isValidSessionId(id)) return null;

  const response = await fetch(`${baseUrl.replace(/\/$/, "")}/api/sessions/${encodeURIComponent(id)}`, {
    headers: { "X-Access-Code": accessCode }
  });
  if (response.status === 404) return null;
  if (!response.ok) {
    throw new Error(`旧测量单读取失败（${response.status}）`);
  }

  const data = await response.json() as { session?: WorktimeSession };
  return data.session ? normalizeSession(data.session) : null;
}

export async function loadOrMigrateSession(
  id: string,
  deps: SessionRepositoryDeps = {
    load: loadStoredSession,
    save: saveStoredSession,
    loadLegacy: loadLegacySession
  }
) {
  if (!isValidSessionId(id)) return null;
  const stored = await deps.load(id);
  if (stored) return stored;

  const legacy = await deps.loadLegacy(id);
  if (!legacy) return null;
  return deps.save(legacy);
}
