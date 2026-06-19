import type { VercelRequest, VercelResponse } from "@vercel/node";
import { apiError, noStore } from "../_lib/http.js";
import { enforceRateLimit } from "../_lib/rateLimit.js";
import { loadOrMigrateSession, saveStoredSession } from "../_lib/sessionRepository.js";
import type { WorktimeSession } from "../../src/types.js";

function requestId(req: VercelRequest) {
  return Array.isArray(req.query.id) ? req.query.id[0] : req.query.id;
}

export default async function handler(req: VercelRequest, res: VercelResponse) {
  const id = requestId(req);
  if (!id) return noStore(res).status(400).json({ error: "缺少测量单 ID" });

  try {
    if (req.method === "GET") {
      if (!await enforceRateLimit(req, res, "read")) return;
      const session = await loadOrMigrateSession(id);
      if (!session) return noStore(res).status(404).json({ error: "测量单不存在" });
      return noStore(res).status(200).json({ session });
    }

    if (req.method === "PUT") {
      if (!await enforceRateLimit(req, res, "save")) return;
      const body = req.body as WorktimeSession | null;
      if (!body || body.id !== id) {
        return noStore(res).status(400).json({ error: "保存内容无效" });
      }
      const saved = await saveStoredSession(body);
      return noStore(res).status(200).json({ session: saved });
    }

    return noStore(res).status(405).json({ error: "接口方法不支持" });
  } catch (error) {
    return apiError(res, error, "测量单处理失败");
  }
}
