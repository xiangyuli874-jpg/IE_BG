import type { VercelRequest, VercelResponse } from "@vercel/node";
import { createSession } from "../../shared/sessionCore.js";
import { apiError, noStore } from "../_lib/http.js";
import { enforceRateLimit } from "../_lib/rateLimit.js";
import { saveStoredSession } from "../_lib/sessionRepository.js";

export default async function handler(req: VercelRequest, res: VercelResponse) {
  if (req.method !== "POST") {
    return noStore(res).status(405).json({ error: "接口方法不支持" });
  }

  try {
    if (!await enforceRateLimit(req, res, "create")) return;
    const body = (req.body ?? {}) as { lineId?: string; modelId?: string; modelName?: string };
    const session = createSession(
      body.lineId ?? "a-line",
      body.modelId ?? "ordinary-washer-dryer",
      body.modelName
    );
    const saved = await saveStoredSession(session);
    return noStore(res).status(201).json({ session: saved });
  } catch (error) {
    return apiError(res, error, "创建测量单失败");
  }
}
