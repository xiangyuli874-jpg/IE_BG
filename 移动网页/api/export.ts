import type { VercelRequest, VercelResponse } from "@vercel/node";
import { buildWorkbookBuffer } from "../shared/excel.js";
import type { ExportRequest } from "../src/types.js";
import { apiError, noStore } from "./_lib/http.js";
import { enforceRateLimit } from "./_lib/rateLimit.js";
import { loadOrMigrateSession } from "./_lib/sessionRepository.js";

export default async function handler(req: VercelRequest, res: VercelResponse) {
  if (req.method !== "POST") {
    return noStore(res).status(405).json({ error: "接口方法不支持" });
  }

  try {
    if (!await enforceRateLimit(req, res, "export")) return;
    const body = req.body as ExportRequest | null;
    if (!body?.sessionId) return noStore(res).status(400).json({ error: "缺少测量单 ID" });
    if (body.scope !== "group" && body.scope !== "all") {
      return noStore(res).status(400).json({ error: "导出范围无效" });
    }

    const session = await loadOrMigrateSession(body.sessionId);
    if (!session) return noStore(res).status(404).json({ error: "测量单不存在" });

    const result = await buildWorkbookBuffer(session, body.scope, body.groupId);
    res.setHeader("Content-Type", "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet");
    res.setHeader("Content-Disposition", `attachment; filename*=UTF-8''${encodeURIComponent(result.fileName)}`);
    res.setHeader("X-Process-Count", String(result.processCount));
    noStore(res).status(200).send(result.buffer);
  } catch (error) {
    return apiError(res, error, "导出失败");
  }
}
