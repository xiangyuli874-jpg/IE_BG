import type { Config } from "@netlify/functions";
import { verifyAccessCode } from "./_shared/auth";
import { buildWorkbookBuffer } from "./_shared/excel";
import { errorResponse } from "./_shared/http";
import { loadSession } from "./_shared/session";
import type { ExportRequest } from "../../src/types";

export default async (req: Request) => {
  try {
    const authError = verifyAccessCode(req);
    if (authError) return authError;

    if (req.method !== "POST") return errorResponse("接口方法不支持", 405);
    const body = (await req.json().catch(() => null)) as ExportRequest | null;
    if (!body?.sessionId) return errorResponse("缺少测量单 ID", 400);
    if (body.scope !== "group" && body.scope !== "all") return errorResponse("导出范围无效", 400);

    const session = await loadSession(body.sessionId);
    if (!session) return errorResponse("测量单不存在", 404);

    const result = await buildWorkbookBuffer(session, body.scope, body.groupId);
    const responseBody = result.buffer.buffer.slice(
      result.buffer.byteOffset,
      result.buffer.byteOffset + result.buffer.byteLength
    ) as ArrayBuffer;
    return new Response(responseBody, {
      status: 200,
      headers: {
        "Content-Type": "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        "Content-Disposition": `attachment; filename*=UTF-8''${encodeURIComponent(result.fileName)}`,
        "X-Process-Count": String(result.processCount),
        "Cache-Control": "no-store"
      }
    });
  } catch (error) {
    return errorResponse(error instanceof Error ? error.message : "导出失败", 500);
  }
};

export const config: Config = {
  path: "/api/export",
  method: ["POST"]
};
