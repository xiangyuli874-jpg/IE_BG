import type { Config, Context } from "@netlify/functions";
import { verifyAccessCode } from "./_shared/auth";
import { createSession, loadSession, saveSession } from "./_shared/session";
import { errorResponse, jsonResponse } from "./_shared/http";
import type { WorktimeSession } from "../../src/types";

export default async (req: Request, context: Context) => {
  try {
    const authError = verifyAccessCode(req);
    if (authError) return authError;

    const id = context.params.id;

    if (req.method === "POST" && !id) {
      const body = (await req.json().catch(() => ({}))) as { lineId?: string; modelId?: string; modelName?: string };
      const session = createSession(body.lineId ?? "a-line", body.modelId ?? "ordinary-washer-dryer", body.modelName);
      const saved = await saveSession(session);
      return jsonResponse({ session: saved }, { status: 201 });
    }

    if (req.method === "GET" && id) {
      const session = await loadSession(id);
      if (!session) return errorResponse("测量单不存在", 404);
      return jsonResponse({ session });
    }

    if (req.method === "PUT" && id) {
      const body = (await req.json().catch(() => null)) as WorktimeSession | null;
      if (!body || body.id !== id) return errorResponse("保存内容无效", 400);
      const saved = await saveSession(body);
      return jsonResponse({ session: saved });
    }

    return errorResponse("接口路径或方法不支持", 405);
  } catch (error) {
    return errorResponse(error instanceof Error ? error.message : "请求处理失败", 500);
  }
};

export const config: Config = {
  path: ["/api/sessions", "/api/sessions/:id"],
  method: ["GET", "POST", "PUT"]
};
