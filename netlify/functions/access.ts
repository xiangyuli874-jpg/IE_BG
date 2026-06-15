import type { Config } from "@netlify/functions";
import { verifyAccessCode } from "./_shared/auth";
import { jsonResponse } from "./_shared/http";

export default async (req: Request) => {
  const authError = verifyAccessCode(req);
  if (authError) return authError;
  return jsonResponse({ ok: true });
};

export const config: Config = {
  path: "/api/access",
  method: ["POST"]
};
