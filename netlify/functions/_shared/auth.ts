import { errorResponse } from "./http";

export const ACCESS_CODE_HEADER = "X-Access-Code";

function configuredAccessCode() {
  const globalWithNetlify = globalThis as typeof globalThis & {
    Netlify?: { env?: { get?: (key: string) => string | undefined } };
  };
  return globalWithNetlify.Netlify?.env?.get?.("SITE_ACCESS_CODE") ?? process.env.SITE_ACCESS_CODE;
}

export function verifyAccessCode(req: Request) {
  const expected = configuredAccessCode();
  if (!expected) {
    return errorResponse("访问码未配置", 503);
  }

  const actual = req.headers.get(ACCESS_CODE_HEADER);
  if (actual !== expected) {
    return errorResponse("访问码错误或已失效", 401);
  }

  return null;
}
