import type { VercelResponse } from "@vercel/node";

export function noStore(res: VercelResponse) {
  res.setHeader("Cache-Control", "no-store");
  return res;
}

export function apiError(res: VercelResponse, error: unknown, fallback: string) {
  const message = error instanceof Error ? error.message : fallback;
  return noStore(res).status(500).json({ error: message });
}
