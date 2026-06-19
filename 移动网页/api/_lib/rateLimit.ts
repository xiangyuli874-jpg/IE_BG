import { Ratelimit } from "@upstash/ratelimit";
import type { Duration } from "@upstash/ratelimit";
import type { VercelRequest, VercelResponse } from "@vercel/node";
import { getRedis } from "./redis.js";

export const RATE_LIMIT_POLICIES = {
  create: { requests: 30, window: "1 h" },
  read: { requests: 300, window: "10 m" },
  save: { requests: 600, window: "10 m" },
  export: { requests: 60, window: "1 h" }
} as const;

type RateLimitPolicy = keyof typeof RATE_LIMIT_POLICIES;
const limiters = new Map<RateLimitPolicy, Ratelimit>();

function requestIp(req: VercelRequest) {
  const forwarded = req.headers["x-forwarded-for"];
  const value = Array.isArray(forwarded) ? forwarded[0] : forwarded;
  return value?.split(",")[0]?.trim() || req.socket.remoteAddress || "unknown";
}

function limiterFor(policyName: RateLimitPolicy) {
  const cached = limiters.get(policyName);
  if (cached) return cached;

  const policy = RATE_LIMIT_POLICIES[policyName];
  const limiter = new Ratelimit({
    redis: getRedis(),
    limiter: Ratelimit.slidingWindow(policy.requests, policy.window as Duration),
    prefix: `worktime:ratelimit:${policyName}`,
    analytics: false
  });
  limiters.set(policyName, limiter);
  return limiter;
}

export async function enforceRateLimit(
  req: VercelRequest,
  res: VercelResponse,
  policyName: RateLimitPolicy
) {
  const result = await limiterFor(policyName).limit(requestIp(req));
  res.setHeader("X-RateLimit-Limit", String(result.limit));
  res.setHeader("X-RateLimit-Remaining", String(result.remaining));
  res.setHeader("X-RateLimit-Reset", String(result.reset));
  if (result.success) return true;

  res.setHeader("Retry-After", String(Math.max(1, Math.ceil((result.reset - Date.now()) / 1000))));
  res.status(429).json({ error: "操作过于频繁，请稍后再试" });
  return false;
}
