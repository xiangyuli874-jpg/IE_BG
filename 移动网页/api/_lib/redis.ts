import { Redis } from "@upstash/redis";

let redisClient: Redis | null = null;

function redisConfig() {
  const url =
    process.env.UPSTASH_REDIS_REST_KV_REST_API_URL ??
    process.env.UPSTASH_REDIS_REST_URL ??
    process.env.KV_REST_API_URL;
  const token =
    process.env.UPSTASH_REDIS_REST_KV_REST_API_TOKEN ??
    process.env.UPSTASH_REDIS_REST_TOKEN ??
    process.env.KV_REST_API_TOKEN;
  if (!url || !token) {
    throw new Error("Vercel Redis 尚未配置");
  }
  return { url, token };
}

export function getRedis() {
  redisClient ??= new Redis(redisConfig());
  return redisClient;
}
