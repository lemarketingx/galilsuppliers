import { existsSync, mkdirSync, readFileSync, writeFileSync } from 'node:fs';
import { dirname, join } from 'node:path';
import { fileURLToPath } from 'node:url';

/* Storage abstraction used by every collection (users, uploads, reset
   requests): a handful of get/set(key, jsonValue) calls.

   On Vercel, once a Redis storage integration is connected to the project
   (Vercel's "KV" product, or the Upstash Redis integration from the
   Marketplace - both are the same underlying service, just under different
   names depending on when/how you added it), Vercel injects REST
   credentials as environment variables and we talk to that - data then
   persists properly across requests, deployments and cold starts (unlike
   the local filesystem, which Vercel's serverless functions cannot reliably
   write to). Different integration flows have named those variables
   differently over time, so both are checked here.

   Without any of those env vars (plain `node local-api-server.mjs`, or a
   self-hosted server with no Redis attached), we fall back to a small JSON
   file per key under .local-data/, so the app still works out of the box
   for local development. */

const kvUrl = process.env.KV_REST_API_URL;
const kvToken = process.env.KV_REST_API_TOKEN;
const upstashUrl = process.env.UPSTASH_REDIS_REST_URL;
const upstashToken = process.env.UPSTASH_REDIS_REST_TOKEN;
const hasRedis = !!((kvUrl && kvToken) || (upstashUrl && upstashToken));

async function createImpl() {
  if (kvUrl && kvToken) {
    const { kv } = await import('@vercel/kv');
    return { get: k => kv.get(k), set: (k, v) => kv.set(k, v) };
  }
  if (upstashUrl && upstashToken) {
    const { Redis } = await import('@upstash/redis');
    const redis = new Redis({ url: upstashUrl, token: upstashToken });
    return { get: k => redis.get(k), set: (k, v) => redis.set(k, v) };
  }
  const __dirname = dirname(fileURLToPath(import.meta.url));
  const DIR = join(__dirname, '..', '.local-data');
  const fileFor = k => join(DIR, `${k}.json`);
  return {
    get: async k => {
      if (!existsSync(fileFor(k))) return null;
      try { return JSON.parse(readFileSync(fileFor(k), 'utf8')); } catch { return null; }
    },
    set: async (k, v) => {
      if (!existsSync(DIR)) mkdirSync(DIR, { recursive: true });
      writeFileSync(fileFor(k), JSON.stringify(v, null, 2));
    },
  };
}

const implPromise = createImpl();
export const kv = {
  get: async k => (await implPromise).get(k),
  set: async (k, v) => (await implPromise).set(k, v),
};
export const usingLocalFallback = !hasRedis;
