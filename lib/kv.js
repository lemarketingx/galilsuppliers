import { existsSync, mkdirSync, readFileSync, writeFileSync } from 'node:fs';
import { dirname, join } from 'node:path';
import { fileURLToPath } from 'node:url';

/* Storage abstraction used by every collection (users, uploads, reset
   requests): a handful of get/set(key, jsonValue) calls.

   On Vercel, once the "KV" (Upstash Redis) storage integration is added to
   the project, Vercel injects KV_REST_API_URL/KV_REST_API_TOKEN and we use
   the real @vercel/kv client - data then persists properly across requests,
   deployments and cold starts (unlike the local filesystem, which Vercel's
   serverless functions cannot reliably write to).

   Without those env vars (plain `node local-api-server.mjs`, or a
   self-hosted server with no KV attached), we fall back to a small JSON
   file per key under .local-data/, so the app still works out of the box
   for local development. */

const hasVercelKV = !!(process.env.KV_REST_API_URL && process.env.KV_REST_API_TOKEN);

async function createImpl() {
  if (hasVercelKV) {
    const { kv } = await import('@vercel/kv');
    return { get: k => kv.get(k), set: (k, v) => kv.set(k, v) };
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
export const usingLocalFallback = !hasVercelKV;
