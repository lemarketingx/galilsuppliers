import 'dotenv/config';
import app from './lib/app.js';
import { usingLocalFallback } from './lib/kv.js';

const PORT = process.env.API_PORT || 4000;
app.listen(PORT, () => {
  console.log(`[local-api] listening on http://localhost:${PORT}`);
  if (usingLocalFallback) console.log('[local-api] No Redis credentials found (KV_REST_API_URL/UPSTASH_REDIS_REST_URL) - using local .local-data/*.json files instead. On Vercel, connect a Redis storage integration so this uses real persistent storage.');
});
