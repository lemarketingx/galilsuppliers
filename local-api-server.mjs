import 'dotenv/config';
import app from './lib/app.js';
import { usingLocalFallback } from './lib/kv.js';

const PORT = process.env.API_PORT || 4000;
app.listen(PORT, () => {
  console.log(`[local-api] listening on http://localhost:${PORT}`);
  if (usingLocalFallback) console.log('[local-api] No KV_REST_API_URL/KV_REST_API_TOKEN found - using local .local-data/*.json files instead of Vercel KV. On Vercel, add the KV storage integration so this uses real persistent storage.');
});
