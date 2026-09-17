import { kv } from './kv.js';

const KEY = 'resetRequests';

export async function listRequests() {
  const requests = (await kv.get(KEY)) || [];
  return [...requests].sort((a, b) => b.requestedAt.localeCompare(a.requestedAt));
}

export async function insertRequest(entry) {
  const requests = (await kv.get(KEY)) || [];
  // avoid piling up duplicate pending requests from the same user
  const next = requests.filter(r => r.username !== entry.username);
  next.push(entry);
  await kv.set(KEY, next);
  return entry;
}

export async function deleteRequest(id) {
  const requests = (await kv.get(KEY)) || [];
  await kv.set(KEY, requests.filter(r => r.id !== id));
}
