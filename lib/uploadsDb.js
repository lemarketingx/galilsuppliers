import { kv } from './kv.js';

const KEY = 'uploads';

export async function listUploads() {
  const uploads = (await kv.get(KEY)) || [];
  return [...uploads].sort((a, b) => b.uploadedAt.localeCompare(a.uploadedAt));
}

export async function insertUpload(entry) {
  const uploads = (await kv.get(KEY)) || [];
  uploads.push(entry);
  await kv.set(KEY, uploads);
  return entry;
}
