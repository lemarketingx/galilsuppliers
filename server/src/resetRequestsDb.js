import { existsSync, mkdirSync, readFileSync, writeFileSync } from 'node:fs';
import { dirname, join } from 'node:path';
import { fileURLToPath } from 'node:url';

const __dirname = dirname(fileURLToPath(import.meta.url));
const DATA_DIR = join(__dirname, '..', 'data');
const FILE = join(DATA_DIR, 'reset-requests.json');

function ensureStore() {
  if (!existsSync(DATA_DIR)) mkdirSync(DATA_DIR, { recursive: true });
  if (!existsSync(FILE)) writeFileSync(FILE, JSON.stringify({ requests: [] }, null, 2));
}

function readAll() {
  ensureStore();
  try { return JSON.parse(readFileSync(FILE, 'utf8')); } catch { return { requests: [] }; }
}

function writeAll(data) {
  writeFileSync(FILE, JSON.stringify(data, null, 2));
}

export function listRequests() {
  return [...readAll().requests].sort((a, b) => b.requestedAt.localeCompare(a.requestedAt));
}

export function insertRequest(entry) {
  const data = readAll();
  // avoid piling up duplicate pending requests from the same user
  data.requests = data.requests.filter(r => r.username !== entry.username);
  data.requests.push(entry);
  writeAll(data);
  return entry;
}

export function deleteRequest(id) {
  const data = readAll();
  data.requests = data.requests.filter(r => r.id !== id);
  writeAll(data);
}
