import { existsSync, mkdirSync, readFileSync, writeFileSync } from 'node:fs';
import { dirname, join } from 'node:path';
import { fileURLToPath } from 'node:url';

const __dirname = dirname(fileURLToPath(import.meta.url));
const DATA_DIR = join(__dirname, '..', 'data');
const LOG_FILE = join(DATA_DIR, 'uploads.json');

function ensureStore() {
  if (!existsSync(DATA_DIR)) mkdirSync(DATA_DIR, { recursive: true });
  if (!existsSync(LOG_FILE)) writeFileSync(LOG_FILE, JSON.stringify({ uploads: [] }, null, 2));
}

function readAll() {
  ensureStore();
  try { return JSON.parse(readFileSync(LOG_FILE, 'utf8')); } catch { return { uploads: [] }; }
}

function writeAll(data) {
  writeFileSync(LOG_FILE, JSON.stringify(data, null, 2));
}

export function listUploads() {
  return [...readAll().uploads].sort((a, b) => b.uploadedAt.localeCompare(a.uploadedAt));
}

export function insertUpload(entry) {
  const data = readAll();
  data.uploads.push(entry);
  writeAll(data);
  return entry;
}
