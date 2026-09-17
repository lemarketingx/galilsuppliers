import { existsSync, mkdirSync, readFileSync, writeFileSync } from 'node:fs';
import { dirname, join } from 'node:path';
import { fileURLToPath } from 'node:url';

const __dirname = dirname(fileURLToPath(import.meta.url));
const DATA_DIR = join(__dirname, '..', 'data');
const USERS_FILE = join(DATA_DIR, 'users.json');

function ensureStore() {
  if (!existsSync(DATA_DIR)) mkdirSync(DATA_DIR, { recursive: true });
  if (!existsSync(USERS_FILE)) writeFileSync(USERS_FILE, JSON.stringify({ users: [] }, null, 2));
}

function readAll() {
  ensureStore();
  try { return JSON.parse(readFileSync(USERS_FILE, 'utf8')); } catch { return { users: [] }; }
}

function writeAll(data) {
  writeFileSync(USERS_FILE, JSON.stringify(data, null, 2));
}

export function listUsers() {
  return readAll().users;
}

export function findUserByUsername(username) {
  return readAll().users.find(u => u.username.toLowerCase() === String(username || '').toLowerCase());
}

export function findUserById(id) {
  return readAll().users.find(u => u.id === id);
}

export function insertUser(user) {
  const data = readAll();
  data.users.push(user);
  writeAll(data);
  return user;
}

export function updateUser(id, patch) {
  const data = readAll();
  const idx = data.users.findIndex(u => u.id === id);
  if (idx === -1) return null;
  data.users[idx] = { ...data.users[idx], ...patch };
  writeAll(data);
  return data.users[idx];
}

export function deleteUser(id) {
  const data = readAll();
  const before = data.users.length;
  data.users = data.users.filter(u => u.id !== id);
  writeAll(data);
  return data.users.length < before;
}
