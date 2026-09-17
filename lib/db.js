import { kv } from './kv.js';

const KEY = 'users';

export async function listUsers() {
  return (await kv.get(KEY)) || [];
}

export async function findUserByUsername(username) {
  const users = await listUsers();
  return users.find(u => u.username.toLowerCase() === String(username || '').toLowerCase());
}

export async function findUserById(id) {
  const users = await listUsers();
  return users.find(u => u.id === id);
}

export async function insertUser(user) {
  const users = await listUsers();
  users.push(user);
  await kv.set(KEY, users);
  return user;
}

export async function updateUser(id, patch) {
  const users = await listUsers();
  const idx = users.findIndex(u => u.id === id);
  if (idx === -1) return null;
  users[idx] = { ...users[idx], ...patch };
  await kv.set(KEY, users);
  return users[idx];
}

export async function deleteUser(id) {
  const users = await listUsers();
  const before = users.length;
  const next = users.filter(u => u.id !== id);
  await kv.set(KEY, next);
  return next.length < before;
}
