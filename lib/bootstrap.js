import { randomUUID } from 'node:crypto';
import { hashPassword } from './auth.js';
import { insertUser, listUsers } from './db.js';

/* Serverless functions have no single "server startup" moment (each
   invocation may be a fresh cold start), so instead of running this once
   like a traditional Express app would, it's called defensively at the top
   of every request and is cheap/idempotent once a user exists. */
let bootstrapped = false;

export async function ensureBootstrapAdmin() {
  if (bootstrapped) return;
  const users = await listUsers();
  if (users.length > 0) { bootstrapped = true; return; }
  const username = process.env.ADMIN_USERNAME || 'admin';
  const password = process.env.ADMIN_PASSWORD || 'change-me-please';
  await insertUser({
    id: randomUUID(),
    username,
    passwordHash: hashPassword(password),
    role: 'admin',
    createdAt: new Date().toISOString(),
  });
  console.log(`[bootstrap] Created initial admin user "${username}". Log in and change the password immediately.`);
  bootstrapped = true;
}
