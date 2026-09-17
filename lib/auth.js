import bcrypt from 'bcryptjs';
import jwt from 'jsonwebtoken';

const JWT_SECRET = process.env.JWT_SECRET;
const JWT_EXPIRES_IN = process.env.JWT_EXPIRES_IN || '12h';

if (!JWT_SECRET || JWT_SECRET === 'change-this-to-a-long-random-secret') {
  console.warn('[auth] WARNING: JWT_SECRET is missing or still set to the example value. Set a real secret as a Vercel environment variable before going live.');
}

export function hashPassword(plain) {
  return bcrypt.hashSync(plain, 10);
}

export function verifyPassword(plain, hash) {
  return bcrypt.compareSync(plain, hash);
}

export function signToken(user) {
  return jwt.sign({ sub: user.id, username: user.username, role: user.role }, JWT_SECRET, { expiresIn: JWT_EXPIRES_IN });
}

export function verifyToken(token) {
  try { return jwt.verify(token, JWT_SECRET); } catch { return null; }
}

export const ROLES = ['viewer', 'uploader', 'admin'];
export const canUpload = role => role === 'uploader' || role === 'admin';
