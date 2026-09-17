import { verifyToken } from './auth.js';
import { findUserById } from './db.js';

export async function requireAuth(req, res, next) {
  const header = req.headers.authorization || '';
  const token = header.startsWith('Bearer ') ? header.slice(7) : null;
  if (!token) return res.status(401).json({ error: 'התחברות נדרשת.' });
  const payload = verifyToken(token);
  if (!payload) return res.status(401).json({ error: 'ההתחברות פגה או אינה תקפה. יש להתחבר מחדש.' });
  const user = await findUserById(payload.sub);
  if (!user) return res.status(401).json({ error: 'המשתמש אינו קיים עוד.' });
  req.user = { id: user.id, username: user.username, role: user.role };
  next();
}

export function requireRole(...roles) {
  return (req, res, next) => {
    if (!req.user || !roles.includes(req.user.role)) {
      return res.status(403).json({ error: 'אין לך הרשאה לבצע פעולה זו.' });
    }
    next();
  };
}
