import { randomUUID } from 'node:crypto';
import { Router } from 'express';
import { findUserByUsername } from '../db.js';
import { signToken, verifyPassword } from '../auth.js';
import { requireAuth } from '../authMiddleware.js';
import { insertRequest } from '../resetRequestsDb.js';

const router = Router();

router.post('/login', async (req, res) => {
  const { username, password } = req.body || {};
  if (!username || !password) return res.status(400).json({ error: 'יש להזין שם משתמש וסיסמה.' });
  const user = await findUserByUsername(username);
  if (!user || !verifyPassword(password, user.passwordHash)) {
    return res.status(401).json({ error: 'שם משתמש או סיסמה שגויים.' });
  }
  const token = signToken(user);
  res.json({ token, user: { id: user.id, username: user.username, role: user.role } });
});

router.get('/me', requireAuth, (req, res) => {
  res.json({ user: req.user });
});

// Public on purpose - always answers the same way so a caller can't use it to
// probe which usernames exist. If the username is real, a reset request is
// logged for an admin to handle from the "ניהול משתמשים" screen.
router.post('/forgot-password', async (req, res) => {
  const { username } = req.body || {};
  if (username) {
    const user = await findUserByUsername(username);
    if (user) await insertRequest({ id: randomUUID(), username: user.username, requestedAt: new Date().toISOString() });
  }
  res.json({ ok: true });
});

export default router;
