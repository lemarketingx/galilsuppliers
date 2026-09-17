import { randomUUID } from 'node:crypto';
import { Router } from 'express';
import { deleteUser, findUserById, findUserByUsername, insertUser, listUsers, updateUser } from '../db.js';
import { hashPassword, ROLES } from '../auth.js';
import { requireAuth, requireRole } from '../middleware/auth.js';

const router = Router();
router.use(requireAuth, requireRole('admin'));

const toPublic = u => ({ id: u.id, username: u.username, role: u.role, createdAt: u.createdAt });

router.get('/', (req, res) => {
  res.json({ users: listUsers().map(toPublic) });
});

router.post('/', (req, res) => {
  const { username, password, role } = req.body || {};
  if (!username || !password) return res.status(400).json({ error: 'יש להזין שם משתמש וסיסמה.' });
  if (password.length < 6) return res.status(400).json({ error: 'הסיסמה חייבת להכיל לפחות 6 תווים.' });
  if (role && !ROLES.includes(role)) return res.status(400).json({ error: 'תפקיד לא תקין.' });
  if (findUserByUsername(username)) return res.status(409).json({ error: 'שם המשתמש כבר קיים.' });
  const user = insertUser({
    id: randomUUID(),
    username: String(username).trim(),
    passwordHash: hashPassword(password),
    role: role || 'viewer',
    createdAt: new Date().toISOString(),
  });
  res.status(201).json({ user: toPublic(user) });
});

router.patch('/:id', (req, res) => {
  const target = findUserById(req.params.id);
  if (!target) return res.status(404).json({ error: 'המשתמש לא נמצא.' });
  const { role, password } = req.body || {};
  const patch = {};
  if (role !== undefined) {
    if (!ROLES.includes(role)) return res.status(400).json({ error: 'תפקיד לא תקין.' });
    if (target.role === 'admin' && role !== 'admin' && target.id === req.user.id) {
      return res.status(400).json({ error: 'לא ניתן להסיר הרשאת מנהל מעצמך.' });
    }
    patch.role = role;
  }
  if (password !== undefined) {
    if (password.length < 6) return res.status(400).json({ error: 'הסיסמה חייבת להכיל לפחות 6 תווים.' });
    patch.passwordHash = hashPassword(password);
  }
  const updated = updateUser(req.params.id, patch);
  res.json({ user: toPublic(updated) });
});

router.delete('/:id', (req, res) => {
  if (req.params.id === req.user.id) return res.status(400).json({ error: 'לא ניתן למחוק את המשתמש המחובר.' });
  const target = findUserById(req.params.id);
  if (!target) return res.status(404).json({ error: 'המשתמש לא נמצא.' });
  deleteUser(req.params.id);
  res.status(204).end();
});

export default router;
