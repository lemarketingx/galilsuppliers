import { randomUUID } from 'node:crypto';
import { Router } from 'express';
import { deleteUser, findUserById, findUserByUsername, insertUser, listUsers, updateUser } from '../db.js';
import { hashPassword, ROLES } from '../auth.js';
import { requireAuth, requireRole } from '../authMiddleware.js';
import { deleteRequest, listRequests } from '../resetRequestsDb.js';

const router = Router();
router.use(requireAuth, requireRole('admin'));

const toPublic = u => ({ id: u.id, username: u.username, role: u.role, createdAt: u.createdAt });

router.get('/', async (req, res) => {
  res.json({ users: (await listUsers()).map(toPublic) });
});

router.post('/', async (req, res) => {
  const { username, password, role } = req.body || {};
  if (!username || !password) return res.status(400).json({ error: 'יש להזין שם משתמש וסיסמה.' });
  if (password.length < 6) return res.status(400).json({ error: 'הסיסמה חייבת להכיל לפחות 6 תווים.' });
  if (role && !ROLES.includes(role)) return res.status(400).json({ error: 'תפקיד לא תקין.' });
  if (await findUserByUsername(username)) return res.status(409).json({ error: 'שם המשתמש כבר קיים.' });
  const user = await insertUser({
    id: randomUUID(),
    username: String(username).trim(),
    passwordHash: hashPassword(password),
    role: role || 'viewer',
    createdAt: new Date().toISOString(),
  });
  res.status(201).json({ user: toPublic(user) });
});

// Password reset requests, submitted from the public "שכחתי סיסמה" screen.
// Registered before "/:id" so "reset-requests" is never swallowed by it.
router.get('/reset-requests', async (req, res) => {
  res.json({ requests: await listRequests() });
});

router.delete('/reset-requests/:id', async (req, res) => {
  await deleteRequest(req.params.id);
  res.status(204).end();
});

router.patch('/:id', async (req, res) => {
  const target = await findUserById(req.params.id);
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
  const updated = await updateUser(req.params.id, patch);
  res.json({ user: toPublic(updated) });
});

router.delete('/:id', async (req, res) => {
  if (req.params.id === req.user.id) return res.status(400).json({ error: 'לא ניתן למחוק את המשתמש המחובר.' });
  const target = await findUserById(req.params.id);
  if (!target) return res.status(404).json({ error: 'המשתמש לא נמצא.' });
  await deleteUser(req.params.id);
  res.status(204).end();
});

export default router;
