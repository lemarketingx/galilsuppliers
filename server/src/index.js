import { randomUUID } from 'node:crypto';
import 'dotenv/config';
import cors from 'cors';
import express from 'express';
import { hashPassword } from './auth.js';
import { insertUser, listUsers } from './db.js';
import authRoutes from './routes/auth.js';
import usersRoutes from './routes/users.js';

function bootstrapAdmin() {
  if (listUsers().length > 0) return;
  const username = process.env.ADMIN_USERNAME || 'admin';
  const password = process.env.ADMIN_PASSWORD || 'change-me-please';
  insertUser({
    id: randomUUID(),
    username,
    passwordHash: hashPassword(password),
    role: 'admin',
    createdAt: new Date().toISOString(),
  });
  console.log(`[bootstrap] Created initial admin user "${username}". Log in and change the password immediately.`);
}

bootstrapAdmin();

const app = express();
const allowedOrigins = (process.env.FRONTEND_ORIGIN || '').split(',').map(s => s.trim()).filter(Boolean);
app.use(cors({
  origin: allowedOrigins.length ? allowedOrigins : true,
}));
app.use(express.json());

app.get('/api/health', (req, res) => res.json({ ok: true }));
app.use('/api/auth', authRoutes);
app.use('/api/users', usersRoutes);

app.use((err, req, res, next) => {
  console.error(err);
  res.status(500).json({ error: 'שגיאת שרת פנימית.' });
});

const PORT = process.env.PORT || 4000;
app.listen(PORT, () => console.log(`[server] Auth server listening on port ${PORT}`));
