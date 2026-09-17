import express from 'express';
import { ensureBootstrapAdmin } from './bootstrap.js';
import authRoutes from './routes/auth.js';
import usersRoutes from './routes/users.js';
import uploadsRoutes from './routes/uploads.js';

const app = express();
app.use(express.json());

// Frontend and API are served from the same Vercel deployment (same
// origin), so no CORS handling is needed here.

app.use(async (req, res, next) => { await ensureBootstrapAdmin(); next(); });

app.get('/api/health', (req, res) => res.json({ ok: true }));
app.use('/api/auth', authRoutes);
app.use('/api/users', usersRoutes);
app.use('/api/uploads', uploadsRoutes);

app.use((err, req, res, next) => {
  console.error(err);
  res.status(500).json({ error: 'שגיאת שרת פנימית.' });
});

export default app;
