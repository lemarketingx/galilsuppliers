import { randomUUID } from 'node:crypto';
import { extname } from 'node:path';
import { Router } from 'express';
import multer from 'multer';
import { insertUpload, listUploads } from '../uploadsDb.js';
import { requireAuth, requireRole } from '../authMiddleware.js';

const ALLOWED_KINDS = ['boq-excel', 'suppliers-excel', 'attachment'];
const ALLOWED_EXTENSIONS = ['.xlsx', '.xls', '.xlsm', '.csv', '.pdf', '.jpg', '.jpeg', '.png', '.dwg', '.doc', '.docx'];
const MAX_SIZE = 15 * 1024 * 1024; // 15MB, matches the client-side attachment limit

// Serverless functions can't write to a persistent local disk, and the
// browser already has the original file locally - so the server's job here
// is only to check permission and log who uploaded what, not to store the
// file bytes. Hence memoryStorage() with no further use of req.file.buffer.
const upload = multer({
  storage: multer.memoryStorage(),
  limits: { fileSize: MAX_SIZE },
  fileFilter: (req, file, cb) => {
    const ext = extname(file.originalname).toLowerCase();
    if (!ALLOWED_EXTENSIONS.includes(ext)) return cb(new Error('סוג קובץ לא נתמך.'));
    cb(null, true);
  },
});

const router = Router();
router.use(requireAuth, requireRole('uploader', 'admin'));

router.post('/', (req, res) => {
  upload.single('file')(req, res, async err => {
    if (err) return res.status(400).json({ error: err.message === 'סוג קובץ לא נתמך.' ? err.message : 'שגיאה בהעלאת הקובץ (ייתכן שהקובץ גדול מדי).' });
    if (!req.file) return res.status(400).json({ error: 'לא התקבל קובץ.' });
    const kind = ALLOWED_KINDS.includes(req.body.kind) ? req.body.kind : 'attachment';
    const entry = {
      id: randomUUID(),
      originalName: req.file.originalname,
      kind,
      size: req.file.size,
      uploadedBy: req.user.username,
      uploadedAt: new Date().toISOString(),
    };
    await insertUpload(entry);
    res.status(201).json({ upload: entry });
  });
});

router.get('/', requireRole('admin'), async (req, res) => {
  res.json({ uploads: await listUploads() });
});

export default router;
