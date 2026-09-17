import { randomUUID } from 'node:crypto';
import { existsSync, mkdirSync } from 'node:fs';
import { dirname, extname, join } from 'node:path';
import { fileURLToPath } from 'node:url';
import { Router } from 'express';
import multer from 'multer';
import { insertUpload, listUploads } from '../uploadsDb.js';
import { requireAuth, requireRole } from '../middleware/auth.js';

const __dirname = dirname(fileURLToPath(import.meta.url));
const UPLOAD_DIR = join(__dirname, '..', '..', 'uploads');
if (!existsSync(UPLOAD_DIR)) mkdirSync(UPLOAD_DIR, { recursive: true });

const ALLOWED_KINDS = ['boq-excel', 'suppliers-excel', 'attachment'];
const ALLOWED_EXTENSIONS = ['.xlsx', '.xls', '.xlsm', '.csv', '.pdf', '.jpg', '.jpeg', '.png', '.dwg', '.doc', '.docx'];
const MAX_SIZE = 15 * 1024 * 1024; // 15MB, matches the client-side attachment limit

const storage = multer.diskStorage({
  destination: (req, file, cb) => cb(null, UPLOAD_DIR),
  filename: (req, file, cb) => cb(null, `${Date.now()}-${randomUUID()}${extname(file.originalname).toLowerCase()}`),
});

const upload = multer({
  storage,
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
  upload.single('file')(req, res, err => {
    if (err) return res.status(400).json({ error: err.message === 'סוג קובץ לא נתמך.' ? err.message : 'שגיאה בהעלאת הקובץ (ייתכן שהקובץ גדול מדי).' });
    if (!req.file) return res.status(400).json({ error: 'לא התקבל קובץ.' });
    const kind = ALLOWED_KINDS.includes(req.body.kind) ? req.body.kind : 'attachment';
    const entry = {
      id: randomUUID(),
      originalName: req.file.originalname,
      storedAs: req.file.filename,
      kind,
      size: req.file.size,
      uploadedBy: req.user.username,
      uploadedAt: new Date().toISOString(),
    };
    insertUpload(entry);
    res.status(201).json({ upload: entry });
  });
});

router.get('/', requireRole('admin'), (req, res) => {
  res.json({ uploads: listUploads() });
});

export default router;
