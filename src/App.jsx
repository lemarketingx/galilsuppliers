import React, { useMemo, useRef, useState, useEffect } from 'react';
import * as XLSX from 'xlsx';
import { Search, Upload, Trash2, Star, StarHalf, Download, RotateCcw, Building2, Phone, Mail, MapPin, Tag, FileSpreadsheet } from 'lucide-react';
import logo from './galil-logo.webp';

const STORAGE_KEY = 'galil_suppliers_directory_v3';

const demoSuppliers = [
  { id: crypto.randomUUID(), name: 'Technobar', field: 'מכשור / ברזים / צנרת', contact: 'Adi', phone: '', email: '', city: 'ישראל', catalog: 'Instrumentation', rating: 4, status: 'מאושר', notes: 'נתוני דוגמה בלבד' },
  { id: crypto.randomUUID(), name: 'ABB', field: 'חשמל / אוטומציה / ציוד תעשייתי', contact: '', phone: '', email: '', city: '', catalog: 'Electrical', rating: 5, status: 'מאושר', notes: 'נתוני דוגמה בלבד' },
  { id: crypto.randomUUID(), name: 'HAPMAN', field: 'מערכות שינוע / ציוד מכני', contact: '', phone: '', email: '', city: 'USA', catalog: 'Mechanical', rating: 4, status: 'בבדיקה', notes: 'נתוני דוגמה בלבד' },
  { id: crypto.randomUUID(), name: 'DELFIN S.R.L', field: 'מערכות העמסה / ציוד תהליך', contact: '', phone: '', email: '', city: 'Italy', catalog: 'Process Equipment', rating: 3, status: 'בבדיקה', notes: 'נתוני דוגמה בלבד' },
];

const engineeringFields = [
  'כל התחומים', 'צנרת', 'חשמל', 'מכשור ובקרה', 'הנדסה אזרחית', 'קונסטרוקציה', 'מכונות', 'ציוד תהליך', 'משאבות', 'ברזים', 'פלדות ומתכות', 'בטון', 'קבלני ביצוע', 'לוחות חשמל', 'כבלים', 'אוטומציה', 'HVAC', 'בטיחות', 'בידוד וצבע', 'שינוע', 'רכש כללי'
];

function normalize(value) {
  return String(value ?? '').trim();
}

function numberOrDefault(value, fallback = 0) {
  const n = Number(String(value ?? '').replace(',', '.'));
  return Number.isFinite(n) ? n : fallback;
}

function getByHeaders(row, options) {
  const keys = Object.keys(row);
  const cleanMap = new Map(keys.map(k => [String(k).replace(/\s+/g, '').replace(/"/g, '').trim(), k]));
  for (const option of options) {
    const clean = String(option).replace(/\s+/g, '').replace(/"/g, '').trim();
    if (row[option] !== undefined) return row[option];
    if (cleanMap.has(clean)) return row[cleanMap.get(clean)];
  }
  return '';
}

function supplierFromRow(row, index) {
  const name = normalize(getByHeaders(row, ['שם ספק', 'ספק', 'שם הספק', 'Supplier', 'Vendor', 'vendor_name'])) || `ספק ${index + 1}`;
  const field = normalize(getByHeaders(row, ['תחום', 'תחום פעילות', 'קטגוריה', 'תיאור תחום', 'דיסציפלינה', 'תאור משאב', 'תיאור משאב', 'Resource Description', 'Category'])) || 'לא סווג';
  return {
    id: crypto.randomUUID(),
    name,
    field,
    contact: normalize(getByHeaders(row, ['איש קשר', 'איש קשר ספק', 'Contact', 'contact_name'])),
    phone: normalize(getByHeaders(row, ['טלפון', 'נייד', 'מספר טלפון', 'Phone', 'Mobile'])),
    email: normalize(getByHeaders(row, ['מייל', 'אימייל', 'דואר אלקטרוני', 'Email', 'email'])),
    city: normalize(getByHeaders(row, ['עיר', 'מיקום', 'מדינה', 'כתובת', 'Location', 'City', 'Country'])),
    catalog: normalize(getByHeaders(row, ['מק״ט', 'מקט', 'מק"ט', 'קוד ספק', 'מספר ספק', 'Catalog', 'SKU'])),
    status: normalize(getByHeaders(row, ['סטטוס', 'Status'])) || 'חדש',
    notes: normalize(getByHeaders(row, ['הערות', 'Notes', 'Remarks', 'תיאור', 'תאור'])),
    rating: Math.min(5, Math.max(0, numberOrDefault(getByHeaders(row, ['דירוג', 'Rating', 'rating']), 0))),
    raw: row,
  };
}

export default function App() {
  const fileRef = useRef(null);
  const [suppliers, setSuppliers] = useState(() => {
    try { return JSON.parse(localStorage.getItem(STORAGE_KEY)) || demoSuppliers; } catch { return demoSuppliers; }
  });
  const [query, setQuery] = useState('');
  const [field, setField] = useState('כל התחומים');
  const [status, setStatus] = useState('הכל');
  const [message, setMessage] = useState('');

  useEffect(() => localStorage.setItem(STORAGE_KEY, JSON.stringify(suppliers)), [suppliers]);

  const availableFields = useMemo(() => {
    const set = new Set(engineeringFields);
    suppliers.forEach(s => s.field && set.add(s.field));
    return Array.from(set);
  }, [suppliers]);

  const filtered = useMemo(() => {
    const q = query.toLowerCase().trim();
    return suppliers.filter(s => {
      const text = `${s.name} ${s.field} ${s.contact} ${s.email} ${s.phone} ${s.city} ${s.catalog} ${s.notes}`.toLowerCase();
      const byQuery = !q || text.includes(q);
      const byField = field === 'כל התחומים' || s.field.includes(field) || field.includes(s.field);
      const byStatus = status === 'הכל' || s.status === status;
      return byQuery && byField && byStatus;
    });
  }, [suppliers, query, field, status]);

  const stats = useMemo(() => ({
    total: suppliers.length,
    approved: suppliers.filter(s => s.status === 'מאושר').length,
    avg: suppliers.length ? (suppliers.reduce((a, s) => a + Number(s.rating || 0), 0) / suppliers.length).toFixed(1) : '0.0',
    fields: new Set(suppliers.map(s => s.field)).size,
  }), [suppliers]);

  async function handleExcelUpload(e) {
    const file = e.target.files?.[0];
    if (!file) return;
    try {
      const data = await file.arrayBuffer();
      const workbook = XLSX.read(data, { type: 'array' });
      const sheet = workbook.Sheets[workbook.SheetNames[0]];
      const rows = XLSX.utils.sheet_to_json(sheet, { defval: '' });
      const parsed = rows.map(supplierFromRow).filter(s => s.name && s.name !== '-');
      setSuppliers(parsed);
      setMessage(`נטענו ${parsed.length} ספקים מהקובץ: ${file.name}`);
    } catch (err) {
      setMessage('שגיאה בקריאת הקובץ. ודא שהקובץ בפורמט XLSX ושיש שורת כותרות תקינה.');
    } finally {
      e.target.value = '';
    }
  }

  function deleteSupplier(id) {
    setSuppliers(prev => prev.filter(s => s.id !== id));
  }

  function setRating(id, rating) {
    setSuppliers(prev => prev.map(s => s.id === id ? { ...s, rating } : s));
  }

  function setSupplierStatus(id, nextStatus) {
    setSuppliers(prev => prev.map(s => s.id === id ? { ...s, status: nextStatus } : s));
  }

  function resetDemo() {
    setSuppliers(demoSuppliers.map(s => ({ ...s, id: crypto.randomUUID() })));
    setMessage('הוחזרו נתוני הדוגמה.');
  }

  function exportExcel() {
    const rows = suppliers.map(s => ({
      'שם ספק': s.name,
      'תחום': s.field,
      'איש קשר': s.contact,
      'טלפון': s.phone,
      'מייל': s.email,
      'מיקום': s.city,
      'מק"ט/קוד': s.catalog,
      'סטטוס': s.status,
      'דירוג': s.rating,
      'הערות': s.notes,
    }));
    const ws = XLSX.utils.json_to_sheet(rows);
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, 'Suppliers');
    XLSX.writeFile(wb, 'galil_suppliers_export.xlsx');
  }

  return (
    <div className="app" dir="rtl">
      <header className="hero">
        <div className="heroInner">
          <img src={logo} className="logo" alt="Galil Group" />
          <div>
            <p className="eyebrow">GALIL GROUP · Supplier Intelligence</p>
            <h1>מאגר ספקים חכם לחברת הנדסה</h1>
            <p className="sub">חיפוש ספקים לפי שם, תחום, סטטוס ודירוג — כולל העלאת Excel לפי הפורמט המצורף.</p>
          </div>
        </div>
      </header>

      <main className="container">
        <section className="toolbar card">
          <div className="uploadBox" onClick={() => fileRef.current?.click()}>
            <Upload size={26} />
            <div>
              <strong>העלאת קובץ Excel</strong>
              <span>עמודות נתמכות: שם ספק/ספק, תחום/תאור משאב, איש קשר, טלפון, מייל, מיקום, סטטוס, דירוג, הערות</span>
            </div>
            <input ref={fileRef} type="file" accept=".xlsx,.xls" onChange={handleExcelUpload} hidden />
          </div>
          <button onClick={exportExcel} className="action secondary"><Download size={18}/> ייצוא Excel</button>
          <button onClick={resetDemo} className="action ghost"><RotateCcw size={18}/> איפוס דוגמה</button>
        </section>

        {message && <div className="message">{message}</div>}

        <section className="statsGrid">
          <Stat title="סה״כ ספקים" value={stats.total} icon={<Building2/>}/>
          <Stat title="ספקים מאושרים" value={stats.approved} icon={<FileSpreadsheet/>}/>
          <Stat title="דירוג ממוצע" value={stats.avg} icon={<Star/>}/>
          <Stat title="תחומי פעילות" value={stats.fields} icon={<Tag/>}/>
        </section>

        <section className="filters card">
          <div className="searchBox">
            <Search size={20}/>
            <input value={query} onChange={e => setQuery(e.target.value)} placeholder="חיפוש לפי שם ספק, תחום, איש קשר, מייל, עיר או הערות..." />
          </div>
          <select value={field} onChange={e => setField(e.target.value)}>
            {availableFields.map(f => <option key={f}>{f}</option>)}
          </select>
          <select value={status} onChange={e => setStatus(e.target.value)}>
            {['הכל','חדש','מאושר','בבדיקה','לא מאושר'].map(s => <option key={s}>{s}</option>)}
          </select>
        </section>

        <section className="resultsHeader">
          <h2>תוצאות חיפוש</h2>
          <span>{filtered.length} מתוך {suppliers.length} ספקים</span>
        </section>

        <section className="grid">
          {filtered.map(supplier => (
            <SupplierCard key={supplier.id} supplier={supplier} onDelete={deleteSupplier} onRate={setRating} onStatus={setSupplierStatus}/>
          ))}
        </section>

        {filtered.length === 0 && <div className="empty card">לא נמצאו ספקים לפי הסינון הנוכחי.</div>}
      </main>
    </div>
  );
}

function Stat({ title, value, icon }) {
  return <div className="stat card"><div className="statIcon">{icon}</div><div><span>{title}</span><strong>{value}</strong></div></div>;
}

function SupplierCard({ supplier, onDelete, onRate, onStatus }) {
  return (
    <article className="supplier card">
      <div className="supplierTop">
        <div>
          <span className="pill">{supplier.field}</span>
          <h3>{supplier.name}</h3>
        </div>
        <button className="delete" onClick={() => onDelete(supplier.id)} title="מחיקת ספק"><Trash2 size={18}/></button>
      </div>

      <div className="rating" aria-label="דירוג ספק">
        {[1,2,3,4,5].map(n => (
          <button key={n} onClick={() => onRate(supplier.id, n)} className={n <= supplier.rating ? 'star active' : 'star'} title={`דרג ${n}`}><Star size={20} fill="currentColor"/></button>
        ))}
        <span>{supplier.rating || 0}/5</span>
      </div>

      <div className="details">
        {supplier.contact && <p><Building2 size={16}/> {supplier.contact}</p>}
        {supplier.phone && <p><Phone size={16}/> {supplier.phone}</p>}
        {supplier.email && <p><Mail size={16}/> {supplier.email}</p>}
        {supplier.city && <p><MapPin size={16}/> {supplier.city}</p>}
        {supplier.catalog && <p><Tag size={16}/> {supplier.catalog}</p>}
      </div>

      {supplier.notes && <p className="notes">{supplier.notes}</p>}

      <div className="statusRow">
        <span>סטטוס:</span>
        <select value={supplier.status} onChange={e => onStatus(supplier.id, e.target.value)}>
          {['חדש','מאושר','בבדיקה','לא מאושר'].map(s => <option key={s}>{s}</option>)}
        </select>
      </div>
    </article>
  );
}
