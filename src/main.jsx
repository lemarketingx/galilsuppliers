import React, { useMemo, useState, useEffect } from 'react';
import { createRoot } from 'react-dom/client';
import * as XLSX from 'xlsx';
import { Upload, Search, Star, Trash2, RefreshCw, Download, Building2, Filter, Pencil, CheckCircle2 } from 'lucide-react';
import './style.css';

const STORAGE_KEY = 'galil_suppliers_v7_fixed';

const DISCIPLINES = [
  'צנרת', 'חשמל', 'מכשור ובקרה', 'הנדסה אזרחית', 'מכונות וציוד', 'מתכת וקונסטרוקציה',
  'בידוד וצבע', 'HVAC ומיזוג', 'בטיחות וכיבוי אש', 'לוגיסטיקה ושילוח', 'כימיקלים וחומרים',
  'שירותי תכנון וייעוץ', 'הדרכות וכנסים', 'IT ותוכנה', 'כללי / אחר'
];

const sampleSuppliers = [
  { id:'demo-1', project:'00802', supplierNo:'51638', name:'ספק צנרת לדוגמה בע״מ', description:'אספקת צינורות, ברזים, אוגנים ואביזרי צנרת לפרויקט תעשייתי', sourceDiscipline:'', discipline:'צנרת', rating:4 },
  { id:'demo-2', project:'00803', supplierNo:'7008', name:'חשמל תעשייתי גליל', description:'לוחות חשמל, כבלים, תעלות ובדיקות חשמל', sourceDiscipline:'', discipline:'חשמל', rating:5 },
  { id:'demo-3', project:'00804', supplierNo:'51421', name:'קבלן בטון ופיתוח', description:'עבודות בטון, חפירה, קונסטרוקציה ועבודות אזרחיות', sourceDiscipline:'', discipline:'הנדסה אזרחית', rating:3 }
];

function normalizeHeader(value = '') {
  return String(value)
    .replace(/["׳'`’‘]/g, '')
    .replace(/[\.\/\-_:()\[\]״]/g, '')
    .replace(/\s+/g, '')
    .trim()
    .toLowerCase();
}

function getValue(row, aliases) {
  const normalized = {};
  Object.keys(row || {}).forEach((key) => {
    normalized[normalizeHeader(key)] = row[key];
  });
  for (const alias of aliases) {
    const key = normalizeHeader(alias);
    if (normalized[key] !== undefined && normalized[key] !== null && String(normalized[key]).trim() !== '') return normalized[key];
  }
  return '';
}

function clean(value) {
  return String(value ?? '').replace(/\s+/g, ' ').trim();
}

function detectDiscipline(name = '', description = '', sourceDiscipline = '') {
  const manual = clean(sourceDiscipline);
  if (manual) return manual;
  const text = `${name} ${description}`.toLowerCase();
  const rules = [
    ['צנרת', ['צנרת','צינור','צינורות','pipe','piping','valve','ברז','ברזים','אוגן','אוגנים','flange','fitting','fittings','משאבה','משאבות','flow','hydrotest','ריתוך צנרת']],
    ['חשמל', ['חשמל','כבל','כבלים','לוח חשמל','לוחות חשמל','ארון חשמל','תעלות כבלים','electric','electrical','power','cable','abb','siemens','schneider','ממסר','שנאי','תאורה']],
    ['מכשור ובקרה', ['מכשור','בקרה','instrument','instrumentation','control','plc','dcs','סקדה','scada','חיישן','sensor','מד לחץ','transmitter','שסתום בקרה']],
    ['הנדסה אזרחית', ['בטון','יציקה','קונסטרוקציה','חפירה','עפר','פיתוח','בינוי','אזרחי','civil','concrete','construction','קבלן','קבלנים','ריצוף','טפסנות','ברזל זיון']],
    ['מכונות וציוד', ['מכונה','מכונות','ציוד','מסוע','מסועים','מנוע','גיר','משאבת','equipment','mechanical','conveyor','gear','motor','bearing','compressor']],
    ['מתכת וקונסטרוקציה', ['מתכת','פלדה','נירוסטה','מסגרות','מסגריה','קונסטרוקציית פלדה','steel','metal','welding','ריתוך','galvanized','גילוון']],
    ['בידוד וצבע', ['בידוד','צבע','צביעה','coating','paint','insulation','sandblast','התזת חול']],
    ['HVAC ומיזוג', ['מיזוג','אוורור','hvac','air condition','מזגן','צילר','chiller','מפוח','fan','duct','תעלות אוויר']],
    ['בטיחות וכיבוי אש', ['בטיחות','כיבוי','אש','ספרינקלרים','sprinkler','fire','safety','גלאי','גילוי אש','מטפים']],
    ['לוגיסטיקה ושילוח', ['שילוח','הובלה','משלוח','לוגיסטיקה','מכס','עמילות','freight','shipping','dhl','ups','הטסה','ים','נמל']],
    ['כימיקלים וחומרים', ['כימיקלים','חומרי גלם','חומר','chemical','chemicals','resin','paint material']],
    ['שירותי תכנון וייעוץ', ['תכנון','ייעוץ','יועץ','הנדסה','שרטוט','אדריכל','פיקוח','ניהול פרויקט','consulting','engineering','design','cad','bim']],
    ['הדרכות וכנסים', ['הדרכה','הדרכות','כנס','קורס','השתלמות','מכללה','training','conference','סדנה','הרשמה']],
    ['IT ותוכנה', ['תוכנה','מחשוב','מחשב','אינטרנט','שרת','רישיון','רשיונות','software','it','license','מערכת','סייבר','ecovadis']]
  ];
  for (const [discipline, keywords] of rules) {
    if (keywords.some((word) => text.includes(word.toLowerCase()))) return discipline;
  }
  return 'כללי / אחר';
}

function parseWorkbookRows(rows) {
  return rows.map((row, index) => {
    const name = clean(getValue(row, ['שם ספק/קבלן', 'שם ספק', 'ספק', 'supplier', 'vendor', 'שם']));
    const description = clean(getValue(row, ['תאור הסעיף/פרק', 'תיאור הסעיף/פרק', 'תאור סעיף', 'תיאור סעיף', 'תיאור', 'תאור', 'description', 'תאור משאב']));
    const sourceDiscipline = clean(getValue(row, ['תחום עיסוק', 'תחום', 'דיסציפלינה', 'discipline', 'category', 'קטגוריה']));
    const supplierNo = clean(getValue(row, ["מס' ספק/קבלן", 'מס ספק/קבלן', 'מספר ספק/קבלן', 'מספר ספק', 'מק״ט', 'מקט', 'supplier number']));
    const project = clean(getValue(row, ['פרויקט', 'project']));
    const phone = clean(getValue(row, ['טלפון', 'נייד', 'phone', 'mobile']));
    const email = clean(getValue(row, ['מייל', 'אימייל', 'דואל', 'email']));
    const contact = clean(getValue(row, ['איש קשר', 'contact', 'contact person']));
    const discipline = detectDiscipline(name, description, sourceDiscipline);
    return {
      id: `${Date.now()}-${index}-${supplierNo || name || Math.random()}`,
      project,
      supplierNo,
      name: name || 'ספק ללא שם',
      description,
      sourceDiscipline,
      discipline,
      contact,
      phone,
      email,
      rating: 0,
      notes: '',
      importedAt: new Date().toLocaleDateString('he-IL')
    };
  }).filter(s => s.name !== 'ספק ללא שם' || s.description || s.supplierNo || s.project);
}

function App() {
  const [suppliers, setSuppliers] = useState(sampleSuppliers);
  const [query, setQuery] = useState('');
  const [discipline, setDiscipline] = useState('הכל');
  const [message, setMessage] = useState('');

  useEffect(() => {
    const saved = localStorage.getItem(STORAGE_KEY);
    if (saved) setSuppliers(JSON.parse(saved));
  }, []);

  useEffect(() => {
    localStorage.setItem(STORAGE_KEY, JSON.stringify(suppliers));
  }, [suppliers]);

  const handleExcelUpload = (event) => {
    const file = event.target.files?.[0];
    if (!file) return;
    const reader = new FileReader();
    reader.onload = (e) => {
      try {
        const data = new Uint8Array(e.target.result);
        const workbook = XLSX.read(data, { type: 'array', cellDates: true, raw: false });
        const sheetName = workbook.SheetNames[0];
        const sheet = workbook.Sheets[sheetName];
        const rows = XLSX.utils.sheet_to_json(sheet, { defval: '', raw: false });
        const parsed = parseWorkbookRows(rows);
        setSuppliers(parsed);
        setMessage(`נטענו ${parsed.length.toLocaleString('he-IL')} ספקים. שם הספק נלקח מעמודת "שם ספק/קבלן" והסיווג נעשה לפי "תחום עיסוק" או זיהוי חכם מהשם/התיאור.`);
      } catch (err) {
        setMessage('שגיאה בקריאת הקובץ. ודא שמדובר בקובץ xlsx תקין עם שורת כותרות.');
      }
    };
    reader.readAsArrayBuffer(file);
    event.target.value = '';
  };

  const filtered = useMemo(() => {
    const q = query.toLowerCase();
    return suppliers.filter(s => {
      const inDiscipline = discipline === 'הכל' || s.discipline === discipline;
      const text = `${s.name} ${s.description} ${s.supplierNo} ${s.project} ${s.contact} ${s.email}`.toLowerCase();
      return inDiscipline && text.includes(q);
    });
  }, [suppliers, query, discipline]);

  const stats = useMemo(() => {
    return DISCIPLINES.map(d => ({ name: d, count: suppliers.filter(s => s.discipline === d).length })).filter(x => x.count > 0);
  }, [suppliers]);

  const updateSupplier = (id, patch) => setSuppliers(prev => prev.map(s => s.id === id ? { ...s, ...patch } : s));
  const deleteSupplier = (id) => setSuppliers(prev => prev.filter(s => s.id !== id));

  const exportExcel = () => {
    const data = suppliers.map(s => ({
      'פרויקט': s.project,
      "מס' ספק/קבלן": s.supplierNo,
      'שם ספק/קבלן': s.name,
      'תאור הסעיף/פרק': s.description,
      'תחום עיסוק': s.discipline,
      'איש קשר': s.contact,
      'טלפון': s.phone,
      'מייל': s.email,
      'דירוג': s.rating,
      'הערות': s.notes
    }));
    const ws = XLSX.utils.json_to_sheet(data);
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, 'Suppliers');
    XLSX.writeFile(wb, 'galil_suppliers_export.xlsx');
  };

  const resetDemo = () => {
    setSuppliers(sampleSuppliers);
    setMessage('המערכת אופסה לנתוני דוגמה.');
  };

  return <div className="app" dir="rtl">
    <header className="hero">
      <div className="hero-content">
        <div className="brand"><div className="logo">GG</div><div><div className="eyebrow">GALIL GROUP · SUPPLIER INTELLIGENCE</div><h1>מאגר ספקים חכם</h1><p>חיפוש ספקים לפי שם, פרויקט ותחום — עם העלאת Excel, סיווג אוטומטי, דירוג ומחיקה.</p></div></div>
        <label className="uploadBtn"><Upload size={20}/> העלאת Excel<input type="file" accept=".xlsx,.xls" onChange={handleExcelUpload}/></label>
      </div>
    </header>

    <main className="wrap">
      {message && <div className="message"><CheckCircle2 size={18}/>{message}</div>}
      <section className="toolbar">
        <div className="searchBox"><Search size={19}/><input value={query} onChange={e=>setQuery(e.target.value)} placeholder="חפש לפי שם ספק, מספר ספק, תיאור, פרויקט או איש קשר"/></div>
        <div className="selectBox"><Filter size={18}/><select value={discipline} onChange={e=>setDiscipline(e.target.value)}><option>הכל</option>{DISCIPLINES.map(d=><option key={d}>{d}</option>)}</select></div>
        <button onClick={exportExcel}><Download size={18}/> ייצוא Excel</button>
        <button onClick={resetDemo} className="secondary"><RefreshCw size={18}/> איפוס דמו</button>
      </section>

      <section className="stats">
        <div className="stat primary"><Building2/><div><strong>{suppliers.length.toLocaleString('he-IL')}</strong><span>ספקים במערכת</span></div></div>
        {stats.slice(0, 8).map(s => <button className="stat" key={s.name} onClick={()=>setDiscipline(s.name)}><strong>{s.count}</strong><span>{s.name}</span></button>)}
      </section>

      <section className="grid">
        {filtered.map(s => <article className="card" key={s.id}>
          <div className="cardTop"><div><span className="badge">{s.discipline}</span><h3>{s.name}</h3><p className="supplierNo">מס׳ ספק: {s.supplierNo || '-' } · פרויקט: {s.project || '-'}</p></div><button className="delete" onClick={()=>deleteSupplier(s.id)} title="מחיקת ספק"><Trash2 size={18}/></button></div>
          <p className="desc">{s.description || 'אין תיאור בקובץ'}</p>
          <div className="editRow"><Pencil size={16}/><select value={s.discipline} onChange={e=>updateSupplier(s.id, { discipline: e.target.value })}>{DISCIPLINES.map(d=><option key={d}>{d}</option>)}</select></div>
          <div className="rating">{[1,2,3,4,5].map(n => <button key={n} onClick={()=>updateSupplier(s.id,{ rating:n })} className={n <= Number(s.rating) ? 'star active' : 'star'}><Star size={20} fill="currentColor"/></button>)}</div>
          <textarea value={s.notes} onChange={e=>updateSupplier(s.id,{ notes:e.target.value })} placeholder="הערות על הספק / איכות / זמינות / מחירים" />
          <div className="meta"><span>{s.contact || 'ללא איש קשר'}</span><span>{s.phone || s.email || 'ללא פרטי קשר'}</span></div>
        </article>)}
      </section>
      {filtered.length === 0 && <div className="empty">לא נמצאו ספקים לפי החיפוש הנוכחי.</div>}
    </main>
  </div>;
}

createRoot(document.getElementById('root')).render(<App/>);
