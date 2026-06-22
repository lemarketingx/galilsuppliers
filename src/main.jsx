import React, { useEffect, useMemo, useRef, useState } from 'react';
import { createRoot } from 'react-dom/client';
import * as XLSX from 'xlsx';
import jsPDF from 'jspdf';
import html2canvas from 'html2canvas';
import { UploadCloud, Calculator, Search, Plus, Trash2, Save, Printer, Download, RotateCcw, Building2, Zap, Pipette, HardHat, BarChart3, FileText, Users, Star, Pencil, CheckCircle2, Database, ClipboardList } from 'lucide-react';
import './style.css';

const BOQ_KEY = 'galil_boq_v5_real_sheets_disciplines';
const SUP_KEY = 'galil_suppliers_v9';
const logo = '/galil-logo.webp';
const fmt = v => new Intl.NumberFormat('he-IL', { style: 'currency', currency: 'ILS', maximumFractionDigits: 0 }).format(Number(v) || 0);
const num = v => { const n = Number(String(v ?? '').replace(/,/g, '').replace(/[₪$€]/g, '').trim()); return Number.isFinite(n) ? n : 0; };
const clean = v => String(v ?? '').replace(/\s+/g, ' ').trim();
const norm = v => String(v ?? '').replace(/[\u200e\u200f]/g, '').replace(/["׳'`’‘״]/g, '').replace(/[\.\/\-_:()\[\]]/g, '').replace(/\s+/g, '').trim().toLowerCase();
function getVal(row, aliases) { const n = {}; Object.keys(row || {}).forEach(k => n[norm(k)] = row[k]); for (const a of aliases) { const v = n[norm(a)]; if (v !== undefined && v !== null && String(v).trim() !== '') return v; } return ''; }

const defaultBoqDisciplines = {
  piping: { name: 'צנרת', icon: Pipette },
  electricity: { name: 'חשמל', icon: Zap },
  fire: { name: 'כיבוי אש', icon: FileText },
  hvac: { name: 'מיזוג אויר', icon: HardHat },
  civil: { name: 'הנדסה אזרחית', icon: Building2 },
  instrumentation: { name: 'מכשור', icon: Calculator }
};
function normalizeBoqDisc(v) {
  const s = clean(v).toLowerCase();
  if (!s) return 'general';
  if (['צנרת', 'piping', 'pipe'].includes(s)) return 'piping';
  if (['חשמל', 'electricity', 'electrical', 'electric'].includes(s)) return 'electricity';
  if (['כיבוי אש', 'כיבוי', 'fire', 'firefighting', 'בטיחות וכיבוי אש'].includes(s)) return 'fire';
  if (['מיזוג אויר', 'מיזוג אוויר', 'מיזוג', 'hvac'].includes(s)) return 'hvac';
  if (['אזרחי', 'הנדסה אזרחית', 'civil', 'construction'].includes(s)) return 'civil';
  if (['מכשור', 'מכשור ובקרה', 'instrumentation', 'instrument', 'control'].includes(s)) return 'instrumentation';
  return 'custom_' + norm(s).slice(0, 40);
}
function makeDisciplineMapFromSheetNames(sheetNames, existing = defaultBoqDisciplines) {
  const map = { ...existing };
  sheetNames.forEach(name => {
    const id = normalizeBoqDisc(name);
    if (!map[id]) map[id] = { name: clean(name), icon: FileText };
  });
  return map;
}
const sampleItems = [
  { disciplineId: 'piping', code: 'P-001', desc: 'אספקה והתקנת צינור CS Sch.40 בקוטר 2"', unit: 'מטר', material: 145, labor: 95, engineering: 18, overhead: 22, supplier: 'ספק דוגמה', validity: '30 יום', notes: 'כולל חיתוך וריתוך', defaultQty: 1 },
  { disciplineId: 'electricity', code: 'E-001', desc: 'אספקה והשחלת כבל חשמל 5x10 N2XY', unit: 'מטר', material: 48, labor: 32, engineering: 6, overhead: 8, supplier: 'ספק דוגמה', validity: '45 יום', notes: 'כולל סימון כבלים', defaultQty: 1 },
  { disciplineId: 'fire', code: 'F-001', desc: 'התקנת מערכת כיבוי אש / ספרינקלרים כולל אביזרים', unit: 'מטר', material: 210, labor: 135, engineering: 30, overhead: 35, supplier: 'ספק כיבוי אש', validity: '30 יום', notes: 'נתון דוגמה', defaultQty: 1 },
  { disciplineId: 'hvac', code: 'H-001', desc: 'התקנת תעלת מיזוג אויר כולל תליות ובידוד בסיסי', unit: 'מטר', material: 180, labor: 120, engineering: 25, overhead: 30, supplier: 'ספק מיזוג', validity: '30 יום', notes: 'נתון דוגמה', defaultQty: 1 },
  { disciplineId: 'civil', code: 'C-001', desc: 'יציקת בטון C30 כולל טפסנות בסיסית וברזל', unit: 'מ״ק', material: 720, labor: 520, engineering: 95, overhead: 120, supplier: 'קבלן דוגמה', validity: '30 יום', notes: 'לא כולל בדיקות מעבדה', defaultQty: 1 },
  { disciplineId: 'instrumentation', code: 'I-001', desc: 'אספקה והתקנת משדר לחץ כולל חיווט ובדיקת לולאה', unit: 'יח׳', material: 1350, labor: 480, engineering: 220, overhead: 120, supplier: 'ספק מכשור', validity: '30 יום', notes: 'נתון דוגמה', defaultQty: 1 }
];
function inferBoqDiscipline(...values) { const t = values.join(' ').toLowerCase(); if (/כיבוי|אש|ספרינקלר|sprinkler|fire|מטף/.test(t)) return 'fire'; if (/מיזוג|אויר|אוויר|hvac|צילר|chiller|מפוח|duct/.test(t)) return 'hvac'; if (/מכשור|בקרה|instrument|control|plc|dcs|scada|חיישן|sensor|transmitter|משדר/.test(t)) return 'instrumentation'; if (/צנרת|צינור|pipe|valve|ברז|flange|אוגן/.test(t)) return 'piping'; if (/חשמל|כבל|לוח|ארון|cable|elect|panel/.test(t)) return 'electricity'; if (/אזרח|בטון|ברזל|חפירה|קבלן|civil|concrete|rebar|עפר/.test(t)) return 'civil'; return 'piping'; }
function mapBoqRow(row, i, forcedDiscipline = '') {
  const section = getVal(row, ['תאור הסעיף/פרק', 'תיאור הסעיף/פרק', 'תאור סעיף', 'תיאור סעיף', 'תיאור', 'תאור', 'Description']);
  const supplier = getVal(row, ['ספק', 'supplier', 'Vendor']);
  const sku = getVal(row, ['מק"ט', 'מק״ט', 'מקט', 'code', 'קוד']);
  const qty = num(getVal(row, ['כמות', 'Qty', 'Quantity'])) || 1;
  const unit = getVal(row, ['מק"ט', 'יחידה', 'יחידת מידה', 'Unit']) || 'יח׳';
  const unitPrice = num(getVal(row, ["מחיר ליח' לפני הנחה", 'מחיר ליח לפני הנחה', 'מחיר יחידה', 'מחיר ליחידה', 'Unit Price', 'price']));
  const totalVat = num(getVal(row, ['מחיר כולל מע"מ', 'מחיר כולל מעמ', 'Total Including VAT', 'סהכ', 'סה״כ']));
  const project = getVal(row, ['פרויקט', 'Project']);
  const projectDesc = getVal(row, ['תאור פרויקט', 'תיאור פרויקט', 'Project Description']);
  const resource = getVal(row, ['תאור משאב', 'תיאור משאב', 'Resource Description']);
  const quoteDate = getVal(row, ['תאריך הגשת הצעת מחיר', 'תאריך', 'validity', 'תוקף']);
  const discRaw = getVal(row, ['discipline', 'Discipline', 'דיסציפלינה', 'תחום', 'פרק']);
  const calc = unitPrice || (totalVat && qty ? totalVat / qty : 0);
  const desc = clean(section || resource || projectDesc || '');
  const disciplineId = forcedDiscipline ? normalizeBoqDisc(forcedDiscipline) : (discRaw ? normalizeBoqDisc(discRaw) : inferBoqDiscipline(section, resource, projectDesc));
  return { id: `${disciplineId}-${sku || i}-${Math.random().toString(36).slice(2)}`, disciplineId, code: String(sku || `XL-${i + 1}`), desc: desc || 'פריט ללא תיאור', unit, material: calc, labor: 0, engineering: 0, overhead: 0, supplier: String(supplier || ''), validity: String(quoteDate || ''), notes: String([project && `פרויקט: ${project}`, projectDesc && `תיאור פרויקט: ${projectDesc}`, resource && `משאב: ${resource}`, totalVat && `מחיר כולל מעמ: ${fmt(totalVat)}`].filter(Boolean).join(' | ')), defaultQty: qty, currency: getVal(row, ['מטבע חוזה', 'מטבע', 'Currency']) || 'ILS', totalIncludingVat: totalVat };
}
const itemTotal = i => num(i.material) + num(i.labor) + num(i.engineering) + num(i.overhead);

const SUP_DISCIPLINES = ['צנרת', 'חשמל', 'מכשור ובקרה', 'הנדסה אזרחית', 'מכונות וציוד', 'מתכת וקונסטרוקציה', 'בידוד וצבע', 'HVAC ומיזוג', 'בטיחות וכיבוי אש', 'לוגיסטיקה ושילוח', 'כימיקלים וחומרים', 'שירותי תכנון וייעוץ', 'הדרכות וכנסים', 'IT ותוכנה', 'כללי / אחר'];
const sampleSuppliers = [
  { id: 'demo-1', project: '00802', supplierNo: '51638', name: 'ספק צנרת לדוגמה בע״מ', description: 'אספקת צינורות, ברזים, אוגנים ואביזרי צנרת', discipline: 'צנרת', rating: 4, contact: '', phone: '', email: '', notes: '' },
  { id: 'demo-2', project: '00803', supplierNo: '7008', name: 'חשמל תעשייתי גליל', description: 'לוחות חשמל, כבלים, תעלות ובדיקות חשמל', discipline: 'חשמל', rating: 5, contact: '', phone: '', email: '', notes: '' },
  { id: 'demo-3', project: '00804', supplierNo: '51421', name: 'קבלן בטון ופיתוח', description: 'עבודות בטון, חפירה וקונסטרוקציה', discipline: 'הנדסה אזרחית', rating: 3, contact: '', phone: '', email: '', notes: '' }
];
function detectSupplierDiscipline(name = '', desc = '', source = '') { const manual = clean(source); if (manual) return manual; const text = `${name} ${desc}`.toLowerCase(); const rules = [['צנרת', ['צנרת', 'צינור', 'pipe', 'valve', 'ברז', 'אוגן']], ['חשמל', ['חשמל', 'כבל', 'לוח חשמל', 'electric', 'cable', 'abb', 'siemens']], ['מכשור ובקרה', ['מכשור', 'בקרה', 'instrument', 'control', 'sensor', 'transmitter']], ['הנדסה אזרחית', ['בטון', 'יציקה', 'קונסטרוקציה', 'חפירה', 'אזרחי', 'civil', 'concrete']], ['HVAC ומיזוג', ['מיזוג', 'אוורור', 'hvac', 'chiller', 'מפוח']], ['בטיחות וכיבוי אש', ['בטיחות', 'כיבוי', 'אש', 'sprinkler', 'fire']]]; for (const [d, ks] of rules) if (ks.some(w => text.includes(w.toLowerCase()))) return d; return 'כללי / אחר'; }
function findHeaderRow(sheet) {
  const range = XLSX.utils.decode_range(sheet['!ref'] || 'A1');
  for (let r = range.s.r; r <= Math.min(range.s.r + 10, range.e.r); r++) {
    let filled = 0;
    for (let c = range.s.c; c <= range.e.c; c++) {
      const cell = sheet[XLSX.utils.encode_cell({ r, c })];
      if (cell && String(cell.v || '').trim()) filled++;
    }
    if (filled >= 3) return r;
  }
  return 0;
}

function parseWorkbook(wb) {
  const targetSheet = wb.SheetNames.find(n => norm(n).includes('datasheet')) || wb.SheetNames[0];
  const ws = wb.Sheets[targetSheet];
  if (!ws) return [];
  const headerRow = findHeaderRow(ws);
  const rows = XLSX.utils.sheet_to_json(ws, { defval: '', raw: false, range: headerRow });
  const seen = new Set();
  return rows
    .filter(row => Object.values(row).some(v => String(v ?? '').trim() !== ''))
    .map((row, i) => {
      const supplierNo = clean(getVal(row, ['מספר ספק', "מס' ספק/קבלן", 'מס ספק/קבלן', 'מספר ספק/קבלן', 'מק״ט', 'מקט', 'supplier number', 'supplierNo']));
      const name = clean(getVal(row, ['שם ספק', 'שם ספק/קבלן', 'ספק', 'supplier', 'vendor', 'שם']));
      const field = clean(getVal(row, ['תחום', 'תחום עיסוק', 'דיסציפלינה', 'discipline', 'category', 'קטגוריה', 'field']));
      const description = clean(getVal(row, ['תחום פעילות מורחב', 'תאור הסעיף/פרק', 'תיאור הסעיף/פרק', 'תיאור', 'תאור', 'description', 'תאור משאב']));
      const address = clean(getVal(row, ['כתובת', 'address']));
      const cityCountry = clean(getVal(row, ['עיר ומדינה', 'עיר', 'city', 'cityCountry']));
      const zip = clean(getVal(row, ['מיקוד', 'zip', 'postal code']));
      const country = clean(getVal(row, ['ארץ', 'מדינה', 'country']));
      const phone = clean(getVal(row, ['מספר טלפון', 'טלפון', 'נייד', 'phone', 'mobile']));
      const fax = clean(getVal(row, ['פקס', 'fax']));
      const contact = clean(getVal(row, ['איש קשר', 'contact', 'contact person']));
      const certainty = clean(getVal(row, ['ודאות', 'certainty']));
      const notes = clean(getVal(row, ['הערות', 'notes', 'הערה']));

      const effectiveField = field || description || '';
      const effectiveDesc = description || '';

      if (supplierNo && seen.has(supplierNo)) return null;
      if (supplierNo) seen.add(supplierNo);

      if (!name && !supplierNo) return null;

      return {
        id: `${Date.now()}-${i}-${supplierNo || Math.random()}`,
        supplierNo,
        name: name || 'ספק ללא שם',
        field: effectiveField,
        description: effectiveDesc,
        discipline: detectSupplierDiscipline(name, effectiveDesc, effectiveField),
        address,
        cityCountry,
        zip,
        country,
        phone,
        fax,
        contact,
        certainty,
        notes,
        rating: 0,
        importedAt: new Date().toLocaleDateString('he-IL')
      };
    })
    .filter(Boolean);
}

function parseSuppliers(rows) { return rows.map((row, i) => { const name = clean(getVal(row, ['שם ספק/קבלן', 'שם ספק', 'ספק', 'supplier', 'vendor', 'שם'])); const desc = clean(getVal(row, ['תחום פעילות מורחב', 'תאור הסעיף/פרק', 'תיאור הסעיף/פרק', 'תיאור', 'תאור', 'description', 'תאור משאב'])); const source = clean(getVal(row, ['תחום', 'תחום עיסוק', 'דיסציפלינה', 'discipline', 'category', 'קטגוריה'])); const supplierNo = clean(getVal(row, ['מספר ספק', "מס' ספק/קבלן", 'מס ספק/קבלן', 'מספר ספק/קבלן', 'מק״ט', 'מקט', 'supplier number'])); const project = clean(getVal(row, ['פרויקט', 'project'])); const phone = clean(getVal(row, ['מספר טלפון', 'טלפון', 'נייד', 'phone', 'mobile'])); const email = clean(getVal(row, ['מייל', 'אימייל', 'דואל', 'email'])); const contact = clean(getVal(row, ['איש קשר', 'contact', 'contact person'])); const field = clean(getVal(row, ['תחום', 'תחום עיסוק', 'field'])); return { id: `${Date.now()}-${i}-${supplierNo || name || Math.random()}`, project, supplierNo, name: name || 'ספק ללא שם', description: desc, field: field || desc || '', discipline: detectSupplierDiscipline(name, desc, source), contact, phone, email, rating: 0, notes: '', importedAt: new Date().toLocaleDateString('he-IL') }; }).filter(s => s.name !== 'ספק ללא שם' || s.description || s.supplierNo || s.project); }

function Shell() { const [tab, setTab] = useState('boq'); return <div className="app" dir="rtl"><header className="top"><div className="brand"><img src={logo} /><div><span>GALIL GROUP</span><h1>מערכת הנדסה ורכש</h1><p>מחירון כתבי כמויות + מאגר ספקים במערכת אחת</p></div></div><nav><button className={tab === 'boq' ? 'active' : ''} onClick={() => setTab('boq')}><ClipboardList size={18} /> מחירון / BOQ</button><button className={tab === 'suppliers' ? 'active' : ''} onClick={() => setTab('suppliers')}><Users size={18} /> מאגר ספקים</button></nav></header>{tab === 'boq' ? <BoqApp /> : <SuppliersApp />}</div>; }

function BoqApp() {
  const inputRef = useRef(null); const reportRef = useRef(null);
  const [items, setItems] = useState(sampleItems); const [boqDisciplines, setBoqDisciplines] = useState(defaultBoqDisciplines); const [newDiscipline, setNewDiscipline] = useState(''); const [cart, setCart] = useState([]); const [query, setQuery] = useState(''); const [disc, setDisc] = useState('all'); const [status, setStatus] = useState('נטען מחירון דוגמה. אפשר להעלות Excel אמיתי.'); const [project, setProject] = useState({ name: 'אומדן פרויקט חדש', customer: 'לקוח / מחלקה', estimator: 'קבוצת גליל', currency: 'ILS' }); const [percent, setPercent] = useState({ management: 7, contingency: 12, profit: 10, discount: 0 }); const [result, setResult] = useState(false);
  useEffect(() => { const s = localStorage.getItem(BOQ_KEY); if (s) { try { const d = JSON.parse(s); setItems(d.items || sampleItems); setBoqDisciplines(d.boqDisciplines || defaultBoqDisciplines); setCart(d.cart || []); setProject(d.project || project); setPercent(d.percent || percent); } catch {} } }, []);
  const filtered = useMemo(() => items.filter(x => (disc === 'all' || x.disciplineId === disc) && `${x.desc} ${x.supplier} ${x.code} ${x.notes}`.toLowerCase().includes(query.toLowerCase())), [items, disc, query]);
  const add = item => setCart(prev => { const ex = prev.find(x => x.code === item.code && x.disciplineId === item.disciplineId && x.desc === item.desc); if (ex) return prev.map(x => x === ex ? { ...x, qty: x.qty + (item.defaultQty || 1) } : x); return [...prev, { ...item, id: `cart-${Date.now()}-${Math.random()}`, qty: item.defaultQty || 1 }]; });
  const upload = e => { const f = e.target.files?.[0]; if (!f) return; const r = new FileReader(); r.onload = ev => { try { const wb = XLSX.read(new Uint8Array(ev.target.result), { type: 'array', cellDates: true, raw: false }); const all = []; const newMap = makeDisciplineMapFromSheetNames(wb.SheetNames, defaultBoqDisciplines); wb.SheetNames.forEach(sheetName => { const ws = wb.Sheets[sheetName]; const rows = XLSX.utils.sheet_to_json(ws, { defval: '', raw: false }); rows.forEach((row, i) => { const mapped = mapBoqRow(row, all.length + i, sheetName); if (mapped.desc !== 'פריט ללא תיאור' || itemTotal(mapped) > 0) all.push(mapped); }); }); setBoqDisciplines(newMap); setItems(all.length ? all : sampleItems); setDisc('all'); setResult(false); setStatus(`נטענו ${all.length.toLocaleString('he-IL')} פריטים מתוך ${wb.SheetNames.length} לשוניות. כל לשונית סווגה כדיסציפלינה.`); } catch (err) { console.error(err); setStatus('שגיאה בקריאת הקובץ. בדוק שהקובץ הוא Excel תקין.'); } }; r.readAsArrayBuffer(f); e.target.value = ''; };
  const addDiscipline = () => { const name = clean(newDiscipline); if (!name) return; const id = normalizeBoqDisc(name); setBoqDisciplines(prev => ({ ...prev, [id]: { name, icon: FileText } })); setDisc(id); setNewDiscipline(''); };
  const deleteDiscipline = id => { if (disc === 'all' || !boqDisciplines[id]) return; setBoqDisciplines(prev => { const next = { ...prev }; delete next[id]; return next; }); setItems(prev => prev.filter(x => x.disciplineId !== id)); setCart(prev => prev.filter(x => x.disciplineId !== id)); setDisc('all'); };
  const totals = useMemo(() => { const material = cart.reduce((s, x) => s + num(x.material) * x.qty, 0); const labor = cart.reduce((s, x) => s + num(x.labor) * x.qty, 0); const eng = cart.reduce((s, x) => s + num(x.engineering) * x.qty, 0); const overhead = cart.reduce((s, x) => s + num(x.overhead) * x.qty, 0); const direct = material + labor + eng + overhead; const management = direct * percent.management / 100; const contingency = direct * percent.contingency / 100; const beforeDiscount = direct + management + contingency; const discount = beforeDiscount * percent.discount / 100; const profit = (beforeDiscount - discount) * percent.profit / 100; return { material, labor, eng, overhead, direct, management, contingency, discount, profit, total: beforeDiscount - discount + profit }; }, [cart, percent]);
  const byDisc = useMemo(() => Object.entries(boqDisciplines).map(([id, d]) => ({ id, name: d.name, total: cart.filter(x => x.disciplineId === id).reduce((s, x) => s + itemTotal(x) * x.qty, 0) })), [cart, boqDisciplines]);
  const save = () => { localStorage.setItem(BOQ_KEY, JSON.stringify({ items, boqDisciplines, cart, project, percent })); alert('הפרויקט נשמר'); };
  const reset = () => { setItems(sampleItems); setBoqDisciplines(defaultBoqDisciplines); setCart([]); setDisc('all'); setResult(false); setStatus('חזרת למחירון דוגמה'); };
  const exportCSV = () => { const data = cart.map(x => ({ 'דיסציפלינה': boqDisciplines[x.disciplineId]?.name || x.disciplineId, 'מק״ט': x.code, 'תיאור': x.desc, 'כמות': x.qty, 'יחידה': x.unit, 'עלות יחידה': itemTotal(x), 'סה״כ': itemTotal(x) * x.qty })); const ws = XLSX.utils.json_to_sheet(data); const wb = XLSX.utils.book_new(); XLSX.utils.book_append_sheet(wb, ws, 'BOQ'); XLSX.writeFile(wb, 'galil-boq.xlsx'); };
  const exportPDF = async () => { setResult(true); setTimeout(async () => { if (!reportRef.current) return; const canvas = await html2canvas(reportRef.current, { scale: 2, useCORS: true }); const img = canvas.toDataURL('image/png'); const pdf = new jsPDF('p', 'mm', 'a4'); const w = 210, h = canvas.height * w / canvas.width; pdf.addImage(img, 'PNG', 0, 0, w, h); pdf.save('galil-boq-report.pdf'); }, 100); };
  return <main className="layout"><section className="left"><div className="panel upload"><div><h2><UploadCloud /> העלאת מחירון Excel</h2><p>המערכת קוראת את כל הלשוניות בקובץ. שם כל לשונית הופך לדיסציפלינה באתר.</p><div className="status"><CheckCircle2 size={16} />{status}</div></div><div className="actions"><input ref={inputRef} type="file" accept=".xlsx,.xls,.xlsm,.csv" hidden onChange={upload} /><button onClick={() => inputRef.current.click()}><UploadCloud size={18} /> העלאת Excel</button><button onClick={reset}><RotateCcw size={18} /> דוגמה</button></div></div><div className="panel"><h2><Database /> פרטי פרויקט</h2><div className="formGrid"><Field label="שם פרויקט" value={project.name} onChange={v => setProject({ ...project, name: v })} /><Field label="לקוח / מחלקה" value={project.customer} onChange={v => setProject({ ...project, customer: v })} /><Field label="עורך אומדן" value={project.estimator} onChange={v => setProject({ ...project, estimator: v })} /><Field label="מטבע" value={project.currency} onChange={v => setProject({ ...project, currency: v })} /></div></div><div className="panel"><div className="catalogTop"><div><h2><Calculator /> מחירון פריטים</h2><p>חיפוש, סינון לפי דיסציפלינה והוספה לסל הפרויקט.</p></div><div className="search"><Search size={18} /><input value={query} onChange={e => setQuery(e.target.value)} placeholder="חיפוש פריט / ספק / מק״ט" /></div></div><div className="chips"><button className={disc === 'all' ? 'selected' : ''} onClick={() => setDisc('all')}>הכל</button>{Object.entries(boqDisciplines).map(([id, d]) => <button key={id} className={disc === id ? 'selected' : ''} onClick={() => setDisc(id)}>{d.name}</button>)}</div><div className="disciplineManage"><input value={newDiscipline} onChange={e => setNewDiscipline(e.target.value)} placeholder="הוסף דיסציפלינה ידנית" /><button onClick={addDiscipline}><Plus size={16} /> הוסף דיסציפלינה</button>{disc !== 'all' && <button className="dangerMini" onClick={() => deleteDiscipline(disc)}><Trash2 size={16} /> מחק דיסציפלינה נבחרת</button>}</div><div className="items">{filtered.map((it, idx) => { const Icon = boqDisciplines[it.disciplineId]?.icon || FileText; return <div className="item" key={`${it.id}-${idx}`}><div className="itemIcon"><Icon size={22} /></div><div className="itemText"><b>{it.desc}</b><span>{boqDisciplines[it.disciplineId]?.name || it.disciplineId} · מק״ט: {it.code} · ספק: {it.supplier || '-'} · יחידה: {it.unit}</span><small>{it.notes}</small></div><div className="price"><b>{fmt(itemTotal(it))}</b><button onClick={() => add(it)}><Plus size={16} /> הוסף</button></div></div>; })}</div></div></section><aside className="panel sticky"><h2><Calculator /> סל חישוב</h2>{cart.length === 0 ? <div className="empty">לא נבחרו פריטים</div> : <div className="cartList">{cart.map(x => <div className="cart" key={x.id}><button onClick={() => setCart(p => p.filter(z => z.id !== x.id))}><Trash2 size={16} /></button><b>{x.desc}</b><div><input type="number" value={x.qty} onChange={e => setCart(p => p.map(z => z.id === x.id ? { ...z, qty: num(e.target.value) } : z))} /><span>{x.unit}</span><strong>{fmt(itemTotal(x) * x.qty)}</strong></div></div>)}</div>}<div className="percentGrid"><Field label="ניהול %" type="number" value={percent.management} onChange={v => setPercent({ ...percent, management: num(v) })} /><Field label="בלתי צפוי %" type="number" value={percent.contingency} onChange={v => setPercent({ ...percent, contingency: num(v) })} /><Field label="רווח %" type="number" value={percent.profit} onChange={v => setPercent({ ...percent, profit: num(v) })} /><Field label="הנחה %" type="number" value={percent.discount} onChange={v => setPercent({ ...percent, discount: num(v) })} /></div><button className="calc" disabled={!cart.length} onClick={() => setResult(true)}>חשב פרויקט</button><div className="sideBtns"><button onClick={save}><Save size={16} /> שמור</button><button onClick={exportCSV}><Download size={16} /> Excel</button><button onClick={exportPDF}><Printer size={16} /> PDF</button></div><div className="totalBox"><span>סה״כ</span><b>{fmt(totals.total)}</b></div></aside>{result && <section className="report" ref={reportRef}><div className="reportBox"><div className="reportHead"><img src={logo} /><div><h2>דוח אומדן פרויקט</h2><p>{project.name} · {new Date().toLocaleDateString('he-IL')}</p></div><button onClick={exportPDF}>PDF</button></div><div className="kpis"><K title="חומרים" value={fmt(totals.material)} /><K title="עבודה" value={fmt(totals.labor)} /><K title="תכנון" value={fmt(totals.eng)} /><K title="סה״כ" value={fmt(totals.total)} big /></div><div className="reportGrid"><div className="box"><h3><BarChart3 /> לפי דיסציפלינה</h3>{byDisc.map(d => { const pct = totals.direct ? Math.round(d.total / totals.direct * 100) : 0; return <div className="bar" key={d.id}><span><b>{d.name}</b><b>{fmt(d.total)} · {pct}%</b></span><i><em style={{ width: pct + '%' }} /></i></div>; })}</div><div className="box"><h3>סיכום מסחרי</h3><Line l="עלות ישירה" v={totals.direct} /><Line l={`ניהול ${percent.management}%`} v={totals.management} /><Line l={`בלתי צפוי ${percent.contingency}%`} v={totals.contingency} /><Line l={`הנחה ${percent.discount}%`} v={-totals.discount} /><Line l={`רווח ${percent.profit}%`} v={totals.profit} /><div className="grand"><span>סה״כ אומדן</span><b>{fmt(totals.total)}</b></div></div></div><table><thead><tr><th>דיסציפלינה</th><th>מק״ט</th><th>תיאור</th><th>כמות</th><th>יחידה</th><th>סה״כ</th></tr></thead><tbody>{cart.map(x => <tr key={x.id}><td>{boqDisciplines[x.disciplineId]?.name || x.disciplineId}</td><td>{x.code}</td><td>{x.desc}</td><td>{x.qty}</td><td>{x.unit}</td><td>{fmt(itemTotal(x) * x.qty)}</td></tr>)}</tbody></table><div className="disclaimer">הנתונים מיועדים לאומדן ראשוני בלבד ודורשים אישור הנדסי/מסחרי לפני שימוש מחייב.</div></div></section>}</main>;
}

function SuppliersApp() {
  const [suppliers, setSuppliers] = useState(sampleSuppliers);
  const [query, setQuery] = useState('');
  const [disc, setDisc] = useState('הכל');
  const [message, setMessage] = useState('טוען מאגר ספקים...');
  const [loaded, setLoaded] = useState(false);

  useEffect(() => {
    const cached = localStorage.getItem(SUP_KEY);
    if (cached) {
      try {
        const data = JSON.parse(cached);
        if (Array.isArray(data) && data.length > 0) {
          setSuppliers(data);
          setMessage(`נטענו ${data.length.toLocaleString('he-IL')} ספקים מהמטמון.`);
          setLoaded(true);
          return;
        }
      } catch {}
    }
    fetch('/suppliers.json')
      .then(r => r.json())
      .then(data => {
        if (data.suppliers && data.suppliers.length > 0) {
          setSuppliers(data.suppliers);
          localStorage.setItem(SUP_KEY, JSON.stringify(data.suppliers));
          setMessage(`נטענו ${data.suppliers.length.toLocaleString('he-IL')} ספקים מהשרת.`);
        } else {
          setMessage('אפשר להעלות Excel בפורמט מאגר הספקים שלך.');
        }
      })
      .catch(() => {
        setMessage('אפשר להעלות Excel בפורמט מאגר הספקים שלך.');
      })
      .finally(() => setLoaded(true));
  }, []);

  useEffect(() => {
    if (loaded) localStorage.setItem(SUP_KEY, JSON.stringify(suppliers));
  }, [suppliers, loaded]);

  const upload = e => {
    const f = e.target.files?.[0];
    if (!f) return;
    const r = new FileReader();
    r.onload = ev => {
      try {
        const wb = XLSX.read(new Uint8Array(ev.target.result), { type: 'array', cellDates: true, raw: false });
        const parsed = parseWorkbook(wb);
        if (parsed.length > 0) {
          setSuppliers(parsed);
          setMessage(`נטענו ${parsed.length.toLocaleString('he-IL')} ספקים מתוך "${wb.SheetNames.join(', ')}".`);
        } else {
          const rows = XLSX.utils.sheet_to_json(wb.Sheets[wb.SheetNames[0]], { defval: '', raw: false });
          const fallback = parseSuppliers(rows);
          setSuppliers(fallback);
          setMessage(`נטענו ${fallback.length.toLocaleString('he-IL')} ספקים (fallback parser).`);
        }
      } catch (err) {
        console.error(err);
        setMessage('שגיאה בקריאת הקובץ');
      }
    };
    r.readAsArrayBuffer(f);
    e.target.value = '';
  };

  const filtered = useMemo(() => suppliers.filter(s =>
    (disc === 'הכל' || s.discipline === disc) &&
    `${s.name} ${s.description || ''} ${s.field || ''} ${s.supplierNo} ${s.project || ''} ${s.contact} ${s.phone || ''} ${s.fax || ''} ${s.address || ''} ${s.cityCountry || ''} ${s.notes || ''}`.toLowerCase().includes(query.toLowerCase())
  ), [suppliers, disc, query]);

  const stats = useMemo(() => SUP_DISCIPLINES.map(d => ({ name: d, count: suppliers.filter(s => s.discipline === d).length })).filter(x => x.count > 0), [suppliers]);
  const update = (id, patch) => setSuppliers(p => p.map(s => s.id === id ? { ...s, ...patch } : s));
  const del = id => setSuppliers(p => p.filter(s => s.id !== id));

  const exportExcel = () => {
    const data = suppliers.map(s => ({
      'מספר ספק': s.supplierNo,
      'שם ספק': s.name,
      'תחום': s.field || s.discipline,
      'תחום פעילות מורחב': s.description,
      'כתובת': s.address || '',
      'עיר ומדינה': s.cityCountry || '',
      'מיקוד': s.zip || '',
      'ארץ': s.country || '',
      'מספר טלפון': s.phone,
      'פקס': s.fax || '',
      'איש קשר': s.contact,
      'ודאות': s.certainty || '',
      'הערות': s.notes,
      'דירוג': s.rating
    }));
    const ws = XLSX.utils.json_to_sheet(data);
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, 'Suppliers');
    XLSX.writeFile(wb, 'galil-suppliers.xlsx');
  };

  return <main className="supPage"><section className="panel supHero"><div><h2><Users /> מאגר ספקים וקבלנים</h2><p>חיפוש לפי שם ספק, תחום, מספר ספק, איש קשר או תיאור. כולל סיווג אוטומטי, מחיקה, דירוג ותיקון ידני.</p><div className="status"><CheckCircle2 size={16} />{message}</div></div><div className="actions"><label className="fileBtn"><UploadCloud size={18} /> העלאת Excel<input type="file" accept=".xlsx,.xls,.csv" hidden onChange={upload} /></label><button onClick={exportExcel}><Download size={18} /> ייצוא Excel</button><button onClick={() => { setSuppliers(sampleSuppliers); setMessage('חזרת לנתוני דוגמה'); }}><RotateCcw size={18} /> דוגמה</button></div></section><section className="supplierControls panel"><div className="search big"><Search size={18} /><input value={query} onChange={e => setQuery(e.target.value)} placeholder="חיפוש ספק לפי שם / תחום / תיאור / איש קשר" /></div><select value={disc} onChange={e => setDisc(e.target.value)}><option>הכל</option>{SUP_DISCIPLINES.map(d => <option key={d}>{d}</option>)}</select></section><section className="stats">{stats.map(s => <div className="stat" key={s.name}><b>{s.count}</b><span>{s.name}</span></div>)}</section><section className="supplierGrid">{filtered.map(s => <article className="supplier" key={s.id}><div className="supplierTop"><div><span>{s.discipline}</span><h3>{s.name}</h3><p>מס׳ ספק: {s.supplierNo || '-'}{s.address ? ` · ${s.address}` : ''}{s.cityCountry ? ` · ${s.cityCountry}` : ''}</p></div><button className="danger" onClick={() => del(s.id)}><Trash2 size={16} /></button></div><p className="desc">{s.description || s.field || 'אין תיאור'}</p><div className="supplierMeta"><span>איש קשר: {s.contact || '-'}</span><span>טלפון: {s.phone || '-'}</span><span>פקס: {s.fax || '-'}</span>{s.certainty && <span>ודאות: {s.certainty}</span>}</div><div className="editRow"><label><Pencil size={14} /> סיווג ידני</label><select value={s.discipline} onChange={e => update(s.id, { discipline: e.target.value })}>{SUP_DISCIPLINES.map(d => <option key={d}>{d}</option>)}</select></div><div className="rating">{[1, 2, 3, 4, 5].map(n => <button key={n} onClick={() => update(s.id, { rating: n })} className={n <= s.rating ? 'on' : ''}><Star size={20} fill="currentColor" /></button>)}</div><textarea value={s.notes || ''} onChange={e => update(s.id, { notes: e.target.value })} placeholder="הערות על הספק" /></article>)}</section></main>;
}
function Field({ label, value, onChange, type = 'text' }) { return <label>{label}<input type={type} value={value} onChange={e => onChange(e.target.value)} /></label>; }
function K({ title, value, big }) { return <div className={'kpi ' + (big ? 'big' : '')}><span>{title}</span><b>{value}</b></div>; }
function Line({ l, v }) { return <div className="line"><span>{l}</span><b>{fmt(v)}</b></div>; }

createRoot(document.getElementById('root')).render(<Shell />);
