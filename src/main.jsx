import React, { useEffect, useMemo, useRef, useState, useCallback } from 'react';
import { createRoot } from 'react-dom/client';
import * as XLSX from 'xlsx';
import jsPDF from 'jspdf';
import html2canvas from 'html2canvas';
import { UploadCloud, Calculator, Search, Plus, Trash2, Save, Printer, Download, RotateCcw, Building2, Zap, Pipette, HardHat, BarChart3, FileText, Users, Star, Pencil, CheckCircle2, Database, ClipboardList, X, Copy, ChevronDown, ChevronUp, Paperclip, Clock, Send, ArrowUpDown, Eye, FolderPlus, Filter, Percent, Hash, Check, Square, CheckSquare, Layers, FileSearch, Loader, Phone, Mail, LayoutDashboard, TrendingUp, Activity, Package, UserPlus } from 'lucide-react';
import * as pdfjsLib from 'pdfjs-dist';
import './style.css';

pdfjsLib.GlobalWorkerOptions.workerSrc = new URL('pdfjs-dist/build/pdf.worker.min.mjs', import.meta.url).href;

const SUP_KEY = 'galil_suppliers_v9';
const PROJ_IDX = 'galil_proj_idx_v2';
const ACTIVE_PROJ = 'galil_active_proj_v2';
const FAV_KEY = 'galil_fav_v1';
const LEGACY_BOQ = 'galil_boq_v5_real_sheets_disciplines';
const VAT_RATE = 0.17;
const COMPANY_NAME = 'שם החברה';
const SYSTEM_TITLE = 'מערכת הנדסה ורכש';

const fmt = (v, cur = 'ILS') => new Intl.NumberFormat('he-IL', { style: 'currency', currency: cur === 'USD' ? 'USD' : cur === 'EUR' ? 'EUR' : 'ILS', maximumFractionDigits: 0 }).format(Number(v) || 0);
const num = v => { const n = Number(String(v ?? '').replace(/,/g, '').replace(/[₪$€]/g, '').trim()); return Number.isFinite(n) ? n : 0; };
const clean = v => String(v ?? '').replace(/\s+/g, ' ').trim();
const norm = v => String(v ?? '').replace(/[‎‏]/g, '').replace(/["׳'`''״]/g, '').replace(/[\.\/\-_:()\[\]]/g, '').replace(/\s+/g, '').trim().toLowerCase();
function getVal(row, aliases) { const n = {}; Object.keys(row || {}).forEach(k => n[norm(k)] = row[k]); for (const a of aliases) { const v = n[norm(a)]; if (v !== undefined && v !== null && String(v).trim() !== '') return v; } return ''; }
let _idCounter = 0;
function uid(prefix) { return `${prefix}-${++_idCounter}-${Math.random().toString(36).slice(2, 8)}`; }

const STATUS_OPTIONS = ['טיוטה', 'בבדיקה', 'מאושר', 'נשלח'];
const STATUS_COLORS = { 'טיוטה': '#94a3b8', 'בבדיקה': '#f59e0b', 'מאושר': '#10b981', 'נשלח': '#6366f1' };

const defaultBoqDisciplines = {
  piping: { name: 'צנרת', icon: 'Pipette' }, electricity: { name: 'חשמל', icon: 'Zap' },
  fire: { name: 'כיבוי אש', icon: 'FileText' }, hvac: { name: 'מיזוג אויר', icon: 'HardHat' },
  civil: { name: 'הנדסה אזרחית', icon: 'Building2' }, instrumentation: { name: 'מכשור', icon: 'Calculator' }
};
const ICON_MAP = { Pipette, Zap, FileText, HardHat, Building2, Calculator };
const getIcon = name => ICON_MAP[name] || FileText;

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
  sheetNames.forEach(name => { const id = normalizeBoqDisc(name); if (!map[id]) map[id] = { name: clean(name), icon: 'FileText' }; });
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
  const unit = getVal(row, ['יחידה', 'יחידת מידה', 'Unit']) || 'יח׳';
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
  return { id: uid('item'), disciplineId, code: String(sku || `XL-${i + 1}`), desc: desc || 'פריט ללא תיאור', unit, material: calc, labor: 0, engineering: 0, overhead: 0, supplier: String(supplier || ''), validity: String(quoteDate || ''), notes: String([project && `פרויקט: ${project}`, projectDesc && `תיאור פרויקט: ${projectDesc}`, resource && `משאב: ${resource}`, totalVat && `מחיר כולל מעמ: ${fmt(totalVat)}`].filter(Boolean).join(' | ')), defaultQty: qty, currency: getVal(row, ['מטבע חוזה', 'מטבע', 'Currency']) || 'ILS', totalIncludingVat: totalVat };
}
const itemTotal = i => num(i.material) + num(i.labor) + num(i.engineering) + num(i.overhead);

const SUP_DISCIPLINES = ['צנרת', 'חשמל', 'מכשור ובקרה', 'הנדסה אזרחית', 'מכונות וציוד', 'מתכת וקונסטרוקציה', 'בידוד וצבע', 'HVAC ומיזוג', 'בטיחות וכיבוי אש', 'לוגיסטיקה ושילוח', 'כימיקלים וחומרים', 'שירותי תכנון וייעוץ', 'הדרכות וכנסים', 'IT ותוכנה', 'כללי / אחר'];
const sampleSuppliers = [
  { id: 'demo-1', project: '00802', supplierNo: '51638', name: 'ספק צנרת לדוגמה בע״מ', description: 'אספקת צינורות, ברזים, אוגנים ואביזרי צנרת', discipline: 'צנרת', rating: 4, contact: '', phone: '', email: '', notes: '' },
  { id: 'demo-2', project: '00803', supplierNo: '7008', name: 'חשמל תעשייתי לדוגמה', description: 'לוחות חשמל, כבלים, תעלות ובדיקות חשמל', discipline: 'חשמל', rating: 5, contact: '', phone: '', email: '', notes: '' },
  { id: 'demo-3', project: '00804', supplierNo: '51421', name: 'קבלן בטון ופיתוח', description: 'עבודות בטון, חפירה וקונסטרוקציה', discipline: 'הנדסה אזרחית', rating: 3, contact: '', phone: '', email: '', notes: '' }
];
function detectSupplierDiscipline(name = '', desc = '', source = '') {
  const manual = clean(source);
  if (manual) { const exact = SUP_DISCIPLINES.find(d => norm(d) === norm(manual)); if (exact) return exact; const partial = SUP_DISCIPLINES.find(d => norm(d).includes(norm(manual)) || norm(manual).includes(norm(d))); if (partial) return partial; }
  const text = `${manual} ${name} ${desc}`.toLowerCase();
  const rules = [['צנרת', ['צנרת', 'צינור', 'pipe', 'valve', 'ברז', 'אוגן']], ['חשמל', ['חשמל', 'כבל', 'לוח חשמל', 'electric', 'cable', 'abb', 'siemens']], ['מכשור ובקרה', ['מכשור', 'בקרה', 'instrument', 'control', 'sensor', 'transmitter']], ['הנדסה אזרחית', ['בטון', 'יציקה', 'קונסטרוקציה', 'חפירה', 'אזרחי', 'civil', 'concrete']], ['מכונות וציוד', ['מכונות', 'ציוד', 'משאבה', 'pump', 'compressor', 'מדחס']], ['מתכת וקונסטרוקציה', ['מתכת', 'קונסטרוקציה', 'steel', 'פלדה', 'ריתוך', 'weld']], ['בידוד וצבע', ['בידוד', 'צבע', 'insulation', 'paint', 'coating']], ['HVAC ומיזוג', ['מיזוג', 'אוורור', 'hvac', 'chiller', 'מפוח']], ['בטיחות וכיבוי אש', ['בטיחות', 'כיבוי', 'אש', 'sprinkler', 'fire']], ['לוגיסטיקה ושילוח', ['לוגיסטיקה', 'שילוח', 'הובלה', 'logistics', 'shipping']], ['כימיקלים וחומרים', ['כימיקל', 'חומר', 'chemical']], ['שירותי תכנון וייעוץ', ['תכנון', 'ייעוץ', 'engineering', 'consulting']], ['הדרכות וכנסים', ['הדרכה', 'כנס', 'training']], ['IT ותוכנה', ['it', 'תוכנה', 'software', 'מחשב']]];
  for (const [d, ks] of rules) if (ks.some(w => text.includes(w.toLowerCase()))) return d; return 'כללי / אחר';
}
function findHeaderRow(sheet) { const range = XLSX.utils.decode_range(sheet['!ref'] || 'A1'); for (let r = range.s.r; r <= Math.min(range.s.r + 10, range.e.r); r++) { let filled = 0; for (let c = range.s.c; c <= range.e.c; c++) { const cell = sheet[XLSX.utils.encode_cell({ r, c })]; if (cell && String(cell.v || '').trim()) filled++; } if (filled >= 3) return r; } return 0; }
function parseWorkbook(wb) {
  const targetSheet = wb.SheetNames.find(n => norm(n).includes('datasheet')) || wb.SheetNames[0];
  const ws = wb.Sheets[targetSheet]; if (!ws) return [];
  const headerRow = findHeaderRow(ws);
  const rows = XLSX.utils.sheet_to_json(ws, { defval: '', raw: false, range: headerRow });
  const seen = new Set();
  return rows.filter(row => Object.values(row).some(v => String(v ?? '').trim() !== '')).map((row, i) => {
    const supplierNo = clean(getVal(row, ['מספר ספק', "מס' ספק/קבלן", 'מס ספק/קבלן', 'מספר ספק/קבלן', 'מק״ט', 'מקט', 'supplier number', 'supplierNo']));
    const name = clean(getVal(row, ['שם ספק', 'שם ספק/קבלן', 'ספק', 'supplier', 'vendor', 'שם']));
    const field = clean(getVal(row, ['תחום', 'תחום עיסוק', 'דיסציפלינה', 'discipline', 'category', 'קטגוריה', 'field']));
    const description = clean(getVal(row, ['תחום פעילות מורחב', 'תאור הסעיף/פרק', 'תיאור הסעיף/פרק', 'תיאור', 'תאור', 'description', 'תאור משאב']));
    const address = clean(getVal(row, ['כתובת', 'address'])); const cityCountry = clean(getVal(row, ['עיר ומדינה', 'עיר', 'city', 'cityCountry']));
    const zip = clean(getVal(row, ['מיקוד', 'zip', 'postal code'])); const country = clean(getVal(row, ['ארץ', 'מדינה', 'country']));
    const phone = clean(getVal(row, ['מספר טלפון', 'טלפון', 'נייד', 'phone', 'mobile'])); const fax = clean(getVal(row, ['פקס', 'fax']));
    const contact = clean(getVal(row, ['איש קשר', 'contact', 'contact person']));
    const email = clean(getVal(row, ['מייל', 'אימייל', 'דואל', 'email', 'דוא"ל', 'mail', 'e-mail']));
    const certainty = clean(getVal(row, ['ודאות', 'certainty'])); const notes = clean(getVal(row, ['הערות', 'notes', 'הערה']));
    if (supplierNo && seen.has(supplierNo)) return null; if (supplierNo) seen.add(supplierNo); if (!name && !supplierNo) return null;
    return { id: uid('sup'), supplierNo, name: name || 'ספק ללא שם', field: field || description || '', description: description || '', discipline: detectSupplierDiscipline(name, description || '', field || description || ''), address, cityCountry, zip, country, phone, fax, email, contact, certainty, notes, rating: 0, importedAt: new Date().toLocaleDateString('he-IL') };
  }).filter(Boolean);
}
function parseSuppliers(rows) { return rows.map((row, i) => { const name = clean(getVal(row, ['שם ספק/קבלן', 'שם ספק', 'ספק', 'supplier', 'vendor', 'שם'])); const desc = clean(getVal(row, ['תחום פעילות מורחב', 'תאור הסעיף/פרק', 'תיאור הסעיף/פרק', 'תיאור', 'תאור', 'description', 'תאור משאב'])); const source = clean(getVal(row, ['תחום', 'תחום עיסוק', 'דיסציפלינה', 'discipline', 'category', 'קטגוריה'])); const supplierNo = clean(getVal(row, ['מספר ספק', "מס' ספק/קבלן", 'מס ספק/קבלן', 'מספר ספק/קבלן', 'מק״ט', 'מקט', 'supplier number'])); const project = clean(getVal(row, ['פרויקט', 'project'])); const phone = clean(getVal(row, ['מספר טלפון', 'טלפון', 'נייד', 'phone', 'mobile'])); const email = clean(getVal(row, ['מייל', 'אימייל', 'דואל', 'email'])); const contact = clean(getVal(row, ['איש קשר', 'contact', 'contact person'])); const field = clean(getVal(row, ['תחום', 'תחום עיסוק', 'field'])); return { id: uid('ps'), project, supplierNo, name: name || 'ספק ללא שם', description: desc, field: field || desc || '', discipline: detectSupplierDiscipline(name, desc, source), contact, phone, email, rating: 0, notes: '', importedAt: new Date().toLocaleDateString('he-IL') }; }).filter(s => s.name !== 'ספק ללא שם' || s.description || s.supplierNo || s.project); }

/* ====== PDF PARSING ====== */
async function extractPdfText(arrayBuffer) {
  const pdf = await pdfjsLib.getDocument({ data: arrayBuffer }).promise;
  const pages = [];
  for (let p = 1; p <= pdf.numPages; p++) {
    const page = await pdf.getPage(p);
    const content = await page.getTextContent();
    const items = content.items.filter(i => i.str?.trim());
    if (!items.length) continue;
    const lines = [];
    let currentLine = []; let lastY = null;
    items.sort((a, b) => {
      const dy = b.transform[5] - a.transform[5];
      if (Math.abs(dy) > 3) return dy;
      return a.transform[4] - b.transform[4];
    });
    for (const item of items) {
      const y = Math.round(item.transform[5]);
      if (lastY !== null && Math.abs(y - lastY) > 3) { if (currentLine.length) lines.push(currentLine); currentLine = []; }
      currentLine.push({ text: item.str.trim(), x: Math.round(item.transform[4]), w: item.width || 0 });
      lastY = y;
    }
    if (currentLine.length) lines.push(currentLine);
    pages.push(lines);
  }
  return pages;
}

function detectPdfColumns(pages) {
  const allLines = pages.flat();
  if (allLines.length < 3) return null;
  const xPositions = {};
  allLines.forEach(line => line.forEach(cell => {
    const bucket = Math.round(cell.x / 15) * 15;
    xPositions[bucket] = (xPositions[bucket] || 0) + 1;
  }));
  const cols = Object.entries(xPositions).filter(([, count]) => count > allLines.length * 0.15).map(([x]) => Number(x)).sort((a, b) => a - b);
  if (cols.length < 2) return null;
  return cols;
}

function pdfLinesToCells(line, cols) {
  if (!cols) return line.map(c => c.text);
  const cells = cols.map(() => '');
  for (const item of line) {
    let best = 0, bestDist = Infinity;
    cols.forEach((cx, i) => { const d = Math.abs(item.x - cx); if (d < bestDist) { bestDist = d; best = i; } });
    cells[best] = (cells[best] ? cells[best] + ' ' : '') + item.text;
  }
  return cells;
}

const NUM_RE = /^[\d,\.]+$/;
const PRICE_RE = /^[₪$€]?\s*[\d,\.]+\s*[₪$€]?$/;
function classifyPdfCells(cells) {
  let descIdx = -1, qtyIdx = -1, priceIdx = -1, unitIdx = -1, codeIdx = -1;
  const unitWords = ['יח', 'מטר', 'מ"א', 'מ״א', 'מ"ר', 'מ״ר', 'מק', 'טון', 'ק"ג', 'קג', 'ליטר', 'סט', 'קומפ', 'unit', 'pcs', 'm', 'kg', 'ton', 'set', 'lot'];
  cells.forEach((c, i) => {
    const v = clean(c).toLowerCase();
    const n = norm(c);
    if (['תיאור', 'תאור', 'description', 'פריט', 'סעיף', 'תאורהסעיף', 'תיאורהסעיף'].some(w => n.includes(norm(w)))) descIdx = i;
    else if (['כמות', 'qty', 'quantity', 'כמ'].some(w => n.includes(norm(w)))) qtyIdx = i;
    else if (['מחיר', 'price', 'עלות', 'סהכ', 'סכום', 'total', 'amount'].some(w => n.includes(norm(w)))) priceIdx = i;
    else if (['יחידה', 'יחמידה', 'unit', 'יח'].some(w => n.includes(norm(w)))) unitIdx = i;
    else if (['מקט', 'מקטט', 'code', 'קוד', 'סעיף', 'פריט', 'item'].some(w => n === norm(w))) codeIdx = i;
  });
  if (descIdx === -1) {
    let longest = -1, longestIdx = -1;
    cells.forEach((c, i) => { if (c.length > longest && i !== qtyIdx && i !== priceIdx && i !== unitIdx && i !== codeIdx) { longest = c.length; longestIdx = i; } });
    if (longestIdx >= 0) descIdx = longestIdx;
  }
  return { descIdx, qtyIdx, priceIdx, unitIdx, codeIdx };
}

function parsePdfRows(pages) {
  const cols = detectPdfColumns(pages);
  const allLines = pages.flat();
  if (allLines.length < 2) return { items: [], rawText: allLines.map(l => l.map(c => c.text).join(' ')).join('\n') };

  const gridLines = allLines.map(line => pdfLinesToCells(line, cols));
  if (gridLines.length < 2) return { items: [], rawText: gridLines.map(r => r.join(' | ')).join('\n') };

  let headerIdx = -1;
  const mapping = { descIdx: -1, qtyIdx: -1, priceIdx: -1, unitIdx: -1, codeIdx: -1 };
  for (let i = 0; i < Math.min(gridLines.length, 8); i++) {
    const m = classifyPdfCells(gridLines[i]);
    if (m.descIdx >= 0 && (m.priceIdx >= 0 || m.qtyIdx >= 0)) { Object.assign(mapping, m); headerIdx = i; break; }
  }
  if (headerIdx === -1) {
    for (let i = 0; i < Math.min(gridLines.length, 5); i++) {
      const hasText = gridLines[i].some(c => c.length > 10 && !NUM_RE.test(c.trim()));
      const hasNum = gridLines[i].some(c => PRICE_RE.test(c.trim()));
      if (hasText && hasNum) {
        let longest = -1, longestIdx = -1, numIdx = -1;
        gridLines[i].forEach((c, j) => { if (c.length > longest && !NUM_RE.test(c.trim())) { longest = c.length; longestIdx = j; } if (PRICE_RE.test(c.trim()) && numIdx === -1) numIdx = j; });
        mapping.descIdx = longestIdx; mapping.priceIdx = numIdx; headerIdx = i - 1;
        break;
      }
    }
  }

  const dataStart = headerIdx + 1;
  const items = [];
  for (let i = dataStart; i < gridLines.length; i++) {
    const row = gridLines[i];
    const desc = mapping.descIdx >= 0 ? clean(row[mapping.descIdx] || '') : row.filter((c, j) => j !== mapping.qtyIdx && j !== mapping.priceIdx).map(c => c.trim()).filter(Boolean).join(' ');
    if (!desc || desc.length < 2) continue;
    const priceRaw = mapping.priceIdx >= 0 ? row[mapping.priceIdx] : '';
    const qtyRaw = mapping.qtyIdx >= 0 ? row[mapping.qtyIdx] : '';
    const unitRaw = mapping.unitIdx >= 0 ? row[mapping.unitIdx] : '';
    const codeRaw = mapping.codeIdx >= 0 ? row[mapping.codeIdx] : '';
    const price = num(priceRaw);
    const qty = num(qtyRaw) || 1;
    if (/^(סה[״"]?כ|total|subtotal|sum|סיכום)$/i.test(norm(desc))) continue;
    if (!price && !desc.match(/[֐-׿a-zA-Z]{3,}/)) continue;

    items.push({
      id: uid('pdf'),
      disciplineId: inferBoqDiscipline(desc),
      code: clean(codeRaw) || `PDF-${items.length + 1}`,
      desc,
      unit: clean(unitRaw) || 'יח׳',
      material: price,
      labor: 0, engineering: 0, overhead: 0,
      supplier: '', validity: '',
      notes: 'יובא מ-PDF',
      defaultQty: qty,
      currency: 'ILS',
      totalIncludingVat: 0
    });
  }

  const rawText = gridLines.map(r => r.join(' | ')).join('\n');
  return { items, rawText };
}

/* ====== PROJECT MANAGEMENT ====== */
const projKey = id => `galil_proj_${id}`;
const loadIdx = () => { try { return JSON.parse(localStorage.getItem(PROJ_IDX)) || []; } catch { return []; } };
const saveIdx = list => localStorage.setItem(PROJ_IDX, JSON.stringify(list));
const loadProj = id => { try { return JSON.parse(localStorage.getItem(projKey(id))); } catch { return null; } };
const saveProj = (id, data) => localStorage.setItem(projKey(id), JSON.stringify(data));
const delProj = id => localStorage.removeItem(projKey(id));

const DISC_TO_SUP = { piping: 'צנרת', electricity: 'חשמל', instrumentation: 'מכשור ובקרה', civil: 'הנדסה אזרחית', hvac: 'HVAC ומיזוג', fire: 'בטיחות וכיבוי אש' };

/* ====== SHELL ====== */
function Shell() {
  const [tab, setTab] = useState('dashboard');
  return <div className="app" dir="rtl">
    <header className="top"><div className="brand"><div className="logoPlaceholder"><Building2 size={32} /></div><div><span>{COMPANY_NAME}</span><h1>{SYSTEM_TITLE}</h1><p>מחירון כתבי כמויות + מאגר ספקים במערכת אחת</p></div></div>
    <nav><button className={tab === 'dashboard' ? 'active' : ''} onClick={() => setTab('dashboard')}><LayoutDashboard size={18} /> דשבורד</button><button className={tab === 'boq' ? 'active' : ''} onClick={() => setTab('boq')}><ClipboardList size={18} /> מחירון / BOQ</button><button className={tab === 'suppliers' ? 'active' : ''} onClick={() => setTab('suppliers')}><Users size={18} /> מאגר ספקים</button></nav></header>
    {tab === 'dashboard' ? <Dashboard onNavigate={setTab} /> : tab === 'boq' ? <BoqApp /> : <SuppliersApp />}
    <footer className="appFooter">© {new Date().getFullYear()} {COMPANY_NAME} · {SYSTEM_TITLE}</footer>
  </div>;
}

/* ====== DASHBOARD ====== */
function Dashboard({ onNavigate }) {
  const [projects, setProjects] = useState([]);
  const [suppliers, setSuppliers] = useState([]);
  useEffect(() => {
    const idx = loadIdx();
    setProjects(idx);
    try { const s = JSON.parse(localStorage.getItem(SUP_KEY)) || []; setSuppliers(s); } catch {}
  }, []);

  const projDetails = useMemo(() => projects.map(p => {
    const d = loadProj(p.id);
    const cartLen = d?.cart?.length || 0;
    const itemsLen = d?.items?.length || 0;
    let total = 0;
    if (d?.cart) d.cart.forEach(x => { total += (num(x.material) + num(x.labor) + num(x.engineering) + num(x.overhead)) * (x.qty || 1); });
    return { ...p, cartLen, itemsLen, total, currency: d?.project?.currency || 'ILS', customer: d?.project?.customer || '', estimator: d?.project?.estimator || '' };
  }), [projects]);

  const totalProjects = projects.length;
  const totalSuppliers = suppliers.length;
  const totalEstimate = projDetails.reduce((s, p) => s + p.total, 0);
  const activeProjects = projDetails.filter(p => p.status === 'בבדיקה' || p.status === 'מאושר').length;
  const topDisciplines = useMemo(() => {
    const c = {};
    suppliers.forEach(s => { c[s.discipline] = (c[s.discipline] || 0) + 1; });
    return Object.entries(c).sort((a, b) => b[1] - a[1]).slice(0, 6);
  }, [suppliers]);
  const recentProjects = projDetails.sort((a, b) => (b.updatedAt || '').localeCompare(a.updatedAt || '')).slice(0, 5);

  return <main className="dashPage">
    <section className="dashWelcome">
      <h2><Activity size={28} /> ברוכים הבאים ל{SYSTEM_TITLE}</h2>
      <p>סקירה כללית של פרויקטים, ספקים ונתונים עדכניים</p>
    </section>

    <section className="dashStats">
      <div className="dashStat" onClick={() => onNavigate('boq')}><Package size={28} /><b>{totalProjects}</b><span>פרויקטים</span></div>
      <div className="dashStat" onClick={() => onNavigate('suppliers')}><Users size={28} /><b>{totalSuppliers}</b><span>ספקים במאגר</span></div>
      <div className="dashStat"><TrendingUp size={28} /><b>{fmt(totalEstimate)}</b><span>סה״כ אומדנים</span></div>
      <div className="dashStat"><Activity size={28} /><b>{activeProjects}</b><span>פרויקטים פעילים</span></div>
    </section>

    <section className="dashGrid">
      <div className="dashCard">
        <h3><ClipboardList size={20} /> פרויקטים אחרונים</h3>
        {recentProjects.length === 0 ? <p className="dashEmpty">אין פרויקטים עדיין. <button onClick={() => onNavigate('boq')}>צור פרויקט ראשון</button></p> :
        <div className="dashTable">
          <table><thead><tr><th>שם</th><th>לקוח</th><th>סטטוס</th><th>פריטים בסל</th><th>סה״כ</th><th>עדכון</th></tr></thead>
          <tbody>{recentProjects.map(p => <tr key={p.id}>
            <td><b>{p.name}</b></td>
            <td>{p.customer || '-'}</td>
            <td><span className="dashStatus" style={{ background: STATUS_COLORS[p.status] || '#94a3b8' }}>{p.status || 'טיוטה'}</span></td>
            <td>{p.cartLen}</td>
            <td>{fmt(p.total, p.currency)}</td>
            <td>{p.updatedAt ? new Date(p.updatedAt).toLocaleDateString('he-IL') : '-'}</td>
          </tr>)}</tbody></table>
        </div>}
        <button className="dashLink" onClick={() => onNavigate('boq')}><ClipboardList size={16} /> עבור למחירון ← </button>
      </div>

      <div className="dashCard">
        <h3><Users size={20} /> ספקים לפי תחום</h3>
        {topDisciplines.length === 0 ? <p className="dashEmpty">אין ספקים. <button onClick={() => onNavigate('suppliers')}>העלה מאגר ספקים</button></p> :
        <div className="dashBars">{topDisciplines.map(([name, count]) => {
          const pct = totalSuppliers ? Math.round(count / totalSuppliers * 100) : 0;
          return <div className="bar" key={name}><span><b>{name}</b><b>{count} ({pct}%)</b></span><i><em style={{ width: pct + '%' }} /></i></div>;
        })}</div>}
        <button className="dashLink" onClick={() => onNavigate('suppliers')}><Users size={16} /> עבור למאגר ספקים ← </button>
      </div>
    </section>

    <section className="dashQuickActions">
      <h3>פעולות מהירות</h3>
      <div className="dashActions">
        <button onClick={() => onNavigate('boq')}><Calculator size={20} /> בנה אומדן חדש</button>
        <button onClick={() => onNavigate('suppliers')}><UserPlus size={20} /> הוסף ספק</button>
        <button onClick={() => onNavigate('boq')}><UploadCloud size={20} /> העלה Excel</button>
        <button onClick={() => onNavigate('boq')}><FileSearch size={20} /> העלה PDF</button>
      </div>
    </section>
  </main>;
}

/* ====== BOQ APP ====== */
function BoqApp() {
  const inputRef = useRef(null); const reportRef = useRef(null); const searchRef = useRef(null); const attachRef = useRef(null); const saveManualRef = useRef(null);
  const [items, setItems] = useState(sampleItems);
  const [boqDisciplines, setBoqDisciplines] = useState(defaultBoqDisciplines);
  const [newDiscipline, setNewDiscipline] = useState('');
  const [cart, setCart] = useState([]);
  const [query, setQuery] = useState('');
  const [disc, setDisc] = useState('all');
  const [status, setStatus] = useState('נטען מחירון דוגמה.');
  const [project, setProject] = useState({ name: 'אומדן פרויקט חדש', customer: 'לקוח / מחלקה', estimator: '', currency: 'ILS', status: 'טיוטה', exchangeRate: 1 });
  const [percent, setPercent] = useState({ management: 7, contingency: 12, profit: 10, discount: 0 });
  const [result, setResult] = useState(false);
  const [loaded, setLoaded] = useState(false);

  // #1 Multi-project
  const [projects, setProjects] = useState([]);
  const [activeId, setActiveId] = useState(null);

  // #4 Checkbox multi-select
  const [selected, setSelected] = useState(new Set());

  // #7 Sort
  const [sortMode, setSortMode] = useState('default');

  // #8 Favorites
  const [favorites, setFavorites] = useState(new Set());
  const [favOnly, setFavOnly] = useState(false);

  // #3 VAT
  const [showVat, setShowVat] = useState(false);

  // #6 Inline cart edit
  const [expandedCart, setExpandedCart] = useState(new Set());

  // #9 Discipline markup
  const [discMarkup, setDiscMarkup] = useState({});
  const [showDiscMarkup, setShowDiscMarkup] = useState(false);

  // #5 Add manual item
  const [showAddForm, setShowAddForm] = useState(false);
  const [newItem, setNewItem] = useState({ desc: '', code: '', unit: 'יח׳', material: 0, labor: 0, engineering: 0, overhead: 0, disciplineId: 'piping', supplier: '' });

  // #14 Attachments + PDF parsing
  const [attachments, setAttachments] = useState([]);
  const [pdfPreview, setPdfPreview] = useState(null); // { items: [], rawText: '', fileName: '' }
  const [pdfParsing, setPdfParsing] = useState(false);

  // #11 Version history
  const [versions, setVersions] = useState([]);
  const [showCompare, setShowCompare] = useState(false);

  // #12 Supplier hints
  const [supplierHints, setSupplierHints] = useState([]);

  // Gather project state into saveable object
  const getState = useCallback(() => ({
    items, boqDisciplines, cart, project, percent, discMarkup,
    attachments: attachments.map(a => ({ ...a, dataUrl: a.dataUrl?.length < 2000000 ? a.dataUrl : null })),
    versions
  }), [items, boqDisciplines, cart, project, percent, discMarkup, attachments, versions]);

  const applyState = useCallback((d) => {
    if (!d) return;
    setItems(d.items || sampleItems);
    setBoqDisciplines(d.boqDisciplines || defaultBoqDisciplines);
    setCart(d.cart || []);
    setProject(d.project || { name: 'פרויקט חדש', customer: '', estimator: '', currency: 'ILS', status: 'טיוטה', exchangeRate: 1 });
    setPercent(d.percent || { management: 7, contingency: 12, profit: 10, discount: 0 });
    setDiscMarkup(d.discMarkup || {});
    setAttachments(d.attachments || []);
    setVersions(d.versions || []);
    setResult(false); setDisc('all'); setQuery(''); setSelected(new Set()); setExpandedCart(new Set());
  }, []);

  // Init
  useEffect(() => {
    const idx = loadIdx();
    setProjects(idx);
    const aid = localStorage.getItem(ACTIVE_PROJ);
    if (aid) { const d = loadProj(aid); if (d) { applyState(d); setActiveId(aid); } else if (idx.length) { const d2 = loadProj(idx[0].id); applyState(d2); setActiveId(idx[0].id); } }
    else if (idx.length) { const d2 = loadProj(idx[0].id); applyState(d2); setActiveId(idx[0].id); }
    else {
      const legacy = localStorage.getItem(LEGACY_BOQ);
      if (legacy) { try { const d = JSON.parse(legacy); setItems(d.items || sampleItems); setBoqDisciplines(d.boqDisciplines || defaultBoqDisciplines); setCart(d.cart || []); setProject(d.project || project); setPercent(d.percent || percent); } catch {} }
      const id = uid('proj');
      setActiveId(id);
      const meta = { id, name: 'פרויקט ראשון', status: 'טיוטה', updatedAt: new Date().toISOString() };
      setProjects([meta]); saveIdx([meta]);
    }
    try { setFavorites(new Set(JSON.parse(localStorage.getItem(FAV_KEY)) || [])); } catch {}
    try { setSupplierHints(JSON.parse(localStorage.getItem(SUP_KEY)) || []); } catch {}
    setLoaded(true);
  }, []);

  // Auto-save
  useEffect(() => {
    if (!loaded || !activeId) return;
    const timer = setTimeout(() => {
      saveProj(activeId, getState());
      const idx = loadIdx().map(p => p.id === activeId ? { ...p, name: project.name, status: project.status, updatedAt: new Date().toISOString() } : p);
      if (!idx.find(p => p.id === activeId)) idx.push({ id: activeId, name: project.name, status: project.status, updatedAt: new Date().toISOString() });
      saveIdx(idx); setProjects(idx);
      localStorage.setItem(ACTIVE_PROJ, activeId);
    }, 800);
    return () => clearTimeout(timer);
  }, [items, boqDisciplines, cart, project, percent, discMarkup, attachments, versions, loaded, activeId]);

  // Save favorites
  useEffect(() => { if (loaded) localStorage.setItem(FAV_KEY, JSON.stringify([...favorites])); }, [favorites, loaded]);

  // #15 Keyboard shortcuts
  useEffect(() => {
    const h = e => {
      if ((e.ctrlKey || e.metaKey) && e.key === 's') { e.preventDefault(); saveManualRef.current(); }
      if ((e.ctrlKey || e.metaKey) && e.key === 'f') { e.preventDefault(); searchRef.current?.focus(); }
      if (e.key === 'Escape') { setQuery(''); setShowAddForm(false); setShowCompare(false); }
    };
    window.addEventListener('keydown', h); return () => window.removeEventListener('keydown', h);
  }, []);

  // Filtering & sorting
  const filtered = useMemo(() => {
    let arr = items.filter(x => (disc === 'all' || x.disciplineId === disc) && `${x.desc} ${x.supplier} ${x.code} ${x.notes}`.toLowerCase().includes(query.toLowerCase()));
    if (favOnly) arr = arr.filter(x => favorites.has(x.code));
    if (sortMode === 'price-asc') arr = [...arr].sort((a, b) => itemTotal(a) - itemTotal(b));
    else if (sortMode === 'price-desc') arr = [...arr].sort((a, b) => itemTotal(b) - itemTotal(a));
    else if (sortMode === 'name') arr = [...arr].sort((a, b) => a.desc.localeCompare(b.desc));
    return arr;
  }, [items, disc, query, favOnly, favorites, sortMode]);

  const add = item => setCart(prev => { const ex = prev.find(x => x.code === item.code && x.disciplineId === item.disciplineId && x.desc === item.desc); if (ex) return prev.map(x => x === ex ? { ...x, qty: x.qty + (item.defaultQty || 1) } : x); return [...prev, { ...item, id: uid('cart'), qty: item.defaultQty || 1, cartNote: '' }]; });
  const addSelected = () => { const toAdd = items.filter(x => selected.has(x.id || x.code)); toAdd.forEach(add); setSelected(new Set()); };
  const toggleSelect = id => setSelected(prev => { const n = new Set(prev); if (n.has(id)) n.delete(id); else n.add(id); return n; });
  const selectAll = () => { if (selected.size === filtered.length) setSelected(new Set()); else setSelected(new Set(filtered.map(x => x.id || x.code))); };

  const upload = e => { const f = e.target.files?.[0]; if (!f) return; const r = new FileReader(); r.onload = ev => { try { const wb = XLSX.read(new Uint8Array(ev.target.result), { type: 'array', cellDates: true, raw: false }); const all = []; const newMap = makeDisciplineMapFromSheetNames(wb.SheetNames, defaultBoqDisciplines); wb.SheetNames.forEach(sheetName => { const ws = wb.Sheets[sheetName]; const rows = XLSX.utils.sheet_to_json(ws, { defval: '', raw: false }); rows.forEach((row, i) => { const mapped = mapBoqRow(row, all.length + i, sheetName); if (mapped.desc !== 'פריט ללא תיאור' || itemTotal(mapped) > 0) all.push(mapped); }); }); setBoqDisciplines(newMap); setItems(all.length ? all : sampleItems); setDisc('all'); setResult(false); setSelected(new Set()); setStatus(`נטענו ${all.length.toLocaleString('he-IL')} פריטים מתוך ${wb.SheetNames.length} לשוניות.`); } catch (err) { console.error(err); setStatus('שגיאה בקריאת הקובץ.'); } }; r.readAsArrayBuffer(f); e.target.value = ''; };
  const addDiscipline = () => { const name = clean(newDiscipline); if (!name) return; const id = normalizeBoqDisc(name); setBoqDisciplines(prev => ({ ...prev, [id]: { name, icon: 'FileText' } })); setDisc(id); setNewDiscipline(''); };
  const deleteDiscipline = id => { if (disc === 'all' || !boqDisciplines[id]) return; if (!confirm('למחוק דיסציפלינה וכל הפריטים שלה?')) return; setBoqDisciplines(prev => { const next = { ...prev }; delete next[id]; return next; }); setItems(prev => prev.filter(x => x.disciplineId !== id)); setCart(prev => prev.filter(x => x.disciplineId !== id)); setDisc('all'); };

  // #5 Add manual item
  const addManualItem = () => {
    if (!newItem.desc) return;
    const item = { ...newItem, id: uid('man'), defaultQty: 1, validity: '', notes: 'פריט ידני', currency: 'ILS' };
    setItems(prev => [...prev, item]);
    setNewItem({ desc: '', code: '', unit: 'יח׳', material: 0, labor: 0, engineering: 0, overhead: 0, disciplineId: disc === 'all' ? 'piping' : disc, supplier: '' });
    setShowAddForm(false); setStatus('נוסף פריט ידני.');
  };

  // #9 Get effective markup per discipline
  const getMarkup = dId => ({ ...percent, ...(discMarkup[dId] || {}) });

  // Totals calculation with per-discipline markup
  const totals = useMemo(() => {
    let material = 0, labor = 0, eng = 0, overhead = 0, direct = 0, management = 0, contingency = 0, discountTotal = 0, profitTotal = 0;
    const byDisc = {};
    cart.forEach(x => {
      const t = itemTotal(x) * x.qty;
      const m = num(x.material) * x.qty, l = num(x.labor) * x.qty, e = num(x.engineering) * x.qty, o = num(x.overhead) * x.qty;
      material += m; labor += l; eng += e; overhead += o;
      const mk = getMarkup(x.disciplineId);
      const d = m + l + e + o;
      management += d * mk.management / 100;
      contingency += d * mk.contingency / 100;
      const bfd = d + d * mk.management / 100 + d * mk.contingency / 100;
      const discAmt = bfd * mk.discount / 100;
      discountTotal += discAmt;
      profitTotal += (bfd - discAmt) * mk.profit / 100;
      if (!byDisc[x.disciplineId]) byDisc[x.disciplineId] = 0;
      byDisc[x.disciplineId] += t;
    });
    direct = material + labor + eng + overhead;
    const beforeDiscount = direct + management + contingency;
    const total = beforeDiscount - discountTotal + profitTotal;
    const rate = num(project.exchangeRate) || 1;
    return { material, labor, eng, overhead, direct, management, contingency, discount: discountTotal, profit: profitTotal, total, totalConverted: total * rate, byDisc };
  }, [cart, percent, discMarkup, project.exchangeRate]);

  const byDiscArr = useMemo(() => Object.entries(boqDisciplines).map(([id, d]) => ({ id, name: d.name, total: totals.byDisc[id] || 0 })), [totals.byDisc, boqDisciplines]);

  // #7 Cart grouped by discipline
  const cartGrouped = useMemo(() => {
    const groups = {};
    cart.forEach(x => { if (!groups[x.disciplineId]) groups[x.disciplineId] = []; groups[x.disciplineId].push(x); });
    return Object.entries(groups).map(([id, items]) => ({ id, name: boqDisciplines[id]?.name || id, items, subtotal: items.reduce((s, x) => s + itemTotal(x) * x.qty, 0) }));
  }, [cart, boqDisciplines]);

  // #12 Supplier matches for a discipline
  const getSupplierMatches = dId => {
    const supDisc = DISC_TO_SUP[dId];
    if (!supDisc) return [];
    return supplierHints.filter(s => s.discipline === supDisc).slice(0, 3);
  };

  // Save manual + version snapshot
  const saveManual = () => {
    const snap = { id: uid('ver'), savedAt: new Date().toISOString(), totalSnapshot: totals.total, cartCount: cart.length };
    setVersions(prev => [...prev.slice(-9), snap]);
    setStatus('הפרויקט נשמר.');
  };
  saveManualRef.current = saveManual;

  // #1 Project management
  const newProject = () => {
    if (activeId) saveProj(activeId, getState());
    const id = uid('proj');
    const meta = { id, name: 'פרויקט חדש', status: 'טיוטה', updatedAt: new Date().toISOString() };
    const idx = [...loadIdx(), meta]; saveIdx(idx); setProjects(idx);
    setActiveId(id); localStorage.setItem(ACTIVE_PROJ, id);
    applyState(null); setItems(sampleItems); setBoqDisciplines(defaultBoqDisciplines);
    setProject({ name: 'פרויקט חדש', customer: '', estimator: '', currency: 'ILS', status: 'טיוטה', exchangeRate: 1 });
    setStatus('נוצר פרויקט חדש.');
  };
  const switchProject = id => {
    if (id === activeId) return;
    if (activeId) saveProj(activeId, getState());
    const d = loadProj(id); applyState(d); setActiveId(id); localStorage.setItem(ACTIVE_PROJ, id);
    setStatus(`נטען פרויקט: ${d?.project?.name || id}`);
  };
  const duplicateProject = () => {
    const id = uid('proj');
    const data = { ...getState(), project: { ...project, name: project.name + ' (העתק)' } };
    saveProj(id, data);
    const meta = { id, name: data.project.name, status: 'טיוטה', updatedAt: new Date().toISOString() };
    const idx = [...loadIdx(), meta]; saveIdx(idx); setProjects(idx);
    switchProject(id); setStatus('הפרויקט שוכפל.');
  };
  const deleteProjectById = id => {
    if (!confirm('למחוק את הפרויקט?')) return;
    delProj(id);
    const idx = loadIdx().filter(p => p.id !== id); saveIdx(idx); setProjects(idx);
    if (id === activeId) { if (idx.length) switchProject(idx[0].id); else newProject(); }
  };

  const reset = () => { setItems(sampleItems); setBoqDisciplines(defaultBoqDisciplines); setCart([]); setDisc('all'); setResult(false); setSelected(new Set()); setStatus('חזרת למחירון דוגמה.'); };

  // #10 Detailed Excel export
  const exportCSV = () => {
    const data = cart.map(x => ({ 'דיסציפלינה': boqDisciplines[x.disciplineId]?.name || x.disciplineId, 'מק״ט': x.code, 'תיאור': x.desc, 'כמות': x.qty, 'יחידה': x.unit, 'חומרים': num(x.material), 'עבודה': num(x.labor), 'תכנון': num(x.engineering), 'תקורה': num(x.overhead), 'עלות ליחידה': itemTotal(x), 'סה״כ שורה': itemTotal(x) * x.qty, 'ספק': x.supplier || '', 'הערה': x.cartNote || '' }));
    data.push({});
    data.push({ 'דיסציפלינה': 'סיכום', 'תיאור': 'עלות ישירה', 'סה״כ שורה': totals.direct });
    data.push({ 'תיאור': `ניהול ${percent.management}%`, 'סה״כ שורה': totals.management });
    data.push({ 'תיאור': `בלתי צפוי ${percent.contingency}%`, 'סה״כ שורה': totals.contingency });
    data.push({ 'תיאור': `הנחה ${percent.discount}%`, 'סה״כ שורה': -totals.discount });
    data.push({ 'תיאור': `רווח ${percent.profit}%`, 'סה״כ שורה': totals.profit });
    if (showVat) data.push({ 'תיאור': `מע"מ ${VAT_RATE * 100}%`, 'סה״כ שורה': totals.total * VAT_RATE });
    data.push({ 'דיסציפלינה': 'סה״כ', 'תיאור': project.name, 'סה״כ שורה': showVat ? totals.total * (1 + VAT_RATE) : totals.total });
    const ws = XLSX.utils.json_to_sheet(data);
    ws['!cols'] = [{ wch: 18 }, { wch: 10 }, { wch: 40 }, { wch: 8 }, { wch: 8 }, { wch: 12 }, { wch: 12 }, { wch: 12 }, { wch: 12 }, { wch: 14 }, { wch: 14 }, { wch: 18 }, { wch: 20 }];
    const wb = XLSX.utils.book_new(); XLSX.utils.book_append_sheet(wb, ws, 'BOQ'); XLSX.writeFile(wb, `boq-${project.name}.xlsx`);
  };
  const exportPDF = async () => { setResult(true); setTimeout(async () => { if (!reportRef.current) return; const canvas = await html2canvas(reportRef.current, { scale: 2, useCORS: true }); const img = canvas.toDataURL('image/png'); const pdf = new jsPDF('p', 'mm', 'a4'); const w = 210, h = canvas.height * w / canvas.width; pdf.addImage(img, 'PNG', 0, 0, w, h); pdf.save(`boq-${project.name}.pdf`); }, 100); };

  // #13 RFQ Export
  const exportRFQ = (supplierName = '') => {
    const data = cart.map(x => ({ 'מק״ט': x.code, 'תיאור': x.desc, 'כמות': x.qty, 'יחידה': x.unit, 'דיסציפלינה': boqDisciplines[x.disciplineId]?.name || '', 'מחיר ליחידה (למילוי)': '', 'סה״כ (למילוי)': '', 'הערות': '' }));
    const ws = XLSX.utils.json_to_sheet(data);
    ws['!cols'] = [{ wch: 10 }, { wch: 45 }, { wch: 8 }, { wch: 8 }, { wch: 16 }, { wch: 16 }, { wch: 14 }, { wch: 20 }];
    const wb = XLSX.utils.book_new(); XLSX.utils.book_append_sheet(wb, ws, 'RFQ');
    XLSX.writeFile(wb, `rfq-${project.name}${supplierName ? '-' + supplierName : ''}.xlsx`);
    setStatus('בקשה להצעת מחיר יוצאה בהצלחה.');
  };

  // #14 Attachments + PDF auto-parse
  const handleAttach = e => {
    const files = Array.from(e.target?.files || e.dataTransfer?.files || []);
    files.forEach(f => {
      if (f.size > 10000000) { alert(`${f.name} גדול מ-10MB, לא נשמר`); return; }

      // PDF auto-parse
      if (f.type === 'application/pdf' || f.name.toLowerCase().endsWith('.pdf')) {
        setPdfParsing(true);
        const abReader = new FileReader();
        abReader.onload = async ev => {
          try {
            const pages = await extractPdfText(ev.target.result);
            const { items, rawText } = parsePdfRows(pages);
            setPdfPreview({ items, rawText, fileName: f.name });
            if (items.length > 0) setStatus(`נמצאו ${items.length} שורות מחירון ב-${f.name}. בדוק ואשר ייבוא.`);
            else setStatus(`לא נמצאו טבלאות ב-${f.name}. טקסט גולמי מוצג למטה.`);
          } catch (err) { console.error('PDF parse error:', err); setStatus(`שגיאה בקריאת ${f.name}`); }
          setPdfParsing(false);
        };
        abReader.readAsArrayBuffer(f);
        // Also store as attachment
        const reader2 = new FileReader();
        reader2.onload = ev2 => setAttachments(prev => [...prev, { id: uid('att'), name: f.name, type: f.type, size: f.size, dataUrl: f.size < 3000000 ? ev2.target.result : null, addedAt: new Date().toLocaleDateString('he-IL') }]);
        reader2.readAsDataURL(f);
        return;
      }

      const reader = new FileReader();
      reader.onload = ev => setAttachments(prev => [...prev, { id: uid('att'), name: f.name, type: f.type, size: f.size, dataUrl: ev.target.result, addedAt: new Date().toLocaleDateString('he-IL') }]);
      reader.readAsDataURL(f);
    });
    if (e.target) e.target.value = '';
  };
  const removeAttach = id => setAttachments(prev => prev.filter(a => a.id !== id));
  const importPdfItems = () => {
    if (!pdfPreview?.items?.length) return;
    const newMap = { ...boqDisciplines };
    pdfPreview.items.forEach(it => { if (!newMap[it.disciplineId]) newMap[it.disciplineId] = { name: it.disciplineId, icon: 'FileText' }; });
    setBoqDisciplines(newMap);
    setItems(prev => [...prev, ...pdfPreview.items]);
    setStatus(`יובאו ${pdfPreview.items.length} פריטים מ-${pdfPreview.fileName} למחירון.`);
    setPdfPreview(null);
  };
  const importPdfReplace = () => {
    if (!pdfPreview?.items?.length) return;
    const newMap = makeDisciplineMapFromSheetNames([], defaultBoqDisciplines);
    pdfPreview.items.forEach(it => { if (!newMap[it.disciplineId]) newMap[it.disciplineId] = { name: it.disciplineId, icon: 'FileText' }; });
    setBoqDisciplines(newMap);
    setItems(pdfPreview.items);
    setDisc('all'); setResult(false); setSelected(new Set());
    setStatus(`${pdfPreview.items.length} פריטים מ-${pdfPreview.fileName} החליפו את המחירון.`);
    setPdfPreview(null);
  };

  const vatAmount = showVat ? totals.total * VAT_RATE : 0;
  const grandTotal = totals.total + vatAmount;
  const cur = project.currency;

  return <main className={'layout' + (result ? ' reportOpen' : '')}>
    {/* #1 Project Bar */}
    <section className="projectBar">
      <div className="projSelect">
        <Layers size={16} />
        <select value={activeId || ''} onChange={e => switchProject(e.target.value)}>
          {projects.map(p => <option key={p.id} value={p.id}>{p.name} {p.status ? `(${p.status})` : ''}</option>)}
        </select>
      </div>
      <div className="projActions">
        <button onClick={newProject} title="פרויקט חדש"><FolderPlus size={15} /></button>
        <button onClick={duplicateProject} title="שכפל"><Copy size={15} /></button>
        <button onClick={() => deleteProjectById(activeId)} title="מחק" className="dangerMini"><Trash2 size={15} /></button>
      </div>
      <div className="projStatus">
        <select value={project.status || 'טיוטה'} onChange={e => setProject(p => ({ ...p, status: e.target.value }))} style={{ borderColor: STATUS_COLORS[project.status] || '#94a3b8' }}>
          {STATUS_OPTIONS.map(s => <option key={s}>{s}</option>)}
        </select>
        {project.status && <span className="statusDot" style={{ background: STATUS_COLORS[project.status] }} />}
      </div>
      <span className="projDate"><Clock size={13} /> {new Date().toLocaleDateString('he-IL')}</span>
    </section>

    <section className="left">
      {/* Upload */}
      <div className="panel upload"><div><h2><UploadCloud /> העלאת מחירון Excel</h2><p>המערכת קוראת את כל הלשוניות בקובץ.</p><div className="status"><CheckCircle2 size={16} />{status}</div></div><div className="actions"><input ref={inputRef} type="file" accept=".xlsx,.xls,.xlsm,.csv" hidden onChange={upload} /><button onClick={() => inputRef.current.click()}><UploadCloud size={18} /> העלאת Excel</button><button onClick={reset}><RotateCcw size={18} /> דוגמה</button></div></div>

      {/* #2 Project details + status + currency */}
      <div className="panel"><h2><Database /> פרטי פרויקט</h2><div className="formGrid">
        <Field label="שם פרויקט" value={project.name} onChange={v => setProject(p => ({ ...p, name: v }))} />
        <Field label="לקוח / מחלקה" value={project.customer} onChange={v => setProject(p => ({ ...p, customer: v }))} />
        <Field label="עורך אומדן" value={project.estimator} onChange={v => setProject(p => ({ ...p, estimator: v }))} />
        <label>מטבע<select value={project.currency} onChange={e => setProject(p => ({ ...p, currency: e.target.value }))}><option value="ILS">ILS ₪</option><option value="USD">USD $</option><option value="EUR">EUR €</option></select></label>
        {project.currency !== 'ILS' && <Field label="שער המרה ל-ILS" type="number" value={project.exchangeRate} onChange={v => setProject(p => ({ ...p, exchangeRate: num(v) || 1 }))} />}
        <label className="vatToggle"><input type="checkbox" checked={showVat} onChange={e => setShowVat(e.target.checked)} /> כולל מע"מ {VAT_RATE * 100}%</label>
        {versions.length > 0 && <label className="verInfo"><Hash size={13} /> גרסה {versions.length} · {versions[versions.length - 1]?.savedAt ? new Date(versions[versions.length - 1].savedAt).toLocaleString('he-IL') : ''}</label>}
      </div></div>

      {/* #14 Attachments + PDF auto-import */}
      <div className="panel attachPanel">
        <h2 onClick={() => attachRef.current?.click()} style={{ cursor: 'pointer' }}><Paperclip /> צרופות ומסמכים ({attachments.length})</h2>
        <input ref={attachRef} type="file" hidden multiple accept=".pdf,.xlsx,.xls,.csv,.jpg,.jpeg,.png,.dwg,.doc,.docx" onChange={handleAttach} />
        <div className="attachZone" onDragOver={e => e.preventDefault()} onDrop={e => { e.preventDefault(); handleAttach(e); }}>
          {pdfParsing ? <div className="pdfLoading"><Loader size={24} className="spin" /><p>מנתח את ה-PDF... מחלץ טקסט וטבלאות</p></div> :
          <><p>גרור קבצים לכאן או <button onClick={() => attachRef.current?.click()}>בחר קבצים</button></p>
          <small><FileSearch size={14} /> קובצי PDF ינותחו אוטומטית - המערכת תחלץ טבלאות מחירים ותייבא למחירון</small></>}
        </div>

        {/* PDF preview panel */}
        {pdfPreview && <div className="pdfPreview">
          <div className="pdfPreviewHead">
            <h3><FileSearch size={18} /> תוצאות ניתוח: {pdfPreview.fileName}</h3>
            {pdfPreview.items.length > 0 ? <span className="pdfCount">{pdfPreview.items.length} שורות זוהו</span> : <span className="pdfNoItems">לא זוהו טבלאות</span>}
          </div>
          {pdfPreview.items.length > 0 && <>
            <div className="pdfTable">
              <table><thead><tr><th>מק״ט</th><th>תיאור</th><th>כמות</th><th>יחידה</th><th>מחיר</th><th>דיסציפלינה</th></tr></thead>
              <tbody>{pdfPreview.items.slice(0, 30).map(it => <tr key={it.id}>
                <td>{it.code}</td><td>{it.desc}</td><td>{it.defaultQty}</td><td>{it.unit}</td><td>{fmt(it.material)}</td><td>{boqDisciplines[it.disciplineId]?.name || it.disciplineId}</td>
              </tr>)}</tbody></table>
              {pdfPreview.items.length > 30 && <small>... ועוד {pdfPreview.items.length - 30} שורות</small>}
            </div>
            <div className="pdfActions">
              <button onClick={importPdfItems}><Plus size={16} /> הוסף למחירון הקיים</button>
              <button onClick={importPdfReplace}><RotateCcw size={16} /> החלף את המחירון</button>
              <button onClick={() => setPdfPreview(null)}><X size={16} /> בטל</button>
            </div>
          </>}
          {pdfPreview.items.length === 0 && pdfPreview.rawText && <div className="pdfRawText">
            <p>טקסט גולמי שחולץ מהמסמך:</p>
            <pre>{pdfPreview.rawText.slice(0, 3000)}{pdfPreview.rawText.length > 3000 ? '\n...(קוצר)' : ''}</pre>
          </div>}
        </div>}

        {attachments.length > 0 && <div className="attachList">{attachments.map(a => <div key={a.id} className="attachItem">
          {a.type?.startsWith('image/') && a.dataUrl && <img src={a.dataUrl} className="attachThumb" />}
          {a.type === 'application/pdf' && <FileText size={20} style={{ color: '#dc2626', flexShrink: 0 }} />}
          <span>{a.name} <small>({(a.size / 1024).toFixed(0)}KB)</small></span>
          <button onClick={() => removeAttach(a.id)}><X size={14} /></button>
        </div>)}</div>}
      </div>

      {/* Catalog */}
      <div className="panel">
        <div className="catalogTop">
          <div><h2><Calculator /> מחירון פריטים ({filtered.length})</h2></div>
          <div className="catalogActions">
            <div className="search"><Search size={18} /><input ref={searchRef} value={query} onChange={e => setQuery(e.target.value)} placeholder="חיפוש פריט / ספק / מק״ט (Ctrl+F)" />{query && <button className="clearBtn" onClick={() => setQuery('')}><X size={16} /></button>}</div>
            <select className="sortSelect" value={sortMode} onChange={e => setSortMode(e.target.value)} title="מיון">
              <option value="default">ברירת מחדל</option><option value="price-asc">מחיר: נמוך לגבוה</option><option value="price-desc">מחיר: גבוה לנמוך</option><option value="name">לפי שם</option>
            </select>
          </div>
        </div>
        <div className="chips">
          <button className={disc === 'all' ? 'selected' : ''} onClick={() => setDisc('all')}>הכל</button>
          {Object.entries(boqDisciplines).map(([id, d]) => <button key={id} className={disc === id ? 'selected' : ''} onClick={() => setDisc(id)}>{d.name}</button>)}
          <button className={favOnly ? 'selected' : ''} onClick={() => setFavOnly(!favOnly)} title="מועדפים"><Star size={14} /></button>
        </div>
        <div className="disciplineManage">
          <input value={newDiscipline} onChange={e => setNewDiscipline(e.target.value)} placeholder="הוסף דיסציפלינה ידנית" />
          <button onClick={addDiscipline}><Plus size={16} /> דיסציפלינה</button>
          {disc !== 'all' && <button className="dangerMini" onClick={() => deleteDiscipline(disc)}><Trash2 size={16} /> מחק</button>}
          <button onClick={() => setShowAddForm(!showAddForm)}><Plus size={16} /> פריט ידני</button>
        </div>

        {/* #5 Add manual item form */}
        {showAddForm && <div className="addItemForm"><div className="formGrid">
          <Field label="תיאור" value={newItem.desc} onChange={v => setNewItem(p => ({ ...p, desc: v }))} />
          <Field label="מק״ט" value={newItem.code} onChange={v => setNewItem(p => ({ ...p, code: v }))} />
          <Field label="יחידה" value={newItem.unit} onChange={v => setNewItem(p => ({ ...p, unit: v }))} />
          <label>דיסציפלינה<select value={newItem.disciplineId} onChange={e => setNewItem(p => ({ ...p, disciplineId: e.target.value }))}>{Object.entries(boqDisciplines).map(([id, d]) => <option key={id} value={id}>{d.name}</option>)}</select></label>
          <Field label="חומרים ₪" type="number" value={newItem.material} onChange={v => setNewItem(p => ({ ...p, material: num(v) }))} />
          <Field label="עבודה ₪" type="number" value={newItem.labor} onChange={v => setNewItem(p => ({ ...p, labor: num(v) }))} />
          <Field label="תכנון ₪" type="number" value={newItem.engineering} onChange={v => setNewItem(p => ({ ...p, engineering: num(v) }))} />
          <Field label="תקורה ₪" type="number" value={newItem.overhead} onChange={v => setNewItem(p => ({ ...p, overhead: num(v) }))} />
        </div><button className="calc" onClick={addManualItem} style={{ marginTop: 8 }}>הוסף למחירון</button></div>}

        {/* #4 Selection bar */}
        {selected.size > 0 && <div className="selectBar"><button onClick={selectAll}>{selected.size === filtered.length ? <><CheckSquare size={15} /> בטל הכל</> : <><Square size={15} /> בחר הכל</>}</button><button onClick={addSelected}><Plus size={15} /> הוסף {selected.size} פריטים לסל</button><span>{selected.size} נבחרו</span></div>}

        <div className="items">
          {filtered.length === 0 && <div className="emptyResults"><b>לא נמצאו פריטים</b><p>{query ? `אין תוצאות עבור "${query}"` : 'אין פריטים בדיסציפלינה זו'}</p></div>}
          {filtered.map((it, idx) => { const Icon = getIcon(boqDisciplines[it.disciplineId]?.icon); const itemId = it.id || it.code; return <div className={'item' + (selected.has(itemId) ? ' itemSelected' : '')} key={`${itemId}-${idx}`}>
            <button className="itemCheck" onClick={() => toggleSelect(itemId)}>{selected.has(itemId) ? <CheckSquare size={18} /> : <Square size={18} />}</button>
            <div className="itemIcon"><Icon size={22} /></div>
            <div className="itemText"><b>{it.desc}</b><span>{boqDisciplines[it.disciplineId]?.name || it.disciplineId} · {it.code} · {it.supplier || '-'} · {it.unit}</span><small>{it.notes}</small></div>
            <div className="price"><b>{fmt(itemTotal(it), cur)}</b>
              <div className="priceActions"><button onClick={() => add(it)}><Plus size={16} /> הוסף</button>
              <button className={favorites.has(it.code) ? 'favBtn on' : 'favBtn'} onClick={() => setFavorites(prev => { const n = new Set(prev); if (n.has(it.code)) n.delete(it.code); else n.add(it.code); return n; })}><Star size={14} fill={favorites.has(it.code) ? 'currentColor' : 'none'} /></button>
              <button className="itemActionBtn" title="שכפל" onClick={() => setItems(prev => [...prev, { ...it, id: uid('dup'), code: it.code + '-copy' }])}><Copy size={14} /></button>
              <button className="itemActionBtn dangerMini" title="מחק" onClick={() => { if (confirm('למחוק פריט מהמחירון?')) setItems(prev => prev.filter(x => x.id !== it.id)); }}><Trash2 size={14} /></button></div>
            </div>
          </div>; })}
        </div>
      </div>
    </section>

    {/* SIDEBAR */}
    <aside className="panel sticky">
      <div className="cartHead"><h2><Calculator /> סל חישוב ({cart.length})</h2>{cart.length > 0 && <button className="clearCartBtn" onClick={() => { if (confirm('לנקות את כל הסל?')) { setCart([]); setResult(false); } }}><Trash2 size={14} /> נקה סל</button>}</div>
      {cart.length === 0 ? <div className="empty">לא נבחרו פריטים</div> : <>
        {/* #7 Cart grouped by discipline */}
        <div className="cartList">{cartGrouped.map(g => <div key={g.id} className="cartGroup">
          <div className="cartGroupHead"><b>{g.name}</b><span>{fmt(g.subtotal, cur)}</span></div>
          {g.items.map(x => {
            const expanded = expandedCart.has(x.id);
            const sups = getSupplierMatches(x.disciplineId);
            return <div className="cart" key={x.id}>
              <div className="cartRow">
                <button onClick={() => { if (confirm('להסיר פריט מהסל?')) setCart(p => p.filter(z => z.id !== x.id)); }}><Trash2 size={14} /></button>
                <b>{x.desc}</b>
                <button className="expandBtn" onClick={() => setExpandedCart(prev => { const n = new Set(prev); if (n.has(x.id)) n.delete(x.id); else n.add(x.id); return n; })}>{expanded ? <ChevronUp size={14} /> : <ChevronDown size={14} />}</button>
              </div>
              <div className="cartQty">
                <input type="number" value={x.qty} min={1} onChange={e => setCart(p => p.map(z => z.id === x.id ? { ...z, qty: num(e.target.value) || 1 } : z))} />
                <span>{x.unit}</span><strong>{fmt(itemTotal(x) * x.qty, cur)}</strong>
              </div>
              {/* #6 Inline edit */}
              {expanded && <div className="cartExpand">
                <MiniField label="חומרים" value={x.material} onChange={v => setCart(p => p.map(z => z.id === x.id ? { ...z, material: num(v) } : z))} />
                <MiniField label="עבודה" value={x.labor} onChange={v => setCart(p => p.map(z => z.id === x.id ? { ...z, labor: num(v) } : z))} />
                <MiniField label="תכנון" value={x.engineering} onChange={v => setCart(p => p.map(z => z.id === x.id ? { ...z, engineering: num(v) } : z))} />
                <MiniField label="תקורה" value={x.overhead} onChange={v => setCart(p => p.map(z => z.id === x.id ? { ...z, overhead: num(v) } : z))} />
                <label className="cartNoteLabel">הערה<input value={x.cartNote || ''} onChange={e => setCart(p => p.map(z => z.id === x.id ? { ...z, cartNote: e.target.value } : z))} placeholder="הערת מהנדס" /></label>
                {/* #12 Supplier hints */}
                {sups.length > 0 && <div className="supHints"><small>ספקים מומלצים:</small>{sups.map(s => <span key={s.id}>{s.name} {s.phone ? `(${s.phone})` : ''} {'★'.repeat(s.rating || 0)}</span>)}</div>}
              </div>}
            </div>;
          })}
        </div>)}</div>
      </>}

      {/* Markup */}
      <div className="percentGrid">
        <Field label="ניהול %" type="number" value={percent.management} onChange={v => setPercent(p => ({ ...p, management: num(v) }))} />
        <Field label="בלתי צפוי %" type="number" value={percent.contingency} onChange={v => setPercent(p => ({ ...p, contingency: num(v) }))} />
        <Field label="רווח %" type="number" value={percent.profit} onChange={v => setPercent(p => ({ ...p, profit: num(v) }))} />
        <Field label="הנחה %" type="number" value={percent.discount} onChange={v => setPercent(p => ({ ...p, discount: num(v) }))} />
      </div>

      {/* #9 Per-discipline markup toggle */}
      <button className="discMarkupToggle" onClick={() => setShowDiscMarkup(!showDiscMarkup)}>
        <Percent size={14} /> אחוזים לפי דיסציפלינה {showDiscMarkup ? <ChevronUp size={14} /> : <ChevronDown size={14} />}
      </button>
      {showDiscMarkup && <div className="discMarkupPanel">{Object.entries(boqDisciplines).map(([id, d]) => {
        const mk = discMarkup[id] || {};
        const set = (k, v) => setDiscMarkup(prev => ({ ...prev, [id]: { ...(prev[id] || {}), [k]: num(v) } }));
        const hasOverride = mk.management !== undefined || mk.contingency !== undefined || mk.profit !== undefined || mk.discount !== undefined;
        return <div key={id} className="discMkRow">
          <b>{d.name} {hasOverride && <span className="overrideBadge">מותאם</span>}</b>
          <div className="discMkInputs">
            <MiniField label="ניהול" value={mk.management ?? percent.management} onChange={v => set('management', v)} />
            <MiniField label="בלת״צ" value={mk.contingency ?? percent.contingency} onChange={v => set('contingency', v)} />
            <MiniField label="רווח" value={mk.profit ?? percent.profit} onChange={v => set('profit', v)} />
            <MiniField label="הנחה" value={mk.discount ?? percent.discount} onChange={v => set('discount', v)} />
          </div>
        </div>;
      })}</div>}

      <button className="calc" disabled={!cart.length} onClick={() => setResult(true)}>חשב פרויקט</button>
      <div className="sideBtns">
        <button onClick={saveManual}><Save size={16} /> שמור (Ctrl+S)</button>
        <button onClick={exportCSV}><Download size={16} /> Excel</button>
        <button onClick={exportPDF}><Printer size={16} /> PDF</button>
        <button onClick={() => exportRFQ()}><Send size={16} /> RFQ</button>
        {versions.length > 1 && <button onClick={() => setShowCompare(!showCompare)}><Eye size={16} /> השוואה</button>}
      </div>

      {/* #11 Version comparison */}
      {showCompare && versions.length > 1 && <div className="compareBox">
        <h3><Eye size={16} /> השוואת גרסאות</h3>
        {versions.slice(-5).map((v, i, arr) => {
          const prev = arr[i - 1];
          const delta = prev ? v.totalSnapshot - prev.totalSnapshot : 0;
          return <div key={v.id} className="compareLine">
            <span>#{i + 1} · {new Date(v.savedAt).toLocaleString('he-IL')}</span>
            <span>{fmt(v.totalSnapshot, cur)} {delta !== 0 && <b style={{ color: delta > 0 ? '#dc2626' : '#16a34a' }}>{delta > 0 ? '+' : ''}{fmt(delta, cur)}</b>}</span>
            <small>{v.cartCount} פריטים</small>
          </div>;
        })}
      </div>}

      <div className="totalBox">
        <span>סה״כ {showVat ? '(כולל מע"מ)' : '(לפני מע"מ)'}</span>
        <b>{fmt(grandTotal, cur)}</b>
        {showVat && <small>מע"מ: {fmt(vatAmount, cur)}</small>}
        {project.currency !== 'ILS' && <small>≈ {fmt(totals.totalConverted)} ₪</small>}
      </div>
    </aside>

    {/* REPORT */}
    {result && <section className="report" ref={reportRef}>
      {/* Floating action bar above report */}
      <div className="reportBar">
        <button onClick={() => setResult(false)}><Pencil size={16} /> חזור לעריכה</button>
        <button onClick={() => { setResult(false); setCart([]); setStatus('הסל נוקה. אפשר להתחיל מחדש.'); }}><Trash2 size={16} /> נקה סל והתחל מחדש</button>
        <button onClick={() => { if (!confirm('האם לאפס את הכל? כל הנתונים ייאבדו.')) return; setResult(false); setCart([]); setItems(sampleItems); setBoqDisciplines(defaultBoqDisciplines); setPercent({ management: 7, contingency: 12, profit: 10, discount: 0 }); setDiscMarkup({}); setProject(p => ({ ...p, name: 'פרויקט חדש', customer: '', status: 'טיוטה' })); setAttachments([]); setVersions([]); setStatus('המערכת אופסה. אפשר להתחיל הכל מאפס.'); }}><RotateCcw size={16} /> אפס הכל מאפס</button>
        <span className="reportBarInfo">{cart.length} פריטים · {fmt(grandTotal, cur)}</span>
      </div>
      <div className="reportBox">
      {/* #16 Print cover */}
      <div className="printCover"><div className="logoPlaceholder printLogo"><Building2 size={48} /></div><h1>{project.name}</h1><p>{project.customer} · {project.estimator} · {new Date().toLocaleDateString('he-IL')}</p><p>סטטוס: {project.status}</p></div>
      <div className="reportHead"><div className="logoPlaceholder"><Building2 size={28} /></div><div><h2>דוח אומדן פרויקט</h2><p>{project.name} · {project.customer} · {new Date().toLocaleDateString('he-IL')}</p><p>סטטוס: {project.status} · עורך: {project.estimator} · מטבע: {project.currency}{versions.length > 0 ? ` · גרסה ${versions.length}` : ''}</p></div><div className="reportActions"><button onClick={exportPDF}>PDF</button><button onClick={() => exportRFQ()}>RFQ</button></div></div>
      <div className="kpis"><K title="חומרים" value={fmt(totals.material, cur)} /><K title="עבודה" value={fmt(totals.labor, cur)} /><K title="תכנון" value={fmt(totals.eng, cur)} /><K title="סה״כ" value={fmt(grandTotal, cur)} big /></div>
      <div className="reportGrid">
        <div className="box"><h3><BarChart3 /> לפי דיסציפלינה</h3>{byDiscArr.map(d => { const pct = totals.direct ? Math.round(d.total / totals.direct * 100) : 0; return <div className="bar" key={d.id}><span><b>{d.name}</b><b>{fmt(d.total, cur)} · {pct}%</b></span><i><em style={{ width: pct + '%' }} /></i></div>; })}</div>
        <div className="box"><h3>סיכום מסחרי</h3><Line l="עלות ישירה" v={totals.direct} c={cur} /><Line l={`ניהול`} v={totals.management} c={cur} /><Line l={`בלתי צפוי`} v={totals.contingency} c={cur} /><Line l={`הנחה`} v={-totals.discount} c={cur} /><Line l={`רווח`} v={totals.profit} c={cur} />{showVat && <Line l={`מע"מ ${VAT_RATE * 100}%`} v={vatAmount} c={cur} />}<div className="grand"><span>סה״כ אומדן</span><b>{fmt(grandTotal, cur)}</b>{project.currency !== 'ILS' && <small>≈ {fmt(totals.totalConverted)} ₪</small>}</div></div>
      </div>
      <table><thead><tr><th>דיסציפלינה</th><th>מק״ט</th><th>תיאור</th><th>ספק</th><th>כמות</th><th>יחידה</th><th>חומרים</th><th>עבודה</th><th>תכנון</th><th>תקורה</th><th>סה״כ</th></tr></thead><tbody>{cart.map(x => <tr key={x.id}><td>{boqDisciplines[x.disciplineId]?.name || x.disciplineId}</td><td>{x.code}</td><td>{x.desc}{x.cartNote ? ` (${x.cartNote})` : ''}</td><td>{x.supplier || '-'}</td><td>{x.qty}</td><td>{x.unit}</td><td>{fmt(num(x.material) * x.qty, cur)}</td><td>{fmt(num(x.labor) * x.qty, cur)}</td><td>{fmt(num(x.engineering) * x.qty, cur)}</td><td>{fmt(num(x.overhead) * x.qty, cur)}</td><td>{fmt(itemTotal(x) * x.qty, cur)}</td></tr>)}</tbody></table>
      <div className="disclaimer">הנתונים מיועדים לאומדן ראשוני בלבד ודורשים אישור הנדסי/מסחרי לפני שימוש מחייב.</div>
      <div className="reportBottomActions">
        <button onClick={() => setResult(false)}><Pencil size={16} /> חזור לעריכה והוסף פריטים</button>
        <button onClick={exportPDF}><Printer size={16} /> ייצוא PDF</button>
        <button onClick={exportCSV}><Download size={16} /> ייצוא Excel</button>
        <button onClick={() => exportRFQ()}><Send size={16} /> RFQ לספק</button>
      </div>
    </div></section>}
  </main>;
}

/* ====== SUPPLIERS APP ====== */
function SuppliersApp() {
  const [suppliers, setSuppliers] = useState(sampleSuppliers);
  const [query, setQuery] = useState(''); const [disc, setDisc] = useState('הכל');
  const [message, setMessage] = useState('טוען מאגר ספקים...'); const [loaded, setLoaded] = useState(false);
  const [visibleCount, setVisibleCount] = useState(50);
  const [showAddSupplier, setShowAddSupplier] = useState(false);
  const [newSup, setNewSup] = useState({ name: '', supplierNo: '', description: '', discipline: 'כללי / אחר', contact: '', phone: '', email: '', address: '', notes: '' });

  useEffect(() => {
    const cached = localStorage.getItem(SUP_KEY);
    if (cached) { try { const data = JSON.parse(cached); if (Array.isArray(data) && data.length > 0) { setSuppliers(data); setMessage(`נטענו ${data.length.toLocaleString('he-IL')} ספקים מהמטמון.`); setLoaded(true); return; } } catch {} }
    fetch('/suppliers.json').then(r => r.json()).then(data => { if (data.suppliers?.length > 0) { setSuppliers(data.suppliers); localStorage.setItem(SUP_KEY, JSON.stringify(data.suppliers)); setMessage(`נטענו ${data.suppliers.length.toLocaleString('he-IL')} ספקים מהשרת.`); } else setMessage('אפשר להעלות Excel.'); }).catch(() => setMessage('אפשר להעלות Excel.')).finally(() => setLoaded(true));
  }, []);
  useEffect(() => { if (loaded) localStorage.setItem(SUP_KEY, JSON.stringify(suppliers)); }, [suppliers, loaded]);

  const upload = e => { const f = e.target.files?.[0]; if (!f) return; const r = new FileReader(); r.onload = ev => { try { const wb = XLSX.read(new Uint8Array(ev.target.result), { type: 'array', cellDates: true, raw: false }); const parsed = parseWorkbook(wb); if (parsed.length > 0) { setSuppliers(parsed); setMessage(`נטענו ${parsed.length.toLocaleString('he-IL')} ספקים.`); } else { const rows = XLSX.utils.sheet_to_json(wb.Sheets[wb.SheetNames[0]], { defval: '', raw: false }); const fallback = parseSuppliers(rows); setSuppliers(fallback); setMessage(`נטענו ${fallback.length.toLocaleString('he-IL')} ספקים (fallback).`); } } catch (err) { console.error(err); setMessage('שגיאה בקריאת הקובץ'); } }; r.readAsArrayBuffer(f); e.target.value = ''; };

  const filtered = useMemo(() => { setVisibleCount(50); return suppliers.filter(s => (disc === 'הכל' || s.discipline === disc) && `${s.name} ${s.description || ''} ${s.field || ''} ${s.supplierNo || ''} ${s.project || ''} ${s.contact || ''} ${s.phone || ''} ${s.email || ''} ${s.fax || ''} ${s.address || ''} ${s.cityCountry || ''} ${s.notes || ''}`.toLowerCase().includes(query.toLowerCase())); }, [suppliers, disc, query]);
  const disciplineCounts = useMemo(() => { const c = {}; for (const s of suppliers) c[s.discipline] = (c[s.discipline] || 0) + 1; return c; }, [suppliers]);
  const stats = useMemo(() => SUP_DISCIPLINES.map(d => ({ name: d, count: disciplineCounts[d] || 0 })).filter(x => x.count > 0), [disciplineCounts]);
  const update = (id, patch) => setSuppliers(p => p.map(s => s.id === id ? { ...s, ...patch } : s));
  const del = id => { if (!confirm('האם למחוק את הספק?')) return; setSuppliers(p => p.filter(s => s.id !== id)); };
  const addSupplier = () => {
    if (!newSup.name.trim()) return;
    setSuppliers(p => [...p, { ...newSup, id: uid('sup'), field: newSup.description, rating: 0, importedAt: new Date().toLocaleDateString('he-IL') }]);
    setNewSup({ name: '', supplierNo: '', description: '', discipline: 'כללי / אחר', contact: '', phone: '', email: '', address: '', notes: '' });
    setShowAddSupplier(false);
    setMessage('ספק חדש נוסף בהצלחה.');
  };
  const duplicateSupplier = s => {
    setSuppliers(p => [...p, { ...s, id: uid('sup'), name: s.name + ' (העתק)' }]);
    setMessage('ספק שוכפל.');
  };
  const exportExcel = () => { const data = suppliers.map(s => ({ 'מספר ספק': s.supplierNo, 'שם ספק': s.name, 'תחום': s.field || s.discipline, 'תחום פעילות מורחב': s.description, 'כתובת': s.address || '', 'עיר ומדינה': s.cityCountry || '', 'מיקוד': s.zip || '', 'ארץ': s.country || '', 'מספר טלפון': s.phone, 'פקס': s.fax || '', 'מייל': s.email || '', 'איש קשר': s.contact, 'ודאות': s.certainty || '', 'הערות': s.notes, 'דירוג': s.rating })); const ws = XLSX.utils.json_to_sheet(data); const wb = XLSX.utils.book_new(); XLSX.utils.book_append_sheet(wb, ws, 'Suppliers'); XLSX.writeFile(wb, 'suppliers-export.xlsx'); };
  const visible = filtered.slice(0, visibleCount);

  if (!loaded) return <main className="supPage"><div className="loadingState"><b>טוען מאגר ספקים...</b><p>אנא המתן</p></div></main>;
  return <main className="supPage">
    <section className="panel supHero"><div><h2><Users /> מאגר ספקים וקבלנים</h2><p>חיפוש, סיווג, דירוג ותיקון ידני. {suppliers.length} ספקים במאגר.</p><div className="status"><CheckCircle2 size={16} />{message}</div></div><div className="actions"><button onClick={() => setShowAddSupplier(!showAddSupplier)}><UserPlus size={18} /> הוסף ספק</button><label className="fileBtn"><UploadCloud size={18} /> העלאת Excel<input type="file" accept=".xlsx,.xls,.csv" hidden onChange={upload} /></label><button onClick={exportExcel}><Download size={18} /> ייצוא</button><button onClick={() => { setSuppliers(sampleSuppliers); setMessage('חזרת לנתוני דוגמה'); }}><RotateCcw size={18} /> דוגמה</button></div></section>

    {/* Add supplier form */}
    {showAddSupplier && <section className="panel addSupForm">
      <h2><UserPlus size={22} /> הוספת ספק חדש</h2>
      <div className="formGrid">
        <Field label="שם ספק *" value={newSup.name} onChange={v => setNewSup(p => ({ ...p, name: v }))} />
        <Field label="מספר ספק" value={newSup.supplierNo} onChange={v => setNewSup(p => ({ ...p, supplierNo: v }))} />
        <Field label="תיאור / תחום פעילות" value={newSup.description} onChange={v => setNewSup(p => ({ ...p, description: v }))} />
        <label>דיסציפלינה<select value={newSup.discipline} onChange={e => setNewSup(p => ({ ...p, discipline: e.target.value }))}>{SUP_DISCIPLINES.map(d => <option key={d}>{d}</option>)}</select></label>
        <Field label="איש קשר" value={newSup.contact} onChange={v => setNewSup(p => ({ ...p, contact: v }))} />
        <Field label="טלפון" value={newSup.phone} onChange={v => setNewSup(p => ({ ...p, phone: v }))} />
        <Field label="אימייל" value={newSup.email} onChange={v => setNewSup(p => ({ ...p, email: v }))} />
        <Field label="כתובת" value={newSup.address} onChange={v => setNewSup(p => ({ ...p, address: v }))} />
      </div>
      <div className="addSupActions">
        <button className="calc" onClick={addSupplier} style={{ flex: 1 }}><Plus size={16} /> הוסף ספק</button>
        <button onClick={() => setShowAddSupplier(false)} style={{ background: '#e2e8f0', color: '#0f172a' }}><X size={16} /> ביטול</button>
      </div>
    </section>}

    <section className="supplierControls panel"><div className="search big"><Search size={18} /><input value={query} onChange={e => setQuery(e.target.value)} placeholder="חיפוש ספק לפי שם, תחום, טלפון, מייל..." />{query && <button className="clearBtn" onClick={() => setQuery('')}><X size={16} /></button>}</div><select value={disc} onChange={e => setDisc(e.target.value)}><option value="הכל">הכל ({suppliers.length})</option>{SUP_DISCIPLINES.map(d => <option key={d} value={d}>{d} ({disciplineCounts[d] || 0})</option>)}</select></section>
    <section className="stats">{stats.map(s => <div className="stat" key={s.name}><b>{s.count}</b><span>{s.name}</span></div>)}</section>
    <section className="supplierGrid">
      {visible.length === 0 && <div className="emptyResults"><b>לא נמצאו ספקים</b><p>{query ? `אין תוצאות עבור "${query}"` : 'אין ספקים בקטגוריה זו'}</p></div>}
      {visible.map(s => <article className="supplier" key={s.id}>
        <div className="supplierTop"><div><span>{s.discipline}</span><h3>{s.name}</h3><p>מס׳ ספק: {s.supplierNo || '-'}{s.address ? ` · ${s.address}` : ''}{s.cityCountry ? ` · ${s.cityCountry}` : ''}</p></div>
        <div className="supTopActions"><button title="שכפל" onClick={() => duplicateSupplier(s)}><Copy size={14} /></button><button className="danger" title="מחק" onClick={() => del(s.id)}><Trash2 size={14} /></button></div></div>
        <p className="desc">{s.description || s.field || 'אין תיאור'}</p>
        <div className="supplierMeta">
          <span><Pencil size={13} /> איש קשר: <input className="inlineEdit" value={s.contact || ''} onChange={e => update(s.id, { contact: e.target.value })} placeholder="שם איש קשר" /></span>
          <span><Phone size={13} /> טלפון: {s.phone ? <a href={`tel:${s.phone}`} className="contactLink">{s.phone}</a> : <input className="inlineEdit" value="" onChange={e => update(s.id, { phone: e.target.value })} placeholder="הוסף טלפון" />}</span>
          <span><Mail size={13} /> מייל: {s.email ? <a href={`mailto:${s.email}`} className="contactLink">{s.email}</a> : <input className="inlineEdit" value="" onChange={e => update(s.id, { email: e.target.value })} placeholder="הוסף מייל" />}</span>
          {s.fax && <span>פקס: {s.fax}</span>}
          {s.certainty && <span>ודאות: {s.certainty}</span>}
          {s.importedAt && <span><Clock size={13} /> נוסף: {s.importedAt}</span>}
        </div>
        <div className="editRow"><label><Pencil size={14} /> סיווג</label><select value={s.discipline} onChange={e => update(s.id, { discipline: e.target.value })}>{SUP_DISCIPLINES.map(d => <option key={d}>{d}</option>)}</select></div>
        <div className="rating">{[1, 2, 3, 4, 5].map(n => <button key={n} onClick={() => update(s.id, { rating: n })} className={n <= s.rating ? 'on' : ''}><Star size={20} fill="currentColor" /></button>)}</div>
        <textarea value={s.notes || ''} onChange={e => update(s.id, { notes: e.target.value })} placeholder="הערות" />
      </article>)}
      {filtered.length > visibleCount && <button className="calc" onClick={() => setVisibleCount(v => v + 50)}>הצג עוד {Math.min(50, filtered.length - visibleCount)} מתוך {filtered.length - visibleCount}</button>}
    </section>
  </main>;
}

/* ====== HELPERS ====== */
function Field({ label, value, onChange, type = 'text' }) { return <label>{label}<input type={type} value={value} onChange={e => onChange(e.target.value)} /></label>; }
function MiniField({ label, value, onChange }) { return <label className="miniField">{label}<input type="number" value={value} onChange={e => onChange(e.target.value)} /></label>; }
function K({ title, value, big }) { return <div className={'kpi ' + (big ? 'big' : '')}><span>{title}</span><b>{value}</b></div>; }
function Line({ l, v, c = 'ILS' }) { return <div className="line"><span>{l}</span><b>{fmt(v, c)}</b></div>; }

createRoot(document.getElementById('root')).render(<Shell />);
