import React, { useEffect, useMemo, useState } from 'react';
import { createRoot } from 'react-dom/client';
import * as XLSX from 'xlsx';
import { Search, Upload, Trash2, Star, Download, Brain, RefreshCcw, Edit3, Save, Building2, Filter } from 'lucide-react';
import './style.css';

const STORAGE_KEY = 'galil_suppliers_v6_smart';
const RULES_KEY = 'galil_suppliers_v6_learning_rules';

const DISCIPLINES = [
  'צנרת', 'חשמל', 'מכשור ובקרה', 'הנדסה אזרחית', 'מכונות וציוד', 'קונסטרוקציה ומתכת',
  'HVAC / מיזוג', 'כיבוי אש ובטיחות', 'בידוד וצבע', 'לוגיסטיקה ושילוח', 'בדיקות ומעבדה', 'כללי'
];

const DEFAULT_SUPPLIERS = [
  { id:'demo-1', name:'Technobar', contact:'Adi', phone:'', email:'', discipline:'מכשור ובקרה', rawField:'Instrumentation, valves and control items', rating:5, status:'מאושר', source:'דוגמה' },
  { id:'demo-2', name:'ABB', contact:'', phone:'', email:'', discipline:'חשמל', rawField:'Electrical panels, drives and automation', rating:4, status:'מאושר', source:'דוגמה' },
  { id:'demo-3', name:'Hapman', contact:'Lucas', phone:'', email:'', discipline:'מכונות וציוד', rawField:'Conveyors, bulk material handling', rating:4, status:'בבדיקה', source:'דוגמה' },
  { id:'demo-4', name:'Delfin S.R.L', contact:'', phone:'', email:'', discipline:'מכונות וציוד', rawField:'Loading systems', rating:3, status:'בבדיקה', source:'דוגמה' },
  { id:'demo-5', name:'Civil Concrete Works Ltd', contact:'', phone:'', email:'', discipline:'הנדסה אזרחית', rawField:'Concrete and civil construction', rating:4, status:'מאושר', source:'דוגמה' },
];

const KEYWORDS = {
  'חשמל': ['electric','electrical','power','cable','cables','panel','switchgear','transformer','ups','motor','drive','vfd','abb','schneider','siemens','eaton','לוח','חשמל','כבל','כבלים','הארקה','שנאי','מתח','מנוע'],
  'צנרת': ['pipe','piping','valve','valves','flange','fitting','flow','pump','hydro','gasket','coupling','npt','cs','ss','pvc','צנרת','צינור','צינורות','ברז','שסתום','אוגן','אביזרי צנרת','ריתוך'],
  'מכשור ובקרה': ['instrument','instrumentation','control','automation','plc','sensor','transmitter','dcs','scada','calibration','emerson','yokogawa','endress','rosemount','מכשור','בקרה','כיול','מד לחץ','מד טמפרטורה','טרנסמיטר','סנסור'],
  'הנדסה אזרחית': ['civil','construction','concrete','rebar','building','earthworks','excavation','foundation','contractor','בטון','ברזל','קבלן','אזרחי','בנייה','חפירה','יסודות','טפסנות','ריצוף'],
  'מכונות וציוד': ['machine','machinery','equipment','conveyor','sieve','mixer','feeder','crusher','gear','bearing','motorized','hapman','allgaier','ציוד','מכונה','מכונות','מסוע','נפה','מערבל','מיסב','מפוח'],
  'קונסטרוקציה ומתכת': ['steel','metal','structure','structural','fabrication','welding','galvanized','platform','handrail','grating','מתכת','פלדה','קונסטרוקציה','מסגרות','ריתוך','מגולוון','פלטפורמה','מעקה'],
  'HVAC / מיזוג': ['hvac','air condition','chiller','ventilation','fan','duct','cooling','מיזוג','אוורור','צילר','תעלה','מפוח','קירור'],
  'כיבוי אש ובטיחות': ['fire','safety','sprinkler','alarm','ppe','extinguisher','כיבוי','בטיחות','ספרינקלר','גלאי','מטף','אזעקה'],
  'בידוד וצבע': ['insulation','painting','paint','coating','sandblast','thermal','בידוד','צבע','צביעה','ציפוי','התזה','ניקוי חול'],
  'לוגיסטיקה ושילוח': ['shipping','freight','logistics','forwarder','dhl','ups','customs','transport','shipment','שילוח','לוגיסטיקה','הובלה','מכס','משלוח','ים','אוויר'],
  'בדיקות ומעבדה': ['test','testing','inspection','lab','ndt','xray','certificate','qa','qc','בדיקה','בדיקות','מעבדה','תעודה','אישור','פיקוח','ndt']
};

function normalizeText(v='') { return String(v || '').toLowerCase().replace(/["'׳״]/g,' ').replace(/\s+/g,' ').trim(); }
function pick(row, names) { for (const n of names) if (row[n] !== undefined && row[n] !== null && String(row[n]).trim() !== '') return row[n]; return ''; }
function number(v, fallback=0) { const n = Number(String(v ?? '').replace(/,/g,'')); return Number.isFinite(n) ? n : fallback; }

function detectDiscipline({ name='', rawField='', category='', notes='' }, learningRules={}) {
  const text = normalizeText([name, rawField, category, notes].join(' '));
  for (const [key, value] of Object.entries(learningRules)) {
    if (key && text.includes(normalizeText(key))) return { discipline:value, confidence:98, reason:`למידה מתיקון קודם: ${key}` };
  }
  let best = { discipline:'כללי', score:0, hits:[] };
  Object.entries(KEYWORDS).forEach(([discipline, words]) => {
    const hits = words.filter(w => text.includes(normalizeText(w)));
    const score = hits.length;
    if (score > best.score) best = { discipline, score, hits };
  });
  const confidence = best.score >= 3 ? 92 : best.score === 2 ? 78 : best.score === 1 ? 58 : 20;
  return { discipline: best.discipline, confidence, reason: best.hits.length ? `זוהו מילים: ${best.hits.slice(0,4).join(', ')}` : 'לא זוהו מילות מפתח — סווג ככללי' };
}

function mapExcelRow(row, i, learningRules) {
  const name = pick(row, ['שם ספק','ספק','Supplier','supplier','Vendor','vendor','שם החברה','חברה']);
  const rawField = pick(row, ['תחום עיסוק','תחום','דיסציפלינה','Discipline','Category','תיאור','תאור','תאור משאב','תיאור משאב','הערות']);
  const contact = pick(row, ['איש קשר','איש קשר ספק','Contact','contact']);
  const phone = pick(row, ['טלפון','נייד','Phone','phone','Mobile','mobile']);
  const email = pick(row, ['מייל','אימייל','Email','email']);
  const address = pick(row, ['כתובת','Address','address','מדינה','Country']);
  const status = pick(row, ['סטטוס','Status','status']) || 'חדש';
  const existingDiscipline = pick(row, ['דיסציפלינה','תחום עיסוק','תחום','Discipline']);
  const detection = existingDiscipline ? { discipline: existingDiscipline, confidence: 100, reason:'נלקח ישירות מעמודת תחום/דיסציפלינה באקסל' } : detectDiscipline({ name, rawField }, learningRules);
  return {
    id: crypto.randomUUID ? crypto.randomUUID() : `row-${Date.now()}-${i}`,
    name: name || `ספק ללא שם ${i+1}`,
    contact, phone, email, address,
    rawField,
    discipline: detection.discipline,
    confidence: detection.confidence,
    reason: detection.reason,
    rating: number(pick(row, ['דירוג','Rating','rating']), 0),
    status,
    source: 'Excel',
    createdAt: new Date().toISOString()
  };
}

function App() {
  const [suppliers, setSuppliers] = useState([]);
  const [learningRules, setLearningRules] = useState({});
  const [query, setQuery] = useState('');
  const [discipline, setDiscipline] = useState('הכל');
  const [status, setStatus] = useState('הכל');
  const [selected, setSelected] = useState(null);

  useEffect(() => {
    const saved = localStorage.getItem(STORAGE_KEY);
    const rules = localStorage.getItem(RULES_KEY);
    setLearningRules(rules ? JSON.parse(rules) : {});
    setSuppliers(saved ? JSON.parse(saved) : DEFAULT_SUPPLIERS);
  }, []);

  useEffect(() => { if (suppliers.length) localStorage.setItem(STORAGE_KEY, JSON.stringify(suppliers)); }, [suppliers]);
  useEffect(() => { localStorage.setItem(RULES_KEY, JSON.stringify(learningRules)); }, [learningRules]);

  const stats = useMemo(() => {
    const total = suppliers.length;
    const approved = suppliers.filter(s => s.status === 'מאושר').length;
    const avg = total ? (suppliers.reduce((a,s)=>a+Number(s.rating||0),0)/total).toFixed(1) : 0;
    return { total, approved, avg };
  }, [suppliers]);

  const filtered = useMemo(() => suppliers.filter(s => {
    const text = normalizeText([s.name,s.discipline,s.rawField,s.contact,s.email,s.phone].join(' '));
    const q = normalizeText(query);
    return (!q || text.includes(q)) && (discipline === 'הכל' || s.discipline === discipline) && (status === 'הכל' || s.status === status);
  }), [suppliers, query, discipline, status]);

  const handleExcel = async (e) => {
    const file = e.target.files?.[0];
    if (!file) return;
    const buffer = await file.arrayBuffer();
    const wb = XLSX.read(buffer, { type:'array' });
    const ws = wb.Sheets[wb.SheetNames[0]];
    const rows = XLSX.utils.sheet_to_json(ws, { defval:'' });
    const parsed = rows.map((row, i) => mapExcelRow(row, i, learningRules));
    setSuppliers(parsed);
  };

  const deleteSupplier = (id) => setSuppliers(prev => prev.filter(s => s.id !== id));
  const updateSupplier = (id, patch) => setSuppliers(prev => prev.map(s => s.id === id ? { ...s, ...patch } : s));

  const changeDiscipline = (supplier, newDiscipline) => {
    updateSupplier(supplier.id, { discipline:newDiscipline, confidence:100, reason:'תוקן ידנית על ידי המשתמש' });
    const key = supplier.name.split(' ').slice(0,2).join(' ');
    if (key.length >= 3) setLearningRules(prev => ({ ...prev, [key]: newDiscipline }));
  };

  const rerunClassification = () => {
    setSuppliers(prev => prev.map(s => {
      const d = detectDiscipline(s, learningRules);
      return { ...s, discipline:d.discipline, confidence:d.confidence, reason:d.reason };
    }));
  };

  const exportExcel = () => {
    const data = suppliers.map(s => ({ 'שם ספק':s.name, 'תחום/דיסציפלינה':s.discipline, 'דירוג':s.rating, 'סטטוס':s.status, 'איש קשר':s.contact, 'טלפון':s.phone, 'מייל':s.email, 'תיאור/תחום עיסוק':s.rawField, 'סיבת סיווג':s.reason }));
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, XLSX.utils.json_to_sheet(data), 'Suppliers');
    XLSX.writeFile(wb, 'galil_suppliers_v6.xlsx');
  };

  const resetDemo = () => setSuppliers(DEFAULT_SUPPLIERS);

  return <div dir="rtl" className="app">
    <header className="hero">
      <div className="heroGlass">
        <div className="brand"><div className="logo">GG</div><div><div className="eyebrow">GALIL GROUP · Supplier Intelligence</div><h1>מאגר ספקים חכם</h1><p>חיפוש ספקים לפי שם ותחום, העלאת Excel, סיווג אוטומטי, דירוג ומחיקה.</p></div></div>
        <div className="actions">
          <label className="uploadBtn"><Upload size={18}/> העלאת Excel<input type="file" accept=".xlsx,.xls,.csv" onChange={handleExcel}/></label>
          <button onClick={rerunClassification}><Brain size={18}/> סווג מחדש</button>
          <button onClick={exportExcel}><Download size={18}/> ייצוא Excel</button>
          <button onClick={resetDemo}><RefreshCcw size={18}/> דוגמה</button>
        </div>
      </div>
    </header>

    <main className="container">
      <section className="statsGrid">
        <div className="stat"><span>סה״כ ספקים</span><strong>{stats.total}</strong></div>
        <div className="stat"><span>ספקים מאושרים</span><strong>{stats.approved}</strong></div>
        <div className="stat"><span>דירוג ממוצע</span><strong>{stats.avg}</strong></div>
        <div className="stat"><span>כללי סיווג שנלמדו</span><strong>{Object.keys(learningRules).length}</strong></div>
      </section>

      <section className="panel controls">
        <div className="searchBox"><Search size={18}/><input value={query} onChange={e=>setQuery(e.target.value)} placeholder="חיפוש לפי שם ספק, תחום, איש קשר, מייל או טלפון"/></div>
        <div className="selectWrap"><Filter size={16}/><select value={discipline} onChange={e=>setDiscipline(e.target.value)}><option>הכל</option>{DISCIPLINES.map(d=><option key={d}>{d}</option>)}</select></div>
        <select value={status} onChange={e=>setStatus(e.target.value)}><option>הכל</option><option>מאושר</option><option>בבדיקה</option><option>חדש</option><option>לא מאושר</option></select>
      </section>

      <section className="grid">
        {filtered.map(s => <article className="card" key={s.id}>
          <div className="cardTop"><div><div className="discipline">{s.discipline}</div><h3>{s.name}</h3></div><button className="danger" onClick={()=>deleteSupplier(s.id)}><Trash2 size={17}/></button></div>
          <p className="desc">{s.rawField || 'אין תיאור תחום עיסוק'}</p>
          <div className="meta"><span>איש קשר: {s.contact || '-'}</span><span>טלפון: {s.phone || '-'}</span><span>מייל: {s.email || '-'}</span></div>
          <div className="confidence"><span>אמון סיווג: {s.confidence || 0}%</span><small>{s.reason}</small></div>
          <div className="row"><select value={s.discipline} onChange={e=>changeDiscipline(s, e.target.value)}>{DISCIPLINES.map(d=><option key={d}>{d}</option>)}</select><select value={s.status} onChange={e=>updateSupplier(s.id,{status:e.target.value})}><option>חדש</option><option>בבדיקה</option><option>מאושר</option><option>לא מאושר</option></select></div>
          <div className="stars">{[1,2,3,4,5].map(n=><button key={n} onClick={()=>updateSupplier(s.id,{rating:n})} className={n <= s.rating ? 'star active':'star'}><Star size={20} fill="currentColor"/></button>)}</div>
          <button className="edit" onClick={()=>setSelected(s)}><Edit3 size={16}/> עריכת פרטים</button>
        </article>)}
      </section>
    </main>

    {selected && <div className="modalBackdrop" onClick={()=>setSelected(null)}><div className="modal" onClick={e=>e.stopPropagation()}><h2>עריכת ספק</h2>{['name','contact','phone','email','rawField','address'].map(key=><label key={key}>{labelFor(key)}<input value={selected[key] || ''} onChange={e=>setSelected({...selected,[key]:e.target.value})}/></label>)}<div className="modalActions"><button onClick={()=>{updateSupplier(selected.id, selected); setSelected(null)}}><Save size={17}/> שמור</button><button onClick={()=>setSelected(null)}>סגור</button></div></div></div>}
  </div>
}
function labelFor(key){return {name:'שם ספק',contact:'איש קשר',phone:'טלפון',email:'מייל',rawField:'תחום עיסוק / תיאור',address:'כתובת'}[key] || key}

createRoot(document.getElementById('root')).render(<App />);
