import React, { useMemo, useState } from 'react';
import { createRoot } from 'react-dom/client';
import * as XLSX from 'xlsx';
import { Search, Upload, FileSpreadsheet, Building2, Phone, Mail, MapPin, Star, Download, Filter, Trash2, Plus, CheckCircle, AlertCircle } from 'lucide-react';
import './style.css';

const domains = [
  'צנרת ואביזרים','חשמל ולוחות','מכשור ובקרה','הנדסה אזרחית','קונסטרוקציה ומתכת','משאבות ומנועים','HVAC ומיזוג','בידוד וצבע','בטיחות וכיבוי אש','לוגיסטיקה ושילוח','קבלני ביצוע','שירותי תכנון','ציוד תעשייתי','כימיקלים וחומרים','כללי'
];

const demoSuppliers = [
  {name:'Technobar', domain:'מכשור ובקרה', contact:'Adi Cohen', phone:'03-0000000', email:'sales@technobar.example', city:'מרכז', status:'מאושר', rating:5, notes:'ברזים, מכשור, ציוד תהליך'},
  {name:'Hapman', domain:'ציוד תעשייתי', contact:'Lucas Whitney', phone:'+1-000-0000', email:'sales@hapman.example', city:'USA', status:'פעיל', rating:4, notes:'מערכות שינוע, פריקת שקים, ציוד תעשייתי'},
  {name:'ABB Israel', domain:'חשמל ולוחות', contact:'Sales Team', phone:'03-0000000', email:'info@abb.example', city:'ישראל', status:'מאושר', rating:4, notes:'ציוד חשמל, בקרה, מנועים'},
  {name:'Delfin S.R.L', domain:'ציוד תעשייתי', contact:'Export Dept.', phone:'+39-000000', email:'export@delfin.example', city:'Italy', status:'פעיל', rating:4, notes:'Loading systems וציוד תעשייתי'},
  {name:'Civil Pro Contractors', domain:'הנדסה אזרחית', contact:'Project Manager', phone:'050-0000000', email:'office@civilpro.example', city:'דרום', status:'בבדיקה', rating:3, notes:'בטון, עבודות עפר, קבלנות אזרחית'},
  {name:'PipeLine Industrial', domain:'צנרת ואביזרים', contact:'Ronen Levi', phone:'052-0000000', email:'sales@pipeline.example', city:'חיפה', status:'מאושר', rating:5, notes:'צנרת תהליך, פלנזים, אביזרי NPT'},
];

function normalizeSupplier(row, index) {
  const get = (...keys) => keys.find(k => row[k] !== undefined && row[k] !== null && String(row[k]).trim() !== '') ? row[keys.find(k => row[k] !== undefined && row[k] !== null && String(row[k]).trim() !== '')] : '';
  return {
    id: `${Date.now()}-${index}`,
    name: String(get('שם ספק','ספק','Supplier','supplier','name','Name') || `ספק ${index + 1}`),
    domain: String(get('תחום','דיסציפלינה','תחום פעילות','Category','category','Domain','domain') || 'כללי'),
    contact: String(get('איש קשר','נציג','Contact','contact') || ''),
    phone: String(get('טלפון','נייד','Phone','phone','Mobile','mobile') || ''),
    email: String(get('אימייל','מייל','Email','email') || ''),
    city: String(get('עיר','מיקום','אזור','City','city','Location','location') || ''),
    status: String(get('סטטוס','Status','status') || 'פעיל'),
    rating: Number(get('דירוג','Rating','rating')) || 3,
    notes: String(get('הערות','תיאור','Notes','notes','Description','description') || ''),
  };
}

function App() {
  const [suppliers, setSuppliers] = useState(demoSuppliers.map((s,i)=>({...s,id:i})));
  const [query, setQuery] = useState('');
  const [domain, setDomain] = useState('הכל');
  const [status, setStatus] = useState('הכל');
  const [selected, setSelected] = useState(null);

  const filtered = useMemo(() => suppliers.filter(s => {
    const q = query.trim().toLowerCase();
    const text = `${s.name} ${s.domain} ${s.contact} ${s.email} ${s.city} ${s.notes}`.toLowerCase();
    return (!q || text.includes(q)) && (domain === 'הכל' || s.domain === domain) && (status === 'הכל' || s.status === status);
  }), [suppliers, query, domain, status]);

  const stats = useMemo(() => ({ total: suppliers.length, approved: suppliers.filter(s=>s.status==='מאושר').length, fields: new Set(suppliers.map(s=>s.domain)).size }), [suppliers]);

  const handleExcelUpload = (e) => {
    const file = e.target.files?.[0];
    if (!file) return;
    const reader = new FileReader();
    reader.onload = (evt) => {
      const data = new Uint8Array(evt.target.result);
      const workbook = XLSX.read(data, { type: 'array' });
      const sheet = workbook.Sheets[workbook.SheetNames[0]];
      const rows = XLSX.utils.sheet_to_json(sheet);
      const parsed = rows.map(normalizeSupplier).filter(s => s.name && s.name !== '-');
      setSuppliers(parsed.length ? parsed : suppliers);
      alert(`נטענו ${parsed.length} ספקים מהאקסל`);
    };
    reader.readAsArrayBuffer(file);
  };

  const exportTemplate = () => {
    const template = [{ 'שם ספק':'ספק לדוגמה', 'תחום':'צנרת ואביזרים', 'איש קשר':'ישראל ישראלי', 'טלפון':'050-0000000', 'אימייל':'sales@example.com', 'עיר':'באר שבע', 'סטטוס':'מאושר', 'דירוג':5, 'הערות':'אביזרי צנרת, ברזים, פלנזים' }];
    const ws = XLSX.utils.json_to_sheet(template);
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, 'Suppliers');
    XLSX.writeFile(wb, 'galil_suppliers_template.xlsx');
  };

  const exportData = () => {
    const ws = XLSX.utils.json_to_sheet(filtered.map(({id,...s})=>s));
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, 'Suppliers');
    XLSX.writeFile(wb, 'galil_suppliers_export.xlsx');
  };

  const resetDemo = () => setSuppliers(demoSuppliers.map((s,i)=>({...s,id:i})));

  return <div className="app" dir="rtl">
    <header className="hero">
      <div className="hero-bg" />
      <div className="hero-inner">
        <div className="brand">
          <div className="logo">GG</div>
          <div>
            <p className="eyebrow">GALIL GROUP · SUPPLIER INTELLIGENCE</p>
            <h1>מאגר ספקים וקבלנים</h1>
            <p className="subtitle">חיפוש ספקים לפי שם, תחום פעילות, סטטוס, מיקום ופרטי קשר — כולל העלאת אקסל מלאה.</p>
          </div>
        </div>
        <div className="upload-card">
          <FileSpreadsheet size={22}/>
          <strong>ייבוא ספקים מאקסל</strong>
          <span>תומך בעמודות: שם ספק, תחום, איש קשר, טלפון, אימייל, עיר, סטטוס, דירוג, הערות</span>
          <label className="upload-btn"><Upload size={18}/> העלה Excel<input type="file" accept=".xlsx,.xls,.csv" onChange={handleExcelUpload}/></label>
        </div>
      </div>
    </header>

    <main className="container">
      <section className="stats-grid">
        <div className="stat"><span>סה״כ ספקים</span><b>{stats.total}</b></div>
        <div className="stat"><span>ספקים מאושרים</span><b>{stats.approved}</b></div>
        <div className="stat"><span>תחומי פעילות</span><b>{stats.fields}</b></div>
      </section>

      <section className="panel">
        <div className="toolbar">
          <div className="searchbox"><Search size={20}/><input value={query} onChange={e=>setQuery(e.target.value)} placeholder="חפש לפי שם ספק, תחום, איש קשר, עיר או הערות"/></div>
          <div className="select-wrap"><Filter size={18}/><select value={domain} onChange={e=>setDomain(e.target.value)}><option>הכל</option>{domains.map(d=><option key={d}>{d}</option>)}</select></div>
          <select value={status} onChange={e=>setStatus(e.target.value)}><option>הכל</option><option>מאושר</option><option>פעיל</option><option>בבדיקה</option><option>חסום</option></select>
        </div>
        <div className="actions">
          <button onClick={exportTemplate}><Download size={17}/> הורד תבנית Excel</button>
          <button onClick={exportData}><Download size={17}/> יצוא תוצאות</button>
          <button onClick={resetDemo} className="ghost"><Trash2 size={17}/> איפוס לדמו</button>
        </div>
      </section>

      <section className="content-grid">
        <div className="supplier-list">
          {filtered.map(s => <article className={`supplier-card ${selected?.id===s.id ? 'active':''}`} key={s.id} onClick={()=>setSelected(s)}>
            <div className="card-head"><div><h3>{s.name}</h3><p>{s.domain}</p></div><StatusBadge status={s.status}/></div>
            <div className="meta"><span><MapPin size={16}/>{s.city || 'לא צוין'}</span><span><Star size={16}/>{s.rating}/5</span></div>
            <p className="notes">{s.notes || 'אין הערות'}</p>
          </article>)}
        </div>
        <aside className="details">
          {selected ? <>
            <div className="details-icon"><Building2/></div>
            <h2>{selected.name}</h2>
            <StatusBadge status={selected.status}/>
            <div className="detail-row"><b>תחום:</b><span>{selected.domain}</span></div>
            <div className="detail-row"><b>איש קשר:</b><span>{selected.contact || '-'}</span></div>
            <div className="detail-row"><b>טלפון:</b><span><Phone size={15}/>{selected.phone || '-'}</span></div>
            <div className="detail-row"><b>אימייל:</b><span><Mail size={15}/>{selected.email || '-'}</span></div>
            <div className="detail-row"><b>מיקום:</b><span>{selected.city || '-'}</span></div>
            <div className="detail-row"><b>דירוג:</b><span>{'★'.repeat(Math.min(5, selected.rating))}</span></div>
            <div className="note-box"><b>הערות:</b><p>{selected.notes || '-'}</p></div>
          </> : <div className="empty"><Plus size={42}/><h2>בחר ספק להצגת פרטים</h2><p>לחץ על כרטיס ספק מהרשימה.</p></div>}
        </aside>
      </section>
    </main>
  </div>
}

function StatusBadge({status}) { const approved = status === 'מאושר'; return <span className={`badge ${approved?'approved':'regular'}`}>{approved ? <CheckCircle size={14}/> : <AlertCircle size={14}/>} {status}</span> }

createRoot(document.getElementById('root')).render(<App />);
