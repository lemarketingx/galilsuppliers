import React, { useEffect, useMemo, useRef, useState } from 'react';
import { createRoot } from 'react-dom/client';
import * as XLSX from 'xlsx';
import jsPDF from 'jspdf';
import html2canvas from 'html2canvas';
import { Search, Upload, Trash2, Star, Calculator, Users, FileSpreadsheet, Download, Printer, Save, RotateCcw, Building2, Filter, Plus, X } from 'lucide-react';
import './style.css';

const logo = '/galil-logo.webp';
const fmt = n => new Intl.NumberFormat('he-IL', { maximumFractionDigits: 0 }).format(Number(n) || 0);
const money = (n, cur='₪') => `${fmt(n)} ${cur || '₪'}`;
const clean = v => String(v ?? '').trim();
const norm = v => clean(v).replace(/\s+/g, '').replace(/["'׳״]/g, '').toLowerCase();
const get = (row, names) => {
  const map = Object.keys(row).reduce((a,k)=>{a[norm(k)] = row[k]; return a;}, {});
  for (const name of names) if (map[norm(name)] !== undefined) return map[norm(name)];
  return '';
};
const toNum = v => {
  if (typeof v === 'number') return v;
  const s = clean(v).replace(/,/g,'').replace(/[₪$€]/g,'');
  return Number(s) || 0;
};
const readExcel = (file, cb) => {
  const reader = new FileReader();
  reader.onload = e => {
    const wb = XLSX.read(new Uint8Array(e.target.result), { type: 'array', cellDates: true });
    const sheet = wb.Sheets[wb.SheetNames[0]];
    cb(XLSX.utils.sheet_to_json(sheet, { defval: '' }));
  };
  reader.readAsArrayBuffer(file);
};
const exportExcel = (rows, name) => {
  const ws = XLSX.utils.json_to_sheet(rows);
  const wb = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(wb, ws, 'Data');
  XLSX.writeFile(wb, name);
};
const disciplines = ['כל התחומים','צנרת','חשמל','הנדסה אזרחית','מכשור ובקרה','מכונות','בטיחות','לוגיסטיקה ושילוח','בידוד וצבע','ציוד תהליך','כללי'];
const detectDiscipline = (...parts) => {
  const t = parts.join(' ').toLowerCase();
  const rules = [
    ['צנרת', ['pipe','piping','valve','flange','fitting','flow','pump','hydro','צנרת','צינור','ברז','שסתום','אוגן','מחבר','אביזרי צנרת']],
    ['חשמל', ['electric','electrical','power','cable','panel','siemens','schneider','abb','חשמל','כבל','לוח','ארון חשמל','מתח','הארקה']],
    ['הנדסה אזרחית', ['civil','concrete','construction','building','contractor','בטון','בינוי','אזרחי','קבלן','עפר','יציקה','קונסטרוקציה']],
    ['מכשור ובקרה', ['instrument','control','automation','plc','sensor','transmitter','מכשור','בקרה','אוטומציה','חיישן','משדר']],
    ['מכונות', ['machine','mechanical','motor','gear','conveyor','bearing','מכונות','מנוע','מסוע','גיר','מיסב']],
    ['בטיחות', ['safety','fire','ppe','בטיחות','כיבוי','אש','מגן']],
    ['לוגיסטיקה ושילוח', ['shipping','freight','dhl','logistics','transport','שילוח','הובלה','לוגיסטיקה','עמילות','מכס']],
    ['בידוד וצבע', ['paint','coating','insulation','sandblast','צבע','צביעה','בידוד','ציפוי','ניקוי חול']],
    ['ציוד תהליך', ['process','silo','tank','vessel','filter','mixer','reactor','מיכל','סילו','פילטר','מערבל','ריאקטור']]
  ];
  for (const [d, words] of rules) if (words.some(w => t.includes(w))) return d;
  return 'כללי';
};

const demoBoq = [
  {id:'b1', section:'צינור CS Sch.40 בקוטר 2" כולל התקנה', supplier:'ספק דוגמה צנרת', sku:'P-002', qty:25, unitPrice:240, total:6000, currency:'₪', project:'T1051', resource:'צנרת תהליך', discipline:'צנרת'},
  {id:'b2', section:'כבל N2XY 5x10 כולל השחלה וחיבור', supplier:'ספק דוגמה חשמל', sku:'E-510', qty:120, unitPrice:88, total:10560, currency:'₪', project:'T1051', resource:'חשמל', discipline:'חשמל'},
  {id:'b3', section:'יציקת בטון C30 כולל טפסנות וברזל', supplier:'קבלן דוגמה אזרחי', sku:'C-030', qty:18, unitPrice:1350, total:24300, currency:'₪', project:'T1051', resource:'עבודות בטון', discipline:'הנדסה אזרחית'}
];
const demoSuppliers = [
  {id:'s1', name:'Technobar', field:'מכשור, צנרת ושסתומים', discipline:'צנרת', contact:'Adi', phone:'', email:'', rating:4, status:'מאושר', notes:'ספק דוגמה'},
  {id:'s2', name:'ABB', field:'Electrical panels, drives and automation', discipline:'חשמל', contact:'', phone:'', email:'', rating:5, status:'מאושר', notes:'ספק דוגמה'},
  {id:'s3', name:'Civil Concrete Ltd', field:'Concrete and civil works', discipline:'הנדסה אזרחית', contact:'', phone:'', email:'', rating:3, status:'בדיקה', notes:'ספק דוגמה'}
];

function App(){
  const [tab,setTab] = useState('boq');
  return <div dir="rtl" className="app">
    <header className="hero">
      <div className="heroGlow" />
      <div className="headerContent">
        <div className="brand"><img src={logo}/><div><div className="eyebrow">GALIL GROUP · Engineering Tools</div><h1>מערכת הנדסית מאוחדת</h1><p>מחירון כתבי כמויות ומאגר ספקים באתר אחד, עם ייבוא Excel, חיפוש, סינון וייצוא.</p></div></div>
        <div className="tabs"><button className={tab==='boq'?'active':''} onClick={()=>setTab('boq')}><Calculator size={18}/> מחירון / BOQ</button><button className={tab==='suppliers'?'active':''} onClick={()=>setTab('suppliers')}><Users size={18}/> רשימת ספקים</button></div>
      </div>
    </header>
    <main>{tab==='boq'?<BoqModule/>:<SuppliersModule/>}</main>
  </div>
}

function BoqModule(){
  const [items,setItems] = useState(()=>JSON.parse(localStorage.getItem('galil_boq_items')||'null') || demoBoq);
  const [cart,setCart] = useState(()=>JSON.parse(localStorage.getItem('galil_boq_cart')||'[]'));
  const [query,setQuery] = useState(''); const [disc,setDisc] = useState('כל התחומים'); const [project,setProject] = useState(''); const reportRef = useRef(null);
  useEffect(()=>localStorage.setItem('galil_boq_items',JSON.stringify(items)),[items]);
  useEffect(()=>localStorage.setItem('galil_boq_cart',JSON.stringify(cart)),[cart]);
  const filtered = useMemo(()=>items.filter(i=>(disc==='כל התחומים'||i.discipline===disc) && [i.section,i.supplier,i.sku,i.resource,i.project].join(' ').toLowerCase().includes(query.toLowerCase())),[items,query,disc]);
  const total = cart.reduce((s,i)=>s+(Number(i.total)||0),0);
  const load = e => { const file=e.target.files?.[0]; if(!file)return; readExcel(file, rows=>{
    const parsed = rows.map((r,i)=>{ const section=clean(get(r,['תאור הסעיף/פרק','תיאור הסעיף/פרק','תיאור','שם פריט'])); const supplier=clean(get(r,['ספק','שם ספק','שם ספק/קבלן'])); const resource=clean(get(r,['תאור משאב','תיאור משאב','תחום עיסוק'])); const total=toNum(get(r,['מחיר כולל מעמ','מחיר כולל מע"מ','מחיר כולל מע״מ','סהכ','סה"כ'])); const unit=toNum(get(r,["מחיר ליח' לפני הנחה",'מחיר ליחידה לפני הנחה','מחיר ליח'])); return {id:'boq-'+Date.now()+'-'+i, section: section||resource||'סעיף ללא תיאור', supplier, sku:clean(get(r,['מקט','מק"ט','מק״ט','Part Number'])), qty:toNum(get(r,['כמות','Qty','Quantity'])) || 1, unitPrice: unit || total, total: total || unit, currency:clean(get(r,['מטבע חוזה','מטבע','Currency'])) || '₪', project:clean(get(r,['פרויקט','Project'])), projectDesc:clean(get(r,['תאור פרויקט','תיאור פרויקט'])), resource, quoteDate:clean(get(r,['תאריך הגשת הצעת מחיר','תאריך הצעה'])), discipline: detectDiscipline(section,supplier,resource)} }); setItems(parsed); setCart([]); }); };
  const add = item => setCart(p=>{const ex=p.find(x=>x.id===item.id); return ex?p.map(x=>x.id===item.id?{...x, selectedQty:(x.selectedQty||x.qty||1)+1,total:(x.unitPrice||x.total)*((x.selectedQty||x.qty||1)+1)}: [...p,{...item,selectedQty:item.qty||1}];});
  const pdf = async()=>{ const el=reportRef.current; if(!el)return; const canvas=await html2canvas(el,{scale:2,useCORS:true,backgroundColor:'#ffffff'}); const img=canvas.toDataURL('image/png'); const doc=new jsPDF('p','mm','a4'); const w=210; const h=canvas.height*w/canvas.width; doc.addImage(img,'PNG',0,0,w,h); doc.save('galil-boq-report.pdf'); };
  return <section className="pageGrid">
    <div className="panel wide"><PanelTitle icon={<FileSpreadsheet/>} title="מחירון כתבי כמויות" sub="ייבוא Excel לפי העמודות שלך, בחירת סעיפים וחישוב אומדן" />
      <div className="toolbar"><label className="upload"><Upload size={18}/> העלאת Excel<input type="file" accept=".xlsx,.xls" onChange={load}/></label><div className="search"><Search size={18}/><input placeholder="חיפוש לפי סעיף, ספק, מק״ט או פרויקט" value={query} onChange={e=>setQuery(e.target.value)}/></div><select value={disc} onChange={e=>setDisc(e.target.value)}>{disciplines.map(d=><option key={d}>{d}</option>)}</select></div>
      <div className="cards">{filtered.map(item=><div className="card" key={item.id}><div><span className="tag">{item.discipline}</span><h3>{item.section}</h3><p>{item.supplier || 'ללא ספק'} · {item.sku || 'ללא מק״ט'} · {item.resource}</p></div><div className="price"><b>{money(item.total,item.currency)}</b><small>כמות: {fmt(item.qty)}</small><button onClick={()=>add(item)}><Plus size={16}/> הוסף</button></div></div>)}</div>
    </div>
    <div className="panel side"><PanelTitle icon={<Calculator/>} title="סל חישוב" sub="הפריטים שנבחרו לפרויקט"/><input className="fullInput" placeholder="שם פרויקט לדוח" value={project} onChange={e=>setProject(e.target.value)}/><div className="cart">{cart.map(i=><div className="cartRow" key={i.id}><span>{i.section}</span><b>{money(i.total,i.currency)}</b><button onClick={()=>setCart(p=>p.filter(x=>x.id!==i.id))}><X size={15}/></button></div>)}</div><div className="totalBox"><span>סה״כ אומדן</span><b>{money(total,'₪')}</b></div><div className="actions"><button onClick={()=>exportExcel(cart,'galil-boq-cart.xlsx')}><Download size={16}/> Excel</button><button onClick={pdf}><Printer size={16}/> PDF</button><button onClick={()=>setCart([])}><RotateCcw size={16}/> נקה</button></div></div>
    <div className="pdfReport" ref={reportRef}><img src={logo}/><h2>Galil Group - BOQ Report</h2><p>פרויקט: {project || 'ללא שם'}</p><table><thead><tr><th>דיסציפלינה</th><th>תיאור סעיף</th><th>ספק</th><th>כמות</th><th>סה״כ</th></tr></thead><tbody>{cart.map(i=><tr key={i.id}><td>{i.discipline}</td><td>{i.section}</td><td>{i.supplier}</td><td>{i.qty}</td><td>{money(i.total,i.currency)}</td></tr>)}</tbody></table><h3>סה״כ: {money(total,'₪')}</h3></div>
  </section>
}

function SuppliersModule(){
  const [suppliers,setSuppliers] = useState(()=>JSON.parse(localStorage.getItem('galil_suppliers')||'null') || demoSuppliers);
  const [query,setQuery] = useState(''); const [disc,setDisc]=useState('כל התחומים');
  useEffect(()=>localStorage.setItem('galil_suppliers',JSON.stringify(suppliers)),[suppliers]);
  const load = e => {const file=e.target.files?.[0]; if(!file)return; readExcel(file, rows=>{ const parsed=rows.map((r,i)=>{ const name=clean(get(r,['שם ספק/קבלן','שם ספק','ספק','Supplier','Vendor','שם'])); const field=clean(get(r,['תחום עיסוק','תחום','תחום פעילות','תיאור','תאור','תאור משאב'])); const notes=clean(get(r,['הערות','Notes','תיאור'])); return {id:'sup-'+Date.now()+'-'+i,name:name||'ספק ללא שם',field,discipline: clean(get(r,['דיסציפלינה','Discipline'])) || detectDiscipline(name,field,notes),contact:clean(get(r,['איש קשר','Contact','שם איש קשר'])),phone:clean(get(r,['טלפון','נייד','Phone'])),email:clean(get(r,['מייל','אימייל','Email'])),rating:toNum(get(r,['דירוג','Rating']))||0,status:clean(get(r,['סטטוס','Status']))||'חדש',notes}; }); setSuppliers(parsed); });};
  const filtered = suppliers.filter(s=>(disc==='כל התחומים'||s.discipline===disc) && [s.name,s.field,s.contact,s.email,s.notes,s.discipline].join(' ').toLowerCase().includes(query.toLowerCase()));
  const update=(id,patch)=>setSuppliers(p=>p.map(s=>s.id===id?{...s,...patch}:s));
  const remove=id=>setSuppliers(p=>p.filter(s=>s.id!==id));
  return <section className="panel suppliersPage"><PanelTitle icon={<Users/>} title="מאגר ספקים" sub="חיפוש לפי שם ספק ותחום, סיווג אוטומטי, דירוג ומחיקה" />
    <div className="toolbar"><label className="upload"><Upload size={18}/> העלאת Excel ספקים<input type="file" accept=".xlsx,.xls" onChange={load}/></label><div className="search"><Search size={18}/><input placeholder="חפש ספק, תחום, איש קשר או הערה" value={query} onChange={e=>setQuery(e.target.value)}/></div><select value={disc} onChange={e=>setDisc(e.target.value)}>{disciplines.map(d=><option key={d}>{d}</option>)}</select><button onClick={()=>exportExcel(suppliers,'galil-suppliers.xlsx')}><Download size={16}/> ייצוא</button></div>
    <div className="stats"><Stat label="ספקים" value={suppliers.length}/><Stat label="בתצוגה" value={filtered.length}/><Stat label="מאושרים" value={suppliers.filter(s=>s.status.includes('מאושר')).length}/></div>
    <div className="supplierGrid">{filtered.map(s=><div className="supplierCard" key={s.id}><div className="supplierHead"><div><span className="tag">{s.discipline}</span><h3>{s.name}</h3><p>{s.field || 'ללא תחום עיסוק'}</p></div><button className="danger" onClick={()=>remove(s.id)}><Trash2 size={17}/></button></div><div className="meta"><span>{s.contact || 'אין איש קשר'}</span><span>{s.phone || 'אין טלפון'}</span><span>{s.email || 'אין מייל'}</span></div><div className="supplierControls"><select value={s.discipline} onChange={e=>update(s.id,{discipline:e.target.value})}>{disciplines.filter(d=>d!=='כל התחומים').map(d=><option key={d}>{d}</option>)}</select><select value={s.status} onChange={e=>update(s.id,{status:e.target.value})}><option>חדש</option><option>בדיקה</option><option>מאושר</option><option>לא פעיל</option></select></div><div className="stars">{[1,2,3,4,5].map(n=><button key={n} onClick={()=>update(s.id,{rating:n})} className={n<=s.rating?'on':''}><Star size={21} fill="currentColor"/></button>)}</div>{s.notes&&<p className="notes">{s.notes}</p>}</div>)}</div>
  </section>
}
function PanelTitle({icon,title,sub}){return <div className="panelTitle"><div className="iconBox">{icon}</div><div><h2>{title}</h2><p>{sub}</p></div></div>}
function Stat({label,value}){return <div className="stat"><span>{label}</span><b>{value}</b></div>}

createRoot(document.getElementById('root')).render(<App/>);
