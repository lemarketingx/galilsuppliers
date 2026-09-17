import React, { createContext, useContext, useEffect, useState } from 'react';
import { LogOut, UserCog, X, Trash2, Plus, ShieldCheck, History, KeyRound } from 'lucide-react';

const API_URL = import.meta.env.VITE_API_URL || '';
const TOKEN_KEY = 'galil_auth_token_v1';
const LOGO_URL = '/brand/galil-logo.png';

async function apiFetch(path, opts = {}) {
  const token = localStorage.getItem(TOKEN_KEY);
  const res = await fetch(`${API_URL}${path}`, {
    ...opts,
    headers: { 'Content-Type': 'application/json', ...(token ? { Authorization: `Bearer ${token}` } : {}), ...(opts.headers || {}) },
  });
  let data = null;
  try { data = await res.json(); } catch {}
  if (!res.ok) throw new Error(data?.error || `שגיאת שרת (${res.status}).`);
  return data;
}

/* Sends the raw file to the auth server first. The server checks the caller's
   token + role (uploader/admin) before accepting it and logs who uploaded
   what, when. Only once this resolves does the caller's existing client-side
   parsing (Excel/PDF) run - without a valid, permitted login the server
   rejects the request and nothing gets processed. */
async function uploadFileToServer(file, kind) {
  const token = localStorage.getItem(TOKEN_KEY);
  if (!token) throw new Error('יש להתחבר לפני העלאת קבצים.');
  const body = new FormData();
  body.append('file', file);
  body.append('kind', kind);
  const res = await fetch(`${API_URL}/api/uploads`, { method: 'POST', headers: { Authorization: `Bearer ${token}` }, body });
  let data = null;
  try { data = await res.json(); } catch {}
  if (!res.ok) throw new Error(data?.error || `שגיאת שרת (${res.status}).`);
  return data.upload;
}

const ROLE_LABELS = { admin: 'מנהל מערכת', uploader: 'מורשה העלאה', viewer: 'צפייה בלבד' };

const AuthContext = createContext(null);
export const useAuth = () => useContext(AuthContext);

export function AuthProvider({ children }) {
  const [user, setUser] = useState(null);
  const [ready, setReady] = useState(false);

  useEffect(() => {
    const token = localStorage.getItem(TOKEN_KEY);
    if (!token) { setReady(true); return; }
    apiFetch('/api/auth/me').then(d => setUser(d.user)).catch(() => localStorage.removeItem(TOKEN_KEY)).finally(() => setReady(true));
  }, []);

  const login = async (username, password) => {
    const d = await apiFetch('/api/auth/login', { method: 'POST', body: JSON.stringify({ username, password }) });
    localStorage.setItem(TOKEN_KEY, d.token);
    setUser(d.user);
  };
  const logout = () => { localStorage.removeItem(TOKEN_KEY); setUser(null); };
  const forgotPassword = username => apiFetch('/api/auth/forgot-password', { method: 'POST', body: JSON.stringify({ username }) });

  const canUpload = !!user && (user.role === 'uploader' || user.role === 'admin');
  const isAdmin = !!user && user.role === 'admin';

  return <AuthContext.Provider value={{ user, ready, login, logout, forgotPassword, canUpload, isAdmin, roleLabel: user ? ROLE_LABELS[user.role] || user.role : '', apiFetch, uploadFile: uploadFileToServer }}>
    {children}
  </AuthContext.Provider>;
}

/* Full-screen gate: nothing in the app renders until the user is signed in. */
export function AuthGate({ children }) {
  const { user, ready } = useAuth();
  if (!ready) return <div className="authSplash"><img src={LOGO_URL} alt="" /></div>;
  if (!user) return <LoginScreen />;
  return children;
}

function LoginScreen() {
  const { login, forgotPassword } = useAuth();
  const [username, setUsername] = useState('');
  const [password, setPassword] = useState('');
  const [error, setError] = useState('');
  const [busy, setBusy] = useState(false);
  const [mode, setMode] = useState('login'); // 'login' | 'forgot'
  const [forgotUsername, setForgotUsername] = useState('');
  const [forgotSent, setForgotSent] = useState(false);

  const submit = async e => {
    e.preventDefault();
    if (!username || !password) return;
    setBusy(true); setError('');
    try { await login(username, password); }
    catch (err) { setError(err.message || 'שגיאה בהתחברות.'); }
    finally { setBusy(false); }
  };

  const submitForgot = async e => {
    e.preventDefault();
    if (!forgotUsername) return;
    setBusy(true); setError('');
    try { await forgotPassword(forgotUsername); setForgotSent(true); }
    catch (err) { setError(err.message || 'שגיאה בשליחת הבקשה.'); }
    finally { setBusy(false); }
  };

  return <div className="loginPage">
    <div className="loginCard">
      <img src={LOGO_URL} alt="קבוצת גליל" className="loginLogo" />
      <h1>{SYSTEM_TITLE_FALLBACK}</h1>

      {mode === 'login' ? <form onSubmit={submit}>
        <label>שם משתמש<input value={username} onChange={e => setUsername(e.target.value)} autoFocus /></label>
        <label>סיסמה<input type="password" value={password} onChange={e => setPassword(e.target.value)} /></label>
        {error && <div className="authError">{error}</div>}
        <button className="calc" type="submit" disabled={busy}>{busy ? 'מתחבר...' : 'התחברות'}</button>
        <button type="button" className="loginForgotLink" onClick={() => { setMode('forgot'); setError(''); }}>שכחתי סיסמה</button>
      </form> : <form onSubmit={submitForgot}>
        {forgotSent ? <p className="authHint">אם שם המשתמש קיים במערכת, בקשתך נשלחה למנהל המערכת והוא יצור איתך קשר עם סיסמה חדשה.</p> : <>
          <p className="authHint">הזן/י שם משתמש - הבקשה תישלח למנהל המערכת לאיפוס הסיסמה.</p>
          <label>שם משתמש<input value={forgotUsername} onChange={e => setForgotUsername(e.target.value)} autoFocus /></label>
          {error && <div className="authError">{error}</div>}
          <button className="calc" type="submit" disabled={busy}>{busy ? 'שולח...' : 'שליחת בקשה'}</button>
        </>}
        <button type="button" className="loginForgotLink" onClick={() => { setMode('login'); setError(''); setForgotSent(false); }}>חזרה להתחברות</button>
      </form>}
    </div>
  </div>;
}

const SYSTEM_TITLE_FALLBACK = 'מערכת הנדסה ורכש';

export function AuthBar() {
  const { user, logout, roleLabel, isAdmin } = useAuth();
  const [showAdmin, setShowAdmin] = useState(false);
  const [showLog, setShowLog] = useState(false);

  return <div className="authBar">
    <span className="authUser"><ShieldCheck size={15} /> {user.username} <em>({roleLabel})</em></span>
    {isAdmin && <button onClick={() => setShowAdmin(true)}><UserCog size={15} /> ניהול משתמשים</button>}
    {isAdmin && <button onClick={() => setShowLog(true)}><History size={15} /> יומן העלאות</button>}
    <button onClick={logout}><LogOut size={15} /> התנתקות</button>
    {showAdmin && <UsersAdmin onClose={() => setShowAdmin(false)} />}
    {showLog && <UploadsLog onClose={() => setShowLog(false)} />}
  </div>;
}

const KIND_LABELS = { 'boq-excel': 'Excel מחירון', 'suppliers-excel': 'Excel ספקים', attachment: 'צרופה' };
const fmtSize = n => n > 1e6 ? `${(n / 1e6).toFixed(1)}MB` : `${Math.round(n / 1024)}KB`;

function UploadsLog({ onClose }) {
  const { apiFetch } = useAuth();
  const [uploads, setUploads] = useState([]);
  const [loading, setLoading] = useState(true);
  const [error, setError] = useState('');

  useEffect(() => {
    apiFetch('/api/uploads').then(d => setUploads(d.uploads)).catch(e => setError(e.message)).finally(() => setLoading(false));
  }, []);

  return <div className="authOverlay" onClick={onClose}>
    <div className="authModal authAdminModal" onClick={e => e.stopPropagation()}>
      <div className="authModalHead"><h2><History size={20} /> יומן העלאות</h2><button onClick={onClose}><X size={18} /></button></div>
      <p className="authHint">רשימת כל הקבצים שהועלו לאתר, מי העלה ומתי.</p>
      {error && <div className="authError">{error}</div>}
      {loading ? <p>טוען...</p> : uploads.length === 0 ? <p>עדיין לא הועלו קבצים.</p> : <div className="authUsersList">
        {uploads.map(u => <div className="authUserRow" key={u.id}>
          <b>{u.originalName}</b>
          <span>{KIND_LABELS[u.kind] || u.kind}</span>
          <span>{fmtSize(u.size)}</span>
          <span>{u.uploadedBy}</span>
          <span>{new Date(u.uploadedAt).toLocaleString('he-IL')}</span>
        </div>)}
      </div>}
    </div>
  </div>;
}

function UsersAdmin({ onClose }) {
  const { apiFetch, user: me } = useAuth();
  const [users, setUsers] = useState([]);
  const [requests, setRequests] = useState([]);
  const [loading, setLoading] = useState(true);
  const [error, setError] = useState('');
  const [form, setForm] = useState({ username: '', password: '', role: 'uploader' });

  const load = () => Promise.all([
    apiFetch('/api/users').then(d => setUsers(d.users)),
    apiFetch('/api/users/reset-requests').then(d => setRequests(d.requests)),
  ]).catch(e => setError(e.message)).finally(() => setLoading(false));
  useEffect(() => { load(); }, []);

  const createUser = async e => {
    e.preventDefault();
    setError('');
    try {
      await apiFetch('/api/users', { method: 'POST', body: JSON.stringify(form) });
      setForm({ username: '', password: '', role: 'uploader' });
      load();
    } catch (err) { setError(err.message); }
  };

  const setRole = async (id, role) => {
    setError('');
    try { await apiFetch(`/api/users/${id}`, { method: 'PATCH', body: JSON.stringify({ role }) }); load(); }
    catch (err) { setError(err.message); }
  };

  const removeUser = async id => {
    if (!confirm('למחוק את המשתמש?')) return;
    setError('');
    try { await apiFetch(`/api/users/${id}`, { method: 'DELETE' }); load(); }
    catch (err) { setError(err.message); }
  };

  const resolveReset = async req => {
    const target = users.find(u => u.username === req.username);
    if (!target) { setError(`המשתמש ${req.username} כבר לא קיים.`); return; }
    const newPassword = prompt(`סיסמה חדשה עבור ${req.username} (6 תווים לפחות):`);
    if (!newPassword) return;
    setError('');
    try {
      await apiFetch(`/api/users/${target.id}`, { method: 'PATCH', body: JSON.stringify({ password: newPassword }) });
      await apiFetch(`/api/users/reset-requests/${req.id}`, { method: 'DELETE' });
      alert(`הסיסמה עבור ${req.username} עודכנה. מסור/י לו את הסיסמה החדשה.`);
      load();
    } catch (err) { setError(err.message); }
  };

  const dismissReset = async id => {
    setError('');
    try { await apiFetch(`/api/users/reset-requests/${id}`, { method: 'DELETE' }); load(); }
    catch (err) { setError(err.message); }
  };

  return <div className="authOverlay" onClick={onClose}>
    <div className="authModal authAdminModal" onClick={e => e.stopPropagation()}>
      <div className="authModalHead"><h2><UserCog size={20} /> ניהול משתמשים והרשאות</h2><button onClick={onClose}><X size={18} /></button></div>
      {error && <div className="authError">{error}</div>}

      {requests.length > 0 && <div className="authResetRequests">
        <h3><KeyRound size={16} /> בקשות איפוס סיסמה ({requests.length})</h3>
        {requests.map(r => <div className="authUserRow" key={r.id}>
          <b>{r.username}</b>
          <span>{new Date(r.requestedAt).toLocaleString('he-IL')}</span>
          <button className="calc" style={{ padding: '6px 12px' }} onClick={() => resolveReset(r)}>אפס סיסמה</button>
          <button className="dangerMini" onClick={() => dismissReset(r.id)} title="התעלם"><X size={14} /></button>
        </div>)}
      </div>}

      <form className="authNewUser" onSubmit={createUser}>
        <input placeholder="שם משתמש" value={form.username} onChange={e => setForm(p => ({ ...p, username: e.target.value }))} />
        <input placeholder="סיסמה (6 תווים לפחות)" type="password" value={form.password} onChange={e => setForm(p => ({ ...p, password: e.target.value }))} />
        <select value={form.role} onChange={e => setForm(p => ({ ...p, role: e.target.value }))}>
          <option value="viewer">צפייה בלבד</option>
          <option value="uploader">מורשה העלאה</option>
          <option value="admin">מנהל מערכת</option>
        </select>
        <button className="calc" type="submit"><Plus size={15} /> הוספת משתמש</button>
      </form>

      {loading ? <p>טוען...</p> : <div className="authUsersList">
        {users.map(u => <div className="authUserRow" key={u.id}>
          <b>{u.username}</b>
          <select value={u.role} onChange={e => setRole(u.id, e.target.value)} disabled={u.id === me.id && u.role === 'admin'}>
            <option value="viewer">צפייה בלבד</option>
            <option value="uploader">מורשה העלאה</option>
            <option value="admin">מנהל מערכת</option>
          </select>
          <button className="dangerMini" onClick={() => removeUser(u.id)} disabled={u.id === me.id} title={u.id === me.id ? 'לא ניתן למחוק את עצמך' : 'מחיקה'}><Trash2 size={14} /></button>
        </div>)}
      </div>}
    </div>
  </div>;
}
