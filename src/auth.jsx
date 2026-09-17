import React, { createContext, useContext, useEffect, useState } from 'react';
import { LogIn, LogOut, Lock, UserCog, X, Trash2, Plus, ShieldCheck } from 'lucide-react';

const API_URL = import.meta.env.VITE_API_URL || '';
const TOKEN_KEY = 'galil_auth_token_v1';

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

  const canUpload = !!user && (user.role === 'uploader' || user.role === 'admin');
  const isAdmin = !!user && user.role === 'admin';

  return <AuthContext.Provider value={{ user, ready, login, logout, canUpload, isAdmin, roleLabel: user ? ROLE_LABELS[user.role] || user.role : '', apiFetch }}>
    {children}
  </AuthContext.Provider>;
}

/* Guard for any upload trigger (button click / drag&drop). Returns true if the
   action may proceed; otherwise opens the login modal (if signed out) or an
   explanatory alert (if signed in without upload permission), and returns false. */
export function useUploadGuard() {
  const { user, canUpload } = useAuth();
  const [showLogin, setShowLogin] = useState(false);
  const guard = () => {
    if (canUpload) return true;
    if (!user) { setShowLogin(true); return false; }
    alert('אין לך הרשאת העלאת קבצים. פנה/י למנהל המערכת כדי לקבל הרשאה.');
    return false;
  };
  return { guard, showLogin, setShowLogin };
}

export function LoginModal({ onClose }) {
  const { login } = useAuth();
  const [username, setUsername] = useState('');
  const [password, setPassword] = useState('');
  const [error, setError] = useState('');
  const [busy, setBusy] = useState(false);

  const submit = async e => {
    e.preventDefault();
    if (!username || !password) return;
    setBusy(true); setError('');
    try { await login(username, password); onClose(); }
    catch (err) { setError(err.message || 'שגיאה בהתחברות.'); }
    finally { setBusy(false); }
  };

  return <div className="authOverlay" onClick={onClose}>
    <form className="authModal" onClick={e => e.stopPropagation()} onSubmit={submit}>
      <div className="authModalHead"><h2><Lock size={20} /> התחברות</h2><button type="button" onClick={onClose}><X size={18} /></button></div>
      <p className="authHint">רק משתמשים עם הרשאת "מורשה העלאה" או "מנהל" יכולים להעלות קבצים לאתר.</p>
      <label>שם משתמש<input value={username} onChange={e => setUsername(e.target.value)} autoFocus /></label>
      <label>סיסמה<input type="password" value={password} onChange={e => setPassword(e.target.value)} /></label>
      {error && <div className="authError">{error}</div>}
      <button className="calc" type="submit" disabled={busy}>{busy ? 'מתחבר...' : 'התחברות'}</button>
    </form>
  </div>;
}

export function AuthBar() {
  const { user, logout, roleLabel, isAdmin } = useAuth();
  const [showLogin, setShowLogin] = useState(false);
  const [showAdmin, setShowAdmin] = useState(false);

  return <div className="authBar">
    {user ? <>
      <span className="authUser"><ShieldCheck size={15} /> {user.username} <em>({roleLabel})</em></span>
      {isAdmin && <button onClick={() => setShowAdmin(true)}><UserCog size={15} /> ניהול משתמשים</button>}
      <button onClick={logout}><LogOut size={15} /> התנתקות</button>
    </> : <button onClick={() => setShowLogin(true)}><LogIn size={15} /> התחברות</button>}
    {showLogin && <LoginModal onClose={() => setShowLogin(false)} />}
    {showAdmin && <UsersAdmin onClose={() => setShowAdmin(false)} />}
  </div>;
}

function UsersAdmin({ onClose }) {
  const { apiFetch, user: me } = useAuth();
  const [users, setUsers] = useState([]);
  const [loading, setLoading] = useState(true);
  const [error, setError] = useState('');
  const [form, setForm] = useState({ username: '', password: '', role: 'uploader' });

  const load = () => apiFetch('/api/users').then(d => setUsers(d.users)).catch(e => setError(e.message)).finally(() => setLoading(false));
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

  return <div className="authOverlay" onClick={onClose}>
    <div className="authModal authAdminModal" onClick={e => e.stopPropagation()}>
      <div className="authModalHead"><h2><UserCog size={20} /> ניהול משתמשים והרשאות</h2><button onClick={onClose}><X size={18} /></button></div>
      {error && <div className="authError">{error}</div>}

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
