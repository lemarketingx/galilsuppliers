import React, { useEffect, useRef, useState } from 'react';
import { ArrowLeft } from 'lucide-react';
import './splash.css';

/* Entry screen shown once per browser session.
   On "כניסה" the logo splits: the purple "d" slides left, the grey "G" slides right,
   and the two halves of the screen open like doors to reveal the system. */
const EXIT_MS = 1600;

export default function Splash({ onEnter }) {
  const [leaving, setLeaving] = useState(false);
  const done = useRef(false);

  const enter = () => {
    if (done.current) return;
    done.current = true;
    const reduce = window.matchMedia && window.matchMedia('(prefers-reduced-motion: reduce)').matches;
    if (reduce) { onEnter(); return; }
    setLeaving(true);
    setTimeout(onEnter, EXIT_MS);
  };

  useEffect(() => {
    const onKey = e => { if (e.key === 'Enter') enter(); };
    window.addEventListener('keydown', onKey);
    return () => window.removeEventListener('keydown', onKey);
  }, []);

  return <div className={'splash' + (leaving ? ' leaving' : '')} role="dialog" aria-modal="true" aria-label="כניסה למערכת הנדסה ורכש">
    <div className="splashDoor splashDoorLeft" />
    <div className="splashDoor splashDoorRight" />
    <div className="splashGlow" />
    <div className="splashContent">
      <div className="splashLogo" aria-label="קבוצת גליל">
        <img src="/brand/galil-d.png" alt="" className="splashPart splashD" draggable="false" />
        <img src="/brand/galil-g.png" alt="" className="splashPart splashG" draggable="false" />
        <img src="/brand/galil-text.png" alt="קבוצת גליל" className="splashPart splashText" draggable="false" />
      </div>
      <div className="splashCopy">
        <h1>מערכת הנדסה ורכש</h1>
        <p>מחירון כתבי כמויות ומאגר ספקים במערכת אחת</p>
      </div>
      <button className="splashEnter" onClick={enter} autoFocus>
        כניסה למערכת <ArrowLeft size={20} />
      </button>
    </div>
    <small className="splashFoot">© {new Date().getFullYear()} גליל הנדסה · מחלקת רכש</small>
  </div>;
}
