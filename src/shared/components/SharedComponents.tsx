import * as React from 'react';
import { STATUS_CFG } from '../models/IProject';

const { useState, useCallback } = React;

// ── Helpers ──────────────────────────────────────────────────────────────────
export const fmtD = (d: string): string =>
  d ? new Date(d + 'T00:00:00').toLocaleDateString('en-AU', { day: '2-digit', month: 'short', year: '2-digit' }) : '—';

export const rfiTot = (r: any): number =>
  (parseFloat(r.model) || 0) + (parseFloat(r.connections) || 0) + (parseFloat(r.checking) || 0) + (parseFloat(r.drawings) || 0) + (parseFloat(r.admin) || 0);

export const isOD = (r: any): boolean =>
  (r.status === 'Open' || r.status === 'Overdue') && r.dateRequired && new Date(r.dateRequired) < new Date();

export const effSt = (r: any): string => isOD(r) ? 'Overdue' : r.status;

export const hrsColor = (pct: number | null, over: boolean): string => {
  if (over || (pct !== null && pct >= 100)) return 'var(--rd)';
  if (pct !== null && pct >= 80) return 'var(--am)';
  return 'var(--gn)';
};

export const hrsRem = (p: any): number | null => p.hrsAllowed > 0 ? p.hrsAllowed - p.hrsUsed : null;
export const hpct = (p: any): number | null => p.hrsAllowed > 0 ? Math.min(100, Math.round((p.hrsUsed / p.hrsAllowed) * 100)) : null;

// ── Tag ──────────────────────────────────────────────────────────────────────
export const Tag: React.FC<{ s: string; small?: boolean }> = ({ s, small }) => {
  const c = STATUS_CFG[s] || { bg: 'rgba(90,110,136,.10)', color: '#4a5e78', bd: '#5a6e88' };
  return (
    <span style={{
      display: 'inline-block', padding: small ? '2px 7px' : '3px 11px', borderRadius: 20,
      fontFamily: 'Montserrat', fontWeight: 700, fontSize: small ? 9.5 : 11.5,
      letterSpacing: '.06em', textTransform: 'uppercase',
      border: `1px solid ${c.bd}`, background: c.bg, color: c.color, whiteSpace: 'nowrap'
    }}>{s}</span>
  );
};

// ── HrsBar ───────────────────────────────────────────────────────────────────
export const HrsBar: React.FC<{ allowed: number; used: number }> = ({ allowed, used }) => {
  const pct = allowed > 0 ? Math.min(100, Math.round((used / allowed) * 100)) : null;
  const over = used > allowed;
  const col = hrsColor(pct, over);
  if (pct === null) return <span style={{ fontFamily: 'Montserrat', fontWeight: 600, fontSize: 12, color: 'var(--t4)' }}>—</span>;
  const rem = Math.abs(used - allowed).toFixed(1);
  return (
    <div style={{ minWidth: 140 }}>
      <div style={{ display: 'flex', justifyContent: 'space-between', fontFamily: 'Montserrat', fontWeight: 600, fontSize: 11.5, marginBottom: 4 }}>
        <span style={{ color: 'var(--t2)' }}>{used}<span style={{ color: 'var(--t4)', fontWeight: 600 }}> / {allowed}h</span></span>
        <span style={{ color: col, fontWeight: over ? 700 : 600 }}>{over ? '+' : ''}{rem}h {over ? 'OVER' : 'left'}</span>
      </div>
      <div style={{ height: 6, background: pct === 0 ? '#ccc' : 'var(--s2)', borderRadius: 3, overflow: 'hidden' }}>
        <div style={{ height: '100%', width: `${pct}%`, background: col, borderRadius: 3, transition: 'width .4s' }} />
      </div>
    </div>
  );
};

// ── RfiBar ───────────────────────────────────────────────────────────────────
export const RfiBar: React.FC<{ allowed: number; used: number }> = ({ allowed, used }) => {
  if (!allowed || allowed === 0) {
    return (
      <div style={{ minWidth: 120 }}>
        <div style={{ fontFamily: 'Montserrat', fontWeight: 600, fontSize: 11.5, color: 'var(--t4)', marginBottom: 4 }}>
          {used > 0 ? <span style={{ color: 'var(--t2)', fontWeight: 600 }}>{used} used</span> : '—'}
          <span style={{ color: 'var(--t4)', fontWeight: 600 }}> / no limit</span>
        </div>
        <div style={{ height: 6, background: '#ccc', borderRadius: 3 }} />
      </div>
    );
  }
  const over = used > allowed;
  const pct = Math.min(100, Math.round((used / allowed) * 100));
  const rem = allowed - used;
  const col = over ? 'var(--rd)' : pct >= 80 ? 'var(--am)' : 'var(--3eg)';
  return (
    <div style={{ minWidth: 120 }}>
      <div style={{ display: 'flex', justifyContent: 'space-between', fontFamily: 'Montserrat', fontWeight: 600, fontSize: 11.5, marginBottom: 4 }}>
        <span style={{ color: 'var(--t2)' }}>{used}<span style={{ color: 'var(--t4)', fontWeight: 600 }}> / {allowed}</span></span>
        <span style={{ color: col, fontWeight: over ? 700 : 600 }}>{over ? `+${Math.abs(rem)} OVER` : `${rem} left`}</span>
      </div>
      <div style={{ height: 6, background: pct === 0 ? '#ccc' : 'var(--s2)', borderRadius: 3, overflow: 'hidden' }}>
        <div style={{ height: '100%', width: `${pct}%`, background: col, borderRadius: 3, transition: 'width .4s' }} />
      </div>
    </div>
  );
};

// ── Section Divider ──────────────────────────────────────────────────────────
export const SDiv: React.FC<{ label: string }> = ({ label }) => (
  <div style={{ display: 'flex', alignItems: 'center', gap: 10, margin: '22px 0 14px' }}>
    <div style={{ width: 4, height: 16, background: 'var(--3eg)', borderRadius: 2, flexShrink: 0 }} />
    <span style={{ fontFamily: 'Montserrat', fontWeight: 800, fontSize: 12.5, letterSpacing: '.12em', textTransform: 'uppercase', color: 'var(--3eg)' }}>{label}</span>
    <div style={{ flex: 1, height: 1, background: 'var(--bd)' }} />
  </div>
);

// ── Stat Card ────────────────────────────────────────────────────────────────
export const Stat: React.FC<{ label: string; value: any; sub?: string; col: string; warn?: boolean }> = ({ label, value, sub, col, warn }) => (
  <div style={{ flex: 1, minWidth: 130, background: 'var(--s0)', border: '1px solid var(--bd)', borderRadius: 10, padding: '18px 20px', borderTop: `4px solid ${col}`, boxShadow: '0 1px 4px rgba(0,0,0,.06)' }}>
    <div style={{ fontFamily: 'Montserrat', fontWeight: 600, fontSize: 11.5, letterSpacing: '.12em', textTransform: 'uppercase', color: 'var(--t4)', marginBottom: 10 }}>{label}</div>
    <div style={{ fontFamily: 'Montserrat', fontWeight: 800, fontSize: 34, color: warn ? 'var(--rd)' : 'var(--t1)', lineHeight: 1 }}>{value}</div>
    {sub && <div style={{ fontFamily: 'Montserrat', fontWeight: 600, fontSize: 12.5, color: col, marginTop: 7 }}>{sub}</div>}
  </div>
);

// ── Slide-over Panel ─────────────────────────────────────────────────────────
export const Panel: React.FC<{ open: boolean; onClose: () => void; title?: string; subtitle?: string; tag?: React.ReactNode; children?: React.ReactNode }> = ({ open, onClose, title, subtitle, tag, children }) => (
  <>
    <div onClick={onClose} style={{ position: 'fixed', inset: 0, background: 'rgba(255,255,255,0.85)', zIndex: 290, display: open ? 'block' : 'none', backdropFilter: 'blur(2px)' }} />
    <div style={{
      position: 'fixed', top: 0, right: 0, bottom: 0, width: 'min(700px,100vw)',
      background: 'var(--s0)', borderLeft: '1px solid var(--bd)', zIndex: 300,
      display: 'flex', flexDirection: 'column',
      transform: open ? 'translateX(0)' : 'translateX(100%)',
      transition: 'transform .22s cubic-bezier(.4,0,.2,1)',
      boxShadow: open ? '-12px 0 40px rgba(0,0,0,.12)' : 'none'
    }}>
      <div style={{ padding: '18px 24px 16px', borderBottom: '1px solid var(--bd)', display: 'flex', alignItems: 'center', gap: 12, flexShrink: 0, background: 'var(--s1)' }}>
        <div style={{ flex: 1 }}>
          <div style={{ fontFamily: 'Montserrat', fontWeight: 800, fontSize: 18, color: 'var(--t1)', letterSpacing: '.02em', lineHeight: 1.1 }}>{title}</div>
          {subtitle && <div style={{ fontFamily: 'Montserrat', fontWeight: 600, fontSize: 12.5, color: 'var(--t4)', marginTop: 4, letterSpacing: '.04em' }}>{subtitle}</div>}
        </div>
        {tag}
        <button onClick={onClose} style={{ background: 'transparent', border: '1px solid var(--bd)', color: 'var(--t3)', width: 32, height: 32, borderRadius: 6, fontSize: 15, display: 'flex', alignItems: 'center', justifyContent: 'center', cursor: 'pointer' }}>✕</button>
      </div>
      <div style={{ flex: 1, overflowY: 'auto', padding: '20px 24px 40px' }}>{children}</div>
    </div>
  </>
);

// ── Form Field ───────────────────────────────────────────────────────────────
export const FF: React.FC<{ label: string; span2?: boolean; children: React.ReactNode }> = ({ label, span2, children }) => (
  <div style={{ gridColumn: span2 ? 'span 2' : 'span 1' }}>
    <label style={{ fontFamily: 'Montserrat', fontWeight: 600, fontSize: 12.5, letterSpacing: '.06em', textTransform: 'uppercase', color: 'var(--t3)', display: 'block', marginBottom: 6 }}>{label}</label>
    {children}
  </div>
);

// ── Icon Button ──────────────────────────────────────────────────────────────
export const IBtn: React.FC<{ onClick: () => void; title?: string; danger?: boolean; children: React.ReactNode }> = ({ onClick, title, danger, children }) => {
  const [h, setH] = useState(false);
  return (
    <button onClick={onClick} title={title}
      onMouseEnter={() => setH(true)} onMouseLeave={() => setH(false)}
      style={{
        background: h ? (danger ? 'var(--rd2)' : 'var(--3eg3)') : 'var(--s0)',
        border: `1px solid ${h ? (danger ? 'var(--rd)' : 'var(--3eg)') : 'var(--bd)'}`,
        color: h ? (danger ? 'var(--rd)' : 'var(--3eg)') : 'var(--t2)',
        padding: '2px 0', borderRadius: 4, fontSize: 10.5, fontFamily: 'Montserrat', fontWeight: 600, cursor: 'pointer', transition: 'all .12s', width: 50, textAlign: 'center'
      }}>{children}</button>
  );
};

// ── Delete Confirmation Modal ────────────────────────────────────────────────
export const DelModal: React.FC<{ open: boolean; label?: string; onConfirm: () => void; onCancel: () => void }> = ({ open, label, onConfirm, onCancel }) => (
  <div style={{ position: 'fixed', inset: 0, background: 'rgba(255,255,255,0.97)', zIndex: 400, display: open ? 'flex' : 'none', alignItems: 'center', justifyContent: 'center', backdropFilter: 'blur(3px)' }}>
    <div style={{ background: 'var(--s2)', border: '1px solid var(--rd)', borderRadius: 4, padding: '28px 32px', maxWidth: 360, width: '90%' }}>
      <div style={{ fontFamily: 'Montserrat', fontWeight: 800, fontSize: 18, color: 'var(--rd)', letterSpacing: '.06em', marginBottom: 10 }}>CONFIRM DELETE</div>
      <div style={{ fontFamily: 'Montserrat', fontSize: 14, color: 'var(--t2)', lineHeight: 1.6, marginBottom: 22 }}>{label || 'This record will be permanently deleted.'}</div>
      <div style={{ display: 'flex', gap: 10 }}>
        <button onClick={onConfirm} style={{ fontFamily: 'Montserrat', fontWeight: 700, fontSize: 14, letterSpacing: '.06em', textTransform: 'uppercase', padding: '8px 20px', background: 'var(--rd)', color: '#1a2030', border: 'none', borderRadius: 2 }}>DELETE</button>
        <button onClick={onCancel} style={{ fontFamily: 'Montserrat', fontSize: 12.5, padding: '8px 18px', background: 'transparent', border: '1px solid var(--bd)', color: 'var(--t2)', borderRadius: 2 }}>CANCEL</button>
      </div>
    </div>
  </div>
);

// ── Toast Hook ───────────────────────────────────────────────────────────────
export function useToast(): { show: (msg: string, type?: string) => void; Toast: React.ReactNode } {
  const [t, setT] = useState<{ msg: string; type: string } | null>(null);
  const show = useCallback((msg: string, type: string = 'success') => {
    setT({ msg, type });
    if (type !== 'error') { setTimeout(() => setT(null), 3500); }
  }, []);
  const Toast = t ? (
    <div onClick={() => setT(null)} style={{
      position: 'fixed', bottom: 24, left: '50%', transform: 'translateX(-50%)',
      background: 'var(--s0)', border: `2px solid ${t.type === 'success' ? 'var(--gn)' : t.type === 'error' ? 'var(--rd)' : 'var(--am)'}`,
      padding: '13px 26px', borderRadius: 8, fontFamily: 'Montserrat', fontWeight: 600, fontSize: 12, zIndex: 9999,
      boxShadow: '0 8px 30px rgba(0,0,0,.5)',
      color: t.type === 'success' ? 'var(--gn)' : t.type === 'error' ? 'var(--rd)' : 'var(--am)',
      whiteSpace: 'pre-wrap', maxWidth: 540, cursor: 'pointer', lineHeight: 1.6
    }}>{t.type === 'error' ? '⚠ ' + t.msg + '\n\nClick to dismiss' : t.msg}</div>
  ) : null;
  return { show, Toast };
}

// ── Primary Button ───────────────────────────────────────────────────────────
export const BtnPrimary: React.FC<{ onClick: () => void; children: React.ReactNode }> = ({ onClick, children }) => (
  <button onClick={onClick} style={{
    fontFamily: 'Montserrat', fontWeight: 700, fontSize: 14, letterSpacing: '.07em', textTransform: 'uppercase',
    padding: '9px 22px', background: 'var(--3eg)', color: '#1a2030', border: 'none', borderRadius: 7,
    boxShadow: '0 2px 8px rgba(45,168,74,.3)', cursor: 'pointer'
  }}>{children}</button>
);

// ── CC Multi-Email Field ─────────────────────────────────────────────────────
export const CcField: React.FC<{ value: string; onChange: (v: string) => void; compact?: boolean }> = ({ value, onChange, compact }) => {
  const emails = value ? value.split(',').map(e => e.trim()).filter(Boolean) : [];
  const [draft, setDraft] = useState('');
  const [focused, setFocused] = useState(false);
  const MAX = 10;

  function addEmail(raw: string): void {
    const e = raw.trim().toLowerCase();
    if (!e || emails.length >= MAX || emails.includes(e)) { setDraft(''); return; }
    onChange([...emails, e].join(','));
    setDraft('');
  }
  function removeEmail(idx: number): void { onChange(emails.filter((_, i) => i !== idx).join(',')); }
  function handleKey(e: React.KeyboardEvent<HTMLInputElement>): void {
    if (e.key === 'Enter' || e.key === ',' || e.key === ';' || e.key === ' ') { e.preventDefault(); addEmail(draft); }
    else if (e.key === 'Backspace' && draft === '' && emails.length > 0) { removeEmail(emails.length - 1); }
  }
  function handleBlur(): void { setFocused(false); if (draft.trim()) addEmail(draft); }
  const atMax = emails.length >= MAX;

  return (
    <div>
      {!compact && (<label style={{ fontFamily: 'Montserrat', fontSize: 11.5, letterSpacing: '.1em', textTransform: 'uppercase', color: 'var(--t4)', display: 'block', marginBottom: 5 }}>
        CC / Secondary Recipients<span style={{ marginLeft: 8, color: atMax ? 'var(--am)' : 'var(--t4)', fontSize: 10.5 }}>{emails.length}/{MAX} addresses</span>
      </label>)}
      <div style={{
        background: 'var(--s2)', border: `1px solid ${focused ? 'var(--3eg)' : 'var(--bd)'}`,
        borderRadius: 2, padding: '6px 8px', minHeight: 38,
        display: 'flex', flexWrap: 'wrap', gap: 5, alignItems: 'center', cursor: 'text', transition: 'border-color .15s'
      }}>
        {emails.map((em, i) => (
          <span key={i} style={{
            display: 'inline-flex', alignItems: 'center', gap: 4,
            background: 'var(--3eg3)', border: '1px solid var(--3eg)',
            borderRadius: 2, padding: '2px 6px 2px 8px',
            fontFamily: 'Montserrat', fontSize: 11.5, color: 'var(--3eg)',
            maxWidth: 220, overflow: 'hidden', textOverflow: 'ellipsis', whiteSpace: 'nowrap'
          }}>
            {em}
            <button type="button" onClick={e => { e.stopPropagation(); removeEmail(i); }} style={{
              background: 'none', border: 'none', color: 'var(--3eg)', cursor: 'pointer',
              fontSize: 12, lineHeight: 1, padding: '0 1px', opacity: 0.7, fontFamily: 'Montserrat', flexShrink: 0
            }}>×</button>
          </span>
        ))}
        {!atMax && <input value={draft} onChange={e => setDraft(e.target.value)} onKeyDown={handleKey} onFocus={() => setFocused(true)} onBlur={handleBlur}
          placeholder={emails.length === 0 ? 'Type email, press Enter or comma to add…' : 'Add another…'}
          style={{ background: 'transparent', border: 'none', outline: 'none', fontFamily: 'Montserrat', fontSize: 12.5, color: 'var(--t1)', flex: 1, minWidth: 180, padding: '2px 4px' }} />}
      </div>
    </div>
  );
};
