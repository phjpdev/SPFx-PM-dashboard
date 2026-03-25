import * as React from 'react';
import { SharePointService } from '../../../shared/services/SharePointService';
import { ITask, ITeamMember, ITaskHistory, DAYS, PROD_TASK_CODES, NON_PROD_TASK_CODES, isNonProd, CHECK_CODES } from '../../../shared/models/ITask';
import { IProject } from '../../../shared/models/IProject';

// ── Helpers ──────────────────────────────────────────────────
const now = (): string => new Date().toISOString();
const fmtTs = (d: string): string => {
  if (!d) return '';
  const x = new Date(d);
  const day = x.getDate();
  const m = ['Jan','Feb','Mar','Apr','May','Jun','Jul','Aug','Sep','Oct','Nov','Dec'][x.getMonth()];
  return `${day} ${m} ${String(x.getHours()).padStart(2,'0')}:${String(x.getMinutes()).padStart(2,'0')}`;
};
const getWeekStart = (date: Date): Date => {
  const d = new Date(date); const day = d.getDay();
  const diff = d.getDate() - day + (day === 0 ? -6 : 1);
  d.setDate(diff); d.setHours(0,0,0,0); return d;
};
const getWeekNumber = (date: Date): number => {
  const d = new Date(Date.UTC(date.getFullYear(), date.getMonth(), date.getDate()));
  const dn = d.getUTCDay() || 7;
  d.setUTCDate(d.getUTCDate() + 4 - dn);
  const ys = new Date(Date.UTC(d.getUTCFullYear(), 0, 1));
  return Math.ceil(((((d as any) - (ys as any)) / 86400000) + 1) / 7); // eslint-disable-line @typescript-eslint/no-explicit-any
};
const getWeeksInYear = (y: number): number => {
  const d = new Date(y, 11, 31);
  return getWeekNumber(d) === 1 ? 52 : getWeekNumber(d);
};
const formatWeekLabel = (mon: Date): string => `Wk ${getWeekNumber(mon)}/${getWeeksInYear(mon.getFullYear())}`;
const formatDateRange = (mon: Date): string => {
  const fri = new Date(mon); fri.setDate(fri.getDate() + 4);
  const m = ['Jan','Feb','Mar','Apr','May','Jun','Jul','Aug','Sep','Oct','Nov','Dec'];
  return `${mon.getDate()} ${m[mon.getMonth()]} — ${fri.getDate()} ${m[fri.getMonth()]}`;
};
const getDayDate = (mon: Date, di: number): string => {
  const d = new Date(mon); d.setDate(d.getDate() + di);
  const m = ['Jan','Feb','Mar','Apr','May','Jun','Jul','Aug','Sep','Oct','Nov','Dec'];
  return `${d.getDate()} ${m[d.getMonth()]}`;
};
const weekKey = (mon: Date): string => mon.toISOString().substring(0, 10);
const round1 = (n: number): number => Math.round(n * 10) / 10;

// ── Styles ───────────────────────────────────────────────────
const C = {
  bg: '#f0f2f5', card: '#ffffff', border: '#e2e6ec', muted: '#3a4a60',
  text: '#111820', dim: '#4a5e78', green: '#2a9e2a', greenLt: '#157a15',
  amber: '#d4880a', amberLt: '#9a5e06', red: '#cc3333', redDk: '#a32020',
  purple: '#534AB7', purpleLt: '#3d2e9e', purpleBg: 'rgba(83,74,183,0.08)',
};
const statusC: Record<string, { bg: string; text: string; bd: string; label: string }> = {
  complete:    { bg: 'rgba(42,158,42,0.12)', text: '#1e8a1e', bd: '#3db63d', label: 'DONE' },
  wip:         { bg: 'rgba(212,136,10,0.12)', text: '#a06808', bd: '#d4880a', label: 'WIP' },
  not_started: { bg: 'rgba(90,110,136,0.10)', text: '#5a6e88', bd: '#8a9bb0', label: 'TODO' },
  blocked:     { bg: 'rgba(204,51,51,0.12)', text: '#b82020', bd: '#cc3333', label: 'BLOCKED' },
  rework:      { bg: 'rgba(204,51,51,0.12)', text: '#b82020', bd: '#cc3333', label: 'REWORK' },
};
const prioC: Record<string, string> = { high: '#cc3333', medium: '#d4880a', low: '#8a9bb0' };

const Badge: React.FC<{ status: string; wipPct: number; reviewStatus: string }> = ({ status, wipPct, reviewStatus }) => {
  const c = statusC[status] || statusC.not_started;
  return (
    <span style={{ display: 'inline-flex', alignItems: 'center', gap: 4 }}>
      <span style={{ display: 'inline-block', padding: '2px 8px', borderRadius: 10, fontSize: 10, fontWeight: 600, letterSpacing: '0.04em', background: c.bg, color: c.text, border: `1px solid ${c.bd}` }}>
        {status === 'wip' ? `WIP ${wipPct}%` : c.label}
      </span>
      {status === 'complete' && reviewStatus === 'accepted' && <span style={{ fontSize: 10, color: '#1e8a1e' }}>&#10003;</span>}
      {status === 'complete' && !reviewStatus && <span style={{ fontSize: 10, color: '#d4880a' }}>&#9679;</span>}
    </span>
  );
};

// ── Props ────────────────────────────────────────────────────
interface TaskBoardProps {
  spService: SharePointService;
  projects: IProject[];
  userDisplayName: string;
  siteUrl: string;
  isManager: boolean;
  toast: (msg: string) => void;
}

// ── Modal Components ─────────────────────────────────────────
const Modal: React.FC<{ onClose: () => void; children: React.ReactNode }> = ({ onClose, children }) => (
  <div style={{ position: 'fixed', inset: 0, background: 'rgba(0,0,0,0.6)', display: 'flex', alignItems: 'center', justifyContent: 'center', zIndex: 1000 }} onClick={onClose}>
    <div onClick={e => e.stopPropagation()} style={{ background: C.card, borderRadius: 10, border: `1px solid ${C.border}`, padding: '24px 28px', maxWidth: 480, width: '92%' }}>{children}</div>
  </div>
);

const CompleteModal: React.FC<{ task: ITask; onConfirm: (hrs: number, note: string) => void; onCancel: () => void }> = ({ task, onConfirm, onCancel }) => {
  const [hours, setHours] = React.useState(task.hoursActual || task.hoursPlanned);
  const [note, setNote] = React.useState('');
  const v = task.hoursPlanned > 0 ? ((hours - task.hoursPlanned) / task.hoursPlanned) * 100 : 0;
  const inp: React.CSSProperties = { padding: '8px 12px', background: '#f5f6f8', border: `1px solid ${v > 20 ? C.redDk : C.border}`, borderRadius: 4, color: C.text, fontSize: 12, width: '100%', boxSizing: 'border-box' };
  return (
    <Modal onClose={onCancel}>
      <div style={{ fontSize: 15, fontWeight: 600, color: C.text, marginBottom: 4 }}>Mark task as complete</div>
      <div style={{ fontSize: 12, color: C.muted, marginBottom: 16 }}>{task.project} — {task.description}</div>
      <div style={{ marginBottom: 12 }}>
        <label style={{ fontSize: 11, color: C.muted, display: 'block', marginBottom: 4 }}>ACTUAL HOURS</label>
        <div style={{ display: 'flex', alignItems: 'center', gap: 10 }}>
          <input type="number" step="0.5" min="0" value={hours} onChange={e => setHours(parseFloat(e.target.value) || 0)} style={{ ...inp, width: 80 }} />
          <span style={{ fontSize: 12, color: C.dim }}>planned: {task.hoursPlanned}h</span>
          {v > 20 && <span style={{ fontSize: 11, color: C.red, background: 'rgba(204,51,51,0.08)', padding: '2px 8px', borderRadius: 4 }}>+{Math.round(v)}% over</span>}
        </div>
      </div>
      <div style={{ marginBottom: 12 }}>
        <label style={{ fontSize: 11, color: C.muted, display: 'block', marginBottom: 4 }}>COMPLETION NOTE</label>
        <input value={note} onChange={e => setNote(e.target.value)} placeholder="What was delivered..." style={inp} />
      </div>
      {CHECK_CODES.includes(task.taskCode) && <div style={{ padding: '8px 12px', background: C.purpleBg, borderRadius: 6, border: `1px solid ${C.purple}`, marginBottom: 12, fontSize: 12, color: C.purpleLt }}>Check task ({task.taskCode}) — cross-references QA checklist for {task.project}.</div>}
      <div style={{ display: 'flex', gap: 10, justifyContent: 'flex-end', marginTop: 16 }}>
        <button onClick={onCancel} style={{ padding: '8px 18px', borderRadius: 5, border: `1px solid ${C.border}`, background: 'transparent', color: '#3a4a60', fontSize: 13, cursor: 'pointer' }}>Cancel</button>
        <button onClick={() => onConfirm(hours, note)} style={{ padding: '8px 18px', borderRadius: 5, border: 'none', background: C.green, color: '#fff', fontSize: 13, fontWeight: 600, cursor: 'pointer' }}>Confirm complete</button>
      </div>
    </Modal>
  );
};

const UndoModal: React.FC<{ task: ITask; onConfirm: (reason: string) => void; onCancel: () => void }> = ({ task, onConfirm, onCancel }) => {
  const [reason, setReason] = React.useState('');
  return (
    <Modal onClose={onCancel}>
      <div style={{ fontSize: 15, fontWeight: 600, color: C.text, marginBottom: 4 }}>Re-open completed task</div>
      <div style={{ padding: '8px 12px', background: 'rgba(212,136,10,0.08)', borderRadius: 6, border: `1px solid ${C.amber}`, marginBottom: 12, fontSize: 12, color: C.amberLt }}>Completed by {task.completedBy} on {fmtTs(task.completedAt)}. Re-opening will be logged.</div>
      <label style={{ fontSize: 11, color: C.muted, display: 'block', marginBottom: 4 }}>REASON (required)</label>
      <input value={reason} onChange={e => setReason(e.target.value)} placeholder="Why re-open?"
        style={{ width: '100%', padding: '8px 12px', background: '#f5f6f8', border: `1px solid ${C.border}`, borderRadius: 4, color: C.text, fontSize: 12, boxSizing: 'border-box', marginBottom: 16 }} />
      <div style={{ display: 'flex', gap: 10, justifyContent: 'flex-end' }}>
        <button onClick={onCancel} style={{ padding: '8px 18px', borderRadius: 5, border: `1px solid ${C.border}`, background: 'transparent', color: '#3a4a60', fontSize: 13, cursor: 'pointer' }}>Cancel</button>
        <button onClick={() => reason.trim() && onConfirm(reason)} disabled={!reason.trim()}
          style={{ padding: '8px 18px', borderRadius: 5, border: 'none', background: reason.trim() ? C.redDk : '#c0c8d4', color: '#fff', fontSize: 13, fontWeight: 600, cursor: reason.trim() ? 'pointer' : 'default', opacity: reason.trim() ? 1 : 0.5 }}>Re-open</button>
      </div>
    </Modal>
  );
};

// ── Task Row ─────────────────────────────────────────────────
interface TaskRowProps {
  task: ITask;
  onUpdate: (t: ITask) => void;
  onEdit: (t: ITask) => void;
  onDelete: (t: ITask) => void;
  isLocked: boolean;
  isManager: boolean;
  currentUserInitials: string;
  isInternal: boolean;
}

const TaskRow: React.FC<TaskRowProps> = ({ task, onUpdate, onEdit, onDelete, isLocked, isManager, currentUserInitials, isInternal }) => {
  const [showComplete, setShowComplete] = React.useState(false);
  const [showUndo, setShowUndo] = React.useState(false);
  const [showHistory, setShowHistory] = React.useState(false);
  const canEdit = !isLocked || isManager;
  const isOwner = task.assignee === currentUserInitials;
  const canComplete = canEdit && task.status !== 'complete';
  const canUndo = canEdit && (isOwner || isManager) && task.status === 'complete' && task.reviewStatus !== 'accepted';
  const overHrs = task.hoursActual > task.hoursPlanned * 1.2;

  const handleComplete = (hours: number, note: string): void => {
    onUpdate({ ...task, status: 'complete', wipPct: 100, hoursActual: hours, completedBy: currentUserInitials, completedAt: now(), completionNote: note,
      history: [...task.history, { action: 'completed', user: currentUserInitials, ts: now(), detail: `${hours}h actual | ${task.hoursPlanned}h planned${note ? ` | ${note}` : ''}` }] });
    setShowComplete(false);
  };
  const handleUndo = (reason: string): void => {
    onUpdate({ ...task, status: 'wip', wipPct: 90, completedBy: '', completedAt: '', completionNote: '', reviewedBy: '', reviewStatus: '',
      history: [...task.history, { action: 're-opened', user: currentUserInitials, ts: now(), detail: `Reason: ${reason}` }] });
    setShowUndo(false);
  };
  const wipTimeout = React.useRef<ReturnType<typeof setTimeout> | null>(null);
  const [localWip, setLocalWip] = React.useState(task.wipPct);
  React.useEffect(() => { setLocalWip(task.wipPct); }, [task.wipPct]);
  const handleWip = (val: number): void => {
    if (!canEdit || task.status === 'complete') return;
    if (val >= 100) { setShowComplete(true); return; }
    setLocalWip(val);
    if (wipTimeout.current) clearTimeout(wipTimeout.current);
    wipTimeout.current = setTimeout(() => {
      onUpdate({ ...task, wipPct: val, status: val > 0 ? 'wip' : 'not_started', hoursActual: round1(task.hoursPlanned * val / 100),
        history: [...task.history, { action: 'progress', user: currentUserInitials, ts: now(), detail: `WIP ${val}%` }] });
    }, 500);
  };

  const hrsText = task.status === 'complete' ? `${task.hoursActual}h` : task.hoursActual > 0 ? `${task.hoursActual}/${task.hoursPlanned}h` : `${task.hoursPlanned}h`;

  return (
    <>
      {showComplete && <CompleteModal task={task} onConfirm={handleComplete} onCancel={() => setShowComplete(false)} />}
      {showUndo && <UndoModal task={task} onConfirm={handleUndo} onCancel={() => setShowUndo(false)} />}
      <div style={{ display: 'grid', gridTemplateColumns: '4px 70px 52px 1fr 36px 50px 120px 110px', gap: 6, padding: '5px 8px', alignItems: 'center', fontSize: 12, borderBottom: `1px solid ${C.card}`, opacity: task.status === 'complete' ? 0.6 : 1, background: isInternal ? 'rgba(90,110,136,0.04)' : 'transparent' }}>
        <div style={{ width: 4, height: 22, borderRadius: 2, background: isInternal ? C.dim : prioC[task.priority] || C.amberLt }} />
        <span style={{ color: isInternal ? C.muted : C.greenLt, fontWeight: 600, fontSize: 11 }}>{task.project}</span>
        <span style={{ color: C.muted, fontSize: 11 }}>{task.taskCode}</span>
        <div>
          <span style={{ color: task.status === 'complete' ? C.muted : C.text, textDecoration: task.status === 'complete' ? 'line-through' : 'none' }}>{task.description}</span>
          {task.completionNote && task.status === 'complete' && <div style={{ fontSize: 10, color: C.greenLt, marginTop: 1 }}>{task.completionNote}</div>}
          {showHistory && (
            <div style={{ background: C.bg, borderRadius: 4, border: `1px solid ${C.border}`, padding: '6px 10px', marginTop: 4, maxHeight: 150, overflowY: 'auto' }}>
              {[...task.history].reverse().map((h, i) => <div key={i} style={{ fontSize: 10, color: '#3a4a60', padding: '2px 0' }}><span style={{ color: C.muted }}>{fmtTs(h.ts)}</span> — {h.action} by {h.user}{h.detail && <span style={{ color: C.dim }}> | {h.detail}</span>}</div>)}
              {task.history.length === 0 && <div style={{ fontSize: 10, color: C.dim }}>No history</div>}
            </div>
          )}
        </div>
        <span style={{ fontSize: 11, color: '#3a4a60', textAlign: 'center' }}>{task.assignee}</span>
        <span style={{ fontSize: 11, color: overHrs ? C.red : '#B4B2A9', textAlign: 'center' }}>{hrsText}</span>
        <div>
          {task.status !== 'complete' && task.status !== 'rework' ? (
            <div style={{ display: 'flex', alignItems: 'center', gap: 6 }}>
              <div
                style={{ position: 'relative', width: 80, height: 20, cursor: canEdit ? 'pointer' : 'default', userSelect: 'none' }}
                onClick={e => {
                  if (!canEdit) return;
                  const rect = (e.currentTarget as HTMLDivElement).getBoundingClientRect();
                  const pct = Math.round(((e.clientX - rect.left) / rect.width) * 100 / 5) * 5;
                  handleWip(Math.max(0, Math.min(100, pct)));
                }}
              >
                <div style={{ position: 'absolute', top: 8, left: 0, width: '100%', height: 4, borderRadius: 2, background: '#3a4a60' }} />
                <div style={{ position: 'absolute', top: 8, left: 0, width: `${localWip}%`, height: 4, borderRadius: 2, background: localWip > 60 ? C.green : C.amber, transition: 'width 0.15s ease' }} />
                <div style={{ position: 'absolute', top: 4, left: `calc(${localWip}% - 6px)`, width: 12, height: 12, borderRadius: '50%', background: localWip > 60 ? C.green : localWip > 0 ? C.amber : '#6b7a8d', border: '2px solid #fff', boxShadow: '0 1px 3px rgba(0,0,0,0.3)', transition: 'left 0.15s ease' }} />
              </div>
              <span style={{ fontSize: 11, fontWeight: 600, color: localWip > 60 ? C.green : localWip > 0 ? C.amber : '#B4B2A9', minWidth: 28 }}>{localWip}%</span>
            </div>
          ) : <Badge status={task.status} wipPct={task.wipPct} reviewStatus={task.reviewStatus} />}
        </div>
        <div style={{ display: 'flex', flexDirection: 'column', gap: 3, alignItems: 'flex-end' }}>
          <div style={{ display: 'flex', gap: 3 }}>
            {canComplete && <button onClick={() => setShowComplete(true)} style={{ padding: '2px 0', width: 42, borderRadius: 3, border: 'none', background: C.green, color: '#fff', fontSize: 10, fontWeight: 700, cursor: 'pointer', textAlign: 'center' }}>Done</button>}
            {isManager && canUndo && <button onClick={() => setShowUndo(true)} style={{ padding: '2px 0', width: 42, borderRadius: 3, border: 'none', background: C.amber, color: '#fff', fontSize: 10, fontWeight: 700, cursor: 'pointer', textAlign: 'center' }}>Undo</button>}
            {isManager && canEdit && <button onClick={() => onEdit(task)} style={{ padding: '2px 0', width: 42, borderRadius: 3, border: 'none', background: C.purple, color: '#fff', fontSize: 10, fontWeight: 700, cursor: 'pointer', textAlign: 'center' }}>Edit</button>}
          </div>
          <div style={{ display: 'flex', gap: 3 }}>
            {isManager && canEdit && <button onClick={() => { if (confirm('Delete task "' + task.description + '"?')) onDelete(task); }} style={{ padding: '2px 0', width: 42, borderRadius: 3, border: `1px solid ${C.redDk}`, background: 'transparent', color: C.redDk, fontSize: 10, fontWeight: 700, cursor: 'pointer', textAlign: 'center' }}>Del</button>}
            <button onClick={() => setShowHistory(!showHistory)} style={{ padding: '2px 0', width: 42, borderRadius: 3, border: `1px solid ${C.muted}`, background: showHistory ? C.muted : 'transparent', color: showHistory ? '#fff' : C.muted, fontSize: 10, fontWeight: 700, cursor: 'pointer', textAlign: 'center' }}>Log</button>
          </div>
        </div>
      </div>
    </>
  );
};

// ── Code Select Helpers ──────────────────────────────────────
const CodeOpts: React.FC<{ codes: typeof PROD_TASK_CODES; value: string; onChange: (v: string) => void; style: React.CSSProperties }> = ({ codes, value, onChange, style }) => (
  <select value={value} onChange={e => onChange(e.target.value)} style={style}>
    {codes.map(g => <optgroup key={g.group} label={g.group}>{g.codes.map(c => <option key={c.id} value={c.id}>{c.id} — {c.label}</option>)}</optgroup>)}
  </select>
);

// ── Edit Task Modal ──────────────────────────────────────────
interface EditTaskModalProps {
  task: ITask;
  team: ITeamMember[];
  activeProjects: IProject[];
  onSave: (t: ITask) => void;
  onCancel: () => void;
}
const EditTaskModal: React.FC<EditTaskModalProps> = ({ task, team, activeProjects, onSave, onCancel }) => {
  const [d, setD] = React.useState<ITask>({ ...task });
  const set = (field: string, val: string | number): void => setD(p => ({ ...p, [field]: val }));
  const isProd = !isNonProd(d.project);
  const ss: React.CSSProperties = { padding: '6px 10px', background: '#f5f6f8', border: `1px solid ${C.border}`, borderRadius: 4, color: C.text, fontSize: 12, fontWeight: 600, width: '100%', boxSizing: 'border-box' as const };
  return (
    <Modal onClose={onCancel}>
      <div style={{ fontSize: 15, fontWeight: 700, color: C.text, marginBottom: 16 }}>Edit Task</div>
      <div style={{ display: 'grid', gridTemplateColumns: '1fr 1fr', gap: 10, marginBottom: 10 }}>
        <div>
          <label style={{ fontSize: 11, fontWeight: 700, color: C.muted, display: 'block', marginBottom: 4 }}>PROJECT</label>
          <select value={d.project} onChange={e => set('project', e.target.value)} style={ss}>
            <option value="3E-INT">3E-INT (Non-production)</option>
            {activeProjects.map(p => <option key={p.projNum} value={p.projNum}>{p.projNum}{p.isEwo ? ' (EWO)' : ''}</option>)}
          </select>
        </div>
        <div>
          <label style={{ fontSize: 11, fontWeight: 700, color: C.muted, display: 'block', marginBottom: 4 }}>TASK CODE</label>
          <CodeOpts codes={isProd ? PROD_TASK_CODES : NON_PROD_TASK_CODES} value={d.taskCode} onChange={v => set('taskCode', v)} style={ss} />
        </div>
      </div>
      <div style={{ marginBottom: 10 }}>
        <label style={{ fontSize: 11, fontWeight: 700, color: C.muted, display: 'block', marginBottom: 4 }}>DESCRIPTION</label>
        <input value={d.description} onChange={e => set('description', e.target.value)} style={ss} />
      </div>
      <div style={{ display: 'grid', gridTemplateColumns: '1fr 1fr 1fr 1fr', gap: 10, marginBottom: 10 }}>
        <div>
          <label style={{ fontSize: 11, fontWeight: 700, color: C.muted, display: 'block', marginBottom: 4 }}>WHO</label>
          <select value={d.assignee} onChange={e => set('assignee', e.target.value)} style={ss}>
            {team.map(t => <option key={t.initials} value={t.initials}>{t.initials}</option>)}
          </select>
        </div>
        <div>
          <label style={{ fontSize: 11, fontWeight: 700, color: C.muted, display: 'block', marginBottom: 4 }}>DAY</label>
          <select value={d.day} onChange={e => set('day', parseInt(e.target.value))} style={ss}>
            {DAYS.map((day, i) => <option key={i} value={i}>{day.substring(0, 3)}</option>)}
          </select>
        </div>
        <div>
          <label style={{ fontSize: 11, fontWeight: 700, color: C.muted, display: 'block', marginBottom: 4 }}>HOURS</label>
          <input type="number" step="0.5" min="0" value={d.hoursPlanned} onChange={e => set('hoursPlanned', parseFloat(e.target.value) || 0)} style={ss} />
        </div>
        <div>
          <label style={{ fontSize: 11, fontWeight: 700, color: C.muted, display: 'block', marginBottom: 4 }}>PRIORITY</label>
          <select value={d.priority} onChange={e => set('priority', e.target.value)} style={ss}>
            <option value="high">High</option><option value="medium">Medium</option><option value="low">Low</option>
          </select>
        </div>
      </div>
      <div style={{ display: 'flex', gap: 10, justifyContent: 'flex-end', marginTop: 16 }}>
        <button onClick={onCancel} style={{ padding: '8px 18px', borderRadius: 5, border: `1px solid ${C.border}`, background: 'transparent', color: C.muted, fontSize: 13, fontWeight: 600, cursor: 'pointer' }}>Cancel</button>
        <button onClick={() => onSave(d)} style={{ padding: '8px 18px', borderRadius: 5, border: 'none', background: C.green, color: '#fff', fontSize: 13, fontWeight: 700, cursor: 'pointer' }}>Save Changes</button>
      </div>
    </Modal>
  );
};

// ── Plan Week Panel ──────────────────────────────────────────
interface PlanRow {
  id: string; project: string; taskCode: string; description: string;
  assignee: string; day: number; hoursPlanned: number; priority: string; isInternal: boolean; wipPct?: number;
}

interface PlanWeekProps {
  lastWeekTasks: ITask[];
  team: ITeamMember[];
  activeProjects: IProject[];
  teamCapacity: number;
  monday: Date;
  currentUserInitials: string;
  onAddTasks: (tasks: Omit<ITask, 'id' | 'spId'>[]) => void;
  onClose: () => void;
}

const PlanWeekPanel: React.FC<PlanWeekProps> = ({ lastWeekTasks, team, activeProjects, teamCapacity, monday, currentUserInitials, onAddTasks, onClose }) => {
  const [rows, setRows] = React.useState<PlanRow[]>([]);
  const [mode, setMode] = React.useState<'production' | 'internal'>('production');
  const unfinished = lastWeekTasks.filter(t => t.status !== 'complete' && !isNonProd(t.project));
  const ss: React.CSSProperties = { padding: '5px 6px', background: '#f5f6f8', border: `1px solid ${C.border}`, borderRadius: 3, color: C.text, fontSize: 11 };

  const addRow = (): void => {
    const defAssignee = team.length > 0 ? team[0].initials : '';
    if (mode === 'production') {
      setRows(p => [...p, { id: `b-${Date.now()}-${Math.random()}`, project: activeProjects[0]?.projNum || '', taskCode: '03a', description: '', assignee: defAssignee, day: 0, hoursPlanned: 4, priority: 'medium', isInternal: false }]);
    } else {
      setRows(p => [...p, { id: `b-${Date.now()}-${Math.random()}`, project: '3E-INT', taskCode: '00c', description: '', assignee: defAssignee, day: 0, hoursPlanned: 1, priority: 'low', isInternal: true }]);
    }
  };
  const updateRow = (id: string, field: string, val: string | number): void => {
    setRows(p => p.map(r => { if (r.id !== id) return r; const u = { ...r, [field]: val }; if (field === 'project') u.isInternal = val === '3E-INT'; return u; }));
  };
  const removeRow = (id: string): void => setRows(p => p.filter(r => r.id !== id));
  const dupRow = (id: string): void => { const s = rows.find(r => r.id === id); if (s) setRows(p => [...p, { ...s, id: `b-${Date.now()}`, day: Math.min(s.day + 1, 4) }]); };

  const carryOne = (task: ITask): void => {
    const rem = round1(task.hoursPlanned * (100 - task.wipPct) / 100);
    setRows(p => [...p, { id: `carry-${task.id}`, project: task.project, taskCode: task.taskCode, description: `${task.description} (carried fwd)`, assignee: task.assignee, day: 0, hoursPlanned: Math.max(rem, 1), priority: task.priority, isInternal: false, wipPct: task.wipPct }]);
  };
  const carryAll = (): void => {
    const newRows = unfinished.filter(t => !rows.some(r => r.description.includes(t.description))).map(t => {
      const rem = round1(t.hoursPlanned * (100 - t.wipPct) / 100);
      return { id: `carry-${t.id}`, project: t.project, taskCode: t.taskCode, description: `${t.description} (carried fwd)`, assignee: t.assignee, day: 0, hoursPlanned: Math.max(rem, 1), priority: t.priority, isInternal: false, wipPct: t.wipPct };
    });
    setRows(p => [...p, ...newRows]);
  };

  const submit = (): void => {
    const wsd = weekKey(monday);
    const valid = rows.filter(r => r.description.trim());
    const tasks: Omit<ITask, 'id' | 'spId'>[] = valid.map(r => {
      const carried = r.wipPct !== undefined && r.wipPct > 0;
      return {
        project: r.project, taskCode: r.taskCode, description: r.description, assignee: r.assignee, day: r.day, weekStartDate: wsd,
        hoursPlanned: r.hoursPlanned, hoursActual: carried ? round1(r.hoursPlanned * r.wipPct! / 100) : 0, wipPct: carried ? r.wipPct! : 0, status: carried ? 'wip' : 'not_started', priority: r.priority,
        completedBy: '', completedAt: '', completionNote: '', reviewedBy: '', reviewStatus: '',
        history: [{ action: 'created', user: currentUserInitials, ts: now(), detail: `${r.hoursPlanned}h planned — weekly planning${carried ? ` (carried fwd at ${r.wipPct}%)` : ''}` }]
      };
    });
    onAddTasks(tasks);
  };

  const prodRows = rows.filter(r => !r.isInternal);
  const intRows = rows.filter(r => r.isInternal);
  const totalH = round1(rows.reduce((s, r) => s + r.hoursPlanned, 0));
  const prodH = round1(prodRows.reduce((s, r) => s + r.hoursPlanned, 0));
  const intH = round1(intRows.reduce((s, r) => s + r.hoursPlanned, 0));
  const perPerson: Record<string, number> = {};
  team.forEach(t => { perPerson[t.initials] = 0; });
  rows.forEach(r => { if (perPerson[r.assignee] !== undefined) perPerson[r.assignee] += r.hoursPlanned; });

  const visRows = mode === 'production' ? prodRows : intRows;
  const isProd = mode === 'production';
  const validCount = rows.filter(r => r.description.trim()).length;

  return (
    <div style={{ background: 'rgba(42,158,42,0.05)', border: `2px solid ${C.green}`, borderRadius: 8, padding: 20, marginBottom: 16 }}>
      <div style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'center', marginBottom: 16 }}>
        <div>
          <div style={{ fontSize: 16, fontWeight: 600, color: C.text }}>Plan week — {formatDateRange(monday)}</div>
          <div style={{ fontSize: 12, color: C.muted, marginTop: 2 }}>Team capacity: {teamCapacity}h production/wk</div>
        </div>
        <button onClick={onClose} style={{ padding: '6px 14px', borderRadius: 4, border: `1px solid ${C.border}`, background: 'transparent', color: C.muted, fontSize: 12, cursor: 'pointer' }}>Cancel</button>
      </div>

      {unfinished.length > 0 && (
        <div style={{ background: C.bg, borderRadius: 6, border: `1px solid ${C.border}`, padding: '12px 16px', marginBottom: 16 }}>
          <div style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'center', marginBottom: 8 }}>
            <span style={{ fontSize: 13, fontWeight: 600, color: C.amberLt }}>{unfinished.length} unfinished from last week</span>
            <button onClick={carryAll} style={{ padding: '5px 12px', borderRadius: 4, border: 'none', background: C.amber, color: C.bg, fontSize: 11, fontWeight: 600, cursor: 'pointer' }}>Carry all forward</button>
          </div>
          {unfinished.map(t => {
            const added = rows.some(r => r.description.includes(t.description));
            return (
              <div key={t.id} style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'center', padding: '5px 0', borderBottom: `1px solid ${C.card}`, opacity: added ? 0.4 : 1 }}>
                <div style={{ display: 'flex', alignItems: 'center', gap: 8, fontSize: 12 }}>
                  <span style={{ color: C.greenLt, fontWeight: 600, fontSize: 11, minWidth: 50 }}>{t.project}</span>
                  <span style={{ color: C.muted, fontSize: 11 }}>{t.taskCode}</span>
                  <span style={{ color: '#1a2030' }}>{t.description}</span>
                  <span style={{ color: C.muted, fontSize: 11 }}>{t.assignee}</span>
                  <Badge status={t.status} wipPct={t.wipPct} reviewStatus={t.reviewStatus} />
                </div>
                {!added ? <button onClick={() => carryOne(t)} style={{ padding: '3px 10px', borderRadius: 3, border: `1px solid ${C.amber}`, background: 'transparent', color: C.amberLt, fontSize: 10, cursor: 'pointer' }}>Carry fwd</button>
                  : <span style={{ fontSize: 10, color: C.greenLt }}>Added</span>}
              </div>
            );
          })}
        </div>
      )}

      <div style={{ display: 'flex', gap: 4, marginBottom: 12 }}>
        <button onClick={() => setMode('production')} style={{ padding: '6px 16px', borderRadius: 4, fontSize: 12, fontWeight: 600, border: 'none', cursor: 'pointer', background: isProd ? C.green : '#262623', color: isProd ? '#fff' : C.muted }}>Production ({prodRows.length})</button>
        <button onClick={() => setMode('internal')} style={{ padding: '6px 16px', borderRadius: 4, fontSize: 12, fontWeight: 600, border: 'none', cursor: 'pointer', background: !isProd ? C.dim : '#262623', color: '#fff' }}>Non-production ({intRows.length})</button>
      </div>

      <div style={{ display: 'grid', gridTemplateColumns: isProd ? '90px 110px 1fr 54px 60px 50px 40px 50px' : '1fr 110px 1fr 54px 60px 50px 50px', gap: 6, padding: '4px 0', fontSize: 10, fontWeight: 600, color: C.dim, letterSpacing: '0.04em' }}>
        <span>{isProd ? 'PROJECT' : '3E-INT'}</span><span>TASK CODE</span><span>DESCRIPTION</span><span>WHO</span><span>DAY</span><span>HRS</span>{isProd && <span>PRI</span>}<span></span>
      </div>

      {visRows.map(r => (
        <div key={r.id} style={{ display: 'grid', gridTemplateColumns: isProd ? '90px 110px 1fr 54px 60px 50px 40px 50px' : '1fr 110px 1fr 54px 60px 50px 50px', gap: 6, padding: '4px 0', alignItems: 'center', borderBottom: `1px solid ${C.card}` }}>
          {isProd ? (
            <select value={r.project} onChange={e => updateRow(r.id, 'project', e.target.value)} style={ss}>
              {activeProjects.map(p => <option key={p.projNum} value={p.projNum}>{p.projNum}{p.isEwo ? ' (EWO)' : ''}</option>)}
            </select>
          ) : <span style={{ fontSize: 11, color: C.muted, padding: '0 6px' }}>3E-INT</span>}
          <CodeOpts codes={isProd ? PROD_TASK_CODES : NON_PROD_TASK_CODES} value={r.taskCode} onChange={v => updateRow(r.id, 'taskCode', v)} style={ss} />
          <input value={r.description} onChange={e => updateRow(r.id, 'description', e.target.value)} placeholder="Description..."
            style={{ padding: '5px 8px', background: '#f5f6f8', border: `1px solid ${C.border}`, borderRadius: 3, color: C.text, fontSize: 11 }} />
          <select value={r.assignee} onChange={e => updateRow(r.id, 'assignee', e.target.value)} style={ss}>
            {team.map(t => <option key={t.initials} value={t.initials}>{t.initials}</option>)}
          </select>
          <select value={r.day} onChange={e => updateRow(r.id, 'day', parseInt(e.target.value))} style={ss}>
            {DAYS.map((d, i) => <option key={i} value={i}>{d.substring(0, 3)}</option>)}
          </select>
          <input type="number" step="0.5" min="0.5" max="12" value={r.hoursPlanned} onChange={e => updateRow(r.id, 'hoursPlanned', parseFloat(e.target.value) || 0)}
            style={{ ...ss, width: '100%', textAlign: 'center' }} />
          {isProd && <select value={r.priority} onChange={e => updateRow(r.id, 'priority', e.target.value)} style={{ ...ss, width: '100%' }}><option value="high">H</option><option value="medium">M</option><option value="low">L</option></select>}
          <div style={{ display: 'flex', gap: 2 }}>
            <button onClick={() => dupRow(r.id)} title="Duplicate to next day" style={{ padding: '2px 5px', borderRadius: 3, border: `1px solid ${C.border}`, background: 'transparent', color: C.muted, fontSize: 10, cursor: 'pointer' }}>+</button>
            <button onClick={() => removeRow(r.id)} style={{ padding: '2px 5px', borderRadius: 3, border: 'none', background: 'transparent', color: C.dim, fontSize: 12, cursor: 'pointer' }}>x</button>
          </div>
        </div>
      ))}

      <button onClick={addRow} style={{ padding: '8px 0', fontSize: 12, fontWeight: 600, color: isProd ? C.greenLt : '#B4B2A9', background: 'transparent', border: `1px dashed ${isProd ? C.green : C.dim}`, borderRadius: 4, cursor: 'pointer', width: '100%', marginTop: 4, textAlign: 'center' }}>+ Add {isProd ? 'production' : 'non-production'} task</button>

      {rows.length > 0 && (
        <div style={{ background: C.bg, borderRadius: 6, padding: '10px 16px', marginTop: 12, marginBottom: 16 }}>
          <div style={{ display: 'flex', gap: 16, alignItems: 'center', flexWrap: 'wrap', fontSize: 12, color: C.muted }}>
            <div><span style={{ fontWeight: 600, color: C.text }}>{rows.length}</span> tasks · <span style={{ fontWeight: 600, color: C.greenLt }}>{prodH}h</span> prod · <span style={{ fontWeight: 600, color: C.muted }}>{intH}h</span> non-prod · <span style={{ fontWeight: 600, color: C.text }}>{totalH}h</span> total</div>
            <div style={{ width: 1, height: 20, background: C.border }} />
            {team.map(t => {
              const h = round1(perPerson[t.initials] || 0);
              if (h === 0) return null;
              const over = h > t.totalHrsPerWeek;
              return <div key={t.initials} style={{ fontSize: 11, color: over ? C.red : C.muted }}>{t.initials}: <span style={{ fontWeight: 600, color: over ? C.red : '#D3D1C7' }}>{h}h</span> / {t.totalHrsPerWeek}h{over && <span style={{ color: C.red }}> !</span>}</div>;
            })}
          </div>
        </div>
      )}

      <div style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'center' }}>
        <div style={{ fontSize: 11, color: C.dim }}>Tasks added to {formatDateRange(monday)}. Edit individually after.</div>
        <button onClick={submit} disabled={validCount === 0}
          style={{ padding: '10px 24px', borderRadius: 6, border: 'none', background: validCount > 0 ? C.green : '#c0c8d4', color: '#fff', fontSize: 14, fontWeight: 600, cursor: validCount > 0 ? 'pointer' : 'default' }}>
          Add {validCount} tasks to week
        </button>
      </div>
    </div>
  );
};

// ═══════════════════════════════════════════════════════════════
// MAIN TASKBOARD
// ═══════════════════════════════════════════════════════════════
const TaskBoard: React.FC<TaskBoardProps> = ({ spService, projects, userDisplayName, isManager, toast }) => {
  const [weekOffset, setWeekOffset] = React.useState(0);
  const [tasks, setTasks] = React.useState<Record<number, ITask[]>>({});
  const [team, setTeam] = React.useState<ITeamMember[]>([]);
  const [filterPerson, setFilterPerson] = React.useState('All');
  const [viewMode, setViewMode] = React.useState<'day' | 'person'>('day');
  const [planningMode, setPlanningMode] = React.useState(false);
  const [loading, setLoading] = React.useState(true);
  const [managerUnlocked, setManagerUnlocked] = React.useState<Record<number, boolean>>({});

  // Derive user initials from display name
  const currentUserInitials = React.useMemo(() => {
    const parts = userDisplayName.trim().split(/\s+/);
    if (parts.length >= 2) return (parts[0][0] + parts[parts.length - 1][0]).toUpperCase();
    return userDisplayName.substring(0, 2).toUpperCase();
  }, [userDisplayName]);

  const monday = React.useMemo(() => {
    const m = getWeekStart(new Date()); m.setDate(m.getDate() + weekOffset * 7); return m;
  }, [weekOffset]);
  const todayIdx = weekOffset === 0 ? new Date().getDay() - 1 : -1;
  const isPast = weekOffset < 0;
  const isLocked = isPast && !managerUnlocked[weekOffset];

  const weekTasks = tasks[weekOffset] || [];
  const lastWeekTasks = tasks[weekOffset - 1] || [];
  const filtered = filterPerson === 'All' ? weekTasks : weekTasks.filter(t => t.assignee === filterPerson);
  const activeProjects = projects;
  const teamCapacity = team.reduce((s, t) => s + t.prodHrsPerWeek, 0);

  // Load team + tasks for current and adjacent weeks
  React.useEffect(() => {
    const load = async (): Promise<void> => {
      try {
        setLoading(true);
        const members = await spService.loadTeamMembers();
        setTeam(members.filter(m => m.isActive));

        // Load tasks for weeks -4 to +1
        const allTasks: Record<number, ITask[]> = {};
        for (let i = -2; i <= 1; i++) {
          const m = getWeekStart(new Date());
          m.setDate(m.getDate() + i * 7);
          const wk = weekKey(m);
          const wkTasks = await spService.loadTasks(wk);
          allTasks[i] = wkTasks;
        }
        setTasks(allTasks);
      } catch (e) {
        toast('Failed to load tasks: ' + String(e));
      } finally {
        setLoading(false);
      }
    };
    load().catch(() => undefined);
  }, []); // eslint-disable-line react-hooks/exhaustive-deps

  const updateTask = async (task: ITask): Promise<void> => {
    try {
      if (task.spId) await spService.updateTask(task.spId, task);
      setTasks(p => ({ ...p, [weekOffset]: (p[weekOffset] || []).map(t => t.id === task.id ? task : t) }));
    } catch (e) { toast('Update failed: ' + String(e)); }
  };

  const addNewTasks = async (newTasks: Omit<ITask, 'id' | 'spId'>[]): Promise<void> => {
    try {
      const created: ITask[] = [];
      for (const t of newTasks) {
        const spId = await spService.addTask(t as ITask);
        created.push({ ...t, id: String(spId), spId } as ITask);
      }
      setTasks(p => ({ ...p, [weekOffset]: [...(p[weekOffset] || []), ...created] }));
      setPlanningMode(false);
      toast(`Added ${created.length} tasks to week`);
    } catch (e) { toast('Failed to add tasks: ' + String(e)); }
  };

  const deleteTask = async (task: ITask): Promise<void> => {
    try {
      if (task.spId) await spService.deleteTask(task.spId);
      setTasks(p => ({ ...p, [weekOffset]: (p[weekOffset] || []).filter(t => t.id !== task.id) }));
      toast('Task deleted');
    } catch (e) { toast('Delete failed: ' + String(e)); }
  };

  const [editTask, setEditTask] = React.useState<ITask | null>(null);

  const saveEditedTask = async (updated: ITask): Promise<void> => {
    await updateTask(updated);
    setEditTask(null);
    toast('Task updated');
  };

  // Computed stats
  const prodTasks = weekTasks.filter(t => !isNonProd(t.project));
  const intTasks = weekTasks.filter(t => isNonProd(t.project));
  const hoursPlanned = round1(prodTasks.reduce((s, t) => s + t.hoursPlanned, 0));
  const hoursActual = round1(prodTasks.reduce((s, t) => s + t.hoursActual, 0));
  const completionPct = prodTasks.length > 0 ? Math.round(prodTasks.reduce((s, t) => s + t.wipPct, 0) / (prodTasks.length * 100) * 100) : 0;
  const completeTasks = prodTasks.filter(t => t.status === 'complete').length;
  const pendingReview = weekTasks.filter(t => t.status === 'complete' && !t.reviewStatus).length;
  const intHours = round1(intTasks.reduce((s, t) => s + t.hoursPlanned, 0));

  // Week tabs
  const weekTabs = React.useMemo(() => {
    const tabs = [];
    for (let i = -2; i <= 1; i++) {
      const m = getWeekStart(new Date()); m.setDate(m.getDate() + i * 7);
      tabs.push({ offset: i, monday: m, label: formatWeekLabel(m), dates: formatDateRange(m), isCurrent: i === 0, isPast: i < 0 });
    }
    return tabs;
  }, []);

  if (loading) return <div style={{ padding: 40, textAlign: 'center', color: C.muted }}>Loading tasks...</div>;

  return (
    <div style={{ color: C.text, fontSize: 12, fontWeight: 600, padding: '0 0 24px' }}>
      {/* Week tabs */}
      <div style={{ display: 'flex', gap: 4, marginBottom: 16, overflowX: 'auto', paddingBottom: 4 }}>
        {weekTabs.map(wt => (
          <button key={wt.offset} onClick={() => { setWeekOffset(wt.offset); setPlanningMode(false); }}
            style={{ padding: '8px 14px', borderRadius: 6, border: wt.offset === weekOffset ? `2px solid ${C.purple}` : `1px solid ${C.border}`, background: wt.offset === weekOffset ? C.purpleBg : wt.isCurrent ? C.card : C.bg, color: wt.offset === weekOffset ? C.purpleLt : wt.isCurrent ? C.text : C.muted, cursor: 'pointer', minWidth: 110, textAlign: 'center', flexShrink: 0, fontFamily: 'inherit', fontSize: 12 }}>
            <div style={{ fontWeight: 600 }}>{wt.label}</div>
            <div style={{ fontSize: 10, fontWeight: 600, marginTop: 2, opacity: 0.7 }}>{wt.dates}</div>
            {wt.isPast && <div style={{ fontSize: 9, fontWeight: 600, marginTop: 2, color: C.purple }}>LOCKED</div>}
            {wt.isCurrent && <div style={{ fontSize: 9, fontWeight: 600, marginTop: 2, color: C.greenLt }}>CURRENT</div>}
          </button>
        ))}
      </div>

      {/* Summary bar */}
      <div style={{ display: 'flex', gap: 16, alignItems: 'center', marginBottom: 16, padding: '12px 16px', background: C.card, borderRadius: 8, flexWrap: 'wrap' }}>
        <div style={{ flex: 1, minWidth: 120 }}>
          <span style={{ fontSize: 22, fontWeight: 700 }}>{hoursActual}</span>
          <span style={{ fontSize: 14, color: C.muted }}> / {hoursPlanned}h</span>
          <div style={{ fontSize: 11, color: C.muted, marginTop: 2 }}>Production hours (capacity: {teamCapacity}h/wk)</div>
        </div>
        <div style={{ width: 1, height: 40, background: C.border }} />
        <div>
          <span style={{ fontSize: 22, fontWeight: 700, color: completionPct >= 80 ? C.greenLt : completionPct >= 40 ? C.amberLt : C.red }}>{completionPct}%</span>
          <div style={{ fontSize: 11, color: C.muted, marginTop: 2 }}>Completion ({completeTasks}/{prodTasks.length})</div>
        </div>
        <div style={{ width: 1, height: 40, background: C.border }} />
        <div>
          <span style={{ fontSize: 16, fontWeight: 600, color: C.muted }}>{intHours}h</span>
          <div style={{ fontSize: 11, color: C.dim, marginTop: 2 }}>Non-production</div>
        </div>
        {pendingReview > 0 && <><div style={{ width: 1, height: 40, background: C.border }} /><span style={{ fontSize: 12, color: C.amberLt, background: 'rgba(212,136,10,0.10)', padding: '4px 10px', borderRadius: 4 }}>{pendingReview} pending review</span></>}
        <div style={{ flex: 1 }} />
        <div style={{ display: 'flex', alignItems: 'center', gap: 8 }}>
          {isLocked && <span style={{ fontSize: 11, color: C.purple }}>LOCKED</span>}
          {isLocked && isManager && <button onClick={() => setManagerUnlocked(p => ({ ...p, [weekOffset]: true }))} style={{ padding: '4px 10px', borderRadius: 4, border: 'none', background: C.purple, color: '#fff', fontSize: 11, fontWeight: 600, cursor: 'pointer' }}>Unlock</button>}
          {isPast && managerUnlocked[weekOffset] && <><span style={{ fontSize: 11, color: C.amberLt }}>UNLOCKED</span><button onClick={() => setManagerUnlocked(p => { const n = { ...p }; delete n[weekOffset]; return n; })} style={{ padding: '4px 10px', borderRadius: 4, border: 'none', background: C.amber, color: C.bg, fontSize: 11, fontWeight: 600, cursor: 'pointer' }}>Re-lock</button></>}
          {!isPast && <span style={{ fontSize: 11, color: C.greenLt }}>{weekOffset === 0 ? 'ACTIVE' : 'FUTURE'}</span>}
        </div>
        {isManager && !isLocked && (
          <button onClick={() => setPlanningMode(!planningMode)}
            style={{ padding: '8px 18px', borderRadius: 6, border: 'none', background: planningMode ? C.redDk : C.green, color: '#fff', fontSize: 13, fontWeight: 600, cursor: 'pointer', whiteSpace: 'nowrap' }}>
            {planningMode ? 'Close planner' : 'Plan week'}
          </button>
        )}
      </div>

      {isLocked && <div style={{ padding: '10px 16px', background: C.purpleBg, borderRadius: 6, border: `1px solid ${C.purple}`, marginBottom: 12, fontSize: 12, color: C.purpleLt }}>This week is locked (read-only).{isManager ? ' You can unlock to make corrections.' : ''}</div>}

      {planningMode && (
        <PlanWeekPanel
          lastWeekTasks={lastWeekTasks}
          team={team}
          activeProjects={activeProjects}
          teamCapacity={teamCapacity}
          monday={monday}
          currentUserInitials={currentUserInitials}
          onAddTasks={addNewTasks}
          onClose={() => setPlanningMode(false)}
        />
      )}

      {/* Filters */}
      <div style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'center', marginBottom: 12 }}>
        <select value={filterPerson} onChange={e => setFilterPerson(e.target.value)}
          style={{ padding: '6px 12px', background: C.card, border: `1px solid ${C.border}`, borderRadius: 4, color: C.text, fontSize: 12, fontWeight: 600, maxWidth: 260 }}>
          <option value="All">All team</option>
          {team.map(t => <option key={t.initials} value={t.initials}>{t.initials} — {t.fullName} ({t.prodHrsPerWeek}h prod)</option>)}
        </select>
        <div style={{ display: 'flex', borderRadius: 4, overflow: 'hidden', border: `1px solid ${C.border}` }}>
          {(['day', 'person'] as const).map(m => (
            <button key={m} onClick={() => setViewMode(m)} style={{ padding: '4px 14px', border: 'none', fontSize: 12, fontWeight: 600, cursor: 'pointer', background: viewMode === m ? C.purple : C.card, color: viewMode === m ? '#fff' : C.muted }}>By {m}</button>
          ))}
        </div>
      </div>

      {/* Column headers */}
      <div style={{ display: 'grid', gridTemplateColumns: '4px 70px 52px 1fr 36px 50px 120px 110px', gap: 6, padding: '4px 8px', fontSize: 10, fontWeight: 600, color: C.dim, letterSpacing: '0.04em', borderBottom: `1px solid ${C.border}` }}>
        <span></span><span>PROJECT</span><span>TASK</span><span>DESCRIPTION</span>
        <span style={{ textAlign: 'center' }}>WHO</span><span style={{ textAlign: 'center' }}>HOURS</span><span>STATUS</span><span style={{ textAlign: 'right' }}>ACTIONS</span>
      </div>

      {/* DAY VIEW */}
      {viewMode === 'day' && DAYS.map((day, di) => {
        const dayAll = filtered.filter(t => t.day === di);
        const dayProd = dayAll.filter(t => !isNonProd(t.project));
        const dayInt = dayAll.filter(t => isNonProd(t.project));
        const isToday = di === todayIdx;
        const dayH = round1(dayAll.reduce((s, t) => s + t.hoursPlanned, 0));
        return (
          <div key={day} style={{ marginBottom: 6, borderRadius: 6, overflow: 'hidden', border: isToday ? `1px solid ${C.purple}` : `1px solid ${C.border}` }}>
            <div style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'center', padding: '8px 12px', background: isToday ? C.purpleBg : C.card }}>
              <div style={{ display: 'flex', alignItems: 'center', gap: 10 }}>
                <span style={{ fontSize: 13, fontWeight: 600, color: isToday ? C.purpleLt : '#D3D1C7' }}>{day}</span>
                <span style={{ fontSize: 11, color: C.dim }}>{getDayDate(monday, di)}</span>
                {isToday && <span style={{ fontSize: 10, fontWeight: 600, color: C.purple, background: 'rgba(83,74,183,0.12)', padding: '1px 8px', borderRadius: 8 }}>TODAY</span>}
              </div>
              <div style={{ fontSize: 11, color: C.muted, display: 'flex', gap: 16 }}>
                <span>{dayProd.length} production</span>
                {dayInt.length > 0 && <span style={{ color: C.dim }}>{dayInt.length} non-prod</span>}
                <span>{dayH}h</span>
                <span>{dayAll.filter(t => t.status === 'complete').length}/{dayAll.length} done</span>
              </div>
            </div>
            {dayProd.length > 0 && <div style={{ padding: '0 4px' }}>{dayProd.map(t => <TaskRow key={t.id} task={t} onUpdate={updateTask} onEdit={setEditTask} onDelete={deleteTask} isLocked={isLocked} isManager={isManager} currentUserInitials={currentUserInitials} isInternal={false} />)}</div>}
            {dayInt.length > 0 && (
              <div style={{ borderTop: `1px dashed ${C.border}`, margin: '0 4px' }}>
                <div style={{ padding: '6px 8px 2px', display: 'flex', alignItems: 'center', gap: 8 }}>
                  <span style={{ fontSize: 10, fontWeight: 600, color: C.dim, letterSpacing: '0.04em' }}>NON-PRODUCTION</span>
                  <div style={{ flex: 1, height: 1, background: C.border }} />
                  <span style={{ fontSize: 10, color: C.dim }}>{round1(dayInt.reduce((s, t) => s + t.hoursPlanned, 0))}h</span>
                </div>
                {dayInt.map(t => <TaskRow key={t.id} task={t} onUpdate={updateTask} onEdit={setEditTask} onDelete={deleteTask} isLocked={isLocked} isManager={isManager} currentUserInitials={currentUserInitials} isInternal={true} />)}
              </div>
            )}
            {dayAll.length === 0 && !isLocked && <div style={{ padding: 12, textAlign: 'center', fontSize: 11, color: C.dim }}>No tasks — use "Plan week" to add</div>}
          </div>
        );
      })}

      {/* PERSON VIEW */}
      {viewMode === 'person' && team.filter(t => filterPerson === 'All' || filterPerson === t.initials).map(person => {
        const pt = filtered.filter(t => t.assignee === person.initials);
        const prodPt = pt.filter(t => !isNonProd(t.project));
        const intPt = pt.filter(t => isNonProd(t.project));
        const ph = round1(pt.reduce((s, t) => s + t.hoursPlanned, 0));
        const prodH2 = round1(prodPt.reduce((s, t) => s + t.hoursPlanned, 0));
        const intH2 = round1(intPt.reduce((s, t) => s + t.hoursPlanned, 0));
        const pct = Math.min(Math.round((ph / person.totalHrsPerWeek) * 100), 100);
        const over = ph > person.totalHrsPerWeek;
        const barCol = over ? C.red : pct > 80 ? C.amber : C.green;
        return (
          <div key={person.initials} style={{ marginBottom: 6, borderRadius: 6, overflow: 'hidden', border: `1px solid ${C.border}` }}>
            <div style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'center', padding: '8px 12px', background: C.card }}>
              <div style={{ display: 'flex', alignItems: 'center', gap: 10 }}>
                <div style={{ width: 28, height: 28, borderRadius: '50%', background: over ? '#501313' : '#1a2a20', border: `1px solid ${over ? C.redDk : C.green}`, display: 'flex', alignItems: 'center', justifyContent: 'center', fontSize: 10, fontWeight: 600, color: over ? '#F09595' : C.greenLt }}>{person.initials}</div>
                <span style={{ fontSize: 13, fontWeight: 600 }}>{person.fullName}</span>
                <span style={{ fontSize: 11, color: C.dim }}>{person.roleType} · {person.productionPct}% prod</span>
              </div>
              <div style={{ display: 'flex', alignItems: 'center', gap: 12, fontSize: 11, color: C.muted }}>
                {prodH2 > 0 && <span>{prodH2}h prod</span>}
                {intH2 > 0 && <span style={{ color: C.dim }}>{intH2}h non-prod</span>}
                <span style={{ color: over ? C.red : C.muted }}>{ph}h / {person.totalHrsPerWeek}h</span>
                <div style={{ width: 60, height: 6, background: C.border, borderRadius: 3, overflow: 'hidden' }}>
                  <div style={{ width: `${pct}%`, height: '100%', borderRadius: 3, background: barCol }} />
                </div>
              </div>
            </div>
            <div style={{ padding: '0 4px' }}>
              {DAYS.map((day, di) => {
                const dayProd2 = prodPt.filter(t => t.day === di);
                const dayInt2 = intPt.filter(t => t.day === di);
                if (dayProd2.length === 0 && dayInt2.length === 0) return null;
                return (
                  <div key={di}>
                    <div style={{ padding: '6px 8px 2px', fontSize: 10, fontWeight: 600, color: C.dim }}>{day} {getDayDate(monday, di)}</div>
                    {dayProd2.map(t => <TaskRow key={t.id} task={t} onUpdate={updateTask} onEdit={setEditTask} onDelete={deleteTask} isLocked={isLocked} isManager={isManager} currentUserInitials={currentUserInitials} isInternal={false} />)}
                    {dayInt2.length > 0 && (
                      <>
                        <div style={{ padding: '4px 8px 2px', display: 'flex', alignItems: 'center', gap: 6 }}>
                          <span style={{ fontSize: 9, color: C.dim, letterSpacing: '0.04em' }}>NON-PROD</span>
                          <div style={{ flex: 1, height: 1, background: C.border }} />
                        </div>
                        {dayInt2.map(t => <TaskRow key={t.id} task={t} onUpdate={updateTask} onEdit={setEditTask} onDelete={deleteTask} isLocked={isLocked} isManager={isManager} currentUserInitials={currentUserInitials} isInternal={true} />)}
                      </>
                    )}
                  </div>
                );
              })}
            </div>
          </div>
        );
      })}

      {weekTasks.length === 0 && !planningMode && <div style={{ padding: '40px 0', textAlign: 'center', color: C.dim, fontSize: 13 }}>No tasks for this week.{isManager ? ' Click "Plan week" to start planning.' : ''}</div>}

      {editTask && (
        <EditTaskModal
          task={editTask}
          team={team}
          activeProjects={activeProjects}
          onSave={saveEditedTask}
          onCancel={() => setEditTask(null)}
        />
      )}
    </div>
  );
};

export default TaskBoard;
