import * as React from 'react';
import { getQuoteAttachments, setQuoteAttachments } from './crmAttachmentStore';
import { DocumentUploadSection } from './crmDocumentUpload';
import {
  loadProjectsFromSharePoint, loadQuoteBudget, loadQuotesFromSharePoint, loadQuotesRemote,
  loadQuotesLocal, mergeQuotesWithLocal,
  saveQuoteBudget,
  upsertQuoteToSharePoint, deleteQuoteFromSharePoint, upsertWipToSharePoint,
  loadQuoteAttachmentsFromSharePoint, syncQuoteAttachmentsToSharePoint, loadQuoteAttachmentIndex,
  getMonthBudgetTarget, getQuoteBudgetForYear, normalizeQuoteBudget, normalizeQuoteBudgetStore,
  setQuoteBudgetForYear, MONTH_LABELS,
  type CrmQuoteBudget, type CrmQuoteBudgetStore,
} from './crmStorage';
import { normalizeCrmQuote } from './crmRfqNormalize';
import { nextProjNum, quoteToCrmProject, quoteToIProject } from './crmProjectHelpers';
import type { SharePointService } from '../../../shared/services/SharePointService';
import type { IProject } from '../../../shared/models/IProject';
import type { CrmAttachment, CrmPerson, CrmCompany, CrmProject, CrmQuote, CrmQuoteStatus, CrmLostReason, CrmRfqDiscipline } from './crmTypes';

const FF = 'Montserrat,sans-serif';
const LS_QUOTES = '3edge-crm-quotes';
const LS_PROJECTS = '3edge-crm-projects';

const C = {
  bg: '#f7f8fa', surface: '#ffffff', border: '#e2e5ea', borderMd: '#cdd1d9',
  text: '#1a2030', sub: '#4a5568', cell: '#364152', muted: '#8a97a8', green: '#2a9e2a',
  red: '#c0392b', purple: '#6c3fbf', thBg: '#f0f2f6', rowHover: '#f5f7fb',
};

const QUOTE_STATUSES: CrmQuoteStatus[] = ['Draft', 'Sent', 'Pending', 'Follow up', 'Lost'];
const DISCIPLINES: CrmRfqDiscipline[] = ['Steel', 'Concrete', 'Both'];
const SOURCES = ['Email', 'Phone', 'Referral', 'Repeat Client', 'Website', 'Other'];
const OWNERS = ['MK', 'SK', 'DC', 'JP'];

const LOST_REASONS: CrmLostReason[] = [
  'Price was too high',
  'Price was too low',
  'Client lost project',
  'Inexperienced',
  'Project been cancelled',
  'Other',
];

const mi: React.CSSProperties = {
  padding: '8px 10px', background: C.surface, border: `1px solid ${C.borderMd}`,
  borderRadius: 4, color: C.text, fontSize: 12.5, fontFamily: FF,
  width: '100%', boxSizing: 'border-box', outline: 'none',
};
const ml: React.CSSProperties = {
  fontSize: 10, fontWeight: 700, color: C.sub, letterSpacing: '.07em',
  textTransform: 'uppercase', marginBottom: 4, display: 'block', fontFamily: FF,
};

const fmtShortDate = (iso: string): string => {
  if (!iso) return '—';
  const d = new Date(iso + 'T00:00:00');
  if (isNaN(d.getTime())) return iso;
  return d.toLocaleDateString('en-AU', { day: 'numeric', month: 'short', year: '2-digit' });
};

const fmtMoney = (n: number): string =>
  '$' + Math.round(n).toLocaleString('en-AU');

const fmtBudget = (n: number): string => {
  if (n >= 1_000_000) {
    const m = n / 1_000_000;
    return '$' + (m % 1 === 0 ? m.toFixed(0) : m.toFixed(2).replace(/\.?0+$/, '')) + 'M';
  }
  if (n >= 10_000) return '$' + Math.round(n / 1000) + 'k';
  return fmtMoney(n);
};

const ensureQuPrefix = (v: string): string => {
  if (!v) return 'QU-';
  return v.toUpperCase().startsWith('QU-') ? v : `QU-${v.replace(/^QU-/i, '')}`;
};

const quoteNumFilled = (v: string): boolean => {
  const t = v.trim();
  return t.length > 0 && t !== 'QU-';
};

const csvCell = (v: string | number): string => {
  const s = String(v ?? '');
  return /[",\n]/.test(s) ? `"${s.replace(/"/g, '""')}"` : s;
};

const exportQuotesCsv = (
  rows: CrmQuote[],
  yr: number,
  companyName: (id: string) => string,
  personName: (id: string) => string,
): void => {
  const headers = ['Quote #', 'RFQ #', 'Project', 'Company', 'Contact', 'Type', 'Date Sent', 'Est Value', 'Hours Est', 'Status', 'Lost Reason'];
  const lines = [
    headers.join(','),
    ...rows.map(item => [
      csvCell(item.quoteNum),
      csvCell(item.rfqNum),
      csvCell(item.projectTitle),
      csvCell(companyName(item.organizationId)),
      csvCell(personName(item.personId)),
      csvCell(item.discipline),
      csvCell(item.quotedDate),
      csvCell(item.projectValue || ''),
      csvCell(item.approximateHours || ''),
      csvCell(item.status),
      csvCell(item.lostReason || ''),
    ].join(',')),
  ];
  const blob = new Blob([lines.join('\n')], { type: 'text/csv;charset=utf-8;' });
  const a = document.createElement('a');
  a.href = URL.createObjectURL(blob);
  a.download = `CRM-Quotes-${yr}.csv`;
  a.click();
  URL.revokeObjectURL(a.href);
};

const loadQuotesLocalFromTab = (): CrmQuote[] => loadQuotesLocal();

const normalizeQuote = (q: CrmQuote): CrmQuote =>
  normalizeCrmQuote(q as CrmQuote & Record<string, unknown>);

const LOST_ARCHIVE_MS = 7 * 24 * 60 * 60 * 1000;

const todayIso = (): string => {
  const t = new Date();
  return `${t.getFullYear()}-${String(t.getMonth() + 1).padStart(2, '0')}-${String(t.getDate()).padStart(2, '0')}`;
};

const shouldAutoArchive = (q: CrmQuote): boolean => {
  if (q.archived || q.status !== 'Lost' || !q.lostAt) return false;
  const lostMs = new Date(q.lostAt + 'T00:00:00').getTime();
  return Date.now() - lostMs >= LOST_ARCHIVE_MS;
};

/** lostAt was inferred from sent/received date — not when the quote was marked lost */
const isLegacyLostAt = (q: CrmQuote): boolean =>
  !!q.lostAt && (q.lostAt === q.quotedDate || q.lostAt === q.dateReceived);

const processQuotes = (raw: CrmQuote[]): CrmQuote[] =>
  raw.map(normalizeQuote).map(q => {
    let next = q;
    if (q.status === 'Lost') {
      if (!q.lostAt || isLegacyLostAt(q)) {
        next = { ...q, lostAt: todayIso(), archived: false, archivedAt: undefined };
      } else if (q.archived && !shouldAutoArchive(q)) {
        next = { ...q, archived: false, archivedAt: undefined };
      }
    }
    if (shouldAutoArchive(next)) {
      next = { ...next, archived: true, archivedAt: todayIso() };
    }
    return next;
  });

const tdBase: React.CSSProperties = {
  padding: '10px',
  borderBottom: `1px solid ${C.border}`,
  verticalAlign: 'middle',
};

const disciplineBadgeStyle = (d: 'Steel' | 'Concrete'): React.CSSProperties => ({
  display: 'inline-block', padding: '1px 6px', borderRadius: 3, fontSize: 9, fontWeight: 700,
  letterSpacing: '.05em', textTransform: 'uppercase', fontFamily: FF,
  background: d === 'Concrete' ? 'rgba(107,79,200,0.12)' : 'rgba(37,99,235,0.12)',
  color: d === 'Concrete' ? '#6b4fc8' : '#2563eb',
  border: `1px solid ${d === 'Concrete' ? '#6b4fc8' : '#2563eb'}`,
});

const DisciplineBadges: React.FC<{ discipline: CrmRfqDiscipline }> = ({ discipline }) => {
  if (discipline === 'Both') {
    return (
      <span style={{ display: 'flex', gap: 3, flexWrap: 'wrap' }}>
        <span style={disciplineBadgeStyle('Steel')}>STEEL</span>
        <span style={disciplineBadgeStyle('Concrete')}>CONCRETE</span>
      </span>
    );
  }
  const kind = discipline === 'Concrete' ? 'Concrete' : 'Steel';
  return <span style={disciplineBadgeStyle(kind)}>{kind.toUpperCase()}</span>;
};

const statusStyle = (s: CrmQuoteStatus): React.CSSProperties => {
  const map: Record<CrmQuoteStatus, { bg: string; color: string }> = {
    Draft:      { bg: '#eef0f3', color: '#5a6578' },
    Sent:       { bg: '#dbeafe', color: '#1d4ed8' },
    Pending:    { bg: '#ffedd5', color: '#c2410c' },
    'Follow up': { bg: '#fef9c3', color: '#a16207' },
    Lost:       { bg: '#fee2e2', color: '#b91c1c' },
  };
  const x = map[s] || map.Draft;
  return { background: x.bg, color: x.color };
};

const badge: React.CSSProperties = {
  fontSize: 10, fontWeight: 700, fontFamily: FF, padding: '3px 8px',
  borderRadius: 3, whiteSpace: 'nowrap', display: 'inline-block',
};

const actionBtn: React.CSSProperties = {
  padding: '4px 8px', borderRadius: 3, fontFamily: FF, fontWeight: 700,
  fontSize: 10, cursor: 'pointer', width: '100%', boxSizing: 'border-box', textAlign: 'center',
  whiteSpace: 'nowrap', lineHeight: 1.2,
};

const DrawingReceivedField: React.FC<{
  label: string;
  checked: boolean;
  date: string;
  onCheck: (v: boolean) => void;
  onDate: (v: string) => void;
}> = ({ label, checked, date, onCheck, onDate }) => (
  <div>
    <label style={{ display: 'flex', alignItems: 'center', gap: 8, fontFamily: FF, fontSize: 12, color: C.text, cursor: 'pointer', marginBottom: checked ? 6 : 0 }}>
      <input
        type="checkbox"
        checked={checked}
        onChange={e => {
          const v = e.target.checked;
          onCheck(v);
          if (!v) onDate('');
        }}
        style={{ width: 16, height: 16, accentColor: C.green, flexShrink: 0 }}
      />
      <span style={{ ...ml, marginBottom: 0, textTransform: 'none', letterSpacing: 0, fontSize: 12 }}>{label}</span>
    </label>
    {checked && (
      <input type="date" value={date} onChange={e => onDate(e.target.value)} style={mi} />
    )}
  </div>
);

const LostReasonModal: React.FC<{
  quote: CrmQuote;
  onConfirm: (reason: CrmLostReason) => void;
  onClose: () => void;
}> = ({ quote, onConfirm, onClose }) => {
  const [reason, setReason] = React.useState<CrmLostReason | ''>('');
  return (
    <div style={{ position: 'fixed', inset: 0, background: 'rgba(0,0,0,.45)', zIndex: 1001, display: 'flex', alignItems: 'center', justifyContent: 'center' }}>
      <div style={{ background: C.surface, borderRadius: 8, width: 440, boxShadow: '0 12px 40px rgba(0,0,0,.18)', border: `1px solid ${C.border}` }}>
        <div style={{ padding: '14px 20px', borderBottom: `1px solid ${C.border}`, background: C.thBg }}>
          <span style={{ fontFamily: FF, fontWeight: 700, fontSize: 14, color: C.text }}>Mark as Lost</span>
        </div>
        <div style={{ padding: '20px 22px' }}>
          <p style={{ fontFamily: FF, fontSize: 12, color: C.sub, margin: '0 0 14px 0', lineHeight: 1.5 }}>
            Mark <strong style={{ color: '#0d9488' }}>{quote.rfqNum}</strong>
            {quote.projectTitle ? ` — ${quote.projectTitle}` : ''} as lost.
          </p>
          <label style={ml}>Reason lost</label>
          <select
            value={reason}
            onChange={e => setReason(e.target.value as CrmLostReason)}
            style={mi}
          >
            <option value="">— Select reason —</option>
            {LOST_REASONS.map(r => <option key={r} value={r}>{r}</option>)}
          </select>
        </div>
        <div style={{ display: 'flex', gap: 10, justifyContent: 'flex-end', padding: '14px 22px', borderTop: `1px solid ${C.border}` }}>
          <button onClick={onClose} style={{ padding: '8px 20px', borderRadius: 4, border: `1px solid ${C.borderMd}`, background: 'transparent', color: C.sub, fontFamily: FF, fontSize: 12, cursor: 'pointer' }}>Cancel</button>
          <button
            onClick={() => { if (reason) onConfirm(reason); }}
            disabled={!reason}
            style={{ padding: '8px 20px', borderRadius: 4, border: 'none', background: reason ? C.red : C.borderMd, color: '#fff', fontFamily: FF, fontWeight: 700, fontSize: 12, cursor: reason ? 'pointer' : 'default' }}
          >
            Mark Lost
          </button>
        </div>
      </div>
    </div>
  );
};

const LinkProjectModal: React.FC<{
  quote: CrmQuote;
  spService: SharePointService;
  onConfirm: (project: IProject) => void;
  onClose: () => void;
}> = ({ quote, spService, onConfirm, onClose }) => {
  const [projects, setProjects] = React.useState<IProject[]>([]);
  const [linkedProjNums, setLinkedProjNums] = React.useState<Set<string>>(new Set());
  const [loading, setLoading] = React.useState(true);
  const [search, setSearch] = React.useState('');
  const [selected, setSelected] = React.useState('');

  React.useEffect(() => {
    let cancelled = false;
    void (async () => {
      try {
        const [dash, crm] = await Promise.all([
          spService.loadProjects(),
          loadProjectsFromSharePoint(spService),
        ]);
        if (cancelled) return;
        const linked = new Set(crm.map(p => p.projNum));
        setLinkedProjNums(linked);
        setProjects(dash.filter(p => !p.isEwo).sort((a, b) => a.projNum.localeCompare(b.projNum, undefined, { numeric: true })));
      } finally {
        if (!cancelled) setLoading(false);
      }
    })();
    return () => { cancelled = true; };
  }, [spService]);

  const q = search.toLowerCase().trim();
  const options = projects.filter(p => {
    if (!q) return true;
    return (
      p.projNum.toLowerCase().includes(q) ||
      p.name.toLowerCase().includes(q) ||
      (p.company || '').toLowerCase().includes(q) ||
      (p.quoteNum || '').toLowerCase().includes(q)
    );
  });

  const selectedProject = projects.find(p => p.projNum === selected);

  return (
    <div style={{ position: 'fixed', inset: 0, background: 'rgba(0,0,0,.45)', zIndex: 1001, display: 'flex', alignItems: 'center', justifyContent: 'center' }}>
      <div style={{ background: C.surface, borderRadius: 8, width: 480, maxHeight: '85vh', display: 'flex', flexDirection: 'column', boxShadow: '0 12px 40px rgba(0,0,0,.18)', border: `1px solid ${C.border}` }}>
        <div style={{ padding: '14px 20px', borderBottom: `1px solid ${C.border}`, background: C.thBg }}>
          <span style={{ fontFamily: FF, fontWeight: 700, fontSize: 14, color: C.text }}>Link to existing 3E project</span>
        </div>
        <div style={{ padding: '20px 22px', overflowY: 'auto' }}>
          <p style={{ fontFamily: FF, fontSize: 12, color: C.sub, margin: '0 0 14px 0', lineHeight: 1.5 }}>
            Link <strong style={{ color: C.purple }}>{quote.quoteNum || quote.rfqNum}</strong>
            {quote.projectTitle ? ` — ${quote.projectTitle}` : ''} to an existing project number. No new 3E number will be created.
          </p>
          {loading ? (
            <p style={{ fontFamily: FF, fontSize: 12, color: C.muted }}>Loading projects…</p>
          ) : (
            <>
              <label style={ml}>Search project</label>
              <input
                value={search}
                onChange={e => setSearch(e.target.value)}
                placeholder="3E-500, name, company…"
                style={{ ...mi, marginBottom: 12 }}
              />
              <label style={ml}>3E project number</label>
              <select value={selected} onChange={e => setSelected(e.target.value)} style={mi}>
                <option value="">— Select project —</option>
                {options.map(p => (
                  <option key={p.projNum} value={p.projNum}>
                    {p.projNum} — {p.name}{p.company ? ` (${p.company})` : ''}{p.quoteNum ? ` · ${p.quoteNum}` : ''}{linkedProjNums.has(p.projNum) ? ' · already in CRM' : ''}
                  </option>
                ))}
              </select>
              {options.length === 0 && !loading && (
                <p style={{ fontFamily: FF, fontSize: 11, color: C.muted, marginTop: 8 }}>
                  No projects match your search.
                </p>
              )}
            </>
          )}
        </div>
        <div style={{ display: 'flex', gap: 10, justifyContent: 'flex-end', padding: '14px 22px', borderTop: `1px solid ${C.border}` }}>
          <button onClick={onClose} style={{ padding: '8px 20px', borderRadius: 4, border: `1px solid ${C.borderMd}`, background: 'transparent', color: C.sub, fontFamily: FF, fontSize: 12, cursor: 'pointer' }}>Cancel</button>
          <button
            onClick={() => { if (selectedProject) onConfirm(selectedProject); }}
            disabled={!selectedProject}
            style={{ padding: '8px 20px', borderRadius: 4, border: 'none', background: selectedProject ? '#0d9488' : C.borderMd, color: '#fff', fontFamily: FF, fontWeight: 700, fontSize: 12, cursor: selectedProject ? 'pointer' : 'default' }}
          >
            Link project
          </button>
        </div>
      </div>
    </div>
  );
};

const QuoteModal: React.FC<{
  initial: CrmQuote;
  spService: SharePointService;
  persons: CrmPerson[];
  companies: CrmCompany[];
  onSave: (q: CrmQuote, attachments: CrmAttachment[], prevAttachments: CrmAttachment[]) => void | Promise<void>;
  onClose: () => void;
}> = ({ initial, spService, persons, companies, onSave, onClose }) => {
  const [d, setD] = React.useState<CrmQuote>(() => ({ ...initial, quoteNum: ensureQuPrefix(initial.quoteNum) }));
  const [attachments, setAttachments] = React.useState<CrmAttachment[]>(() => getQuoteAttachments(initial.id));
  const [awaitingLostReason, setAwaitingLostReason] = React.useState<CrmQuote | null>(null);
  const initialAttachmentsRef = React.useRef<CrmAttachment[]>(getQuoteAttachments(initial.id));
  const set = <K extends keyof CrmQuote>(k: K, v: CrmQuote[K]): void => setD(p => ({ ...p, [k]: v }));
  const grid2: React.CSSProperties = { display: 'grid', gridTemplateColumns: '1fr 1fr', gap: '14px 18px' };

  React.useEffect(() => {
    let cancelled = false;
    void (async () => {
      const remote = await loadQuoteAttachmentsFromSharePoint(spService, initial);
      if (cancelled) return;
      if (remote !== null) {
        setAttachments(remote);
        setQuoteAttachments(initial.id, remote);
        initialAttachmentsRef.current = remote;
      }
    })();
    return () => { cancelled = true; };
  }, [spService, initial.id, initial.spId]);

  const sortedCompanies = React.useMemo(
    () => [...companies].sort((a, b) => a.name.localeCompare(b.name, 'en', { sensitivity: 'base' })),
    [companies],
  );

  const personsForCo = d.organizationId
    ? persons.filter(p => p.organizationId === d.organizationId)
    : persons;

  const onPersonChange = (personId: string): void => {
    const person = persons.find(p => p.id === personId);
    const co = person?.organizationId ? companies.find(c => c.id === person.organizationId) : undefined;
    setD(p => ({
      ...p,
      personId,
      organizationId: person?.organizationId || p.organizationId,
      companyAddress: co?.address || p.companyAddress,
    }));
  };

  const onCompanyChange = (organizationId: string): void => {
    const personStillValid = persons.some(p => p.id === d.personId && p.organizationId === organizationId);
    const co = companies.find(c => c.id === organizationId);
    setD(p => ({
      ...p,
      organizationId,
      personId: personStillValid ? p.personId : '',
      companyAddress: co?.address || '',
    }));
  };

  const quoteNumLocked = quoteNumFilled(initial.quoteNum);
  const canSave = quoteNumFilled(d.quoteNum) && (d.status !== 'Sent' || !!d.quotedDate);

  return (
    <div style={{ position: 'fixed', inset: 0, background: 'rgba(0,0,0,.45)', zIndex: 1000, display: 'flex', alignItems: 'center', justifyContent: 'center' }}>
      <div style={{ background: C.surface, borderRadius: 8, width: 680, maxHeight: '92vh', overflowY: 'auto', boxShadow: '0 12px 40px rgba(0,0,0,.18)', border: `1px solid ${C.border}` }}>
        <div style={{ padding: '14px 20px', borderBottom: `1px solid ${C.border}`, display: 'flex', justifyContent: 'space-between', alignItems: 'center', background: C.thBg }}>
          <span style={{ fontFamily: FF, fontWeight: 700, fontSize: 14, color: C.text }}>Edit Quote — {d.rfqNum}</span>
          <button onClick={onClose} style={{ background: 'none', border: 'none', color: C.muted, cursor: 'pointer', fontSize: 20 }}>×</button>
        </div>
        <div style={{ padding: '20px 22px' }}>
          <div style={{ marginBottom: 14 }}>
            <label style={ml}>Company</label>
            <select value={d.organizationId} onChange={e => onCompanyChange(e.target.value)} style={mi}>
              <option value="">— Select company —</option>
              {sortedCompanies.map(c => <option key={c.id} value={c.id}>{c.name}</option>)}
            </select>
          </div>
          <div style={{ marginBottom: 14 }}>
            <label style={ml}>Company Address</label>
            <input
              value={d.companyAddress}
              onChange={e => set('companyAddress', e.target.value)}
              style={mi}
              placeholder="Company address"
            />
          </div>
          <div style={{ marginBottom: 14 }}>
            <label style={ml}>Contact Person</label>
            <select value={d.personId} onChange={e => onPersonChange(e.target.value)} style={mi}>
              <option value="">— Select person —</option>
              {personsForCo.map(p => <option key={p.id} value={p.id}>{p.name}</option>)}
            </select>
          </div>
          <div style={{ marginBottom: 14 }}>
            <label style={ml}>Project Title</label>
            <input value={d.projectTitle} onChange={e => set('projectTitle', e.target.value)} style={mi} />
          </div>
          <div style={{ marginBottom: 14 }}>
            <label style={ml}>Project Address</label>
            <input value={d.projectAddress} onChange={e => set('projectAddress', e.target.value)} style={mi} />
          </div>
          <div style={grid2}>
            <div>
              <label style={ml}>Discipline</label>
              <select value={d.discipline} onChange={e => set('discipline', e.target.value as CrmRfqDiscipline)} style={mi}>
                {DISCIPLINES.map(x => <option key={x} value={x}>{x === 'Both' ? 'Steel & Concrete' : x}</option>)}
              </select>
            </div>
            <div>
              <label style={ml}>Source</label>
              <select value={d.source} onChange={e => set('source', e.target.value)} style={mi}>
                {SOURCES.map(s => <option key={s} value={s}>{s}</option>)}
              </select>
            </div>
            <div>
              <label style={ml}>Date Received</label>
              <input type="date" value={d.dateReceived} onChange={e => set('dateReceived', e.target.value)} style={mi} />
            </div>
            <div>
              <label style={ml}>RFQ Required By</label>
              <input type="date" value={d.quoteRequiredBy} onChange={e => set('quoteRequiredBy', e.target.value)} style={mi} />
            </div>
            <DrawingReceivedField
              label="Engineer drawing received"
              checked={d.engineerDrawingReceived}
              date={d.engineerDrawingDate}
              onCheck={v => set('engineerDrawingReceived', v)}
              onDate={v => set('engineerDrawingDate', v)}
            />
            <DrawingReceivedField
              label="Architect drawing received"
              checked={d.architectDrawingReceived}
              date={d.architectDrawingDate}
              onCheck={v => set('architectDrawingReceived', v)}
              onDate={v => set('architectDrawingDate', v)}
            />
            <div>
              <label style={ml}>Revision version Eng</label>
              <input value={d.revisionVersionEng} onChange={e => set('revisionVersionEng', e.target.value)} style={mi} placeholder="e.g. Rev A" />
            </div>
            <div>
              <label style={ml}>Revision version Arch</label>
              <input value={d.revisionVersionArch} onChange={e => set('revisionVersionArch', e.target.value)} style={mi} placeholder="e.g. Rev A" />
            </div>
            <div>
              <label style={ml}>Project Value ($)</label>
              <input type="number" min={0} step={100} value={d.projectValue || ''} onChange={e => set('projectValue', Number(e.target.value) || 0)} style={mi} />
            </div>
            <div>
              <label style={ml}>Assigned To</label>
              <select value={d.assignedTo} onChange={e => set('assignedTo', e.target.value)} style={mi}>
                {OWNERS.map(o => <option key={o} value={o}>{o}</option>)}
              </select>
            </div>
            <div>
              <label style={ml}>RFI allowed</label>
              <input type="number" min={0} step={1} value={d.rfiAllowed || ''} onChange={e => set('rfiAllowed', Number(e.target.value) || 0)} style={mi} />
            </div>
            <div>
              <label style={ml}>Est Hours</label>
              <input type="number" min={0} step={1} value={d.approximateHours || ''} onChange={e => set('approximateHours', Number(e.target.value) || 0)} style={mi} />
            </div>
            <div>
              <label style={ml}>Status</label>
              <select
                value={d.status}
                onChange={e => {
                  const status = e.target.value as CrmQuoteStatus;
                  setD(p => ({ ...p, status }));
                }}
                style={mi}
              >
                {QUOTE_STATUSES.map(s => <option key={s} value={s}>{s}</option>)}
              </select>
            </div>
            <div>
              {(d.status === 'Sent' || d.status === 'Lost' || d.status === 'Follow up' || d.status === 'Pending') && (
                <>
                  <label style={ml}>Date quote sent</label>
                  <input
                    type="date"
                    value={d.quotedDate}
                    onChange={e => set('quotedDate', e.target.value)}
                    style={mi}
                  />
                </>
              )}
            </div>
            <div>
              {initial.status === 'Lost' && (
                <>
                  <label style={ml}>Lost Reason</label>
                  <select value={d.lostReason || ''} onChange={e => set('lostReason', e.target.value)} style={mi}>
                    <option value="">— Select reason —</option>
                    {LOST_REASONS.map(r => <option key={r} value={r}>{r}</option>)}
                  </select>
                </>
              )}
            </div>
          </div>
          <div style={{ marginTop: 14, marginBottom: 14 }}>
            <label style={{ display: 'flex', alignItems: 'center', gap: 8, fontFamily: FF, fontSize: 12, color: C.text, cursor: 'pointer' }}>
              <input
                type="checkbox"
                checked={d.createQuoteXero}
                onChange={e => set('createQuoteXero', e.target.checked)}
                style={{ width: 16, height: 16, accentColor: C.green }}
              />
              Create quote in Xero
            </label>
          </div>
          <div style={{ marginBottom: 14 }}>
            <label style={ml}>Quote # (Xero) *</label>
            <div style={{ display: 'flex', border: `1px solid ${C.borderMd}`, borderRadius: 4, overflow: 'hidden', background: quoteNumLocked ? C.thBg : C.surface }}>
              <span style={{ padding: '8px 10px', fontFamily: FF, fontWeight: 700, fontSize: 12.5, color: C.purple, background: C.thBg, borderRight: `1px solid ${C.borderMd}`, whiteSpace: 'nowrap' }}>QU-</span>
              <input
                value={d.quoteNum.startsWith('QU-') ? d.quoteNum.slice(3) : d.quoteNum}
                onChange={e => set('quoteNum', 'QU-' + e.target.value.replace(/^QU-/i, ''))}
                readOnly={quoteNumLocked}
                style={{ ...mi, border: 'none', borderRadius: 0, background: quoteNumLocked ? C.thBg : C.surface, cursor: quoteNumLocked ? 'default' : 'text' }}
                placeholder="0490"
              />
            </div>
          </div>
          <div>
            <label style={ml}>Notes</label>
            <textarea value={d.notes} onChange={e => set('notes', e.target.value)} rows={4} style={{ ...mi, resize: 'vertical', minHeight: 80 }} />
          </div>
          <DocumentUploadSection attachments={attachments} onChange={setAttachments} />
        </div>
        <div style={{ display: 'flex', gap: 10, justifyContent: 'flex-end', padding: '14px 22px', borderTop: `1px solid ${C.border}` }}>
          <button onClick={onClose} style={{ padding: '8px 20px', borderRadius: 4, border: `1px solid ${C.borderMd}`, background: 'transparent', color: C.sub, fontFamily: FF, fontSize: 12, cursor: 'pointer' }}>Cancel</button>
          <button
            onClick={() => {
              if (canSave) {
                const prepared: CrmQuote = { ...d, quoteNum: quoteNumFilled(d.quoteNum) ? d.quoteNum.trim() : '' };
                if (prepared.status === 'Lost' && !prepared.lostReason) {
                  setAwaitingLostReason(prepared);
                  return;
                }
                setQuoteAttachments(d.id, attachments);
                void onSave(prepared, attachments, initialAttachmentsRef.current);
              }
            }}
            style={{ padding: '8px 20px', borderRadius: 4, border: 'none', background: canSave ? C.green : C.borderMd, color: '#fff', fontFamily: FF, fontWeight: 700, fontSize: 12, cursor: canSave ? 'pointer' : 'default' }}
          >
            Save
          </button>
        </div>
      </div>
      {awaitingLostReason && (
        <LostReasonModal
          quote={awaitingLostReason}
          onConfirm={reason => {
            setAwaitingLostReason(null);
            setQuoteAttachments(awaitingLostReason.id, attachments);
            void onSave({ ...awaitingLostReason, lostReason: reason }, attachments, initialAttachmentsRef.current);
          }}
          onClose={() => setAwaitingLostReason(null)}
        />
      )}
    </div>
  );
};

const KpiCard: React.FC<{
  label: string; value: string; sub: string; accent: string; onClick?: () => void;
}> = ({ label, value, sub, accent, onClick }) => (
  <div
    onClick={onClick}
    style={{
      flex: '1 1 120px', background: C.surface, border: `1px solid ${C.border}`, borderRadius: 6,
      overflow: 'hidden', minWidth: 100, cursor: onClick ? 'pointer' : 'default',
    }}
  >
    <div style={{ height: 3, background: accent }} />
    <div style={{ padding: '12px 14px' }}>
      <div style={{ fontFamily: FF, fontSize: 9, fontWeight: 700, letterSpacing: '.08em', color: C.muted, textTransform: 'uppercase' }}>{label}</div>
      <div style={{ fontFamily: FF, fontSize: 22, fontWeight: 700, color: C.text, marginTop: 4 }}>{value}</div>
      <div style={{ fontFamily: FF, fontSize: 11, color: C.muted, marginTop: 2 }}>{sub}</div>
    </div>
  </div>
);

const BudgetKpiCard: React.FC<{
  label: string;
  actual: number;
  target: number;
  pct: number | null;
  accent: string;
  onClick: () => void;
}> = ({ label, actual, target, pct, accent, onClick }) => (
  <div
    onClick={onClick}
    style={{
      flex: '1 1 130px', background: C.surface, border: `1px solid ${C.border}`, borderRadius: 6,
      overflow: 'hidden', minWidth: 110, cursor: 'pointer',
    }}
  >
    <div style={{ height: 3, background: accent }} />
    <div style={{ padding: '12px 14px' }}>
      <div style={{ fontFamily: FF, fontSize: 9, fontWeight: 700, letterSpacing: '.08em', color: C.muted, textTransform: 'uppercase' }}>{label}</div>
      <div style={{ fontFamily: FF, fontSize: 22, fontWeight: 700, color: C.text, marginTop: 4, lineHeight: 1.2 }}>
        {fmtMoney(actual)}
      </div>
      <div style={{ fontFamily: FF, fontSize: 11, color: C.sub, marginTop: 4 }}>
        of {target > 0 ? fmtBudget(target) : '—'} budget
      </div>
      <div style={{ fontFamily: FF, fontSize: 10, color: C.muted, marginTop: 2 }}>
        {pct !== null ? `${pct}% achieved` : 'click to set budget'}
      </div>
    </div>
  </div>
);

const BudgetEditModal: React.FC<{
  budgetStore: CrmQuoteBudgetStore;
  editYear: number;
  yearOptions: number[];
  onSave: (store: CrmQuoteBudgetStore) => void;
  onClose: () => void;
}> = ({ budgetStore, editYear, yearOptions, onSave, onClose }) => {
  const [year, setYear] = React.useState(editYear);
  const [b, setB] = React.useState(() => getQuoteBudgetForYear(budgetStore, editYear));

  React.useEffect(() => {
    setYear(editYear);
    setB(getQuoteBudgetForYear(budgetStore, editYear));
  }, [budgetStore, editYear]);

  const onYearChange = (y: number): void => {
    setYear(y);
    setB(getQuoteBudgetForYear(budgetStore, y));
  };

  const setMonthTarget = (idx: number, value: number): void => {
    setB(p => {
      const months = [...(p.monthTargets || Array(12).fill(0))];
      months[idx] = value;
      return { ...p, year, monthTargets: months, monthTarget: months[new Date().getMonth()] || 0 };
    });
  };

  const years = React.useMemo(() => {
    const set = new Set(yearOptions);
    set.add(year);
    set.add(new Date().getFullYear());
    return Array.from(set).sort((a, b) => b - a);
  }, [yearOptions, year]);

  return (
    <div style={{ position: 'fixed', inset: 0, background: 'rgba(0,0,0,.45)', zIndex: 1001, display: 'flex', alignItems: 'center', justifyContent: 'center' }}>
      <div style={{ background: C.surface, borderRadius: 8, width: 520, maxHeight: '90vh', overflowY: 'auto', boxShadow: '0 12px 40px rgba(0,0,0,.18)', border: `1px solid ${C.border}` }}>
        <div style={{ padding: '14px 20px', borderBottom: `1px solid ${C.border}`, background: C.thBg }}>
          <span style={{ fontFamily: FF, fontWeight: 700, fontSize: 14, color: C.text }}>Quote Budget Targets</span>
        </div>
        <div style={{ padding: '20px 22px' }}>
          <div style={{ marginBottom: 18 }}>
            <label style={ml}>Budget year</label>
            <select value={year} onChange={e => onYearChange(Number(e.target.value))} style={mi}>
              {years.map(y => <option key={y} value={y}>{y}</option>)}
            </select>
          </div>
          <div style={{ marginBottom: 18 }}>
            <label style={ml}>Year budget ($)</label>
            <input type="number" min={0} step={1000} value={b.yearTarget || ''} onChange={e => setB(p => ({ ...p, year, yearTarget: Number(e.target.value) || 0 }))} style={mi} />
          </div>
          <label style={{ ...ml, marginBottom: 10 }}>Monthly budgets ($)</label>
          <div style={{ display: 'grid', gridTemplateColumns: '1fr 1fr', gap: '12px 16px' }}>
            {MONTH_LABELS.map((label, idx) => (
              <div key={label}>
                <label style={{ ...ml, fontSize: 9 }}>{label}</label>
                <input
                  type="number"
                  min={0}
                  step={500}
                  value={b.monthTargets?.[idx] || ''}
                  onChange={e => setMonthTarget(idx, Number(e.target.value) || 0)}
                  style={mi}
                />
              </div>
            ))}
          </div>
        </div>
        <div style={{ display: 'flex', gap: 10, justifyContent: 'flex-end', padding: '14px 22px', borderTop: `1px solid ${C.border}` }}>
          <button onClick={onClose} style={{ padding: '8px 20px', borderRadius: 4, border: `1px solid ${C.borderMd}`, background: 'transparent', color: C.sub, fontFamily: FF, fontSize: 12, cursor: 'pointer' }}>Cancel</button>
          <button onClick={() => onSave(setQuoteBudgetForYear(budgetStore, { ...b, year }))} style={{ padding: '8px 20px', borderRadius: 4, border: 'none', background: C.green, color: '#fff', fontFamily: FF, fontWeight: 700, fontSize: 12, cursor: 'pointer' }}>Save</button>
        </div>
      </div>
    </div>
  );
};

const CrmQuotesTab: React.FC<{
  spService: SharePointService;
  persons: CrmPerson[];
  companies: CrmCompany[];
  seedQuotes?: CrmQuote[] | null;
  onSeedApplied?: () => void;
  onQuoteWon?: (projects: CrmProject[], quotes: CrmQuote[]) => void;
}> = ({ spService, persons, companies, seedQuotes, onSeedApplied, onQuoteWon }) => {
  const [quotes, setQuotes] = React.useState<CrmQuote[]>([]);
  const [wonProjects, setWonProjects] = React.useState<CrmProject[]>([]);
  const [budgetStore, setBudgetStore] = React.useState<CrmQuoteBudgetStore>(() => normalizeQuoteBudgetStore(undefined));
  const [selectedYear, setSelectedYear] = React.useState(() => new Date().getFullYear());
  const [budgetModal, setBudgetModal] = React.useState(false);
  const [ready, setReady] = React.useState(false);
  const [search, setSearch] = React.useState('');
  const [statusFilter, setStatusFilter] = React.useState('all');
  const [dateFrom, setDateFrom] = React.useState('');
  const [dateTo, setDateTo] = React.useState('');
  const [archiveView, setArchiveView] = React.useState(false);
  const [modal, setModal] = React.useState<CrmQuote | null>(null);
  const [lostQuote, setLostQuote] = React.useState<CrmQuote | null>(null);
  const [linkQuote, setLinkQuote] = React.useState<CrmQuote | null>(null);
  const [pendingDelId, setPendingDelId] = React.useState<string | null>(null);
  const [attTick, setAttTick] = React.useState(0);
  const [quotesWithAttachments, setQuotesWithAttachments] = React.useState<Set<string>>(new Set());
  const delTimerRef = React.useRef<ReturnType<typeof setTimeout> | null>(null);
  const quotesRef = React.useRef(quotes);
  const lastLocalEditAtRef = React.useRef(0);
  const pendingSyncAttemptedRef = React.useRef(new Set<string>());
  quotesRef.current = quotes;

  const touchLocalEdit = (): void => { lastLocalEditAtRef.current = Date.now(); };

  const refreshAttachmentIndex = React.useCallback(async (): Promise<void> => {
    try {
      const idx = await loadQuoteAttachmentIndex(spService, quotesRef.current);
      setQuotesWithAttachments(idx);
    } catch { /* ignore */ }
  }, [spService]);

  const pushPendingQuotesToSharePoint = React.useCallback((rows: CrmQuote[]): void => {
    for (const q of rows) {
      if (q.spId || pendingSyncAttemptedRef.current.has(q.id)) continue;
      pendingSyncAttemptedRef.current.add(q.id);
      void upsertQuoteToSharePoint(spService, q).then(saved => {
        if (!saved.spId) return;
        setQuotes(cur => {
          const updated = cur.map(x => x.id === saved.id ? { ...x, spId: saved.spId } : x);
          quotesRef.current = updated;
          try { localStorage.setItem(LS_QUOTES, JSON.stringify(updated)); } catch { /* ignore */ }
          return updated;
        });
      }).catch(() => undefined);
    }
  }, [spService]);

  const syncQuoteChanges = React.useCallback((before: CrmQuote[], after: CrmQuote[]): void => {
    const beforeMap = new Map(before.map(q => [q.id, q]));
    for (const q of after) {
      const p = beforeMap.get(q.id);
      if (!p || JSON.stringify(p) !== JSON.stringify(q)) {
        void upsertQuoteToSharePoint(spService, q).then(withSpId => {
          if (withSpId.spId !== q.spId) {
            setQuotes(cur => {
              const updated = cur.map(x => x.id === withSpId.id ? { ...x, spId: withSpId.spId } : x);
              quotesRef.current = updated;
              try { localStorage.setItem(LS_QUOTES, JSON.stringify(updated)); } catch { /* ignore */ }
              return updated;
            });
          }
        }).catch(() => undefined);
      }
    }
  }, [spService]);

  const applyProcessedQuotes = React.useCallback((raw: CrmQuote[]): void => {
    const before = quotesRef.current;
    const processed = processQuotes(raw);
    quotesRef.current = processed;
    setQuotes(processed);
    try { localStorage.setItem(LS_QUOTES, JSON.stringify(processed)); } catch { /* ignore */ }
    if (before.length > 0 && JSON.stringify(before) !== JSON.stringify(processed)) {
      syncQuoteChanges(before, processed);
    }
    pushPendingQuotesToSharePoint(processed);
  }, [syncQuoteChanges, pushPendingQuotesToSharePoint]);

  const reload = React.useCallback(async (): Promise<void> => {
    try {
      const skipQuotePull = Date.now() - lastLocalEditAtRef.current < 15000;
      if (!skipQuotePull) {
        const remote = await loadQuotesRemote(spService);
        if (remote !== null) {
          const merged = mergeQuotesWithLocal(quotesRef.current, remote);
          const processed = processQuotes(merged);
          if (JSON.stringify(processed) !== JSON.stringify(quotesRef.current)) {
            const before = quotesRef.current;
            quotesRef.current = processed;
            setQuotes(processed);
            try { localStorage.setItem(LS_QUOTES, JSON.stringify(processed)); } catch { /* ignore */ }
            syncQuoteChanges(before, processed);
          }
        }
      }
      const projs = await loadProjectsFromSharePoint(spService);
      setWonProjects(projs);
      void refreshAttachmentIndex();
    } catch {
      if (Date.now() - lastLocalEditAtRef.current >= 15000) {
        setQuotes(processQuotes(loadQuotesLocalFromTab()));
      }
    }
  }, [spService, syncQuoteChanges, refreshAttachmentIndex]);

  React.useEffect(() => {
    let cancelled = false;
    void (async () => {
      try {
        const [data, projs, b] = await Promise.all([
          loadQuotesFromSharePoint(spService),
          loadProjectsFromSharePoint(spService),
          loadQuoteBudget(spService),
        ]);
        if (!cancelled) {
          const merged = mergeQuotesWithLocal(quotesRef.current, data);
          applyProcessedQuotes(merged);
          setWonProjects(projs);
          setBudgetStore(b);
        }
      } catch {
        if (!cancelled) {
          applyProcessedQuotes(processQuotes(quotesRef.current.length ? quotesRef.current : loadQuotesLocalFromTab()));
        }
      } finally {
        if (!cancelled) setReady(true);
      }
    })();
    return () => { cancelled = true; };
  }, [spService, applyProcessedQuotes, refreshAttachmentIndex]);

  React.useEffect(() => {
    if (!ready) return;
    void refreshAttachmentIndex();
  }, [ready, attTick, refreshAttachmentIndex]);

  React.useEffect(() => {
    if (!seedQuotes?.length) return;
    touchLocalEdit();
    applyProcessedQuotes(seedQuotes);
    setReady(true);
    onSeedApplied?.();
    void Promise.all([
      loadProjectsFromSharePoint(spService),
      loadQuoteBudget(spService),
    ]).then(([projs, b]) => {
      setWonProjects(projs);
      setBudgetStore(b);
    });
  }, [seedQuotes, spService, applyProcessedQuotes, onSeedApplied]);

  React.useEffect(() => {
    if (!ready) return;
    const iv = setInterval(() => { void reload(); }, 12000);
    const onFocus = (): void => { void reload(); };
    window.addEventListener('focus', onFocus);
    return () => {
      clearInterval(iv);
      window.removeEventListener('focus', onFocus);
    };
  }, [ready, reload]);

  const companyName = (id: string): string => companies.find(c => c.id === id)?.name || '—';
  const personName = (id: string): string => persons.find(p => p.id === id)?.name || '—';

  const quoteListYear = (q: CrmQuote): string => {
    const ref = q.quotedDate || q.dateReceived;
    return ref.length >= 4 ? ref.substring(0, 4) : String(selectedYear);
  };

  const quoteFilterDate = (q: CrmQuote): string => q.quotedDate || q.dateReceived || '';

  const yearOptions = React.useMemo(() => {
    const yrs = new Set<number>();
    yrs.add(new Date().getFullYear());
    yrs.add(selectedYear);
    quotes.forEach(q => {
      const y = Number(quoteListYear(q));
      if (!isNaN(y) && y > 2000) yrs.add(y);
    });
    wonProjects.forEach(p => {
      if (p.wonDate?.length >= 4) yrs.add(Number(p.wonDate.substring(0, 4)));
    });
    Object.keys(budgetStore.byYear).forEach(y => {
      const n = Number(y);
      if (!isNaN(n)) yrs.add(n);
    });
    return Array.from(yrs).sort((a, b) => b - a);
  }, [quotes, wonProjects, budgetStore, selectedYear]);

  const year = selectedYear;
  const calendarYear = new Date().getFullYear();
  const month = year === calendarYear ? new Date().getMonth() : 0;
  const budget = getQuoteBudgetForYear(budgetStore, year);
  const quotesForYear = quotes.filter(q => quoteListYear(q) === String(year));
  const yearQuotes = quotesForYear.filter(q => !q.archived);
  const archivedQuotes = quotesForYear.filter(q => q.archived && q.status === 'Lost');
  const wonForYear = wonProjects.filter(p => p.wonDate.startsWith(String(year)));
  const monthBudgetTarget = getMonthBudgetTarget(budget, month);

  const stats = React.useMemo(() => {
    const sent = yearQuotes.filter(q => q.status === 'Sent');
    const lost = yearQuotes.filter(q => q.status === 'Lost');
    const totalHours = yearQuotes.reduce((s, q) => s + (q.approximateHours || 0), 0);
    const quoteValue = (q: { projectValue?: number }): number => q.projectValue || 0;
    const quoteInMonth = (q: CrmQuote, mo: number): boolean => {
      const ref = q.quotedDate || q.dateReceived;
      if (!ref || !ref.startsWith(String(year))) return false;
      const d = new Date(ref + 'T00:00:00');
      return !isNaN(d.getTime()) && d.getMonth() === mo;
    };
    // All quote est values for the year (every status) + won/linked quotes no longer on the list
    const yearActual = quotesForYear.reduce((s, q) => s + quoteValue(q), 0)
      + wonForYear.reduce((s, p) => s + quoteValue(p), 0);
    const monthActual = quotesForYear.filter(q => quoteInMonth(q, month)).reduce((s, q) => s + quoteValue(q), 0)
      + wonForYear
        .filter(p => new Date(p.wonDate + 'T00:00:00').getMonth() === month)
        .reduce((s, p) => s + quoteValue(p), 0);
    const yearPct = budget.yearTarget > 0 ? Math.round((yearActual / budget.yearTarget) * 100) : null;
    const monthPct = monthBudgetTarget > 0 ? Math.round((monthActual / monthBudgetTarget) * 100) : null;
    const won = wonForYear.length;
    const decided = won + lost.length;
    const winRate = decided > 0 ? Math.round((won / decided) * 100) : 0;
    return {
      total: yearQuotes.length,
      sent: sent.length,
      lost: lost.length,
      won,
      winRate,
      totalHours,
      yearActual,
      monthActual,
      yearPct,
      monthPct,
    };
  }, [yearQuotes, quotesForYear, wonForYear, budget, month, monthBudgetTarget, year]);

  const q = search.toLowerCase();
  const listSource = archiveView ? archivedQuotes : yearQuotes;
  const filtered = listSource.filter(item => {
    if (!archiveView && statusFilter !== 'all' && item.status !== statusFilter) return false;
    if (dateFrom || dateTo) {
      const ref = quoteFilterDate(item);
      if (!ref) return false;
      if (dateFrom && ref < dateFrom) return false;
      if (dateTo && ref > dateTo) return false;
    }
    if (!q) return true;
    return (
      item.quoteNum.toLowerCase().includes(q) ||
      item.rfqNum.toLowerCase().includes(q) ||
      item.projectTitle.toLowerCase().includes(q) ||
      companyName(item.organizationId).toLowerCase().includes(q)
    );
  });

  const persistQuotes = (next: CrmQuote[]): void => {
    touchLocalEdit();
    const processed = processQuotes(next);
    const prev = quotesRef.current;
    quotesRef.current = processed;
    setQuotes(processed);
    try { localStorage.setItem(LS_QUOTES, JSON.stringify(processed)); } catch { /* ignore */ }

    const prevMap = new Map(prev.map(q => [q.id, q]));
    const nextMap = new Map(processed.map(q => [q.id, q]));

    for (const q of processed) {
      const p = prevMap.get(q.id);
      if (!p || JSON.stringify(p) !== JSON.stringify(q)) {
        void upsertQuoteToSharePoint(spService, q).then(withSpId => {
          if (withSpId.spId !== q.spId) {
            setQuotes(cur => {
              const updated = cur.map(x => x.id === withSpId.id ? { ...x, spId: withSpId.spId } : x);
              quotesRef.current = updated;
              try { localStorage.setItem(LS_QUOTES, JSON.stringify(updated)); } catch { /* ignore */ }
              return updated;
            });
          }
        }).catch(() => undefined);
      }
    }

    for (const q of prev) {
      if (!nextMap.has(q.id)) void deleteQuoteFromSharePoint(spService, q).catch(() => undefined);
    }
  };

  const saveQuote = async (
    item: CrmQuote,
    attachments: CrmAttachment[],
    prevAttachments: CrmAttachment[],
  ): Promise<void> => {
    let saved = normalizeQuote(item);
    const prev = quotesRef.current.find(x => x.id === saved.id);
    if (saved.status === 'Lost') {
      if (prev?.status !== 'Lost' || !saved.lostAt || isLegacyLostAt(saved)) {
        saved = { ...saved, lostAt: todayIso(), archived: false, archivedAt: undefined };
      }
    }
    persistQuotes(quotesRef.current.map(x => x.id === saved.id ? saved : x));
    const withSp = await upsertQuoteToSharePoint(spService, { ...saved, spId: saved.spId ?? prev?.spId });
    if (withSp.spId && withSp.spId !== saved.spId) {
      setQuotes(cur => {
        const updated = cur.map(x => x.id === withSp.id ? { ...x, spId: withSp.spId } : x);
        quotesRef.current = updated;
        try { localStorage.setItem(LS_QUOTES, JSON.stringify(updated)); } catch { /* ignore */ }
        return updated;
      });
    }
    if (withSp.spId) {
      await syncQuoteAttachmentsToSharePoint(spService, withSp, attachments, prevAttachments);
    } else {
      setQuoteAttachments(saved.id, attachments);
    }
    await refreshAttachmentIndex();
    setAttTick(t => t + 1);
    setModal(null);
  };

  const markLost = (item: CrmQuote, reason: CrmLostReason): void => {
    persistQuotes(quotesRef.current.map(x => x.id === item.id
      ? {
        ...x,
        status: 'Lost' as CrmQuoteStatus,
        lostReason: reason,
        lostAt: todayIso(),
        archived: false,
        archivedAt: undefined,
      }
      : x));
    setLostQuote(null);
  };

  const markWon = (item: CrmQuote): void => {
    if (!confirm(`Mark ${item.rfqNum} as WON and create a NEW project number for ${item.projectTitle || ''}?`)) return;
    void (async () => {
      try {
        const [crmProjects, dashProjects] = await Promise.all([
          loadProjectsFromSharePoint(spService),
          spService.loadProjects().catch(() => []),
        ]);
        const projNum = nextProjNum(crmProjects, dashProjects);
        const iProject = quoteToIProject(item, projNum, persons, companies);
        let spId: number | undefined;
        try {
          spId = await spService.addProject(iProject);
        } catch { /* CRM project still created if SP list unavailable */ }
        const crmProject = quoteToCrmProject(item, projNum, spId);
        const nextProjects = [...crmProjects, crmProject];
        const nextQuotes = quotesRef.current.filter(q => q.id !== item.id);
        persistQuotes(nextQuotes);
        try { localStorage.setItem(LS_PROJECTS, JSON.stringify(nextProjects)); } catch { /* ignore */ }
        void upsertWipToSharePoint(spService, crmProject).catch(() => undefined);
        setWonProjects(nextProjects);
        onQuoteWon?.(nextProjects, nextQuotes);
      } catch {
        alert('Could not create project. Please try again.');
      }
    })();
  };

  const linkToProject = (item: CrmQuote, project: IProject): void => {
    void (async () => {
      try {
        const crmProjects = await loadProjectsFromSharePoint(spService);
        if (crmProjects.some(p => p.projNum === project.projNum)) {
          alert(`${project.projNum} is already linked in CRM.`);
          return;
        }
        const wonDate = project.startDate && /^\d{4}-\d{2}-\d{2}/.test(project.startDate)
          ? project.startDate.substring(0, 10)
          : todayIso();
        const crmProject: CrmProject = {
          ...quoteToCrmProject(item, project.projNum, project.spId),
          wonDate,
        };
        const nextProjects = [...crmProjects, crmProject];
        const nextQuotes = quotesRef.current.filter(q => q.id !== item.id);
        persistQuotes(nextQuotes);
        try { localStorage.setItem(LS_PROJECTS, JSON.stringify(nextProjects)); } catch { /* ignore */ }
        void upsertWipToSharePoint(spService, crmProject).catch(() => undefined);
        if (project.spId && item.quoteNum && !project.quoteNum) {
          try {
            await spService.updateProject(project.spId, { ...project, quoteNum: item.quoteNum });
          } catch { /* dashboard quote # optional */ }
        }
        setWonProjects(nextProjects);
        setLinkQuote(null);
        onQuoteWon?.(nextProjects, nextQuotes);
      } catch {
        alert('Could not link project. Please try again.');
      }
    })();
  };

  const deleteQuote = (id: string): void => {
    if (pendingDelId === id) {
      if (delTimerRef.current) clearTimeout(delTimerRef.current);
      setPendingDelId(null);
      persistQuotes(quotesRef.current.filter(x => x.id !== id));
    } else {
      setPendingDelId(id);
      if (delTimerRef.current) clearTimeout(delTimerRef.current);
      delTimerRef.current = setTimeout(() => setPendingDelId(null), 3000);
    }
  };

  if (!ready) {
    return <div style={{ padding: 24, fontFamily: FF, fontSize: 13, color: C.muted }}>Loading quotes…</div>;
  }

  return (
    <>
      <div style={{ display: 'flex', gap: 12, flexWrap: 'wrap', padding: '16px 0 12px 0' }}>
        <KpiCard label="Total Quotes" accent={C.purple} value={String(stats.total)} sub={`${year} year`} />
        <KpiCard label="Sent" accent="#3b82c4" value={String(stats.sent)} sub="awaiting response" />
        <KpiCard
          label="Win Rate"
          accent={C.green}
          value={`${stats.winRate}%`}
          sub={`${stats.won} won · ${stats.lost} lost`}
        />
        <KpiCard label="Lost" accent={C.red} value={String(stats.lost)} sub="declined" />
        <BudgetKpiCard
          label="Quote Budget Year"
          actual={stats.yearActual}
          target={budget.yearTarget}
          pct={stats.yearPct}
          accent="#b36a00"
          onClick={() => setBudgetModal(true)}
        />
        <BudgetKpiCard
          label="Quote Budget Month"
          actual={stats.monthActual}
          target={monthBudgetTarget}
          pct={stats.monthPct}
          accent="#d97706"
          onClick={() => setBudgetModal(true)}
        />
        <KpiCard label="Hours Est" accent="#0d9488" value={String(stats.totalHours)} sub="total estimated hours" />
      </div>

      <div style={{ display: 'flex', alignItems: 'center', gap: 10, flexWrap: 'nowrap', paddingBottom: 12 }}>
        <select
          value={selectedYear}
          onChange={e => setSelectedYear(Number(e.target.value))}
          style={{ padding: '7px 10px', borderRadius: 4, border: `1px solid ${C.borderMd}`, background: C.surface, fontFamily: FF, fontSize: 12, color: C.text, width: 90, flexShrink: 0, boxSizing: 'border-box' }}
        >
          {yearOptions.map(y => <option key={y} value={y}>{y}</option>)}
        </select>
        <input
          value={search}
          onChange={e => setSearch(e.target.value)}
          placeholder="Search quotes…"
          style={{ padding: '7px 12px', borderRadius: 4, border: `1px solid ${C.borderMd}`, background: C.surface, fontFamily: FF, fontSize: 12, width: 200, flexShrink: 0, outline: 'none', boxSizing: 'border-box' }}
        />
        <select
          value={statusFilter}
          onChange={e => { setStatusFilter(e.target.value); setArchiveView(false); }}
          style={{ padding: '7px 10px', borderRadius: 4, border: `1px solid ${C.borderMd}`, background: C.surface, fontFamily: FF, fontSize: 12, color: C.text, width: 150, flexShrink: 0, boxSizing: 'border-box' }}
          disabled={archiveView}
        >
          <option value="all">All statuses</option>
          {QUOTE_STATUSES.map(s => <option key={s} value={s}>{s}</option>)}
        </select>
        <div style={{ display: 'flex', alignItems: 'center', gap: 6, flexShrink: 0 }}>
          <input
            type="date"
            value={dateFrom}
            onChange={e => setDateFrom(e.target.value)}
            title="Quoted from"
            style={{ padding: '6px 8px', borderRadius: 4, border: `1px solid ${C.borderMd}`, background: C.surface, fontFamily: FF, fontSize: 12, color: C.text, width: 130, boxSizing: 'border-box' }}
          />
          <span style={{ fontFamily: FF, fontSize: 12, color: C.muted }}>–</span>
          <input
            type="date"
            value={dateTo}
            onChange={e => setDateTo(e.target.value)}
            title="Quoted to"
            min={dateFrom || undefined}
            style={{ padding: '6px 8px', borderRadius: 4, border: `1px solid ${C.borderMd}`, background: C.surface, fontFamily: FF, fontSize: 12, color: C.text, width: 130, boxSizing: 'border-box' }}
          />
          {(dateFrom || dateTo) && (
            <button
              type="button"
              onClick={() => { setDateFrom(''); setDateTo(''); }}
              style={{ padding: '6px 8px', borderRadius: 4, border: `1px solid ${C.borderMd}`, background: C.surface, fontFamily: FF, fontSize: 11, color: C.sub, cursor: 'pointer', whiteSpace: 'nowrap' }}
            >
              Clear
            </button>
          )}
        </div>
        <div style={{ flex: 1 }} />
        <button
          onClick={() => setArchiveView(v => !v)}
          style={{
            padding: '7px 16px', borderRadius: 4, fontFamily: FF, fontWeight: 700, fontSize: 12,
            cursor: 'pointer', flexShrink: 0, whiteSpace: 'nowrap',
            border: archiveView ? `2px solid ${C.sub}` : `1px solid ${C.borderMd}`,
            background: archiveView ? C.thBg : C.surface,
            color: archiveView ? C.text : C.sub,
          }}
        >
          Archive{archivedQuotes.length > 0 ? ` (${archivedQuotes.length})` : ''}
        </button>
        <button
          onClick={() => exportQuotesCsv(archiveView ? archivedQuotes : yearQuotes, year, companyName, personName)}
          style={{ padding: '7px 16px', borderRadius: 4, border: `1px solid ${C.borderMd}`, background: C.surface, fontFamily: FF, fontWeight: 700, fontSize: 12, color: C.text, cursor: 'pointer', flexShrink: 0, whiteSpace: 'nowrap' }}
        >
          Export All
        </button>
      </div>

      <div style={{ background: C.surface, border: `1px solid ${C.border}`, borderRadius: '0 0 8px 8px', overflowX: 'auto' }}>
        <table style={{ width: '100%', borderCollapse: 'collapse', minWidth: 960 }}>
          <thead>
            <tr>
              {['Quote #', 'RFQ #', 'Project', 'Company', 'Contact', 'Type', 'Quoted', 'Est Value', 'Hours Est', 'Status', 'Actions'].map((h, i) => (
                <th key={h} style={{ padding: '9px 10px', textAlign: 'left', fontFamily: FF, fontWeight: 700, fontSize: 10, letterSpacing: '.06em', textTransform: 'uppercase', color: C.sub, background: C.thBg, borderBottom: `2px solid ${C.borderMd}`, whiteSpace: 'nowrap', ...(i < 2 ? { minWidth: 72 } : {}) }}>
                  {h}
                </th>
              ))}
            </tr>
          </thead>
          <tbody>
            {filtered.length === 0 ? (
              <tr>
                <td colSpan={11} style={{ padding: 48, textAlign: 'center', fontFamily: FF, fontSize: 13, color: C.muted }}>
                  {quotes.length === 0
                    ? 'No quotes yet — move an RFQ at Ready to Quote stage using the Quote button.'
                    : archiveView
                      ? 'No archived quotes yet — lost quotes are archived automatically after 1 week.'
                      : 'No results match your search.'}
                </td>
              </tr>
            ) : filtered.map(item => (
              <tr
                key={item.id}
                style={{ verticalAlign: 'middle' }}
                onMouseEnter={e => { (e.currentTarget as HTMLTableRowElement).style.background = C.rowHover; }}
                onMouseLeave={e => { (e.currentTarget as HTMLTableRowElement).style.background = 'transparent'; }}
              >
                <td style={{ ...tdBase, fontFamily: FF, fontSize: 12, fontWeight: 700, color: C.purple, whiteSpace: 'nowrap' }}>
                  <span style={{ display: 'inline-flex', alignItems: 'center', gap: 5 }}>
                    <span>{item.quoteNum || '—'}</span>
                    {quotesWithAttachments.has(item.id) && (
                      <span title="Has attachments" style={{ fontSize: 12, lineHeight: 1, color: C.sub }} aria-label="Has attachments">📎</span>
                    )}
                  </span>
                </td>
                <td style={{ ...tdBase, fontFamily: FF, fontSize: 12, fontWeight: 600, color: '#0d9488', whiteSpace: 'nowrap' }}>{item.rfqNum}</td>
                <td style={{ ...tdBase, fontFamily: FF, fontSize: 12, fontWeight: 600, color: C.text }}>{item.projectTitle || '—'}</td>
                <td style={{ ...tdBase, fontFamily: FF, fontSize: 12, fontWeight: 600, color: C.text }}>{companyName(item.organizationId)}</td>
                <td style={{ ...tdBase, fontFamily: FF, fontSize: 12, fontWeight: 600, color: C.text }}>{personName(item.personId)}</td>
                <td style={tdBase}>
                  <DisciplineBadges discipline={item.discipline} />
                </td>
                <td style={{ ...tdBase, fontFamily: FF, fontSize: 12, fontWeight: 600, color: C.text, whiteSpace: 'nowrap' }}>
                  {item.quotedDate ? fmtShortDate(item.quotedDate) : '—'}
                </td>
                <td style={{ ...tdBase, fontFamily: FF, fontSize: 12, fontWeight: 600, color: C.text }}>{item.projectValue ? fmtMoney(item.projectValue) : '—'}</td>
                <td style={{ ...tdBase, fontFamily: FF, fontSize: 12, fontWeight: 600, color: C.sub }}>{item.approximateHours ? String(item.approximateHours) : '—'}</td>
                <td style={tdBase}>
                  <span style={{ ...badge, ...statusStyle(item.status) }}>{item.archived ? 'ARCHIVED' : item.status.toUpperCase()}</span>
                </td>
                <td style={{ ...tdBase, padding: '8px 10px 8px 6px' }}>
                  <div style={{ display: 'flex', flexDirection: 'column', gap: 4, minWidth: 64 }}>
                    <button onClick={() => setModal({ ...item })} style={{ ...actionBtn, border: 'none', background: C.purple, color: '#fff' }}>Edit</button>
                    {!item.archived && item.status !== 'Lost' && (
                      <>
                        <button onClick={() => markWon(item)} style={{ ...actionBtn, border: 'none', background: C.green, color: '#fff' }}>WON</button>
                        <button onClick={() => setLinkQuote(item)} style={{ ...actionBtn, border: 'none', background: '#0d9488', color: '#fff' }}>Link 3E</button>
                        <button onClick={() => setLostQuote(item)} style={{ ...actionBtn, border: `1px solid ${C.red}`, background: 'transparent', color: C.red }}>LOST</button>
                      </>
                    )}
                    <button
                      onClick={() => deleteQuote(item.id)}
                      style={{ ...actionBtn, fontSize: 9, border: `1px solid ${pendingDelId === item.id ? C.red : C.borderMd}`, background: pendingDelId === item.id ? C.red : 'transparent', color: pendingDelId === item.id ? '#fff' : C.muted }}
                    >{pendingDelId === item.id ? 'Sure?' : 'Del'}</button>
                  </div>
                </td>
              </tr>
            ))}
          </tbody>
        </table>
        <div style={{ padding: '8px 14px', borderTop: `1px solid ${C.border}`, fontFamily: FF, fontSize: 11, color: C.muted, background: C.thBg }}>
          {filtered.length} of {listSource.length} quote{listSource.length !== 1 ? 's' : ''} ({year}{archiveView ? ', archived — lost 1+ week' : ''}{(dateFrom || dateTo) ? `, ${dateFrom || '…'} – ${dateTo || '…'}` : ''})
        </div>
      </div>

      {modal && (
        <QuoteModal
          initial={modal}
          spService={spService}
          persons={persons}
          companies={companies}
          onSave={saveQuote}
          onClose={() => setModal(null)}
        />
      )}
      {lostQuote && (
        <LostReasonModal
          quote={lostQuote}
          onConfirm={reason => markLost(lostQuote, reason)}
          onClose={() => setLostQuote(null)}
        />
      )}
      {linkQuote && (
        <LinkProjectModal
          quote={linkQuote}
          spService={spService}
          onConfirm={project => linkToProject(linkQuote, project)}
          onClose={() => setLinkQuote(null)}
        />
      )}
      {budgetModal && (
        <BudgetEditModal
          budgetStore={budgetStore}
          editYear={selectedYear}
          yearOptions={yearOptions}
          onSave={store => {
            const normalized = normalizeQuoteBudgetStore(store);
            setBudgetStore(normalized);
            void saveQuoteBudget(spService, normalized);
            setBudgetModal(false);
          }}
          onClose={() => setBudgetModal(false)}
        />
      )}
    </>
  );
};

export default CrmQuotesTab;
