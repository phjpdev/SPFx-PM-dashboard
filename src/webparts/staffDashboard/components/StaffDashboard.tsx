import * as React from 'react';
import { IStaffDashboardProps } from './IStaffDashboardProps';
import {
  Tag, HrsBar, RfiBar, SDiv, Stat, Panel, FF, IBtn, DelModal, useToast,
  BtnPrimary, CcField, fmtD, rfiTot, isOD, effSt, hrsColor, hrsRem, hpct
} from '../../../shared/components/SharedComponents';
import { IProject, IRfi, PROJ_STATUSES, RFI_STATUSES, RFI_TYPES, RFI_RESPONSES } from '../../../shared/models/IProject';
import { SharePointService } from '../../../shared/services/SharePointService';
import styles from '../../../shared/styles/dashboard.module.scss';

// ── Types ────────────────────────────────────────────────────────────────────
type Tab = 'projects' | 'rfis';
type SortDir = 'asc' | 'desc';
interface SortCfg { key: string; dir: SortDir }

// ── Blank templates ──────────────────────────────────────────────────────────
const BLANK_PROJ = (): IProject => ({
  id: '', projNum: '', name: '', discipline: '', status: 'Active', year: new Date().getFullYear(),
  hrsAllowed: 0, hrsUsed: 0, rfisAllowed: 0, quoteNum: '', contact: '', company: '',
  email: '', mobile: '', clientNum: '', clientp0: '', startDate: '', finishDate: '', ifaDate: '',
  ifcDate: '', detailers: '', teamLead: '', teamMembers: '', notes: '', invoices: [], isEwo: false, ewoNum: '', parentId: null
});

const BLANK_RFI = (): IRfi => ({
  id: '', rfiNum: '', rfiSeq: 0, projectId: '', projectName: '', rfiType: '',
  status: 'Open', submittedTo: '', toCompany: '', by: '', byCompany: '', cc: '',
  dateIssued: new Date().toISOString().substring(0, 10), dateRequired: '', description: '',
  attachments: '', clientRfi: '', dateReceived: '', response: 'Pending', responseDesc: '',
  sentBy: '', sentByCompany: '', impacted: 'No', ewoRef: '', ewoCcn: '', tracked: false,
  model: 0, connections: 0, checking: 0, drawings: 0, admin: 0, revision: 'A', email: ''
});

// ── Input / Select helpers ───────────────────────────────────────────────────
const inp = (val: string | number, onChange: (v: string) => void, type = 'text', placeholder = ''): JSX.Element => (
  <input type={type} value={val} onChange={e => onChange(e.target.value)} placeholder={placeholder}
    style={{ fontFamily: 'Montserrat', fontSize: 13, padding: '8px 10px', border: '1px solid var(--bd)', borderRadius: 2, background: 'var(--s2)', color: 'var(--t1)', width: '100%', outline: 'none' }} />
);

const sel = (val: string, onChange: (v: string) => void, opts: string[]): JSX.Element => (
  <select value={val} onChange={e => onChange(e.target.value)}
    style={{ fontFamily: 'Montserrat', fontSize: 13, padding: '8px 10px', border: '1px solid var(--bd)', borderRadius: 2, background: 'var(--s2)', color: 'var(--t1)', width: '100%', outline: 'none', cursor: 'pointer' }}>
    {opts.map(o => <option key={o} value={o}>{o}</option>)}
  </select>
);

const txa = (val: string, onChange: (v: string) => void, rows = 3): JSX.Element => (
  <textarea value={val} onChange={e => onChange(e.target.value)} rows={rows}
    style={{ fontFamily: 'Montserrat', fontSize: 13, padding: '8px 10px', border: '1px solid var(--bd)', borderRadius: 2, background: 'var(--s2)', color: 'var(--t1)', width: '100%', outline: 'none', resize: 'vertical', lineHeight: 1.6 }} />
);

// ── Th (sortable header) ─────────────────────────────────────────────────────
const Th: React.FC<{ label: string; sortKey?: string; sort?: SortCfg; onSort?: (k: string) => void; style?: React.CSSProperties }> =
  ({ label, sortKey, sort, onSort, style }) => {
    const active = sort && sortKey && sort.key === sortKey;
    return (
      <th onClick={sortKey && onSort ? () => onSort(sortKey) : undefined}
        style={{
          fontFamily: 'Montserrat', fontWeight: 700, fontSize: 11.5, letterSpacing: '.1em',
          textTransform: 'uppercase', color: active ? 'var(--3eg)' : 'var(--t4)',
          padding: '10px 14px', background: 'var(--s1)', borderBottom: '2px solid var(--bd)',
          textAlign: 'left', whiteSpace: 'nowrap', cursor: sortKey ? 'pointer' : 'default',
          userSelect: 'none', ...style
        }}>
        {label}{active ? (sort!.dir === 'asc' ? ' ▲' : ' ▼') : ''}
      </th>
    );
  };

// ── Row detail item ──────────────────────────────────────────────────────────
const DRow: React.FC<{ label: string; value?: string | number | null }> = ({ label, value }) => (
  <div style={{ display: 'grid', gridTemplateColumns: '140px 1fr', gap: 8, padding: '7px 0', borderBottom: '1px solid var(--bd)' }}>
    <span style={{ fontFamily: 'Montserrat', fontWeight: 600, fontSize: 11.5, letterSpacing: '.08em', textTransform: 'uppercase', color: 'var(--t4)' }}>{label}</span>
    <span style={{ fontFamily: 'Montserrat', fontSize: 13, color: 'var(--t1)' }}>{value || '—'}</span>
  </div>
);

// ── Main Component ───────────────────────────────────────────────────────────
const StaffDashboard: React.FC<IStaffDashboardProps> = ({ siteUrl, userDisplayName, spHttpClient }) => {
  const svc = React.useMemo(() => new SharePointService(siteUrl, spHttpClient), [siteUrl]);
  const { show: showToast, Toast } = useToast();

  // ── State ──────────────────────────────────────────────────────────────────
  const [tab, setTab] = React.useState<Tab>('projects');
  const [projects, setProjects] = React.useState<IProject[]>([]);
  const [rfis, setRfis] = React.useState<IRfi[]>([]);
  const [loading, setLoading] = React.useState(true);
  const [clock, setClock] = React.useState('');

  // Project filters / sort
  const [pSearch, setPSearch] = React.useState('');
  const [pStatus, setPStatus] = React.useState('All');
  const [pYear, setPYear] = React.useState('All');
  const [pSort, setPSort] = React.useState<SortCfg>({ key: 'projNum', dir: 'asc' });

  // RFI filters / sort
  const [rSearch, setRSearch] = React.useState('');
  const [rStatus, setRStatus] = React.useState('All');
  const [rType, setRType] = React.useState('All');
  const [rProject, setRProject] = React.useState('All');
  const [rSort, setRSort] = React.useState<SortCfg>({ key: 'rfiSeq', dir: 'asc' });

  // Expanded EWO rows
  const [expandedEwo, setExpandedEwo] = React.useState<Set<string>>(new Set());

  // Project panel
  const [projPanel, setProjPanel] = React.useState<'detail' | 'form' | null>(null);
  const [selProj, setSelProj] = React.useState<IProject | null>(null);
  const [projForm, setProjForm] = React.useState<IProject>(BLANK_PROJ());
  const [projSaving, setProjSaving] = React.useState(false);

  // RFI panel
  const [rfiPanel, setRfiPanel] = React.useState<'detail' | 'form' | null>(null);
  const [selRfi, setSelRfi] = React.useState<IRfi | null>(null);
  const [rfiForm, setRfiForm] = React.useState<IRfi>(BLANK_RFI());
  const [rfiSaving, setRfiSaving] = React.useState(false);

  // Delete modal
  const [delTarget, setDelTarget] = React.useState<{ type: 'project' | 'rfi'; item: IProject | IRfi } | null>(null);

  // Time Doctor import state
  const [tdImporting, setTdImporting] = React.useState(false);

  // ── Clock ──────────────────────────────────────────────────────────────────
  React.useEffect(() => {
    const tick = (): void => {
      const now = new Date();
      setClock(now.toLocaleTimeString('en-AU', { hour: '2-digit', minute: '2-digit', second: '2-digit' }));
    };
    tick();
    const id = setInterval(tick, 1000);
    return () => clearInterval(id);
  }, []);

  // ── Load data ──────────────────────────────────────────────────────────────
  const loadAll = React.useCallback(async () => {
    setLoading(true);
    try {
      const [ps, rs] = await Promise.all([svc.loadProjects(), svc.loadRfis()]);
      setProjects(ps);
      setRfis(rs);
    } catch (e: any) {
      showToast('Failed to load data: ' + (e.message || e), 'error');
    } finally {
      setLoading(false);
    }
  }, [svc]);

  React.useEffect(() => { void loadAll(); }, [loadAll]);

  // ── Derived: years for filter ──────────────────────────────────────────────
  const years = React.useMemo(() => {
    const s = new Set(projects.map(p => String(p.year)));
    return ['All', ...Array.from(s).sort((a, b) => Number(b) - Number(a))];
  }, [projects]);

  const rfiProjectNames = React.useMemo(() => {
    const s = new Set(rfis.map(r => r.projectName).filter(Boolean));
    return ['All', ...Array.from(s).sort()];
  }, [rfis]);

  // ── Sort helper ────────────────────────────────────────────────────────────
  function applySort<T>(arr: T[], cfg: SortCfg): T[] {
    return [...arr].sort((a: any, b: any) => {
      let av = a[cfg.key], bv = b[cfg.key];
      if (typeof av === 'string') av = av.toLowerCase();
      if (typeof bv === 'string') bv = bv.toLowerCase();
      if (av < bv) return cfg.dir === 'asc' ? -1 : 1;
      if (av > bv) return cfg.dir === 'asc' ? 1 : -1;
      return 0;
    });
  }

  function toggleSort(cur: SortCfg, key: string, setCfg: (c: SortCfg) => void): void {
    if (cur.key === key) setCfg({ key, dir: cur.dir === 'asc' ? 'desc' : 'asc' });
    else setCfg({ key, dir: 'asc' });
  }

  // ── Filtered/sorted projects ───────────────────────────────────────────────
  const filteredProjects = React.useMemo(() => {
    const q = pSearch.toLowerCase();
    let ps = projects.filter(p => {
      if (pStatus !== 'All' && p.status !== pStatus) return false;
      if (pYear !== 'All' && String(p.year) !== pYear) return false;
      if (q && !p.projNum.toLowerCase().includes(q) && !p.name.toLowerCase().includes(q) && !p.company.toLowerCase().includes(q)) return false;
      return true;
    });
    return applySort(ps, pSort);
  }, [projects, pSearch, pStatus, pYear, pSort]);

  // Top-level projects (non-EWO or EWO with no parent visible) for table
  const topProjects = React.useMemo(() =>
    filteredProjects.filter(p => !p.isEwo || p.parentId == null),
    [filteredProjects]
  );

  const ewoChildren = React.useMemo(() => {
    const m: Record<string, IProject[]> = {};
    filteredProjects.filter(p => p.isEwo && p.parentId != null).forEach(p => {
      if (!m[p.parentId!]) m[p.parentId!] = [];
      m[p.parentId!].push(p);
    });
    return m;
  }, [filteredProjects]);

  // ── Filtered/sorted RFIs ───────────────────────────────────────────────────
  const filteredRfis = React.useMemo(() => {
    const q = rSearch.toLowerCase();
    let rs = rfis.filter(r => {
      if (rStatus !== 'All' && effSt(r) !== rStatus) return false;
      if (rType !== 'All' && r.rfiType !== rType) return false;
      if (rProject !== 'All' && r.projectName !== rProject) return false;
      if (q && !r.rfiNum.toLowerCase().includes(q) && !r.description.toLowerCase().includes(q) && !r.projectName.toLowerCase().includes(q)) return false;
      return true;
    });
    return applySort(rs, rSort);
  }, [rfis, rSearch, rStatus, rType, rProject, rSort]);

  // Group RFIs by project
  const rfisByProject = React.useMemo(() => {
    const groups: { projName: string; items: IRfi[] }[] = [];
    const seen: Record<string, number> = {};
    filteredRfis.forEach(r => {
      const key = r.projectName || '(No Project)';
      if (seen[key] === undefined) {
        seen[key] = groups.length;
        groups.push({ projName: key, items: [] });
      }
      groups[seen[key]].items.push(r);
    });
    return groups;
  }, [filteredRfis]);

  // ── Project stat cards ─────────────────────────────────────────────────────
  const pStats = React.useMemo(() => {
    const active = projects.filter(p => p.status === 'Active').length;
    const ob = projects.filter(p => p.status === 'Over Budget' || (p.hrsAllowed > 0 && p.hrsUsed > p.hrsAllowed)).length;
    const complete = projects.filter(p => p.status === 'Complete').length;
    const onHold = projects.filter(p => p.status === 'On Hold').length;
    return { active, ob, complete, onHold };
  }, [projects]);

  // ── RFI stat cards ─────────────────────────────────────────────────────────
  const rStats = React.useMemo(() => {
    const open = rfis.filter(r => effSt(r) === 'Open').length;
    const overdue = rfis.filter(r => isOD(r)).length;
    const closed = rfis.filter(r => r.status === 'Closed').length;
    const total = rfis.length;
    return { open, overdue, closed, total };
  }, [rfis]);

  // ── CRUD: Project ──────────────────────────────────────────────────────────
  const openNewProject = React.useCallback(() => {
    setProjForm(BLANK_PROJ());
    setSelProj(null);
    setProjPanel('form');
  }, []);

  const openEditProject = React.useCallback((p: IProject) => {
    setProjForm({ ...p });
    setSelProj(p);
    setProjPanel('form');
  }, []);

  const openDetailProject = React.useCallback((p: IProject) => {
    setSelProj(p);
    setProjPanel('detail');
  }, []);

  const saveProject = React.useCallback(async () => {
    if (!projForm.projNum.trim()) { showToast('Project number is required', 'error'); return; }
    if (!projForm.name.trim()) { showToast('Project name is required', 'error'); return; }
    setProjSaving(true);
    try {
      if (selProj && selProj.spId) {
        await svc.updateProject(selProj.spId, projForm);
        showToast('Project updated successfully');
      } else {
        const newId = await svc.addProject(projForm);
        showToast('Project created successfully');
        const updatedForm = { ...projForm, spId: newId };
        setProjForm(updatedForm);
      }
      await loadAll();
      setProjPanel(null);
    } catch (e: any) {
      showToast('Save failed: ' + (e.message || e), 'error');
    } finally {
      setProjSaving(false);
    }
  }, [projForm, selProj, svc, loadAll]);

  const confirmDeleteProject = React.useCallback((p: IProject) => {
    setDelTarget({ type: 'project', item: p });
  }, []);

  const execDelete = React.useCallback(async () => {
    if (!delTarget) return;
    try {
      if (delTarget.type === 'project') {
        const p = delTarget.item as IProject;
        if (!p.spId) { showToast('Cannot delete: no SharePoint ID', 'error'); return; }
        await svc.deleteProject(p.spId);
        showToast('Project deleted');
        setProjPanel(null);
      } else {
        const r = delTarget.item as IRfi;
        if (!r.spId) { showToast('Cannot delete: no SharePoint ID', 'error'); return; }
        await svc.deleteRfi(r.spId);
        showToast('RFI deleted');
        setRfiPanel(null);
      }
      await loadAll();
    } catch (e: any) {
      showToast('Delete failed: ' + (e.message || e), 'error');
    } finally {
      setDelTarget(null);
    }
  }, [delTarget, svc, loadAll]);

  // ── CRUD: RFI ──────────────────────────────────────────────────────────────
  const openNewRfi = React.useCallback((projId?: string, projName?: string) => {
    const blank = BLANK_RFI();
    if (projId) { blank.projectId = projId; blank.projectName = projName || ''; }
    // Auto-assign next RFI seq
    const maxSeq = rfis.reduce((m, r) => Math.max(m, r.rfiSeq || 0), 0);
    blank.rfiSeq = maxSeq + 1;
    // Auto-assign rfiNum
    blank.rfiNum = `RFI-${String(blank.rfiSeq).padStart(3, '0')}`;
    setRfiForm(blank);
    setSelRfi(null);
    setRfiPanel('form');
  }, [rfis]);

  const openEditRfi = React.useCallback((r: IRfi) => {
    setRfiForm({ ...r });
    setSelRfi(r);
    setRfiPanel('form');
  }, []);

  const openDetailRfi = React.useCallback((r: IRfi) => {
    setSelRfi(r);
    setRfiPanel('detail');
  }, []);

  const saveRfi = React.useCallback(async () => {
    if (!rfiForm.projectId.trim()) { showToast('Project is required', 'error'); return; }
    if (!rfiForm.description.trim()) { showToast('Description is required', 'error'); return; }
    setRfiSaving(true);
    try {
      if (selRfi && selRfi.spId) {
        await svc.updateRfi(selRfi.spId, rfiForm);
        showToast('RFI updated successfully');
      } else {
        const newId = await svc.addRfi(rfiForm);
        showToast('RFI created successfully');
        const updatedRfi = { ...rfiForm, spId: newId };
        setRfiForm(updatedRfi);
      }
      await loadAll();
      setRfiPanel(null);
    } catch (e: any) {
      showToast('Save failed: ' + (e.message || e), 'error');
    } finally {
      setRfiSaving(false);
    }
  }, [rfiForm, selRfi, svc, loadAll]);

  const confirmDeleteRfi = React.useCallback((r: IRfi) => {
    setDelTarget({ type: 'rfi', item: r });
  }, []);

  // ── Time Doctor XLS Import ─────────────────────────────────────────────────
  const handleTimeDoctorImport = React.useCallback((e: React.ChangeEvent<HTMLInputElement>) => {
    const file = e.target.files?.[0];
    if (!file) return;
    const XLSX = (window as any).XLSX;
    if (!XLSX) { showToast('XLSX library not loaded. Please ensure SheetJS is available.', 'error'); return; }
    setTdImporting(true);
    const reader = new FileReader();
    reader.onload = async (ev) => {
      try {
        const data = ev.target?.result;
        const wb = XLSX.read(data, { type: 'binary' });
        const ws = wb.Sheets[wb.SheetNames[0]];
        const rows: any[] = XLSX.utils.sheet_to_json(ws, { defval: '' });

        // Expected columns: "Project" (or project number), "Hours" (decimal hours)
        const updates: { projNum: string; hours: number }[] = [];
        rows.forEach((row: any) => {
          const projNum = String(row['Project'] || row['project'] || row['Project Number'] || '').trim();
          const hrs = parseFloat(row['Hours'] || row['hours'] || row['Total Hours'] || '0') || 0;
          if (projNum && hrs > 0) updates.push({ projNum, hours: hrs });
        });

        if (updates.length === 0) {
          showToast('No valid rows found in XLS. Expected "Project" and "Hours" columns.', 'error');
          setTdImporting(false);
          return;
        }

        // Aggregate by project number
        const agg: Record<string, number> = {};
        updates.forEach(u => { agg[u.projNum] = (agg[u.projNum] || 0) + u.hours; });

        let updated = 0;
        let notFound: string[] = [];
        for (const [pNum, hrs] of Object.entries(agg)) {
          const proj = projects.find(p => p.projNum === pNum);
          if (proj && proj.spId) {
            const updated_proj = { ...proj, hrsUsed: parseFloat((proj.hrsUsed + hrs).toFixed(2)) };
            await svc.updateProject(proj.spId, updated_proj);
            updated++;
          } else {
            notFound.push(pNum);
          }
        }

        await loadAll();
        let msg = `Time Doctor import complete. Updated ${updated} project(s).`;
        if (notFound.length > 0) msg += `\nNot found: ${notFound.join(', ')}`;
        showToast(msg, updated > 0 ? 'success' : 'warn');
      } catch (err: any) {
        showToast('Import failed: ' + (err.message || err), 'error');
      } finally {
        setTdImporting(false);
      }
    };
    reader.readAsBinaryString(file);
    e.target.value = '';
  }, [projects, svc, loadAll]);

  // ── PDF Generation for RFI ─────────────────────────────────────────────────
  const generateRfiPdf = React.useCallback((r: IRfi) => {
    const jspdf = (window as any).jspdf;
    if (!jspdf) { showToast('jsPDF not loaded. Please ensure jsPDF is available.', 'error'); return; }
    try {
      const { jsPDF } = jspdf;
      const doc = new jsPDF({ orientation: 'portrait', unit: 'mm', format: 'a4' });
      const pgW = 210;
      let y = 18;

      // Header
      doc.setFillColor(42, 158, 42);
      doc.rect(0, 0, pgW, 28, 'F');
      doc.setFont('helvetica', 'bold');
      doc.setFontSize(18);
      doc.setTextColor(255, 255, 255);
      doc.text('3 EDGE DESIGN', 14, 12);
      doc.setFontSize(10);
      doc.setFont('helvetica', 'normal');
      doc.text('REQUEST FOR INFORMATION', 14, 20);
      doc.setFont('helvetica', 'bold');
      doc.setFontSize(13);
      doc.text(r.rfiNum || '', pgW - 14, 16, { align: 'right' });
      doc.setFontSize(9);
      doc.setFont('helvetica', 'normal');
      doc.text(`Rev: ${r.revision || 'A'}`, pgW - 14, 22, { align: 'right' });

      y = 36;
      doc.setTextColor(30, 30, 50);

      const addSection = (title: string): void => {
        doc.setFillColor(240, 242, 245);
        doc.rect(12, y - 4, pgW - 24, 8, 'F');
        doc.setFont('helvetica', 'bold');
        doc.setFontSize(9);
        doc.setTextColor(42, 100, 42);
        doc.text(title.toUpperCase(), 14, y + 1);
        doc.setTextColor(30, 30, 50);
        y += 8;
      };

      const addRow = (label: string, value: string, half = false): void => {
        doc.setFont('helvetica', 'bold');
        doc.setFontSize(8);
        doc.setTextColor(100, 110, 130);
        doc.text(label.toUpperCase(), 14, y);
        doc.setFont('helvetica', 'normal');
        doc.setTextColor(30, 30, 50);
        doc.setFontSize(9);
        const maxW = half ? (pgW / 2) - 20 : pgW - 28;
        const lines = doc.splitTextToSize(value || '—', maxW);
        doc.text(lines, 55, y);
        y += Math.max(6, lines.length * 5);
      };

      addSection('Project Information');
      addRow('Project', r.projectName);
      addRow('Project ID', r.projectId);
      addRow('RFI Type', r.rfiType);
      y += 2;

      addSection('Submission Details');
      addRow('Submitted To', r.submittedTo);
      addRow('Company', r.toCompany);
      addRow('Submitted By', r.by);
      addRow('By Company', r.byCompany);
      addRow('Date Issued', fmtD(r.dateIssued));
      addRow('Date Required', fmtD(r.dateRequired));
      addRow('Client RFI Ref', r.clientRfi);
      y += 2;

      addSection('Description');
      doc.setFont('helvetica', 'normal');
      doc.setFontSize(9);
      const descLines = doc.splitTextToSize(r.description || '—', pgW - 28);
      doc.text(descLines, 14, y);
      y += descLines.length * 5 + 4;

      addSection('Response');
      addRow('Status', effSt(r));
      addRow('Response', r.response);
      addRow('Date Received', fmtD(r.dateReceived));
      if (r.responseDesc) addRow('Response Notes', r.responseDesc);
      y += 2;

      if (r.tracked) {
        addSection('Hours Breakdown');
        addRow('Model', String(r.model || 0) + ' hrs');
        addRow('Connections', String(r.connections || 0) + ' hrs');
        addRow('Checking', String(r.checking || 0) + ' hrs');
        addRow('Drawings', String(r.drawings || 0) + ' hrs');
        addRow('Admin', String(r.admin || 0) + ' hrs');
        addRow('TOTAL', String(rfiTot(r).toFixed(1)) + ' hrs');
        y += 2;
      }

      if (r.impacted === 'Yes') {
        addSection('EWO / Impact');
        addRow('Impacted', 'Yes');
        if (r.ewoRef) addRow('EWO Ref', r.ewoRef);
        if (r.ewoCcn) addRow('CCN', r.ewoCcn);
      }

      // Footer
      doc.setFontSize(8);
      doc.setTextColor(150, 160, 180);
      doc.text(`Generated ${new Date().toLocaleDateString('en-AU')} — 3 Edge Design Staff Dashboard`, pgW / 2, 290, { align: 'center' });

      doc.save(`${r.rfiNum || 'RFI'}-Rev${r.revision || 'A'}.pdf`);
      showToast(`PDF saved: ${r.rfiNum}-Rev${r.revision || 'A'}.pdf`);
    } catch (err: any) {
      showToast('PDF generation failed: ' + (err.message || err), 'error');
    }
  }, []);

  // ── Set project from RFI form selector ────────────────────────────────────
  const handleRfiProjectChange = React.useCallback((projId: string) => {
    const p = projects.find(pr => pr.id === projId || pr.projNum === projId);
    setRfiForm(prev => ({
      ...prev,
      projectId: projId,
      projectName: p?.name || prev.projectName
    }));
  }, [projects]);

  // ── Render: Project Form ───────────────────────────────────────────────────
  const renderProjectForm = (): JSX.Element => {
    const fp = projForm;
    const set = (k: keyof IProject) => (v: string) => setProjForm(prev => ({ ...prev, [k]: v }));
    const setNum = (k: keyof IProject) => (v: string) => setProjForm(prev => ({ ...prev, [k]: parseFloat(v) || 0 }));
    const setBool = (k: keyof IProject) => (v: string) => setProjForm(prev => ({ ...prev, [k]: v === 'true' }));
    return (
      <div style={{ display: 'grid', gridTemplateColumns: '1fr 1fr', gap: '14px 20px' }}>
        <FF label="Project Number" ><input type="text" value={fp.projNum} onChange={e => setProjForm(p => ({ ...p, projNum: e.target.value }))} style={{ fontFamily: 'Montserrat', fontSize: 13, padding: '8px 10px', border: '1px solid var(--bd)', borderRadius: 2, background: 'var(--s2)', color: 'var(--t1)', width: '100%', outline: 'none' }} /></FF>
        <FF label="Project Name"><input type="text" value={fp.name} onChange={e => setProjForm(p => ({ ...p, name: e.target.value }))} style={{ fontFamily: 'Montserrat', fontSize: 13, padding: '8px 10px', border: '1px solid var(--bd)', borderRadius: 2, background: 'var(--s2)', color: 'var(--t1)', width: '100%', outline: 'none' }} /></FF>
        <FF label="Status">{sel(fp.status, set('status'), PROJ_STATUSES)}</FF>
        <FF label="Year">{inp(fp.year, set('year'), 'number')}</FF>
        <FF label="Hrs Allowed">{inp(fp.hrsAllowed, setNum('hrsAllowed'), 'number')}</FF>
        <FF label="Hrs Used">{inp(fp.hrsUsed, setNum('hrsUsed'), 'number')}</FF>
        <FF label="RFIs Allowed">{inp(fp.rfisAllowed, setNum('rfisAllowed'), 'number')}</FF>
        <FF label="Quote Number">{inp(fp.quoteNum, set('quoteNum'))}</FF>
        <FF label="Contact">{inp(fp.contact, set('contact'))}</FF>
        <FF label="Company">{inp(fp.company, set('company'))}</FF>
        <FF label="Email">{inp(fp.email, set('email'), 'email')}</FF>
        <FF label="Mobile">{inp(fp.mobile, set('mobile'))}</FF>
        <FF label="Client Number">{inp(fp.clientNum, set('clientNum'))}</FF>
        <FF label="Start Date">{inp(fp.startDate, set('startDate'), 'date')}</FF>
        <FF label="Finish Date">{inp(fp.finishDate, set('finishDate'), 'date')}</FF>
        <FF label="IFA Date">{inp(fp.ifaDate, set('ifaDate'), 'date')}</FF>
        <FF label="IFC Date">{inp(fp.ifcDate, set('ifcDate'), 'date')}</FF>
        <FF label="Detailers" span2>{inp(fp.detailers, set('detailers'))}</FF>
        <FF label="Is EWO">
          {sel(fp.isEwo ? 'true' : 'false', setBool('isEwo'), ['false', 'true'])}
        </FF>
        {fp.isEwo && <>
          <FF label="EWO Number">{inp(fp.ewoNum, set('ewoNum'))}</FF>
          <FF label="Parent Project">
            <select value={fp.parentId || ''} onChange={e => setProjForm(p => ({ ...p, parentId: e.target.value || null }))}
              style={{ fontFamily: 'Montserrat', fontSize: 13, padding: '8px 10px', border: '1px solid var(--bd)', borderRadius: 2, background: 'var(--s2)', color: 'var(--t1)', width: '100%', outline: 'none', cursor: 'pointer' }}>
              <option value="">— None —</option>
              {projects.filter(p => !p.isEwo && p.projNum !== fp.projNum).map(p => (
                <option key={p.id} value={p.id}>{p.projNum} — {p.name}</option>
              ))}
            </select>
          </FF>
        </>}
        <div style={{ gridColumn: 'span 2', display: 'flex', gap: 10, paddingTop: 12, borderTop: '1px solid var(--bd)', marginTop: 4 }}>
          <BtnPrimary onClick={saveProject}>{projSaving ? 'Saving…' : selProj ? 'Update Project' : 'Create Project'}</BtnPrimary>
          <button onClick={() => setProjPanel(null)} style={{ fontFamily: 'Montserrat', fontSize: 12.5, padding: '9px 18px', background: 'transparent', border: '1px solid var(--bd)', color: 'var(--t2)', borderRadius: 7, cursor: 'pointer' }}>Cancel</button>
          {selProj && <button onClick={() => confirmDeleteProject(selProj)} style={{ fontFamily: 'Montserrat', fontWeight: 700, fontSize: 12.5, padding: '9px 18px', background: 'var(--rd2)', border: '1px solid var(--rd)', color: 'var(--rd)', borderRadius: 7, cursor: 'pointer', marginLeft: 'auto' }}>Delete</button>}
        </div>
      </div>
    );
  };

  // ── Render: Project Detail ─────────────────────────────────────────────────
  const renderProjectDetail = (p: IProject): JSX.Element => {
    const projRfis = rfis.filter(r => r.projectId === p.id || r.projectId === p.projNum);
    const openRfis = projRfis.filter(r => effSt(r) === 'Open').length;
    const overdueRfis = projRfis.filter(r => isOD(r)).length;
    return (
      <div>
        <div style={{ display: 'flex', gap: 10, marginBottom: 20, flexWrap: 'wrap' }}>
          <IBtn onClick={() => openEditProject(p)}>Edit</IBtn>
          <IBtn onClick={() => confirmDeleteProject(p)} danger>Delete</IBtn>
          <IBtn onClick={() => { openNewRfi(p.id, p.name); setProjPanel(null); }}>+ RFI</IBtn>
        </div>
        <SDiv label="Project Info" />
        <DRow label="Number" value={p.projNum} />
        <DRow label="Name" value={p.name} />
        <DRow label="Status" value={p.status} />
        <DRow label="Year" value={p.year} />
        <DRow label="Quote #" value={p.quoteNum} />
        <DRow label="Client #" value={p.clientNum} />
        <SDiv label="Hours" />
        <div style={{ padding: '12px 0' }}>
          <HrsBar allowed={p.hrsAllowed} used={p.hrsUsed} />
          <div style={{ marginTop: 8, fontFamily: 'Montserrat', fontSize: 12, color: 'var(--t4)' }}>
            Remaining: <strong style={{ color: hrsColor(hpct(p), p.hrsUsed > p.hrsAllowed) }}>{hrsRem(p) !== null ? hrsRem(p) + 'h' : '—'}</strong>
          </div>
        </div>
        <SDiv label="RFIs" />
        <div style={{ display: 'flex', gap: 16, padding: '10px 0 16px', flexWrap: 'wrap' }}>
          <div style={{ fontFamily: 'Montserrat', fontSize: 13 }}>Total: <strong>{projRfis.length}</strong></div>
          <div style={{ fontFamily: 'Montserrat', fontSize: 13 }}>Open: <strong style={{ color: 'var(--gn)' }}>{openRfis}</strong></div>
          {overdueRfis > 0 && <div style={{ fontFamily: 'Montserrat', fontSize: 13 }}>Overdue: <strong style={{ color: 'var(--rd)' }}>{overdueRfis}</strong></div>}
        </div>
        <RfiBar allowed={p.rfisAllowed} used={projRfis.length} />
        <SDiv label="Contact" />
        <DRow label="Contact" value={p.contact} />
        <DRow label="Company" value={p.company} />
        <DRow label="Email" value={p.email} />
        <DRow label="Mobile" value={p.mobile} />
        <SDiv label="Dates" />
        <DRow label="Start" value={fmtD(p.startDate)} />
        <DRow label="Finish" value={fmtD(p.finishDate)} />
        <DRow label="IFA" value={fmtD(p.ifaDate)} />
        <DRow label="IFC" value={fmtD(p.ifcDate)} />
        {p.detailers && <><SDiv label="Team" /><DRow label="Detailers" value={p.detailers} /></>}
        {p.isEwo && <><SDiv label="EWO" /><DRow label="EWO Number" value={p.ewoNum} /></>}
      </div>
    );
  };

  // ── Render: RFI Form ───────────────────────────────────────────────────────
  const renderRfiForm = (): JSX.Element => {
    const f = rfiForm;
    const set = (k: keyof IRfi) => (v: string) => setRfiForm(prev => ({ ...prev, [k]: v }));
    const setNum = (k: keyof IRfi) => (v: string) => setRfiForm(prev => ({ ...prev, [k]: parseFloat(v) || 0 }));
    return (
      <div style={{ display: 'grid', gridTemplateColumns: '1fr 1fr', gap: '14px 20px' }}>
        <FF label="Project">
          <select value={f.projectId} onChange={e => handleRfiProjectChange(e.target.value)}
            style={{ fontFamily: 'Montserrat', fontSize: 13, padding: '8px 10px', border: '1px solid var(--bd)', borderRadius: 2, background: 'var(--s2)', color: 'var(--t1)', width: '100%', outline: 'none', cursor: 'pointer' }}>
            <option value="">— Select Project —</option>
            {projects.map(p => <option key={p.id} value={p.id}>{p.projNum} — {p.name}</option>)}
          </select>
        </FF>
        <FF label="RFI Number">{inp(f.rfiNum, set('rfiNum'))}</FF>
        <FF label="RFI Seq">{inp(f.rfiSeq, setNum('rfiSeq'), 'number')}</FF>
        <FF label="Revision">{inp(f.revision || 'A', set('revision'))}</FF>
        <FF label="Type">{sel(f.rfiType || RFI_TYPES[0], set('rfiType'), RFI_TYPES)}</FF>
        <FF label="Status">{sel(f.status, set('status'), RFI_STATUSES)}</FF>
        <FF label="Date Issued">{inp(f.dateIssued, set('dateIssued'), 'date')}</FF>
        <FF label="Date Required">{inp(f.dateRequired, set('dateRequired'), 'date')}</FF>
        <FF label="Submitted To">{inp(f.submittedTo, set('submittedTo'))}</FF>
        <FF label="To Company">{inp(f.toCompany, set('toCompany'))}</FF>
        <FF label="Submitted By">{inp(f.by, set('by'))}</FF>
        <FF label="By Company">{inp(f.byCompany, set('byCompany'))}</FF>
        <FF label="Email">{inp(f.email || '', set('email'), 'email')}</FF>
        <FF label="Client RFI Ref">{inp(f.clientRfi, set('clientRfi'))}</FF>
        <FF label="Description" span2>{txa(f.description, set('description'), 4)}</FF>
        <FF label="CC" span2>
          <CcField value={f.cc} onChange={v => setRfiForm(prev => ({ ...prev, cc: v }))} />
        </FF>
        <FF label="Response">{sel(f.response || 'Pending', set('response'), RFI_RESPONSES)}</FF>
        <FF label="Date Received">{inp(f.dateReceived, set('dateReceived'), 'date')}</FF>
        <FF label="Sent By">{inp(f.sentBy, set('sentBy'))}</FF>
        <FF label="Sent By Company">{inp(f.sentByCompany, set('sentByCompany'))}</FF>
        <FF label="Response Notes" span2>{txa(f.responseDesc || '', set('responseDesc'), 3)}</FF>
        <FF label="Impacted">{sel(f.impacted || 'No', set('impacted'), ['No', 'Yes'])}</FF>
        <FF label="EWO Ref">{inp(f.ewoRef || '', set('ewoRef'))}</FF>
        <FF label="EWO CCN">{inp(f.ewoCcn, set('ewoCcn'))}</FF>
        <FF label="Track Hours">
          <select value={f.tracked ? 'true' : 'false'} onChange={e => setRfiForm(prev => ({ ...prev, tracked: e.target.value === 'true' }))}
            style={{ fontFamily: 'Montserrat', fontSize: 13, padding: '8px 10px', border: '1px solid var(--bd)', borderRadius: 2, background: 'var(--s2)', color: 'var(--t1)', width: '100%', outline: 'none', cursor: 'pointer' }}>
            <option value="false">No</option>
            <option value="true">Yes</option>
          </select>
        </FF>
        {f.tracked && <>
          <FF label="Model Hrs">{inp(f.model, setNum('model'), 'number')}</FF>
          <FF label="Connections Hrs">{inp(f.connections, setNum('connections'), 'number')}</FF>
          <FF label="Checking Hrs">{inp(f.checking, setNum('checking'), 'number')}</FF>
          <FF label="Drawings Hrs">{inp(f.drawings, setNum('drawings'), 'number')}</FF>
          <FF label="Admin Hrs">{inp(f.admin, setNum('admin'), 'number')}</FF>
          <FF label="Total">
            <div style={{ fontFamily: 'Montserrat', fontWeight: 700, fontSize: 14, color: 'var(--3eg)', padding: '8px 0' }}>{rfiTot(f).toFixed(1)} hrs</div>
          </FF>
        </>}
        <div style={{ gridColumn: 'span 2', display: 'flex', gap: 10, paddingTop: 12, borderTop: '1px solid var(--bd)', marginTop: 4 }}>
          <BtnPrimary onClick={saveRfi}>{rfiSaving ? 'Saving…' : selRfi ? 'Update RFI' : 'Create RFI'}</BtnPrimary>
          <button onClick={() => setRfiPanel(null)} style={{ fontFamily: 'Montserrat', fontSize: 12.5, padding: '9px 18px', background: 'transparent', border: '1px solid var(--bd)', color: 'var(--t2)', borderRadius: 7, cursor: 'pointer' }}>Cancel</button>
          {selRfi && <button onClick={() => confirmDeleteRfi(selRfi)} style={{ fontFamily: 'Montserrat', fontWeight: 700, fontSize: 12.5, padding: '9px 18px', background: 'var(--rd2)', border: '1px solid var(--rd)', color: 'var(--rd)', borderRadius: 7, cursor: 'pointer', marginLeft: 'auto' }}>Delete</button>}
        </div>
      </div>
    );
  };

  // ── Render: RFI Detail ─────────────────────────────────────────────────────
  const renderRfiDetail = (r: IRfi): JSX.Element => (
    <div>
      <div style={{ display: 'flex', gap: 10, marginBottom: 20, flexWrap: 'wrap' }}>
        <IBtn onClick={() => openEditRfi(r)}>Edit</IBtn>
        <IBtn onClick={() => confirmDeleteRfi(r)} danger>Delete</IBtn>
        <IBtn onClick={() => generateRfiPdf(r)}>PDF</IBtn>
      </div>
      <SDiv label="RFI Info" />
      <DRow label="RFI Number" value={r.rfiNum} />
      <DRow label="Revision" value={r.revision} />
      <DRow label="Project" value={r.projectName} />
      <DRow label="Type" value={r.rfiType} />
      <DRow label="Status" value={effSt(r)} />
      <DRow label="Client RFI" value={r.clientRfi} />
      <SDiv label="Submission" />
      <DRow label="Submitted To" value={r.submittedTo} />
      <DRow label="Company" value={r.toCompany} />
      <DRow label="Submitted By" value={r.by} />
      <DRow label="By Company" value={r.byCompany} />
      <DRow label="Date Issued" value={fmtD(r.dateIssued)} />
      <DRow label="Date Required" value={fmtD(r.dateRequired)} />
      <DRow label="Email" value={r.email} />
      {r.cc && <DRow label="CC" value={r.cc} />}
      <SDiv label="Description" />
      <div style={{ fontFamily: 'Montserrat', fontSize: 13, color: 'var(--t1)', lineHeight: 1.7, padding: '8px 0', whiteSpace: 'pre-wrap' }}>{r.description || '—'}</div>
      <SDiv label="Response" />
      <DRow label="Response" value={r.response} />
      <DRow label="Date Received" value={fmtD(r.dateReceived)} />
      <DRow label="Sent By" value={r.sentBy} />
      <DRow label="Sent By Company" value={r.sentByCompany} />
      {r.responseDesc && <DRow label="Notes" value={r.responseDesc} />}
      {(r.impacted === 'Yes' || r.ewoRef) && <>
        <SDiv label="EWO / Impact" />
        <DRow label="Impacted" value={r.impacted} />
        {r.ewoRef && <DRow label="EWO Ref" value={r.ewoRef} />}
        {r.ewoCcn && <DRow label="CCN" value={r.ewoCcn} />}
      </>}
      {r.tracked && <>
        <SDiv label="Hours" />
        <DRow label="Model" value={r.model + ' hrs'} />
        <DRow label="Connections" value={r.connections + ' hrs'} />
        <DRow label="Checking" value={r.checking + ' hrs'} />
        <DRow label="Drawings" value={r.drawings + ' hrs'} />
        <DRow label="Admin" value={r.admin + ' hrs'} />
        <DRow label="Total" value={rfiTot(r).toFixed(1) + ' hrs'} />
      </>}
    </div>
  );

  // ── Render: Header ─────────────────────────────────────────────────────────
  const renderHeader = (): JSX.Element => (
    <div style={{
      background: 'var(--hdr)', borderBottom: '3px solid var(--3eg)',
      padding: '0 28px', display: 'flex', alignItems: 'center', gap: 0,
      height: 56, flexShrink: 0
    }}>
      <div style={{ display: 'flex', alignItems: 'baseline', gap: 8, marginRight: 32 }}>
        <span style={{ fontFamily: 'Montserrat', fontWeight: 900, fontSize: 19, letterSpacing: '.12em', color: '#fff' }}>3</span>
        <span style={{ fontFamily: 'Montserrat', fontWeight: 900, fontSize: 19, letterSpacing: '.12em', color: 'var(--3eg)' }}>EDGE</span>
        <span style={{ fontFamily: 'Montserrat', fontWeight: 400, fontSize: 12.5, letterSpacing: '.18em', textTransform: 'uppercase', color: 'var(--t4)', marginLeft: 2 }}>STAFF</span>
      </div>
      {(['projects', 'rfis'] as Tab[]).map(t => (
        <button key={t} onClick={() => setTab(t)} style={{
          fontFamily: 'Montserrat', fontWeight: 700, fontSize: 12.5, letterSpacing: '.1em', textTransform: 'uppercase',
          padding: '0 20px', height: 56, background: 'transparent', border: 'none',
          borderBottom: tab === t ? '3px solid var(--3eg)' : '3px solid transparent',
          color: tab === t ? '#fff' : 'var(--t3)', cursor: 'pointer', transition: 'all .15s'
        }}>
          {t === 'projects' ? 'Project Tracker' : 'RFI Tracker'}
        </button>
      ))}
      <div style={{ flex: 1 }} />
      <div style={{ display: 'flex', alignItems: 'center', gap: 10 }}>
        <label style={{ display: 'flex', alignItems: 'center', gap: 8, cursor: tdImporting ? 'wait' : 'pointer', opacity: tdImporting ? 0.6 : 1 }}
          title="Import Time Doctor XLS">
          <span style={{ fontFamily: 'Montserrat', fontWeight: 700, fontSize: 11, letterSpacing: '.08em', textTransform: 'uppercase', color: 'var(--t4)', whiteSpace: 'nowrap' }}>
            {tdImporting ? 'Importing…' : 'TD Import'}
          </span>
          <input type="file" accept=".xls,.xlsx" style={{ display: 'none' }} onChange={handleTimeDoctorImport} disabled={tdImporting} />
          <span style={{ fontFamily: 'Montserrat', fontWeight: 700, fontSize: 11.5, padding: '5px 12px', background: 'rgba(42,158,42,0.15)', border: '1px solid var(--3eg)', color: 'var(--3eg)', borderRadius: 5, cursor: 'pointer' }}>XLS</span>
        </label>
        <button onClick={loadAll} disabled={loading} title="Refresh data"
          style={{ fontFamily: 'Montserrat', fontWeight: 700, fontSize: 11.5, padding: '5px 12px', background: 'rgba(42,158,42,0.15)', border: '1px solid var(--3eg)', color: 'var(--3eg)', borderRadius: 5, cursor: 'pointer', opacity: loading ? 0.6 : 1 }}>
          {loading ? '…' : '↻'}
        </button>
        <div style={{ fontFamily: 'Montserrat', fontWeight: 700, fontSize: 13, color: 'var(--3eg)', letterSpacing: '.04em', minWidth: 76, textAlign: 'right' }}>{clock}</div>
        <div style={{ fontFamily: 'Montserrat', fontSize: 11.5, color: 'var(--t4)', marginLeft: 4, whiteSpace: 'nowrap' }}>{userDisplayName}</div>
      </div>
    </div>
  );

  // ── Render: Project Table ──────────────────────────────────────────────────
  const renderProjectRow = (p: IProject, isChild = false): JSX.Element => {
    const children = ewoChildren[p.id] || [];
    const hasChildren = children.length > 0;
    const expanded = expandedEwo.has(p.id);
    const projRfiCount = rfis.filter(r => r.projectId === p.id || r.projectId === p.projNum).length;
    return (
      <React.Fragment key={p.id}>
        <tr className={isChild ? styles.ewoRow : ''} style={{
          background: isChild ? 'rgba(107,79,200,0.04)' : 'var(--s1)',
          borderBottom: '1px solid var(--bd)', cursor: 'pointer',
          transition: 'background .1s'
        }}
          onMouseEnter={e => (e.currentTarget.style.background = isChild ? 'rgba(107,79,200,0.09)' : 'var(--s2)')}
          onMouseLeave={e => (e.currentTarget.style.background = isChild ? 'rgba(107,79,200,0.04)' : 'var(--s1)')}
        >
          <td style={{ padding: '10px 14px', fontFamily: 'Montserrat', fontWeight: 700, fontSize: 13, color: 'var(--3eg)', whiteSpace: 'nowrap' }}
            onClick={() => openDetailProject(p)}>
            {isChild && <span style={{ marginRight: 6, color: 'var(--pu)', fontSize: 11 }}>└</span>}
            {hasChildren && (
              <span onClick={e => { e.stopPropagation(); setExpandedEwo(s => { const n = new Set(s); if (n.has(p.id)) { n.delete(p.id); } else { n.add(p.id); } return n; }); }}
                style={{ marginRight: 6, cursor: 'pointer', color: 'var(--t4)', fontSize: 11 }}>
                {expanded ? '▼' : '▶'}
              </span>
            )}
            {p.projNum}
          </td>
          <td style={{ padding: '10px 14px', fontFamily: 'Montserrat', fontSize: 13.5, color: 'var(--t1)', maxWidth: 260 }} onClick={() => openDetailProject(p)}>
            <div style={{ overflow: 'hidden', textOverflow: 'ellipsis', whiteSpace: 'nowrap' }}>{p.name}</div>
            {p.company && <div style={{ fontFamily: 'Montserrat', fontSize: 11, color: 'var(--t4)', marginTop: 2 }}>{p.company}</div>}
          </td>
          <td style={{ padding: '10px 14px' }} onClick={() => openDetailProject(p)}><Tag s={p.status} /></td>
          <td style={{ padding: '10px 14px' }} onClick={() => openDetailProject(p)}><HrsBar allowed={p.hrsAllowed} used={p.hrsUsed} /></td>
          <td style={{ padding: '10px 14px' }} onClick={() => openDetailProject(p)}>
            <RfiBar allowed={p.rfisAllowed} used={projRfiCount} />
          </td>
          <td style={{ padding: '10px 14px', fontFamily: 'Montserrat', fontSize: 12, color: 'var(--t4)' }} onClick={() => openDetailProject(p)}>
            {p.detailers || '—'}
          </td>
          <td style={{ padding: '10px 14px', fontFamily: 'Montserrat', fontSize: 12, color: 'var(--t4)', whiteSpace: 'nowrap' }} onClick={() => openDetailProject(p)}>
            {fmtD(p.finishDate)}
          </td>
          <td style={{ padding: '8px 10px', whiteSpace: 'nowrap' }}>
            <div style={{ display: 'flex', gap: 5 }}>
              <IBtn onClick={() => openDetailProject(p)} title="View">👁</IBtn>
              <IBtn onClick={() => openEditProject(p)} title="Edit">✏</IBtn>
              <IBtn onClick={() => { openNewRfi(p.id, p.name); }} title="New RFI">+RFI</IBtn>
              <IBtn onClick={() => confirmDeleteProject(p)} danger title="Delete">🗑</IBtn>
            </div>
          </td>
        </tr>
        {hasChildren && expanded && children.map(c => renderProjectRow(c, true))}
      </React.Fragment>
    );
  };

  // ── Render: Project Tab ────────────────────────────────────────────────────
  const renderProjectsTab = (): JSX.Element => (
    <div className={styles.fade}>
      {/* Stat cards */}
      <div style={{ display: 'flex', gap: 14, marginBottom: 22, flexWrap: 'wrap' }}>
        <Stat label="Active" value={pStats.active} col="var(--3eg)" />
        <Stat label="Complete" value={pStats.complete} col="var(--bl)" />
        <Stat label="On Hold" value={pStats.onHold} col="var(--pu)" />
        <Stat label="Over Budget" value={pStats.ob} col="var(--rd)" warn={pStats.ob > 0} />
        <Stat label="Total Projects" value={projects.length} col="var(--t4)" />
      </div>

      {/* Toolbar */}
      <div style={{ display: 'flex', gap: 10, marginBottom: 16, flexWrap: 'wrap', alignItems: 'center' }}>
        <input value={pSearch} onChange={e => setPSearch(e.target.value)} placeholder="Search projects…"
          style={{ fontFamily: 'Montserrat', fontSize: 13, padding: '8px 12px', border: '1px solid var(--bd)', borderRadius: 6, background: 'var(--s1)', color: 'var(--t1)', width: 220, outline: 'none' }} />
        <select value={pStatus} onChange={e => setPStatus(e.target.value)}
          style={{ fontFamily: 'Montserrat', fontSize: 13, padding: '8px 10px', border: '1px solid var(--bd)', borderRadius: 6, background: 'var(--s1)', color: 'var(--t1)', outline: 'none', cursor: 'pointer' }}>
          <option value="All">All Statuses</option>
          {PROJ_STATUSES.map(s => <option key={s} value={s}>{s}</option>)}
        </select>
        <select value={pYear} onChange={e => setPYear(e.target.value)}
          style={{ fontFamily: 'Montserrat', fontSize: 13, padding: '8px 10px', border: '1px solid var(--bd)', borderRadius: 6, background: 'var(--s1)', color: 'var(--t1)', outline: 'none', cursor: 'pointer' }}>
          {years.map(y => <option key={y} value={y}>{y === 'All' ? 'All Years' : y}</option>)}
        </select>
        <div style={{ flex: 1 }} />
        <BtnPrimary onClick={openNewProject}>+ New Project</BtnPrimary>
      </div>

      {/* Table */}
      <div style={{ background: 'var(--s1)', border: '1px solid var(--bd)', borderRadius: 10, overflow: 'hidden', boxShadow: '0 1px 4px rgba(0,0,0,.06)' }}>
        <div style={{ overflowX: 'auto' }}>
          <table>
            <thead>
              <tr>
                <Th label="Proj #" sortKey="projNum" sort={pSort} onSort={k => toggleSort(pSort, k, setPSort)} />
                <Th label="Name" sortKey="name" sort={pSort} onSort={k => toggleSort(pSort, k, setPSort)} />
                <Th label="Status" sortKey="status" sort={pSort} onSort={k => toggleSort(pSort, k, setPSort)} />
                <Th label="Hours" />
                <Th label="RFIs" />
                <Th label="Detailers" sortKey="detailers" sort={pSort} onSort={k => toggleSort(pSort, k, setPSort)} />
                <Th label="Finish" sortKey="finishDate" sort={pSort} onSort={k => toggleSort(pSort, k, setPSort)} />
                <Th label="Actions" />
              </tr>
            </thead>
            <tbody>
              {loading ? (
                <tr><td colSpan={8} style={{ padding: '40px', textAlign: 'center', fontFamily: 'Montserrat', color: 'var(--t4)', fontSize: 14 }}>Loading…</td></tr>
              ) : topProjects.length === 0 ? (
                <tr><td colSpan={8} style={{ padding: '40px', textAlign: 'center', fontFamily: 'Montserrat', color: 'var(--t4)', fontSize: 14 }}>No projects found</td></tr>
              ) : topProjects.map(p => renderProjectRow(p))}
            </tbody>
          </table>
        </div>
        <div style={{ padding: '10px 16px', borderTop: '1px solid var(--bd)', fontFamily: 'Montserrat', fontSize: 11.5, color: 'var(--t4)' }}>
          {filteredProjects.length} project{filteredProjects.length !== 1 ? 's' : ''} shown
        </div>
      </div>
    </div>
  );

  // ── Render: RFI Tab ────────────────────────────────────────────────────────
  const renderRfisTab = (): JSX.Element => (
    <div className={styles.fade}>
      {/* Stat cards */}
      <div style={{ display: 'flex', gap: 14, marginBottom: 22, flexWrap: 'wrap' }}>
        <Stat label="Open" value={rStats.open} col="var(--3eg)" />
        <Stat label="Overdue" value={rStats.overdue} col="var(--rd)" warn={rStats.overdue > 0} />
        <Stat label="Closed" value={rStats.closed} col="var(--bl)" />
        <Stat label="Total RFIs" value={rStats.total} col="var(--t4)" />
      </div>

      {/* Toolbar */}
      <div style={{ display: 'flex', gap: 10, marginBottom: 16, flexWrap: 'wrap', alignItems: 'center' }}>
        <input value={rSearch} onChange={e => setRSearch(e.target.value)} placeholder="Search RFIs…"
          style={{ fontFamily: 'Montserrat', fontSize: 13, padding: '8px 12px', border: '1px solid var(--bd)', borderRadius: 6, background: 'var(--s1)', color: 'var(--t1)', width: 220, outline: 'none' }} />
        <select value={rStatus} onChange={e => setRStatus(e.target.value)}
          style={{ fontFamily: 'Montserrat', fontSize: 13, padding: '8px 10px', border: '1px solid var(--bd)', borderRadius: 6, background: 'var(--s1)', color: 'var(--t1)', outline: 'none', cursor: 'pointer' }}>
          <option value="All">All Statuses</option>
          {RFI_STATUSES.map(s => <option key={s} value={s}>{s}</option>)}
        </select>
        <select value={rType} onChange={e => setRType(e.target.value)}
          style={{ fontFamily: 'Montserrat', fontSize: 13, padding: '8px 10px', border: '1px solid var(--bd)', borderRadius: 6, background: 'var(--s1)', color: 'var(--t1)', outline: 'none', cursor: 'pointer' }}>
          <option value="All">All Types</option>
          {RFI_TYPES.map(t => <option key={t} value={t}>{t}</option>)}
        </select>
        <select value={rProject} onChange={e => setRProject(e.target.value)}
          style={{ fontFamily: 'Montserrat', fontSize: 13, padding: '8px 10px', border: '1px solid var(--bd)', borderRadius: 6, background: 'var(--s1)', color: 'var(--t1)', outline: 'none', cursor: 'pointer' }}>
          {rfiProjectNames.map(n => <option key={n} value={n}>{n === 'All' ? 'All Projects' : n}</option>)}
        </select>
        <div style={{ flex: 1 }} />
        <BtnPrimary onClick={() => openNewRfi()}>+ New RFI</BtnPrimary>
      </div>

      {/* Grouped table */}
      {loading ? (
        <div style={{ padding: '40px', textAlign: 'center', fontFamily: 'Montserrat', color: 'var(--t4)', fontSize: 14 }}>Loading…</div>
      ) : filteredRfis.length === 0 ? (
        <div style={{ padding: '40px', textAlign: 'center', fontFamily: 'Montserrat', color: 'var(--t4)', fontSize: 14 }}>No RFIs found</div>
      ) : rfisByProject.map(group => (
        <div key={group.projName} style={{ marginBottom: 22, background: 'var(--s1)', border: '1px solid var(--bd)', borderRadius: 10, overflow: 'hidden', boxShadow: '0 1px 4px rgba(0,0,0,.06)' }}>
          <div style={{ padding: '10px 16px', background: 'var(--s2)', borderBottom: '2px solid var(--3eg)', display: 'flex', alignItems: 'center', gap: 12 }}>
            <span style={{ fontFamily: 'Montserrat', fontWeight: 800, fontSize: 13, color: 'var(--t1)', letterSpacing: '.04em' }}>{group.projName}</span>
            <span style={{ fontFamily: 'Montserrat', fontSize: 11.5, color: 'var(--t4)' }}>{group.items.length} RFI{group.items.length !== 1 ? 's' : ''}</span>
            <span style={{ fontFamily: 'Montserrat', fontSize: 11, color: 'var(--am)' }}>
              {group.items.filter(r => isOD(r)).length > 0 ? `⚠ ${group.items.filter(r => isOD(r)).length} overdue` : ''}
            </span>
          </div>
          <div style={{ overflowX: 'auto' }}>
            <table>
              <thead>
                <tr>
                  <Th label="RFI #" sortKey="rfiNum" sort={rSort} onSort={k => toggleSort(rSort, k, setRSort)} />
                  <Th label="Type" sortKey="rfiType" sort={rSort} onSort={k => toggleSort(rSort, k, setRSort)} />
                  <Th label="Status" sortKey="status" sort={rSort} onSort={k => toggleSort(rSort, k, setRSort)} />
                  <Th label="Description" />
                  <Th label="Submitted To" sortKey="submittedTo" sort={rSort} onSort={k => toggleSort(rSort, k, setRSort)} />
                  <Th label="Issued" sortKey="dateIssued" sort={rSort} onSort={k => toggleSort(rSort, k, setRSort)} />
                  <Th label="Required" sortKey="dateRequired" sort={rSort} onSort={k => toggleSort(rSort, k, setRSort)} />
                  <Th label="Hours" />
                  <Th label="Actions" />
                </tr>
              </thead>
              <tbody>
                {group.items.map(r => {
                  const overdue = isOD(r);
                  return (
                    <tr key={r.id}
                      style={{ background: overdue ? 'rgba(204,51,51,0.04)' : 'var(--s1)', borderBottom: '1px solid var(--bd)', cursor: 'pointer', transition: 'background .1s' }}
                      onMouseEnter={e => (e.currentTarget.style.background = overdue ? 'rgba(204,51,51,0.09)' : 'var(--s2)')}
                      onMouseLeave={e => (e.currentTarget.style.background = overdue ? 'rgba(204,51,51,0.04)' : 'var(--s1)')}
                    >
                      <td style={{ padding: '9px 14px', fontFamily: 'Montserrat', fontWeight: 700, fontSize: 13, color: 'var(--3eg)', whiteSpace: 'nowrap' }} onClick={() => openDetailRfi(r)}>
                        {r.rfiNum}
                        {r.revision && r.revision !== 'A' && <span style={{ fontFamily: 'Montserrat', fontWeight: 500, fontSize: 10.5, color: 'var(--t4)', marginLeft: 4 }}>Rev{r.revision}</span>}
                      </td>
                      <td style={{ padding: '9px 14px', fontFamily: 'Montserrat', fontSize: 12, color: 'var(--t3)', whiteSpace: 'nowrap' }} onClick={() => openDetailRfi(r)}>{r.rfiType || '—'}</td>
                      <td style={{ padding: '9px 14px' }} onClick={() => openDetailRfi(r)}><Tag s={effSt(r)} /></td>
                      <td style={{ padding: '9px 14px', fontFamily: 'Montserrat', fontSize: 13, color: 'var(--t1)', maxWidth: 280 }} onClick={() => openDetailRfi(r)}>
                        <div style={{ overflow: 'hidden', textOverflow: 'ellipsis', whiteSpace: 'nowrap' }}>{r.description || '—'}</div>
                      </td>
                      <td style={{ padding: '9px 14px', fontFamily: 'Montserrat', fontSize: 12, color: 'var(--t3)', whiteSpace: 'nowrap' }} onClick={() => openDetailRfi(r)}>
                        {r.submittedTo || '—'}
                        {r.toCompany && <div style={{ fontFamily: 'Montserrat', fontSize: 11, color: 'var(--t4)' }}>{r.toCompany}</div>}
                      </td>
                      <td style={{ padding: '9px 14px', fontFamily: 'Montserrat', fontSize: 12, color: 'var(--t4)', whiteSpace: 'nowrap' }} onClick={() => openDetailRfi(r)}>{fmtD(r.dateIssued)}</td>
                      <td style={{ padding: '9px 14px', fontFamily: 'Montserrat', fontSize: 12, color: overdue ? 'var(--rd)' : 'var(--t4)', fontWeight: overdue ? 700 : 400, whiteSpace: 'nowrap' }} onClick={() => openDetailRfi(r)}>{fmtD(r.dateRequired)}</td>
                      <td style={{ padding: '9px 14px', fontFamily: 'Montserrat', fontSize: 12, color: 'var(--t4)', whiteSpace: 'nowrap' }} onClick={() => openDetailRfi(r)}>
                        {r.tracked ? <span style={{ color: 'var(--3eg)', fontWeight: 600 }}>{rfiTot(r).toFixed(1)}h</span> : '—'}
                      </td>
                      <td style={{ padding: '8px 10px', whiteSpace: 'nowrap' }}>
                        <div style={{ display: 'flex', gap: 5 }}>
                          <IBtn onClick={() => openDetailRfi(r)} title="View">👁</IBtn>
                          <IBtn onClick={() => openEditRfi(r)} title="Edit">✏</IBtn>
                          <IBtn onClick={() => generateRfiPdf(r)} title="PDF">PDF</IBtn>
                          <IBtn onClick={() => confirmDeleteRfi(r)} danger title="Delete">🗑</IBtn>
                        </div>
                      </td>
                    </tr>
                  );
                })}
              </tbody>
            </table>
          </div>
        </div>
      ))}
      <div style={{ fontFamily: 'Montserrat', fontSize: 11.5, color: 'var(--t4)', marginTop: 8 }}>
        {filteredRfis.length} RFI{filteredRfis.length !== 1 ? 's' : ''} shown across {rfisByProject.length} project{rfisByProject.length !== 1 ? 's' : ''}
      </div>
    </div>
  );

  // ── Main render ────────────────────────────────────────────────────────────
  return (
    <div className={styles.dashboardRoot}>
      {renderHeader()}
      <div style={{ flex: 1, padding: '24px 28px', overflowY: 'auto' }}>
        {tab === 'projects' ? renderProjectsTab() : renderRfisTab()}
      </div>

      {/* Project panel */}
      <Panel
        open={projPanel !== null}
        onClose={() => setProjPanel(null)}
        title={projPanel === 'form' ? (selProj ? 'Edit Project' : 'New Project') : (selProj?.projNum || 'Project Detail')}
        subtitle={projPanel === 'detail' ? selProj?.name : undefined}
        tag={projPanel === 'detail' && selProj ? <Tag s={selProj.status} /> : undefined}
      >
        {projPanel === 'form' && renderProjectForm()}
        {projPanel === 'detail' && selProj && renderProjectDetail(selProj)}
      </Panel>

      {/* RFI panel */}
      <Panel
        open={rfiPanel !== null}
        onClose={() => setRfiPanel(null)}
        title={rfiPanel === 'form' ? (selRfi ? 'Edit RFI' : 'New RFI') : (selRfi?.rfiNum || 'RFI Detail')}
        subtitle={rfiPanel === 'detail' ? selRfi?.projectName : undefined}
        tag={rfiPanel === 'detail' && selRfi ? <Tag s={effSt(selRfi)} /> : undefined}
      >
        {rfiPanel === 'form' && renderRfiForm()}
        {rfiPanel === 'detail' && selRfi && renderRfiDetail(selRfi)}
      </Panel>

      {/* Delete modal */}
      <DelModal
        open={delTarget !== null}
        label={delTarget?.type === 'project'
          ? `Delete project "${(delTarget.item as IProject).projNum} — ${(delTarget.item as IProject).name}"? This cannot be undone.`
          : `Delete RFI "${(delTarget?.item as IRfi)?.rfiNum}"? This cannot be undone.`}
        onConfirm={execDelete}
        onCancel={() => setDelTarget(null)}
      />

      {Toast}
    </div>
  );
};

export default StaffDashboard;
