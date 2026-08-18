import * as React from 'react';
import {
  Tag, HrsBar, RfiBar, Stat, Panel, FF, IBtn, DelModal, useToast,
  BtnPrimary, SDiv, CcField, fmtD, rfiTot, effSt, isOD, RememberSenderField
} from '../../../shared/components/SharedComponents';
import { IProject, IRfi, PROJ_STATUSES, RFI_STATUSES, RFI_TYPES, RFI_RESPONSES, isProjectDelivered, withProjectDeliveryOnSave, withProjectArchiveToggle } from '../../../shared/models/IProject';
import { SharePointService } from '../../../shared/services/SharePointService';
import styles from './ManagerDashboard.module.scss';
import type { IManagerDashboardProps } from './IManagerDashboardProps';

import { jsPDF } from 'jspdf';
import * as XLSX from 'xlsx';
import TaskBoard from './TaskBoard';
import ChecklistBoard from './ChecklistBoard';
import CrmBoard from './CrmBoard';

// ── Assets ────────────────────────────────────────────────────────────────────
import logoImg from '../assets/3edge-logo.png';
import { drawLetterhead, drawPdfBg } from '../../../shared/utils/pdfLetterhead';
import { applyProjectDefaultsById } from '../../../shared/utils/rfiProjectDefaults';
import { applySenderDefaultsToRfi, RFI_DEFAULT_BY_COMPANY, saveSenderDefaults } from '../../../shared/utils/rfiSenderDefaults';
import { rfiPdfFileName } from '../../../shared/utils/rfiPdfFilename';
import { ITeamMember } from '../../../shared/models/ITask';

// ── Montserrat local fonts ─────────────────────────────────────────────────────
import _fExtraLight from '../assets/Montserrat-ExtraLight.ttf';
import _fBold from '../assets/Montserrat-Bold.ttf';
import _fBoldI from '../assets/Montserrat-BoldItalic.ttf';
import _fExtraBold from '../assets/Montserrat-ExtraBold.ttf';
import _fExtraBoldI from '../assets/Montserrat-ExtraBoldItalic.ttf';
import _fBlack from '../assets/Montserrat-Black.ttf';
import _fBlackI from '../assets/Montserrat-BlackItalic.ttf';
const IMG_LOGO_DASH: string = logoImg;

(function injectMontserrat(): void {
  const id = '3edge-montserrat';
  if (document.getElementById(id)) return;
  const s = document.createElement('style');
  s.id = id;
  s.textContent = [
    `@font-face{font-family:'Montserrat';font-weight:200;font-style:normal;src:url('${_fExtraLight}') format('truetype')}`,
    `@font-face{font-family:'Montserrat';font-weight:700;font-style:normal;src:url('${_fBold}') format('truetype')}`,
    `@font-face{font-family:'Montserrat';font-weight:700;font-style:italic;src:url('${_fBoldI}') format('truetype')}`,
    `@font-face{font-family:'Montserrat';font-weight:800;font-style:normal;src:url('${_fExtraBold}') format('truetype')}`,
    `@font-face{font-family:'Montserrat';font-weight:800;font-style:italic;src:url('${_fExtraBoldI}') format('truetype')}`,
    `@font-face{font-family:'Montserrat';font-weight:900;font-style:normal;src:url('${_fBlack}') format('truetype')}`,
    `@font-face{font-family:'Montserrat';font-weight:900;font-style:italic;src:url('${_fBlackI}') format('truetype')}`,
  ].join('');
  document.head.appendChild(s);
})();

// ── Types ─────────────────────────────────────────────────────────────────────
type Mod = 'projects' | 'rfis' | 'ewos' | 'tasks' | 'checklist' | 'crm';
type SDir = 'asc' | 'desc';
type Role = 'manager' | 'staff';
type SpMode = 'live' | 'local' | 'detecting';

interface PanelState {
  type: 'projDetail' | 'projForm' | 'rfiDetail' | 'rfiForm' | 'ewoForm' | 'ewoDetail' | null;
  proj?: IProject | null;
  rfi?: IRfi | null;
  parentProj?: IProject | null;
}

interface DelState {
  open: boolean;
  label: string;
  onConfirm: () => void;
}

// ── Empty factories ────────────────────────────────────────────────────────────
const emptyProj = (): IProject => ({
  id: '', spId: undefined, projNum: '', name: '', discipline: 'Steel', status: 'Active', year: new Date().getFullYear(),
  hrsAllowed: 0, hrsUsed: 0, rfisAllowed: 0, quoteNum: '', contact: '', company: '',
  email: '', mobile: '', clientNum: '', clientp0: '', startDate: '', finishDate: '', ifaDate: '', ifcDate: '',
  detailers: '', teamLead: '', teamMembers: '', notes: '', invoices: [], isEwo: false, ewoNum: '', parentId: null
});

const emptyRfi = (): IRfi => ({
  id: '', spId: undefined, rfiNum: '', rfiSeq: 0, projectId: '', projectName: '',
  rfiType: RFI_TYPES[0], status: 'Open', submittedTo: '', toCompany: '', by: '', byCompany: '',
  cc: '', dateIssued: new Date().toISOString().substring(0, 10), dateRequired: '',
  description: '', attachments: '', clientRfi: '', dateReceived: '', response: 'Pending',
  responseDesc: '', sentBy: '', sentByCompany: '', impacted: 'No', ewoRef: '', ewoCcn: '',
  tracked: false, model: 0, connections: 0, checking: 0, drawings: 0, admin: 0,
  revision: 'A', email: ''
});

// ── Inline style helpers ───────────────────────────────────────────────────────
const inp: React.CSSProperties = {
  fontFamily: 'Montserrat', fontSize: 13, fontWeight: 600, padding: '8px 12px',
  border: '1px solid var(--bd)', borderRadius: 2,
  background: 'var(--s2)', color: 'var(--t1)', width: '100%', outline: 'none'
};
const selStyle: React.CSSProperties = { ...inp, cursor: 'pointer' };

// ── PDF Generator ─────────────────────────────────────────────────────────────
function generateRfiPdf(rfi: IRfi, proj: IProject | undefined): Blob | undefined {
  try {
    // eslint-disable-next-line @typescript-eslint/no-explicit-any
    const doc: any = new jsPDF({ orientation: 'portrait', unit: 'mm', format: 'a4' });
    const pw = 210; const ph = 297;
    const ml = 15; const mr = 15; const tw = pw - ml - mr;
    let y = drawLetterhead(doc, pw, ph, 'REQUEST FOR INFORMATION', rfi.rfiNum + '  |  Revision: ' + (rfi.revision || 'A'));

    // Helper: section header
    const sectionHeader = (title: string): void => {
      doc.setFillColor(240, 242, 245);
      doc.rect(ml, y, tw, 7, 'F');
      doc.setDrawColor(208, 213, 222);
      doc.rect(ml, y, tw, 7, 'S');
      doc.setFillColor(42, 158, 42);
      doc.rect(ml, y, 3, 7, 'F');
      doc.setFontSize(8);
      doc.setFont('helvetica', 'bold');
      doc.setTextColor(26, 32, 48);
      doc.text(title, ml + 5, y + 4.5);
      y += 9;
    };

    // Helper: two-col row
    const row2 = (l1: string, v1: string, l2: string, v2: string): void => {
      const cw = tw / 2;
      doc.setFontSize(7.5);
      doc.setFont('helvetica', 'bold');
      doc.setTextColor(90, 110, 136);
      doc.text(l1.toUpperCase(), ml, y + 3.5);
      doc.setFont('helvetica', 'normal');
      doc.setTextColor(26, 32, 48);
      doc.text(String(v1 || '—'), ml + 28, y + 3.5, { maxWidth: cw - 30 });
      doc.setFont('helvetica', 'bold');
      doc.setTextColor(90, 110, 136);
      doc.text(l2.toUpperCase(), ml + cw, y + 3.5);
      doc.setFont('helvetica', 'normal');
      doc.setTextColor(26, 32, 48);
      doc.text(String(v2 || '—'), ml + cw + 28, y + 3.5, { maxWidth: cw - 30 });
      doc.setDrawColor(208, 213, 222);
      doc.line(ml, y + 6, ml + tw, y + 6);
      y += 7;
    };

    // Helper: full-width row
    const row1 = (label: string, value: string, bold?: boolean): void => {
      doc.setFontSize(7.5);
      doc.setFont('helvetica', 'bold');
      doc.setTextColor(90, 110, 136);
      doc.text(label.toUpperCase(), ml, y + 3.5);
      doc.setFont('helvetica', bold ? 'bold' : 'normal');
      doc.setTextColor(bold ? 42 : 26, bold ? 158 : 32, bold ? 42 : 48);
      doc.text(String(value || '—'), ml + 40, y + 3.5, { maxWidth: tw - 42 });
      doc.setDrawColor(208, 213, 222);
      doc.line(ml, y + 6, ml + tw, y + 6);
      y += 7;
    };

    // Helper: text block
    const textBlock = (label: string, value: string): void => {
      doc.setFontSize(7.5);
      doc.setFont('helvetica', 'bold');
      doc.setTextColor(90, 110, 136);
      doc.text(label.toUpperCase(), ml, y + 3.5);
      y += 6;
      if (value) {
        doc.setFont('helvetica', 'normal');
        doc.setTextColor(26, 32, 48);
        doc.setFontSize(8);
        const lines = doc.splitTextToSize(value, tw);
        doc.text(lines, ml, y + 4);
        y += lines.length * 5 + 2;
      } else {
        y += 4;
      }
      doc.setDrawColor(208, 213, 222);
      doc.line(ml, y, ml + tw, y);
      y += 3;
    };

    // Part A
    sectionHeader('PART A — REQUEST INFORMATION');
    row2('Project #', proj ? proj.projNum : rfi.projectId, 'Project Name', proj ? proj.name : rfi.projectName);
    row2('RFI Number', rfi.rfiNum, 'RFI Type', rfi.rfiType);
    row2('Date Issued', fmtD(rfi.dateIssued), 'Date Required', fmtD(rfi.dateRequired));
    row2('Submitted To', rfi.submittedTo, 'To Company', rfi.toCompany);
    row2('Prepared By', rfi.by, 'Company', rfi.byCompany);
    if (rfi.cc) row1('CC', rfi.cc);
    y += 3;

    // Part B
    sectionHeader('PART B — DESCRIPTION');
    textBlock('Description', rfi.description);
    if (rfi.attachments) row1('Attachments', rfi.attachments);
    y += 3;

    // Parts C & D
    sectionHeader('PARTS C & D — CLIENT RESPONSE');
    row2('Client RFI #', rfi.clientRfi, 'Date Received', fmtD(rfi.dateReceived));
    row2('Response', rfi.response, 'Status', rfi.status);
    row2('Sent By', rfi.sentBy, 'Sent By Company', rfi.sentByCompany);
    if (rfi.responseDesc) textBlock('Response Details', rfi.responseDesc);
    y += 3;

    // Part E
    sectionHeader('PART E — IMPACT ASSESSMENT');
    row2('Schedule Impact', rfi.impacted, 'EWO Reference', rfi.ewoRef || '—');
    if (rfi.impacted === 'Yes') {
      const total = (rfi.model || 0) + (rfi.connections || 0) + (rfi.checking || 0) + (rfi.drawings || 0) + (rfi.admin || 0);
      row2('Model Hrs', String(rfi.model || 0), 'Connections Hrs', String(rfi.connections || 0));
      row2('Checking Hrs', String(rfi.checking || 0), 'Drawings Hrs', String(rfi.drawings || 0));
      row2('Admin Hrs', String(rfi.admin || 0), 'Total Impact Hrs', String(total));
    }

    // Portrait: footer already in background image
    return doc.output('blob') as Blob;
  } catch (e) {
    console.error('PDF generation error:', e);
    alert('PDF generation failed.');
    return undefined;
  }
}

// ── Export All Projects PDF ───────────────────────────────────────────────────
function generateAllProjectsPdf(projects: IProject[], rfis: IRfi[]): void {
  try {
    // eslint-disable-next-line @typescript-eslint/no-explicit-any
    const doc: any = new jsPDF({ orientation: 'portrait', unit: 'mm', format: 'a4' });
    const pw = 210; const ph = 297;
    const ml = 10; const mr = 10; const tw = pw - ml - mr;
    let y = drawLetterhead(doc, pw, ph, 'PROJECT LIST', projects.length + ' projects  |  ' + new Date().toLocaleDateString('en-AU'));
    const cols = [
      { label: 'PROJECT #', w: 22 }, { label: 'QUOTE #', w: 18 }, { label: 'NAME', w: 42 },
      { label: 'COMPANY', w: 32 }, { label: 'CONTACT', w: 22 }, { label: 'HRS USED', w: 18 },
      { label: 'HRS ALLOWED', w: 22 }, { label: 'START', w: 20 }, { label: 'FINISH', w: 20 },
      { label: 'RFIS', w: 12 }, { label: 'EWOS', w: 12 }, { label: 'STATUS', w: 18 }
    ];
    const totalW = cols.reduce((s, c) => s + c.w, 0);
    const scale = tw / totalW;
    const scaledCols = cols.map(c => ({ ...c, w: c.w * scale }));

    const drawHeader = (): void => {
      doc.setFillColor(240, 242, 245);
      let x = ml;
      scaledCols.forEach(c => {
        doc.rect(x, y, c.w, 7, 'F');
        x += c.w;
      });
      doc.setDrawColor(208, 213, 222);
      doc.rect(ml, y, tw, 7, 'S');
      doc.setFontSize(6.5); doc.setFont('helvetica', 'bold'); doc.setTextColor(90, 110, 136);
      x = ml;
      scaledCols.forEach(c => {
        doc.text(c.label, x + 2, y + 4.5);
        x += c.w;
      });
      y += 8;
    };

    drawHeader();

    const mainProjects = projects.filter(p => !p.isEwo);
    let rowIdx = 0;
    mainProjects.forEach(p => {
      if (y + 7 > ph - 12) { doc.addPage(); y = 10; drawHeader(); rowIdx = 0; }
      // Alternating row background
      if (rowIdx % 2 === 1) { doc.setFillColor(245, 247, 250); doc.rect(ml, y, tw, 7, 'F'); }
      rowIdx++;
      const ewoCount = projects.filter(e => e.isEwo && e.parentId === p.id).length;
      const rfiCount = rfis.filter(r => r.projectId === p.id).length;
      const vals = [
        p.projNum, p.quoteNum || '—', p.name || '—', p.company || '—', p.contact || '—',
        String(p.hrsUsed), String(p.hrsAllowed || '—'), fmtD(p.startDate), fmtD(p.finishDate),
        String(rfiCount), String(ewoCount), p.status
      ];
      doc.setFontSize(7); doc.setFont('helvetica', 'normal'); doc.setTextColor(26, 32, 48);
      let x = ml;
      vals.forEach((v, i) => {
        if (i === 0) { doc.setFont('helvetica', 'bold'); doc.setTextColor(42, 158, 42); }
        else if (i === 11) { doc.setFont('helvetica', 'bold'); doc.setTextColor(v === 'Active' ? 42 : v === 'Complete' ? 46 : 90, v === 'Active' ? 158 : v === 'Complete' ? 109 : 110, v === 'Active' ? 42 : v === 'Complete' ? 180 : 136); }
        else { doc.setFont('helvetica', 'normal'); doc.setTextColor(40, 50, 65); }
        const txt = doc.splitTextToSize(String(v), scaledCols[i].w - 3);
        doc.text(txt[0] || '—', x + 2, y + 4);
        x += scaledCols[i].w;
      });
      doc.setDrawColor(220, 225, 230);
      doc.line(ml, y + 6, ml + tw, y + 6);
      y += 7;
    });

    // Footer on all pages
    const pages = doc.internal.getNumberOfPages();
    for (let i = 2; i <= pages; i++) { doc.setPage(i); drawPdfBg(doc, pw, ph); }

    doc.save('All_Projects_' + new Date().toISOString().substring(0, 10) + '.pdf');
  } catch (e) {
    console.error('Export all projects PDF error:', e);
    alert('PDF generation failed.');
  }
}

// ── Export All RFIs PDF ───────────────────────────────────────────────────────
function generateAllRfisPdf(rfis: IRfi[], projects: IProject[]): void {
  try {
    // eslint-disable-next-line @typescript-eslint/no-explicit-any
    const doc: any = new jsPDF({ orientation: 'portrait', unit: 'mm', format: 'a4' });
    const pw = 210; const ph = 297; const ml = 10; const mr = 10; const tw = pw - ml - mr;
    let y = drawLetterhead(doc, pw, ph, 'RFI LIST', rfis.length + ' RFIs  |  ' + new Date().toLocaleDateString('en-AU'));
    const cols = [
      { label: 'RFI #', w: 28 }, { label: 'PROJECT', w: 24 }, { label: 'PROJECT NAME', w: 35 },
      { label: 'TYPE', w: 22 }, { label: 'STATUS', w: 18 }, { label: 'ISSUED', w: 18 },
      { label: 'REQUIRED', w: 18 }, { label: 'TO', w: 20 }, { label: 'COMPANY', w: 22 },
      { label: 'RESPONSE', w: 20 }, { label: 'IMPACT', w: 14 }, { label: 'DESCRIPTION', w: 38 }
    ];
    const totalW = cols.reduce((s, c) => s + c.w, 0);
    const scale = tw / totalW;
    const sc = cols.map(c => ({ ...c, w: c.w * scale }));
    const drawHdr = (): void => {
      doc.setFillColor(240, 242, 245); let x = ml;
      sc.forEach(c => { doc.rect(x, y, c.w, 7, 'F'); x += c.w; });
      doc.setDrawColor(208, 213, 222); doc.rect(ml, y, tw, 7, 'S');
      doc.setFontSize(6.5); doc.setFont('helvetica', 'bold'); doc.setTextColor(90, 110, 136);
      x = ml; sc.forEach(c => { doc.text(c.label, x + 2, y + 4.5); x += c.w; }); y += 8;
    };
    drawHdr();
    rfis.forEach(r => {
      if (y + 7 > ph - 12) { doc.addPage(); y = 10; drawHdr(); }
      const proj = projects.find(p => p.id === r.projectId);
      const vals = [r.rfiNum, proj ? proj.projNum : r.projectId, proj ? proj.name : r.projectName,
        r.rfiType, r.status, fmtD(r.dateIssued), fmtD(r.dateRequired), r.submittedTo || '—',
        r.toCompany || '—', r.response || '—', r.impacted === 'Yes' ? 'Yes' : 'No',
        (r.description || '—').substring(0, 60)];
      doc.setFontSize(7); let x = ml;
      vals.forEach((v, i) => {
        if (i === 0) { doc.setFont('helvetica', 'bold'); doc.setTextColor(37, 99, 235); }
        else { doc.setFont('helvetica', 'normal'); doc.setTextColor(26, 32, 48); }
        const txt = doc.splitTextToSize(String(v), sc[i].w - 3);
        doc.text(txt[0] || '—', x + 2, y + 4); x += sc[i].w;
      });
      doc.setDrawColor(220, 225, 230); doc.line(ml, y + 6, ml + tw, y + 6); y += 7;
    });
    const pages = doc.internal.getNumberOfPages();
    for (let i = 2; i <= pages; i++) { doc.setPage(i); drawPdfBg(doc, pw, ph); }
    doc.save('All_RFIs_' + new Date().toISOString().substring(0, 10) + '.pdf');
  } catch (e) { console.error(e); alert('PDF generation failed.'); }
}

// ── Export All EWOs PDF ──────────────────────────────────────────────────────
function generateAllEwosPdf(ewos: IProject[], projects: IProject[]): void {
  try {
    // eslint-disable-next-line @typescript-eslint/no-explicit-any
    const doc: any = new jsPDF({ orientation: 'portrait', unit: 'mm', format: 'a4' });
    const pw = 210; const ph = 297; const ml = 10; const mr = 10; const tw = pw - ml - mr;
    let y = drawLetterhead(doc, pw, ph, 'EWO LIST', ewos.length + ' EWOs  |  ' + new Date().toLocaleDateString('en-AU'));
    const cols = [
      { label: 'EWO #', w: 30 }, { label: 'PARENT', w: 20 }, { label: 'NAME', w: 35 },
      { label: 'COMPANY', w: 28 }, { label: 'CONTACT', w: 22 }, { label: 'HRS USED', w: 18 },
      { label: 'HRS ALLOWED', w: 22 }, { label: 'START', w: 18 }, { label: 'FINISH', w: 18 },
      { label: 'STATUS', w: 18 }
    ];
    const totalW = cols.reduce((s, c) => s + c.w, 0);
    const scale = tw / totalW;
    const sc = cols.map(c => ({ ...c, w: c.w * scale }));
    const drawHdr = (): void => {
      doc.setFillColor(240, 242, 245); let x = ml;
      sc.forEach(c => { doc.rect(x, y, c.w, 7, 'F'); x += c.w; });
      doc.setDrawColor(208, 213, 222); doc.rect(ml, y, tw, 7, 'S');
      doc.setFontSize(6.5); doc.setFont('helvetica', 'bold'); doc.setTextColor(90, 110, 136);
      x = ml; sc.forEach(c => { doc.text(c.label, x + 2, y + 4.5); x += c.w; }); y += 8;
    };
    drawHdr();
    ewos.forEach(e => {
      if (y + 7 > ph - 12) { doc.addPage(); y = 10; drawHdr(); }
      const parent = projects.find(p => p.id === e.parentId);
      const vals = [e.ewoNum || e.projNum, parent ? parent.projNum : '—', e.name || '—',
        e.company || '—', e.contact || '—', String(e.hrsUsed), String(e.hrsAllowed || '—'),
        fmtD(e.startDate), fmtD(e.finishDate), e.status];
      doc.setFontSize(7); let x = ml;
      vals.forEach((v, i) => {
        if (i === 0) { doc.setFont('helvetica', 'bold'); doc.setTextColor(212, 136, 10); }
        else { doc.setFont('helvetica', 'normal'); doc.setTextColor(26, 32, 48); }
        const txt = doc.splitTextToSize(String(v), sc[i].w - 3);
        doc.text(txt[0] || '—', x + 2, y + 4); x += sc[i].w;
      });
      doc.setDrawColor(220, 225, 230); doc.line(ml, y + 6, ml + tw, y + 6); y += 7;
    });
    const pages = doc.internal.getNumberOfPages();
    for (let i = 2; i <= pages; i++) { doc.setPage(i); drawPdfBg(doc, pw, ph); }
    doc.save('All_EWOs_' + new Date().toISOString().substring(0, 10) + '.pdf');
  } catch (e) { console.error(e); alert('PDF generation failed.'); }
}

// ── Project Form ───────────────────────────────────────────────────────────────
interface ProjFormProps {
  initial: IProject;
  isNew: boolean;
  projects: IProject[];
  spService: SharePointService;
  siteUrl: string;
  onSave: (p: IProject, addedFiles: File[], removedFiles: string[]) => void;
  onCancel: () => void;
}

const ProjForm: React.FC<ProjFormProps> = ({ initial, isNew, projects, spService, siteUrl, onSave, onCancel }) => {
  const [dupError, setDupError] = React.useState('');
  const [valError, setValError] = React.useState('');
  const [pendingFiles, setPendingFiles] = React.useState<File[]>([]);
  const [existingAttachments, setExistingAttachments] = React.useState<{ FileName: string; ServerRelativeUrl: string }[]>([]);
  const [removedFiles, setRemovedFiles] = React.useState<string[]>([]);
  const noteFileRef = React.useRef<HTMLInputElement>(null);
  const projListName = spService.getProjectListName();
  const siteOrigin = siteUrl.replace(/\/sites\/.*/, '');

  React.useEffect(() => {
    if (!isNew && initial.spId) {
      spService.getAttachments(initial.spId, projListName).then(setExistingAttachments).catch(() => undefined);
    }
  }, [isNew, initial.spId, projListName]);

  const isImageName = (name: string): boolean => /\.(png|jpe?g|gif|webp|bmp|svg)$/i.test(name);

  // Auto-calculate next available project number for new projects
  const nextNum = React.useMemo(() => {
    const nums = projects.map(p => {
      const m = p.projNum.match(/^3E-(\d+)$/i);
      return m ? parseInt(m[1], 10) : 0;
    });
    const max = nums.length > 0 ? Math.max(...nums) : 499;
    return String(max + 1);
  }, [projects]);

  const [d, setD] = React.useState<IProject>(() => {
    if (isNew && !initial.projNum) {
      return { ...initial, projNum: '3E-' + nextNum };
    }
    return { ...initial };
  });

  const set = <K extends keyof IProject>(k: K, v: IProject[K]): void => {
    setD(prev => ({ ...prev, [k]: v }));
    if (k === 'projNum') setDupError('');
  };

  const usedNums = React.useMemo(() => {
    const s = new Set(projects.map(p => p.projNum.toUpperCase()));
    // Exclude current project's original number when editing
    if (!isNew && initial.projNum) s.delete(initial.projNum.toUpperCase());
    return s;
  }, [projects, isNew, initial.projNum]);

  const handleSave = (): void => {
    // Validation
    const missing: string[] = [];
    if (!d.projNum || d.projNum === '3E-') missing.push('Project #');
    if (!d.name) missing.push('Project Name');
    if (!d.company) missing.push('Company');
    if (!d.contact) missing.push('Contact');
    if (!d.teamLead) missing.push('Team Lead');
    if (!d.startDate) missing.push('Start Date');
    if (!d.finishDate) missing.push('Finish Date');
    if (!d.hrsAllowed || d.hrsAllowed <= 0) missing.push('Hours Allowed');
    if (d.rfisAllowed === undefined || d.rfisAllowed === null || String(d.rfisAllowed) === '') missing.push('RFIs Allowed');
    if (missing.length > 0) {
      setValError('Required: ' + missing.join(', '));
      return;
    }
    setValError('');
    if (usedNums.has(d.projNum.toUpperCase())) {
      setDupError('Project # ' + d.projNum + ' is already in use. Choose a different number.');
      return;
    }
    onSave(d, pendingFiles, removedFiles);
  };

  return (
    <div>
      <SDiv label={d.isEwo ? 'EWO (Extra Work Order) Details' : 'Project Details'} />
      <div style={{ display: 'grid', gridTemplateColumns: '1fr 1fr', gap: '14px 18px' }}>
        <FF label="Project #">
          <div style={{ display: 'flex', alignItems: 'center', border: '1px solid var(--bd)', borderRadius: 6, overflow: 'hidden', background: 'var(--s1)' }}>
            <span style={{ padding: '0 8px', fontFamily: 'Montserrat', fontWeight: 700, fontSize: 13, color: 'var(--3eg)', background: 'var(--s2)', borderRight: '1px solid var(--bd)', height: '100%', display: 'flex', alignItems: 'center', whiteSpace: 'nowrap' }}>3E-</span>
            <input style={{ ...inp, border: 'none', borderRadius: 0, flex: 1, minWidth: 0 }}
              value={d.projNum.startsWith('3E-') ? d.projNum.slice(3) : d.projNum}
              onChange={e => set('projNum', '3E-' + e.target.value.replace(/^3E-/i, ''))}
              placeholder="500" />
          </div>
        </FF>
        <FF label="Quote #">
          <div style={{ display: 'flex', alignItems: 'center', border: '1px solid var(--bd)', borderRadius: 6, overflow: 'hidden', background: 'var(--s1)' }}>
            <span style={{ padding: '0 8px', fontFamily: 'Montserrat', fontWeight: 700, fontSize: 13, color: 'var(--3eg)', background: 'var(--s2)', borderRight: '1px solid var(--bd)', height: '100%', display: 'flex', alignItems: 'center', whiteSpace: 'nowrap' }}>QU-</span>
            <input style={{ ...inp, border: 'none', borderRadius: 0, flex: 1, minWidth: 0 }}
              value={d.quoteNum.startsWith('QU-') ? d.quoteNum.slice(3) : d.quoteNum}
              onChange={e => set('quoteNum', 'QU-' + e.target.value.replace(/^QU-/i, ''))}
              placeholder="2601" />
          </div>
        </FF>
        <FF label="Project Name">
          <input style={inp} value={d.name} onChange={e => set('name', e.target.value)} placeholder="Project name" />
        </FF>
        <FF label="Discipline">
          <select style={selStyle} value={d.discipline || ''} onChange={e => set('discipline', e.target.value)}>
            <option value="Steel">Steel</option>
            <option value="Concrete">Concrete</option>
            <option value="Steel & Concrete">Steel & Concrete</option>
          </select>
        </FF>
        <FF label="Company">
          <input style={inp} value={d.company} onChange={e => set('company', e.target.value)} />
        </FF>
        <FF label="Contact">
          <input style={inp} value={d.contact} onChange={e => set('contact', e.target.value)} />
        </FF>
        <FF label="Email">
          <input style={inp} type="email" value={d.email} onChange={e => set('email', e.target.value)} />
        </FF>
        <FF label="Mobile">
          <input style={inp} value={d.mobile} onChange={e => set('mobile', e.target.value)} />
        </FF>
        <FF label="Client Ref #">
          <input style={inp} value={d.clientNum} onChange={e => set('clientNum', e.target.value)} />
        </FF>
        <FF label="Status">
          <select style={selStyle} value={d.status} onChange={e => set('status', e.target.value)}>
            {PROJ_STATUSES.map(s => <option key={s} value={s}>{s}</option>)}
          </select>
        </FF>
        <FF label="Client PO#">
          <input style={inp} value={d.clientp0} onChange={e => set('clientp0', e.target.value)} />
        </FF>
        <FF label="Project Address">
          <input style={inp} value={d.projectAddress || ''} onChange={e => set('projectAddress', e.target.value)} />
        </FF>
      </div>

      <SDiv label="Dates" />
      <div style={{ display: 'grid', gridTemplateColumns: '1fr 1fr', gap: '14px 18px' }}>
        <FF label="Start Date">
          <input style={inp} type="date" value={d.startDate} onChange={e => set('startDate', e.target.value)} />
        </FF>
        <FF label="Finish Date">
          <input style={inp} type="date" value={d.finishDate} onChange={e => set('finishDate', e.target.value)} />
        </FF>
        <FF label="IFA Date">
          <input style={inp} type="date" value={d.ifaDate} onChange={e => set('ifaDate', e.target.value)} />
        </FF>
        <FF label="IFC Date">
          <input style={inp} type="date" value={d.ifcDate} onChange={e => set('ifcDate', e.target.value)} />
        </FF>
      </div>

      <SDiv label="Schedule" />
      <div style={{ display: 'grid', gridTemplateColumns: '1fr 1fr', gap: '14px 18px' }}>
        <FF label="Assigned Team Lead">
          <input style={inp} value={d.teamLead} onChange={e => set('teamLead', e.target.value)} placeholder="Team lead name" />
        </FF>
        <FF label="Assigned Team Members">
          <input style={inp} value={d.teamMembers} onChange={e => set('teamMembers', e.target.value)} placeholder="Comma-separated names" />
        </FF>
      </div>

      <SDiv label="Hours & RFIs" />
      <div style={{ display: 'grid', gridTemplateColumns: '1fr 1fr 1fr', gap: '14px 18px' }}>
        <FF label="Hours Allowed">
          <input style={inp} type="number" step="0.5" value={d.hrsAllowed} onChange={e => set('hrsAllowed', Number(e.target.value))} />
        </FF>
        <FF label="Hours Used">
          {/* Read-only: the nightly Time Doctor sync owns this field and recomputes
              it as hrsBaseline + logged hours, so anything typed here is discarded. */}
          <input
            style={{ ...inp, opacity: 0.65, cursor: 'not-allowed' }}
            type="number"
            value={d.hrsUsed}
            readOnly
            title="Synced nightly from Time Doctor — not editable here."
          />
        </FF>
        <FF label="RFIs Allowed" required>
          <input style={inp} type="number" value={d.rfisAllowed} onChange={e => set('rfisAllowed', Number(e.target.value))} />
        </FF>
      </div>

      <SDiv label="Notes" />
      <FF label="Project Notes">
        <textarea style={{ ...inp, minHeight: 170, resize: 'vertical', lineHeight: 1.5 }} value={d.notes} onChange={e => set('notes', e.target.value)} placeholder="Add notes..." />
      </FF>
      <FF label="Note Attachments (images / screenshots)">
        <div>
          <input ref={noteFileRef} type="file" accept="image/*" multiple style={{ display: 'none' }}
            onChange={e => {
              if (e.target.files) {
                setPendingFiles(prev => [...prev, ...Array.from(e.target.files!)]);
                e.target.value = '';
              }
            }} />
          <button type="button" onClick={() => noteFileRef.current?.click()}
            style={{ ...inp, cursor: 'pointer', background: 'var(--s2)', border: '1px dashed var(--bd)', padding: '8px 12px', fontSize: 12, color: 'var(--t3)', width: '100%', textAlign: 'left' }}>
            + Click to attach pictures or screenshots...
          </button>
          {(existingAttachments.length > 0 || pendingFiles.length > 0) && (
            <div style={{ marginTop: 10, display: 'flex', flexWrap: 'wrap', gap: 10 }}>
              {existingAttachments.map((f, i) => {
                const url = siteOrigin + f.ServerRelativeUrl;
                const isImg = isImageName(f.FileName);
                return (
                  <div key={'ex' + i} style={{ position: 'relative', border: '1px solid var(--bd)', borderRadius: 4, background: 'var(--s2)', padding: 4, width: isImg ? 96 : 'auto', maxWidth: 220 }}>
                    {isImg ? (
                      <a href={url} target="_blank" rel="noopener noreferrer">
                        <img src={url} alt={f.FileName} style={{ width: 88, height: 88, objectFit: 'cover', borderRadius: 2, display: 'block' }} />
                      </a>
                    ) : (
                      <a href={url} target="_blank" rel="noopener noreferrer" style={{ display: 'block', padding: '6px 10px', fontSize: 11.5, color: 'var(--3eg)', fontFamily: 'Montserrat', textDecoration: 'none' }}>
                        {f.FileName}
                      </a>
                    )}
                    <button type="button" onClick={() => { setRemovedFiles(prev => [...prev, f.FileName]); setExistingAttachments(prev => prev.filter((_, j) => j !== i)); }}
                      title="Remove"
                      style={{ position: 'absolute', top: 2, right: 2, background: 'rgba(232,69,69,0.9)', border: 'none', color: '#fff', borderRadius: 3, width: 18, height: 18, fontSize: 12, lineHeight: 1, cursor: 'pointer', fontWeight: 700 }}>
                      ×
                    </button>
                  </div>
                );
              })}
              {pendingFiles.map((f, i) => {
                const objUrl = URL.createObjectURL(f);
                const isImg = f.type.startsWith('image/') || isImageName(f.name);
                return (
                  <div key={'pd' + i} style={{ position: 'relative', border: '1px solid var(--3eg)', borderRadius: 4, background: 'var(--3eg3)', padding: 4, width: isImg ? 96 : 'auto', maxWidth: 220 }}>
                    {isImg ? (
                      <img src={objUrl} alt={f.name} style={{ width: 88, height: 88, objectFit: 'cover', borderRadius: 2, display: 'block' }} />
                    ) : (
                      <span style={{ display: 'block', padding: '6px 10px', fontSize: 11.5, color: 'var(--3eg)', fontFamily: 'Montserrat' }}>{f.name}</span>
                    )}
                    <button type="button" onClick={() => setPendingFiles(prev => prev.filter((_, j) => j !== i))}
                      title="Remove"
                      style={{ position: 'absolute', top: 2, right: 2, background: 'rgba(232,69,69,0.9)', border: 'none', color: '#fff', borderRadius: 3, width: 18, height: 18, fontSize: 12, lineHeight: 1, cursor: 'pointer', fontWeight: 700 }}>
                      ×
                    </button>
                  </div>
                );
              })}
            </div>
          )}
        </div>
      </FF>

      <SDiv label="Invoices" />
      {(d.invoices.length > 0 ? d.invoices : []).map((inv, idx) => (
        <div key={idx} style={{ display: 'grid', gridTemplateColumns: '1fr 70px 1fr 1fr auto auto', gap: '8px', alignItems: 'end', marginBottom: 8 }}>
          <FF label={idx === 0 ? 'Invoice Number' : ''}>
            <input style={inp} value={inv.invNumber} onChange={e => { const invs = [...d.invoices]; invs[idx] = { ...invs[idx], invNumber: e.target.value }; set('invoices', invs); }} placeholder="INV-001" />
          </FF>
          <FF label={idx === 0 ? 'Pct (%)' : ''}>
            <input style={inp} type="number" min={0} max={100} value={inv.invPct ?? 0} onChange={e => { const invs = [...d.invoices]; invs[idx] = { ...invs[idx], invPct: Number(e.target.value) }; set('invoices', invs); }} placeholder="0" />
          </FF>
          <FF label={idx === 0 ? 'Progress Claim' : ''}>
            <input style={inp} value={inv.invProgressClaim ?? ''} onChange={e => { const invs = [...d.invoices]; invs[idx] = { ...invs[idx], invProgressClaim: e.target.value }; set('invoices', invs); }} placeholder="e.g. Claim 1" />
          </FF>
          <FF label={idx === 0 ? 'Invoice Date' : ''}>
            <input style={inp} type="date" value={inv.invDate} onChange={e => { const invs = [...d.invoices]; invs[idx] = { ...invs[idx], invDate: e.target.value }; set('invoices', invs); }} />
          </FF>
          <label style={{ display: 'flex', alignItems: 'center', gap: 6, fontFamily: 'Montserrat', fontSize: 11, fontWeight: 600, color: 'var(--t2)', cursor: 'pointer', paddingBottom: 2 }}>
            <input type="checkbox" checked={!!inv.invPaid} onChange={e => { const invs = [...d.invoices]; invs[idx] = { ...invs[idx], invPaid: e.target.checked }; set('invoices', invs); }} style={{ width: 15, height: 15, cursor: 'pointer' }} />
            Paid
          </label>
          <button onClick={() => { const invs = d.invoices.filter((_, i) => i !== idx); set('invoices', invs); }} style={{ background: 'transparent', border: '1px solid var(--rd)', color: 'var(--rd)', borderRadius: 4, width: 26, height: 26, fontSize: 13, cursor: 'pointer', fontWeight: 700 }}>×</button>
        </div>
      ))}
      {d.invoices.length < 10 && (
        <button onClick={() => set('invoices', [...d.invoices, { invNumber: '', invPct: 0, invProgressClaim: '', invDate: '', invPaid: false }])} style={{ fontFamily: 'Montserrat', fontSize: 11, fontWeight: 600, padding: '6px 14px', background: 'transparent', border: '1px dashed var(--bd)', color: 'var(--t3)', borderRadius: 5, cursor: 'pointer', marginTop: 4 }}>+ Add Invoice</button>
      )}

      {valError && <div style={{ color: 'var(--rd)', fontFamily: 'Montserrat', fontSize: 12.5, marginTop: 12, fontWeight: 600 }}>{valError}</div>}
      {dupError && <div style={{ color: 'var(--am)', fontFamily: 'Montserrat', fontSize: 12.5, marginTop: 12, fontWeight: 600 }}>{dupError}</div>}
      <div style={{ display: 'flex', gap: 10, marginTop: 28, paddingTop: 16, borderTop: '1px solid var(--bd)' }}>
        <BtnPrimary onClick={handleSave}>{isNew ? 'CREATE PROJECT' : 'SAVE CHANGES'}</BtnPrimary>
        <button onClick={onCancel} style={{ fontFamily: 'Montserrat', fontSize: 12.5, padding: '9px 18px', background: 'transparent', border: '1px solid var(--bd)', color: 'var(--t2)', borderRadius: 7, cursor: 'pointer' }}>Cancel</button>
      </div>
    </div>
  );
};

// ── EWO Form ──────────────────────────────────────────────────────────────────
interface EwoFormProps {
  initial: IProject;
  isNew: boolean;
  projects: IProject[];
  onSave: (p: IProject) => void;
  onCancel: () => void;
}

const EwoForm: React.FC<EwoFormProps> = ({ initial, isNew, projects, onSave, onCancel }) => {
  const [ewoValError, setEwoValError] = React.useState('');
  const parentProjects = projects.filter(p => !p.isEwo);
  const allEwos = projects.filter(p => p.isEwo);

  const [d, setD] = React.useState<IProject>(() => {
    return { ...initial, isEwo: true };
  });

  const set = <K extends keyof IProject>(k: K, v: IProject[K]): void => {
    setD(prev => ({ ...prev, [k]: v }));
  };

  const onParentChange = (parentId: string): void => {
    const parent = parentProjects.find(p => p.id === parentId);
    const updates: Partial<IProject> = { parentId: parentId || null };
    if (isNew && parent) {
      const count = allEwos.filter(e => e.parentId === parentId).length;
      const seq = String(count + 1).padStart(3, '0');
      updates.projNum = parent.projNum + '-EWO-' + seq;
      updates.ewoNum = parent.projNum + '-EWO-' + seq;
      // Inherit company details from parent project
      updates.company = parent.company;
      updates.contact = parent.contact;
      updates.email = parent.email;
      updates.mobile = parent.mobile;
      updates.clientNum = parent.clientNum;
    }
    setD(prev => ({ ...prev, ...updates }));
  };

  const inp: React.CSSProperties = { fontFamily: 'Montserrat', fontSize: 13, padding: '8px 10px', border: '1px solid var(--bd)', borderRadius: 6, background: 'var(--s1)', color: 'var(--t1)', width: '100%', boxSizing: 'border-box' };
  const selStyle: React.CSSProperties = { ...inp, appearance: 'auto' as React.CSSProperties['appearance'] };

  return (
    <div>
      <SDiv label="EWO Details" />
      <div style={{ display: 'grid', gridTemplateColumns: '1fr 1fr', gap: '14px 18px' }}>
        <FF label="Parent Project" span2>
          <select style={selStyle} value={d.parentId || ''} onChange={e => onParentChange(e.target.value)}>
            <option value="">— Select parent project —</option>
            {parentProjects.map(p => (
              <option key={p.id} value={p.id}>{p.projNum} — {p.name}</option>
            ))}
          </select>
        </FF>
        <FF label="EWO Number">
          <input style={{ ...inp, background: 'var(--s2)', color: 'var(--t3)' }} value={d.ewoNum || d.projNum} readOnly />
        </FF>
        <FF label="Quote #">
          <div style={{ display: 'flex', alignItems: 'center', border: '1px solid var(--bd)', borderRadius: 6, overflow: 'hidden', background: 'var(--s1)' }}>
            <span style={{ padding: '0 8px', fontFamily: 'Montserrat', fontWeight: 700, fontSize: 13, color: 'var(--3eg)', background: 'var(--s2)', borderRight: '1px solid var(--bd)', height: '100%', display: 'flex', alignItems: 'center', whiteSpace: 'nowrap' }}>QU-</span>
            <input style={{ ...inp, border: 'none', borderRadius: 0, flex: 1, minWidth: 0 }}
              value={d.quoteNum.startsWith('QU-') ? d.quoteNum.slice(3) : d.quoteNum}
              onChange={e => set('quoteNum', 'QU-' + e.target.value.replace(/^QU-/i, ''))} />
          </div>
        </FF>
        <FF label="Project Name">
          <input style={inp} value={d.name} onChange={e => set('name', e.target.value)} placeholder="EWO name" />
        </FF>
        <FF label="Discipline">
          <select style={selStyle} value={d.discipline || ''} onChange={e => set('discipline', e.target.value)}>
            <option value="Steel">Steel</option>
            <option value="Concrete">Concrete</option>
            <option value="Steel & Concrete">Steel & Concrete</option>
          </select>
        </FF>
        <FF label="Company">
          <input style={inp} value={d.company} onChange={e => set('company', e.target.value)} />
        </FF>
        <FF label="Contact">
          <input style={inp} value={d.contact} onChange={e => set('contact', e.target.value)} />
        </FF>
        <FF label="Email">
          <input style={inp} type="email" value={d.email} onChange={e => set('email', e.target.value)} />
        </FF>
        <FF label="Mobile">
          <input style={inp} value={d.mobile} onChange={e => set('mobile', e.target.value)} />
        </FF>
        <FF label="Client Ref #">
          <input style={inp} value={d.clientNum} onChange={e => set('clientNum', e.target.value)} />
        </FF>
        <FF label="Status">
          <select style={selStyle} value={d.status} onChange={e => set('status', e.target.value)}>
            {PROJ_STATUSES.map(s => <option key={s} value={s}>{s}</option>)}
          </select>
        </FF>
        <FF label="Project Address">
          <input style={inp} value={d.projectAddress || ''} onChange={e => set('projectAddress', e.target.value)} />
        </FF>
      </div>

      <SDiv label="Dates" />
      <div style={{ display: 'grid', gridTemplateColumns: '1fr 1fr', gap: '14px 18px' }}>
        <FF label="Start Date">
          <input style={inp} type="date" value={d.startDate} onChange={e => set('startDate', e.target.value)} />
        </FF>
        <FF label="Finish Date">
          <input style={inp} type="date" value={d.finishDate} onChange={e => set('finishDate', e.target.value)} />
        </FF>
      </div>

      <SDiv label="Notes" />
      <div style={{ display: 'grid', gridTemplateColumns: '1fr', gap: '14px' }}>
        <FF label="Notes">
          <textarea style={{ ...inp, minHeight: 60 }} value={d.notes} onChange={e => set('notes', e.target.value)} placeholder="Additional notes..." />
        </FF>
      </div>

      <SDiv label="Hours" />
      <div style={{ display: 'grid', gridTemplateColumns: '1fr 1fr 1fr', gap: '14px 18px' }}>
        <FF label="Hours Allowed">
          <input style={inp} type="number" value={d.hrsAllowed} onChange={e => set('hrsAllowed', Number(e.target.value))} />
        </FF>
        <FF label="Hours Used">
          {/* Editable, unlike a project's. An EWO's projNum is "3E-531-EWO-001",
              which never equals the "3e-531" code the sync matches on, and the sync
              skips isEwo rows anyway — so hand entry is the only source EWO hours
              have. */}
          <input
            style={inp}
            type="number"
            value={d.hrsUsed}
            onChange={e => set('hrsUsed', Number(e.target.value))}
            title="EWO hours are entered by hand — the nightly Time Doctor sync does not track EWOs."
          />
        </FF>
        <FF label="Detailers">
          <input style={inp} value={d.detailers} onChange={e => set('detailers', e.target.value)} placeholder="Comma-separated" />
        </FF>
      </div>

      {ewoValError && <div style={{ color: 'var(--rd)', fontFamily: 'Montserrat', fontSize: 12.5, marginTop: 12, fontWeight: 600 }}>{ewoValError}</div>}
      <div style={{ display: 'flex', gap: 10, marginTop: 28, paddingTop: 16, borderTop: '1px solid var(--bd)' }}>
        <BtnPrimary onClick={() => {
          const missing: string[] = [];
          if (!d.parentId) missing.push('Parent Project');
          if (!d.name) missing.push('Project Name');
          if (!d.company) missing.push('Company');
          if (!d.contact) missing.push('Contact');
          if (missing.length > 0) { setEwoValError('Required: ' + missing.join(', ')); return; }
          setEwoValError('');
          onSave(d);
        }}>{isNew ? 'CREATE EWO' : 'SAVE CHANGES'}</BtnPrimary>
        <button onClick={onCancel} style={{ fontFamily: 'Montserrat', fontSize: 12.5, padding: '9px 18px', background: 'transparent', border: '1px solid var(--bd)', color: 'var(--t2)', borderRadius: 7, cursor: 'pointer' }}>Cancel</button>
      </div>
    </div>
  );
};

// ── Project Detail ─────────────────────────────────────────────────────────────
interface ProjDetailProps {
  proj: IProject;
  rfis: IRfi[];
  isManager: boolean;
  onEdit: () => void;
  onDelete: () => void;
  onNewRfi: () => void;
  onViewRfi: (r: IRfi) => void;
}

const ProjDetail: React.FC<ProjDetailProps> = ({ proj, rfis, isManager, onEdit, onDelete, onNewRfi, onViewRfi }) => {
  const projRfis = rfis.filter(r => r.projectId === proj.id);
  const open = projRfis.filter(r => effSt(r) === 'Open' || effSt(r) === 'Partially Open (Revise and Resend)').length;
  const overdue = projRfis.filter(r => isOD(r)).length;

  const rowItem = (label: string, value: string | number | boolean | null | undefined, highlight?: boolean): JSX.Element => {
    const v = (value === null || value === undefined || value === '') ? '—' : String(value);
    return (
      <div style={{ display: 'flex', padding: '9px 0', borderBottom: '1px solid var(--bd)', gap: 12 }}>
        <span style={{ fontFamily: 'Montserrat', fontWeight: 600, fontSize: 11.5, color: 'var(--t4)', textTransform: 'uppercase', letterSpacing: '.07em', minWidth: 130, flexShrink: 0 }}>{label}</span>
        <span style={{ fontFamily: 'Montserrat', fontWeight: highlight ? 700 : 500, fontSize: 13, color: highlight ? 'var(--3eg)' : 'var(--t1)' }}>{v}</span>
      </div>
    );
  };

  return (
    <div>
      <div style={{ display: 'flex', gap: 10, marginBottom: 18, flexWrap: 'wrap' }}>
        {isManager && <IBtn onClick={onEdit} title="Edit project">Edit</IBtn>}
        {isManager && <IBtn onClick={onNewRfi} title="Create RFI for this project">+ New RFI</IBtn>}
        {isManager && <IBtn onClick={onDelete} danger title="Delete project">Delete</IBtn>}
      </div>

      <SDiv label="Overview" />
      {rowItem('Project #', proj.projNum, true)}
      {rowItem('Quote #', proj.quoteNum)}
      {rowItem('Name', proj.name)}
      {rowItem('Address', proj.projectAddress || '—')}
      {rowItem('Company', proj.company)}
      {rowItem('Contact', proj.contact)}
      {rowItem('Email', proj.email)}
      {rowItem('Mobile', proj.mobile)}
      {rowItem('Client Ref', proj.clientNum)}
      {rowItem('Status', proj.status)}
      {rowItem('Year', proj.year)}
      {rowItem('Detailers', proj.detailers)}

      <SDiv label="Dates" />
      {rowItem('Start Date', fmtD(proj.startDate))}
      {rowItem('Finish Date', fmtD(proj.finishDate))}
      {rowItem('IFA Date', fmtD(proj.ifaDate))}
      {rowItem('IFC Date', fmtD(proj.ifcDate))}

      <SDiv label="Hours" />
      <div style={{ marginBottom: 12 }}>
        <HrsBar allowed={proj.hrsAllowed} used={proj.hrsUsed} />
      </div>
      {rowItem('Hours Allowed', proj.hrsAllowed)}
      {rowItem('Hours Used', proj.hrsUsed)}

      {proj.isEwo && (
        <React.Fragment>
          <SDiv label="EWO Details" />
          {rowItem('EWO Number', proj.ewoNum)}
          {rowItem('Parent Project', proj.parentId || '—')}
        </React.Fragment>
      )}

      <SDiv label={'RFIs (' + projRfis.length + ')'} />
      {projRfis.length === 0
        ? <div style={{ fontFamily: 'Montserrat', fontSize: 13, color: 'var(--t4)', padding: '10px 0' }}>No RFIs for this project.</div>
        : (
          <div>
            <div style={{ display: 'flex', gap: 16, marginBottom: 12 }}>
              <span style={{ fontFamily: 'Montserrat', fontSize: 12, fontWeight: 600, color: 'var(--t3)' }}>
                Total: {projRfis.length} &nbsp;|&nbsp;
                Open: <span style={{ color: open > 0 ? 'var(--am)' : 'var(--t3)' }}>{open}</span> &nbsp;|&nbsp;
                Overdue: <span style={{ color: overdue > 0 ? 'var(--rd)' : 'var(--t3)' }}>{overdue}</span>
              </span>
            </div>
            {projRfis.map(r => (
              <div key={r.id} onClick={() => onViewRfi(r)}
                style={{ display: 'flex', alignItems: 'center', gap: 12, padding: '9px 12px', borderRadius: 6, background: 'var(--s2)', marginBottom: 6, cursor: 'pointer', border: '1px solid var(--bd)' }}>
                <span style={{ fontFamily: 'Montserrat', fontWeight: 700, fontSize: 12.5, color: 'var(--t1)', minWidth: 80 }}>{r.rfiNum}</span>
                <span style={{ fontFamily: 'Montserrat', fontSize: 12, color: 'var(--t3)', flex: 1, overflow: 'hidden', textOverflow: 'ellipsis', whiteSpace: 'nowrap' }}>{r.rfiType}</span>
                <Tag s={effSt(r)} />
              </div>
            ))}
          </div>
        )}
    </div>
  );
};

// ── RFI Form ───────────────────────────────────────────────────────────────────
interface RfiFormProps {
  initial: IRfi;
  isNew: boolean;
  projects: IProject[];
  rfis: IRfi[];
  userDisplayName: string;
  teamMembers: ITeamMember[];
  onSave: (r: IRfi, files: File[]) => void;
  onCancel: () => void;
}

const RfiForm: React.FC<RfiFormProps> = ({ initial, isNew, projects, rfis, userDisplayName, teamMembers, onSave, onCancel }) => {
  const [d, setD] = React.useState<IRfi>(() => {
    let rfi = { ...initial };
    if (isNew && initial.projectId) {
      rfi = applyProjectDefaultsById(rfi, initial.projectId, projects, rfis, { byCompany: RFI_DEFAULT_BY_COMPANY });
    }
    if (isNew) {
      rfi = applySenderDefaultsToRfi(rfi, userDisplayName, teamMembers);
    }
    return rfi;
  });
  const [rememberSender, setRememberSender] = React.useState(true);
  const [rfiValError, setRfiValError] = React.useState('');
  const [pendingFiles, setPendingFiles] = React.useState<File[]>([]);
  const fileRef = React.useRef<HTMLInputElement>(null);

  const set = <K extends keyof IRfi>(k: K, v: IRfi[K]): void => {
    setD(prev => ({ ...prev, [k]: v }));
  };

  const onProjectChange = (projId: string): void => {
    if (isNew && projId) {
      setD(prev => applySenderDefaultsToRfi(
        applyProjectDefaultsById(prev, projId, projects, rfis, { byCompany: RFI_DEFAULT_BY_COMPANY }),
        userDisplayName,
        teamMembers
      ));
      return;
    }
    const p = projects.find(x => x.id === projId);
    setD(prev => ({ ...prev, projectId: projId, projectName: p ? p.name : '' }));
  };

  const totalImpact = (d.model || 0) + (d.connections || 0) + (d.checking || 0) + (d.drawings || 0) + (d.admin || 0);

  return (
    <div>
      {/* Part A */}
      <SDiv label="Part A — Request Information" />
      <div style={{ display: 'grid', gridTemplateColumns: '1fr 1fr', gap: '14px 18px' }}>
        <FF label="Project" span2>
          <select style={selStyle} value={d.projectId} onChange={e => onProjectChange(e.target.value)}>
            <option value="">— Select project —</option>
            {projects.map(p => (
              <option key={p.id} value={p.id}>{p.projNum} — {p.name}</option>
            ))}
          </select>
        </FF>
        <FF label="RFI Number">
          <input style={inp} value={d.rfiNum} onChange={e => set('rfiNum', e.target.value)} placeholder="e.g. 2601-RFI-001" />
        </FF>
        <FF label="RFI Type">
          <select style={selStyle} value={d.rfiType} onChange={e => set('rfiType', e.target.value)}>
            {RFI_TYPES.map(t => <option key={t} value={t}>{t}</option>)}
          </select>
        </FF>
        <FF label="Revision">
          <input style={inp} value={d.revision || 'A'} onChange={e => set('revision', e.target.value)} placeholder="A" />
        </FF>
        <FF label="Status">
          <select style={selStyle} value={d.status} onChange={e => set('status', e.target.value)}>
            {RFI_STATUSES.map(s => <option key={s} value={s}>{s}</option>)}
          </select>
        </FF>
        <FF label="Date Issued">
          <input style={inp} type="date" value={d.dateIssued} onChange={e => set('dateIssued', e.target.value)} />
        </FF>
        <FF label="Date Required">
          <input style={inp} type="date" value={d.dateRequired} onChange={e => set('dateRequired', e.target.value)} />
        </FF>
        <FF label="Submitted To">
          <input style={inp} value={d.submittedTo} onChange={e => set('submittedTo', e.target.value)} />
        </FF>
        <FF label="To Company">
          <input style={inp} value={d.toCompany} onChange={e => set('toCompany', e.target.value)} />
        </FF>
        <FF label="Prepared By">
          <input style={inp} value={d.by} onChange={e => set('by', e.target.value)} />
        </FF>
        <FF label="By Company">
          <input style={inp} value={d.byCompany} onChange={e => set('byCompany', e.target.value)} />
        </FF>
        {isNew && (
          <RememberSenderField checked={rememberSender} onChange={setRememberSender} />
        )}
        <FF label="Email">
          <input style={inp} type="email" value={d.email || ''} onChange={e => set('email', e.target.value)} />
        </FF>
        <FF label="CC">
          <CcField value={d.cc} onChange={v => set('cc', v)} compact />
        </FF>
      </div>

      {/* Part B */}
      <SDiv label="Part B — Description" />
      <div style={{ display: 'grid', gridTemplateColumns: '1fr', gap: '14px' }}>
        <FF label="Description">
          <textarea style={{ ...inp, minHeight: 100 }} value={d.description} onChange={e => set('description', e.target.value)} />
        </FF>
        <FF label="Attachments">
          <div>
            <input ref={fileRef} type="file" multiple style={{ display: 'none' }}
              onChange={e => {
                if (e.target.files) {
                  setPendingFiles(prev => [...prev, ...Array.from(e.target.files!)]);
                  e.target.value = '';
                }
              }} />
            <button type="button" onClick={() => fileRef.current?.click()}
              style={{ ...inp, cursor: 'pointer', background: 'var(--s2)', border: '1px dashed var(--bd)', padding: '8px 12px', fontSize: 12, color: 'var(--t3)', width: '100%', textAlign: 'left' }}>
              + Click to attach files...
            </button>
            {pendingFiles.length > 0 && (
              <div style={{ marginTop: 6, display: 'flex', flexWrap: 'wrap', gap: 6 }}>
                {pendingFiles.map((f, i) => (
                  <span key={i} style={{
                    display: 'inline-flex', alignItems: 'center', gap: 4,
                    background: 'var(--3eg3)', border: '1px solid var(--3eg)', borderRadius: 2,
                    padding: '2px 6px 2px 8px', fontSize: 11.5, color: 'var(--3eg)', fontFamily: 'Montserrat'
                  }}>
                    {f.name}
                    <button type="button" onClick={() => setPendingFiles(prev => prev.filter((_, j) => j !== i))}
                      style={{ background: 'none', border: 'none', cursor: 'pointer', color: 'var(--am)', fontSize: 14, padding: 0, lineHeight: 1 }}>
                      &times;
                    </button>
                  </span>
                ))}
              </div>
            )}
          </div>
        </FF>
      </div>

      {/* Parts C & D */}
      <SDiv label="Parts C & D — Client Response" />
      <div style={{ display: 'grid', gridTemplateColumns: '1fr 1fr', gap: '14px 18px' }}>
        <FF label="Client RFI #">
          <input style={inp} value={d.clientRfi} onChange={e => set('clientRfi', e.target.value)} />
        </FF>
        <FF label="Date Received">
          <input style={inp} type="date" value={d.dateReceived} onChange={e => set('dateReceived', e.target.value)} />
        </FF>
        <FF label="Response">
          <select style={selStyle} value={d.response} onChange={e => set('response', e.target.value)}>
            {RFI_RESPONSES.map(r => <option key={r} value={r}>{r}</option>)}
          </select>
        </FF>
        <FF label="RFI Status">
          <select style={selStyle} value={d.status} onChange={e => set('status', e.target.value)}>
            {RFI_STATUSES.map(s => <option key={s} value={s}>{s}</option>)}
          </select>
        </FF>
        <FF label="Sent By">
          <input style={inp} value={d.sentBy} onChange={e => set('sentBy', e.target.value)} />
        </FF>
        <FF label="Sent By Company">
          <input style={inp} value={d.sentByCompany} onChange={e => set('sentByCompany', e.target.value)} />
        </FF>
        <FF label="Response Description" span2>
          <textarea style={{ ...inp, minHeight: 80 }} value={d.responseDesc || ''} onChange={e => set('responseDesc', e.target.value)} />
        </FF>
      </div>

      {/* Part E */}
      <SDiv label="Part E — Impact Assessment" />
      <div style={{ display: 'grid', gridTemplateColumns: '1fr 1fr', gap: '14px 18px' }}>
        <FF label="Schedule Impacted?">
          <select style={selStyle} value={d.impacted} onChange={e => set('impacted', e.target.value)}>
            <option value="No">No</option>
            <option value="Yes">Yes</option>
          </select>
        </FF>
        <FF label="EWO Reference">
          <input style={inp} value={d.ewoRef || ''} onChange={e => set('ewoRef', e.target.value)} />
        </FF>
        {d.impacted === 'Yes' && (
          <React.Fragment>
            <FF label="Model Hours">
              <input style={inp} type="number" step="0.5" value={d.model} onChange={e => set('model', Number(e.target.value))} />
            </FF>
            <FF label="Connections Hours">
              <input style={inp} type="number" step="0.5" value={d.connections} onChange={e => set('connections', Number(e.target.value))} />
            </FF>
            <FF label="Checking Hours">
              <input style={inp} type="number" step="0.5" value={d.checking} onChange={e => set('checking', Number(e.target.value))} />
            </FF>
            <FF label="Drawings Hours">
              <input style={inp} type="number" step="0.5" value={d.drawings} onChange={e => set('drawings', Number(e.target.value))} />
            </FF>
            <FF label="Admin Hours">
              <input style={inp} type="number" step="0.5" value={d.admin} onChange={e => set('admin', Number(e.target.value))} />
            </FF>
            <FF label="Total Impact Hours">
              <div style={{ padding: '9px 12px', background: 'var(--s1)', border: '1px solid var(--bd)', borderRadius: 2, fontFamily: 'Montserrat', fontWeight: 700, fontSize: 14, color: 'var(--3eg)' }}>
                {totalImpact.toFixed(1)}h
              </div>
            </FF>
          </React.Fragment>
        )}
      </div>

      {rfiValError && <div style={{ color: 'var(--rd)', fontFamily: 'Montserrat', fontSize: 12.5, marginTop: 12, fontWeight: 600 }}>{rfiValError}</div>}
      <div style={{ display: 'flex', gap: 10, marginTop: 28, paddingTop: 16, borderTop: '1px solid var(--bd)' }}>
        <BtnPrimary onClick={() => {
          const missing: string[] = [];
          if (!d.projectId) missing.push('Project');
          if (!d.rfiNum) missing.push('RFI Number');
          if (!d.rfiType) missing.push('RFI Type');
          if (!d.submittedTo) missing.push('Submitted To');
          if (!d.description) missing.push('Description');
          if (missing.length > 0) { setRfiValError('Required: ' + missing.join(', ')); return; }
          setRfiValError('');
          if (isNew && rememberSender && userDisplayName.trim()) {
            saveSenderDefaults(userDisplayName.trim(), { by: d.by, byCompany: d.byCompany });
          }
          onSave(d, pendingFiles);
        }}>{isNew ? 'CREATE RFI' : 'SAVE CHANGES'}</BtnPrimary>
        <button onClick={onCancel} style={{ fontFamily: 'Montserrat', fontSize: 12.5, padding: '9px 18px', background: 'transparent', border: '1px solid var(--bd)', color: 'var(--t2)', borderRadius: 7, cursor: 'pointer' }}>Cancel</button>
      </div>
    </div>
  );
};

// ── RFI Detail ─────────────────────────────────────────────────────────────────
interface RfiDetailProps {
  rfi: IRfi;
  proj: IProject | undefined;
  isManager: boolean;
  siteUrl: string;
  spService: SharePointService;
  onSendEmail: (to: string, cc: string, subject: string, body: string, pdfFileName: string) => Promise<void>;
  onEdit: () => void;
  onNotify?: (message: string) => void;
}

const RfiDetail: React.FC<RfiDetailProps> = ({ rfi, proj, isManager, siteUrl, spService, onSendEmail, onEdit, onNotify }) => {
  const total = rfiTot(rfi);
  const st = effSt(rfi);
  const [attachFiles, setAttachFiles] = React.useState<{ FileName: string; ServerRelativeUrl: string }[]>([]);

  React.useEffect(() => {
    if (rfi.spId) {
      spService.getAttachments(rfi.spId).then(setAttachFiles).catch(() => undefined);
    }
  }, [rfi.spId]);

  const row = (label: string, value: string | number | boolean | null | undefined, highlight?: boolean): JSX.Element => {
    const v = (value === null || value === undefined || value === '') ? '—' : String(value);
    return (
      <div style={{ display: 'flex', padding: '9px 0', borderBottom: '1px solid var(--bd)', gap: 12 }}>
        <span style={{ fontFamily: 'Montserrat', fontWeight: 600, fontSize: 11.5, color: 'var(--t4)', textTransform: 'uppercase', letterSpacing: '.07em', minWidth: 130, flexShrink: 0 }}>{label}</span>
        <span style={{ fontFamily: 'Montserrat', fontWeight: highlight ? 700 : 500, fontSize: 13, color: highlight ? 'var(--3eg)' : 'var(--t1)', wordBreak: 'break-word' }}>{v}</span>
      </div>
    );
  };

  const handleSendToClient = (): void => {
    const blob = generateRfiPdf(rfi, proj);
    if (!blob) return;

    const fileName = rfiPdfFileName(rfi.rfiNum);
    const url = URL.createObjectURL(blob);
    const a = document.createElement('a');
    a.href = url;
    a.download = fileName;
    a.click();
    setTimeout(() => URL.revokeObjectURL(url), 3000);

    const recipients = rfi.email || '';
    if (!recipients) {
      onNotify?.('PDF downloaded. Add a recipient email on the RFI to open an email draft.');
      return;
    }
    const subject = 'RFI ' + rfi.rfiNum + ' — ' + (proj ? proj.name : rfi.projectName);
    const company = rfi.byCompany || RFI_DEFAULT_BY_COMPANY;
    const body =
      'Dear ' + (rfi.submittedTo || 'Client') + ',<br><br>' +
      'Please find attached RFI <strong>' + rfi.rfiNum + '</strong> for your review and response.<br><br>' +
      '<strong>Project:</strong> ' + (proj ? proj.name : rfi.projectName) + '<br>' +
      '<strong>RFI Type:</strong> ' + rfi.rfiType + '<br>' +
      '<strong>Date Issued:</strong> ' + fmtD(rfi.dateIssued) + '<br>' +
      '<strong>Date Required:</strong> ' + fmtD(rfi.dateRequired) + '<br><br>' +
      '<strong>Description:</strong><br>' + (rfi.description || '—') + '<br><br>' +
      'Please respond by ' + fmtD(rfi.dateRequired) + '.<br><br>' +
      'Kind regards,<br>' + (rfi.by || '') + '<br>' + company + '<br><br>' +
      '<em>Please attach the downloaded PDF file (' + fileName + ') before sending. You may add further attachments as needed.</em>';
    onSendEmail(recipients, rfi.cc || '', subject, body, fileName).catch(console.error);
  };

  return (
    <div>
      <div style={{ display: 'flex', gap: 10, marginBottom: 18, flexWrap: 'wrap' }}>
        {isManager && <IBtn onClick={onEdit} title="Edit RFI">Edit</IBtn>}
        <button onClick={handleSendToClient} style={{
          fontFamily: 'Montserrat', fontWeight: 700, fontSize: 12.5, letterSpacing: '.06em',
          textTransform: 'uppercase', padding: '5px 14px', borderRadius: 5, cursor: 'pointer',
          background: 'var(--3eg3)', border: '1px solid var(--3eg)', color: 'var(--3eg)',
          display: 'flex', alignItems: 'center', gap: 6
        }}>
          Send to Client
        </button>
      </div>

      <SDiv label="Part A — Request Information" />
      {row('Project', proj ? (proj.projNum + ' — ' + proj.name) : rfi.projectName, true)}
      {row('RFI Number', rfi.rfiNum, true)}
      {row('RFI Type', rfi.rfiType)}
      {row('Revision', rfi.revision || 'A')}
      {row('Status', st)}
      {row('Date Issued', fmtD(rfi.dateIssued))}
      {row('Date Required', fmtD(rfi.dateRequired))}
      {row('Submitted To', rfi.submittedTo)}
      {row('To Company', rfi.toCompany)}
      {row('Prepared By', rfi.by)}
      {row('By Company', rfi.byCompany)}
      {rfi.email ? row('Email', rfi.email) : null}
      {rfi.cc ? row('CC', rfi.cc) : null}
      {rfi.emailSentDate ? row('Email Sent', fmtD(rfi.emailSentDate)) : null}

      <SDiv label="Part B — Description" />
      <div style={{ padding: '12px 0', borderBottom: '1px solid var(--bd)' }}>
        <div style={{ fontFamily: 'Montserrat', fontWeight: 600, fontSize: 11.5, color: 'var(--t4)', textTransform: 'uppercase', letterSpacing: '.07em', marginBottom: 8 }}>Description</div>
        <div style={{ fontFamily: 'Montserrat', fontSize: 13, color: 'var(--t1)', lineHeight: 1.7, whiteSpace: 'pre-wrap' }}>{rfi.description || '—'}</div>
      </div>
      {attachFiles.length > 0 && (
        <div style={{ padding: '12px 0', borderBottom: '1px solid var(--bd)' }}>
          <div style={{ fontFamily: 'Montserrat', fontWeight: 600, fontSize: 11.5, color: 'var(--t4)', textTransform: 'uppercase', letterSpacing: '.07em', marginBottom: 8 }}>Attachments</div>
          <div style={{ display: 'flex', flexWrap: 'wrap', gap: 8 }}>
            {attachFiles.map((f, i) => (
              <a key={i} href={siteUrl.replace(/\/sites\/.*/, '') + f.ServerRelativeUrl} target="_blank" rel="noopener noreferrer"
                style={{ display: 'inline-flex', alignItems: 'center', gap: 4, background: 'var(--3eg3)', border: '1px solid var(--3eg)', borderRadius: 3, padding: '4px 10px', fontSize: 12, color: 'var(--3eg)', fontFamily: 'Montserrat', textDecoration: 'none', cursor: 'pointer' }}>
                {f.FileName}
              </a>
            ))}
          </div>
        </div>
      )}

      <SDiv label="Parts C & D — Client Response" />
      {row('Client RFI #', rfi.clientRfi)}
      {row('Date Received', fmtD(rfi.dateReceived))}
      {row('Response', rfi.response)}
      {row('Sent By', rfi.sentBy)}
      {row('Sent By Company', rfi.sentByCompany)}
      {rfi.responseDesc ? (
        <div style={{ padding: '12px 0', borderBottom: '1px solid var(--bd)' }}>
          <div style={{ fontFamily: 'Montserrat', fontWeight: 600, fontSize: 11.5, color: 'var(--t4)', textTransform: 'uppercase', letterSpacing: '.07em', marginBottom: 8 }}>Response Details</div>
          <div style={{ fontFamily: 'Montserrat', fontSize: 13, color: 'var(--t1)', lineHeight: 1.7, whiteSpace: 'pre-wrap' }}>{rfi.responseDesc}</div>
        </div>
      ) : null}

      <SDiv label="Part E — Impact Assessment" />
      {row('Schedule Impacted', rfi.impacted)}
      {rfi.ewoRef ? row('EWO Reference', rfi.ewoRef) : null}
      {rfi.impacted === 'Yes' ? (
        <React.Fragment>
          {row('Model Hours', rfi.model)}
          {row('Connections Hours', rfi.connections)}
          {row('Checking Hours', rfi.checking)}
          {row('Drawings Hours', rfi.drawings)}
          {row('Admin Hours', rfi.admin)}
          {row('Total Impact Hours', total.toFixed(1) + 'h', true)}
        </React.Fragment>
      ) : null}
    </div>
  );
};

const fmtTdImport = (raw: string): string => {
  const [iso, ...nameParts] = raw.split('|');
  const who = nameParts.join('|') || '';
  const d = new Date(iso);
  if (isNaN(d.getTime())) return raw;
  const fmt = (tz: string, label: string): string => {
    const opts: Intl.DateTimeFormatOptions = { timeZone: tz, weekday: 'short', day: 'numeric', month: 'short', hour: 'numeric', minute: '2-digit', hour12: true };
    return `${label} ${new Intl.DateTimeFormat('en-AU', opts).format(d)}`;
  };
  const ts = fmt('Australia/Sydney', 'AUS');
  return who ? `${ts} by ${who}` : ts;
};

/** Compact "7 Aug, 9:25 pm" — used inside tooltip sentences. */
const fmtTdShort = (iso: string): string => {
  const d = new Date(iso);
  if (isNaN(d.getTime())) return iso;
  return new Intl.DateTimeFormat('en-AU', {
    timeZone: 'Australia/Sydney', day: 'numeric', month: 'short', hour: 'numeric', minute: '2-digit', hour12: true
  }).format(d);
};

/** Just "7 Aug" — the header chip has room for little more. Time lives in the tooltip. */
const fmtTdDay = (iso: string): string => {
  const d = new Date(iso);
  if (isNaN(d.getTime())) return '';
  return new Intl.DateTimeFormat('en-AU', { timeZone: 'Australia/Sydney', day: 'numeric', month: 'short' }).format(d);
};

/** Counts written by the nightly sync into 3Edge_Settings/tdSyncReport. */
interface TdSyncReport {
  at?: string;
  entries?: number;
  matched?: number;
  unmatched?: number;
  warnings?: number;
  errors?: number;
  updates?: { timelogCreated?: number; timelogUpdated?: number; projectsPatched?: number };
}

type TdTone = 'ok' | 'warn' | 'bad';

/**
 * Real health of the nightly Time Doctor sync.
 *
 * The chip used to be hard-coded green off `lastTdImport` alone, which only
 * proves the job STARTED. It stayed green through eight consecutive nights where
 * the sync ran but could not record its result, so nobody noticed. Green now
 * requires a report that is present, readable, recent, no older than the run it
 * describes, and error-free.
 *
 * `unmatched` is deliberately not an alarm — the "00 - Non Production Tasks"
 * bucket is legitimately unmatched every single night. It is shown in the
 * tooltip instead.
 */
const tdBadge = (lastImport: string, reportRaw: string | null): { tone: TdTone; status: string; when: string; title: string } => {
  const [iso, ...rest] = lastImport.split('|');
  const who = rest.join('|') || 'unknown';
  // Split rather than one string: the header renders `status` unshrinkable and lets
  // `when` ellipsise away, so a narrow window loses the date but never the state.
  const when = fmtTdDay(iso);
  const ranAt = Date.parse(iso);
  const lines = [`Last run: ${fmtTdImport(lastImport)}`];

  let rep: TdSyncReport | null = null;
  if (reportRaw) { try { rep = JSON.parse(reportRaw) as TdSyncReport; } catch (_e) { rep = null; } }
  if (rep) {
    lines.push(`Entries ${rep.entries ?? '?'} · matched ${rep.matched ?? '?'} · unmatched ${rep.unmatched ?? '?'}`);
    lines.push(`Warnings ${rep.warnings ?? 0} · errors ${rep.errors ?? 0}`);
    if (rep.updates) {
      lines.push(`Wrote ${rep.updates.timelogCreated ?? 0} new / ${rep.updates.timelogUpdated ?? 0} updated log rows, patched ${rep.updates.projectsPatched ?? 0} project(s)`);
    }
  }
  // Montserrat as shipped has no glyph for U+2713 or U+26A0 — both map to .notdef,
  // so they render from a fallback face or as tofu. Colour carries the state; these
  // are ASCII reinforcement.
  const bad = (why: string): { tone: TdTone; status: string; when: string; title: string } =>
    ({ tone: 'bad', status: 'TD Sync !', when, title: lines.concat(why).join('\n') });
  const warn = (why: string, status = 'TD Sync ~'): { tone: TdTone; status: string; when: string; title: string } =>
    ({ tone: 'warn', status, when, title: lines.concat(why).join('\n') });

  if (isNaN(ranAt)) return bad('Last-run timestamp is unreadable.');
  const staleH = (Date.now() - ranAt) / 3600000;

  // A manual XLS import or an hours reset writes lastTdImport but no report — flag
  // it rather than letting it masquerade as a healthy automatic run.
  if (who !== 'Auto (Time Doctor)') {
    if (staleH > 36) return bad(`Last change was a manual edit by ${who}, and the nightly sync has not run since.`);
    return warn(`Set by hand (${who}), not the nightly sync. Tonight's run will recompute these hours.`, 'TD Manual');
  }
  if (!reportRaw) return bad('No sync report found — the run cannot confirm it succeeded.');
  if (!rep) return bad('Sync report is unreadable.');

  const repAt = rep.at ? Date.parse(rep.at) : NaN;
  if (isNaN(repAt)) return bad('Sync report has no valid timestamp.');
  // One minute of slack: both settings rows are written from the same instant, in
  // separate try blocks, so either one failing on its own shows up as a mismatch.
  if (repAt < ranAt - 60000) return bad(`Report (${fmtTdShort(rep.at as string)}) is older than the run — the sync ran but could not record its result.`);
  if (repAt > ranAt + 60000) return bad(`Report (${fmtTdShort(rep.at as string)}) is newer than the run stamp — the sync could not record when it last ran.`);
  if (staleH > 36) return bad('No sync in over 36 hours.');
  if ((rep.errors ?? 0) > 0) return bad(`${rep.errors} error(s) in the last run.`);
  // A run that pulled nothing writes a clean, error-free report. Hours are frozen,
  // not correct — but a genuine shutdown week looks the same, so warn, never fail.
  if (rep.entries === 0) return warn('The last run found no Time Doctor activity at all — hours have not moved.');
  if ((rep.warnings ?? 0) > 0) return warn(`${rep.warnings} warning(s) — some projects were skipped.`);
  return { tone: 'ok', status: 'TD Sync OK', when, title: lines.join('\n') };
};

/**
 * A 3E project code anywhere in a Time Doctor project name. Tolerates "3E531",
 * "3E - 531" and en/em-dash variants; `\d{3,}` keeps incidental text ("Bay 3 E 12")
 * from reading as a code. Mirrors PROJ_CODE_RE in 3edge-tdsync/src/sync.js.
 */
const PROJ_CODE_RE = /\b3\s*e\s*[-‐-―_]?\s*(\d{3,})\b/i;

/**
 * Match a Time Doctor project name to a dashboard project.
 *
 * Ranked, not OR-ed — mirrors matchProject() in 3edge-tdsync/src/sync.js so the
 * manual import and the nightly sync always agree. A 3E code in the TD name is
 * authoritative: it beats any name similarity, and if it names no project the row
 * is left unmatched rather than guessed onto a lookalike. Name matching applies
 * only when there is no code, and only when it is unambiguous.
 */
const matchTdProject = (tdName: string, projects: IProject[]): IProject | undefined => {
  const projNumMatch = tdName.match(PROJ_CODE_RE);
  if (projNumMatch) {
    const code = `3e-${projNumMatch[1]}`;
    const exact = projects.filter(p => (p.projNum || '').trim().toLowerCase().replace(/\s+/g, '') === code);
    return exact.find(p => !p.isEwo) || exact[0];
  }
  const td = tdName.trim().toLowerCase();
  const named = projects.filter(p => {
    const name = (p.name || '').trim().toLowerCase();
    // Very short names match far too much to be trusted on their own.
    if (name.length < 5 || !(p.projNum || '').trim()) return false;
    return td.indexOf(name) >= 0 || name.indexOf(td) >= 0;
  });
  const nonEwo = named.filter(p => !p.isEwo);
  if (nonEwo.length === 1) return nonEwo[0];
  if (named.length === 1) return named[0];
  return undefined;
};

// ── Time Doctor Import Modal ───────────────────────────────────────────────────
interface TdImportModalProps {
  projects: IProject[];
  onClose: () => void;
  onApply: (updates: Array<{ projId: string; hrsUsed: number }>) => void;
  onResetHours: () => void;
  lastImport?: string;
}

interface TdPreviewRow {
  projId: string;
  projName: string;
  hrsUsed: number;
  current: number;
}

const TdImportModal: React.FC<TdImportModalProps> = ({ projects, onClose, onApply, onResetHours, lastImport }) => {
  const [preview, setPreview] = React.useState<TdPreviewRow[]>([]);
  const [error, setError] = React.useState('');
  const [parsed, setParsed] = React.useState(false);

  // Parse "17h 00m", "54h 01m", "22m", "0m", or numeric values to decimal hours
  const parseHrsMin = (val: unknown): number => {
    const s = String(val || '').trim();
    if (!s || s === '0m' || s === '0') return 0;
    // Try "Xh Ym" format
    const hm = s.match(/(\d+)\s*h\s*(\d+)\s*m/i);
    if (hm) return parseInt(hm[1], 10) + parseInt(hm[2], 10) / 60;
    // Try "Xh" only
    const hOnly = s.match(/^(\d+)\s*h$/i);
    if (hOnly) return parseInt(hOnly[1], 10);
    // Try "Xm" only
    const mOnly = s.match(/^(\d+)\s*m$/i);
    if (mOnly) return parseInt(mOnly[1], 10) / 60;
    // Try plain number
    const n = parseFloat(s);
    return isNaN(n) ? 0 : n;
  };

  const handleFile = (f: File | null): void => {
    if (!f) return;
    setError('');
    setParsed(false);
    setPreview([]);
    const reader = new FileReader();
    reader.onload = (ev) => {
      try {
        const data = new Uint8Array(ev.target!.result as ArrayBuffer);
        const wb = XLSX.read(data, { type: 'array' });
        const ws = wb.Sheets[wb.SheetNames[0]];
        const rows: unknown[][] = XLSX.utils.sheet_to_json(ws, { header: 1 });
        let hRow = -1;
        let projCol = -1;
        let hrsCol = -1;
        for (let i = 0; i < Math.min(rows.length, 10); i++) {
          const row = rows[i] as unknown[];
          for (let j = 0; j < row.length; j++) {
            const cell = String(row[j] || '').toLowerCase().trim();
            if (cell === 'project' || cell === 'project name') projCol = j;
            if (cell.indexOf('time tracked') >= 0 || cell.indexOf('hour') >= 0 ||
                cell.indexOf('total') >= 0 || cell.indexOf('duration') >= 0 ||
                cell.indexOf('tracked') >= 0 || cell.indexOf('worked') >= 0) hrsCol = j;
          }
          if (projCol >= 0 && hrsCol >= 0) { hRow = i; break; }
        }
        if (hRow < 0 || projCol < 0 || hrsCol < 0) {
          setError('Could not find Project / Hours columns. Ensure the XLS has "Project" and "Time Tracked" (or "Hours") headers.');
          return;
        }
        // Aggregate hours by project
        const aggMap: Record<string, number> = {};
        for (let i = hRow + 1; i < rows.length; i++) {
          const row = rows[i] as unknown[];
          const projRaw = String(row[projCol] || '').trim();
          const hrs = parseHrsMin(row[hrsCol]);
          if (!projRaw || hrs === 0) continue;
          aggMap[projRaw] = (aggMap[projRaw] || 0) + hrs;
        }
        // Match aggregated projects to dashboard projects
        const updates: TdPreviewRow[] = [];
        for (const [xlsName, totalHrs] of Object.entries(aggMap)) {
          // Matches on the 3E-XXX code in names like "01 - 3E-500 SAMPLE TASK".
          const match = matchTdProject(xlsName, projects);
          if (match) {
            const existing = updates.find(u => u.projId === match.id);
            if (existing) { existing.hrsUsed += totalHrs; }
            else { updates.push({ projId: match.id, projName: match.projNum + ' — ' + match.name, hrsUsed: Math.round(totalHrs * 10) / 10, current: match.hrsUsed }); }
          }
        }
        if (updates.length === 0) {
          setError('No matching projects found. Ensure project names in the XLS contain project numbers (e.g. "3E-500").');
          return;
        }
        // Round hours
        updates.forEach(u => { u.hrsUsed = Math.round(u.hrsUsed * 10) / 10; });
        setPreview(updates);
        setParsed(true);
      } catch (e) {
        setError('Failed to parse XLS file: ' + String(e));
      }
    };
    reader.readAsArrayBuffer(f);
  };

  return (
    <div style={{ position: 'fixed', inset: 0, background: 'rgba(240,242,245,0.97)', zIndex: 500, display: 'flex', alignItems: 'center', justifyContent: 'center', backdropFilter: 'blur(3px)' }}>
      <div style={{ background: 'var(--s1)', border: '1px solid var(--bd)', borderRadius: 10, padding: '28px 32px', maxWidth: 560, width: '95%', boxShadow: '0 16px 60px rgba(0,0,0,.18)' }}>
        <div style={{ display: 'flex', alignItems: 'center', justifyContent: 'space-between', marginBottom: 20 }}>
          <div style={{ fontFamily: 'Montserrat', fontWeight: 800, fontSize: 18, color: 'var(--t1)' }}>Import Time Doctor XLS</div>
          <button onClick={onClose} style={{ background: 'transparent', border: '1px solid var(--bd)', color: 'var(--t3)', width: 32, height: 32, borderRadius: 6, fontSize: 15, cursor: 'pointer' }}>x</button>
        </div>
        <div style={{ fontFamily: 'Montserrat', fontSize: 12.5, color: 'var(--t3)', marginBottom: 18, lineHeight: 1.6 }}>
          Select a Time Doctor XLS export. The importer will match project numbers and update Hours Used.
        </div>
        <div style={{ marginBottom: 16 }}>
          <input type="file" accept=".xls,.xlsx,.csv" onChange={e => handleFile(e.target.files ? e.target.files[0] : null)}
            style={{ fontFamily: 'Montserrat', fontSize: 13, color: 'var(--t1)' }} />
        </div>
        {lastImport && (
          <div style={{ fontFamily: 'Montserrat', fontSize: 12, fontWeight: 700, color: 'var(--t2)', marginBottom: 14 }}>
            Last import: {fmtTdImport(lastImport)}
          </div>
        )}
        {error && (
          <div style={{ background: 'var(--rd2)', border: '1px solid var(--rd)', borderRadius: 4, padding: '10px 14px', fontFamily: 'Montserrat', fontSize: 12.5, color: 'var(--rd)', marginBottom: 14 }}>
            {error}
          </div>
        )}
        {parsed && preview.length > 0 && (
          <div>
            <div style={{ fontFamily: 'Montserrat', fontWeight: 700, fontSize: 12.5, color: 'var(--t3)', textTransform: 'uppercase', letterSpacing: '.07em', marginBottom: 10 }}>
              {preview.length} project{preview.length !== 1 ? 's' : ''} will be updated:
            </div>
            <div style={{ maxHeight: 260, overflowY: 'auto', border: '1px solid var(--bd)', borderRadius: 4, marginBottom: 18 }}>
              <table style={{ width: '100%', borderCollapse: 'collapse', fontSize: 12.5, fontFamily: 'Montserrat' }}>
                <thead>
                  <tr style={{ background: 'var(--s2)', borderBottom: '1px solid var(--bd)' }}>
                    <th style={{ padding: '8px 12px', textAlign: 'left', fontWeight: 700, color: 'var(--t3)' }}>Project</th>
                    <th style={{ padding: '8px 12px', textAlign: 'right', fontWeight: 700, color: 'var(--t3)' }}>Current</th>
                    <th style={{ padding: '8px 12px', textAlign: 'right', fontWeight: 700, color: 'var(--t3)' }}>New Total</th>
                    <th style={{ padding: '8px 12px', textAlign: 'right', fontWeight: 700, color: 'var(--t3)' }}>Change</th>
                  </tr>
                </thead>
                <tbody>
                  {preview.map((u, i) => {
                    const newTotal = u.current + u.hrsUsed;
                    return (
                      <tr key={i} style={{ borderBottom: '1px solid var(--bd)' }}>
                        <td style={{ padding: '7px 12px', color: 'var(--t1)', fontWeight: 600 }}>{u.projName}</td>
                        <td style={{ padding: '7px 12px', textAlign: 'right', color: 'var(--t3)' }}>{u.current}h</td>
                        <td style={{ padding: '7px 12px', textAlign: 'right', color: 'var(--t1)', fontWeight: 700 }}>{newTotal.toFixed(1)}h</td>
                        <td style={{ padding: '7px 12px', textAlign: 'right', color: 'var(--am)', fontWeight: 600 }}>+{u.hrsUsed.toFixed(1)}h</td>
                      </tr>
                    );
                  })}
                </tbody>
              </table>
            </div>
            <div style={{ fontFamily: 'Montserrat', fontSize: 11.5, lineHeight: 1.5, color: 'var(--am)', background: 'var(--am2)', border: '1px solid var(--am)', borderRadius: 6, padding: '8px 12px', marginBottom: 12 }}>
              Hours are maintained automatically by the nightly Time Doctor sync
              (hours already worked + hours logged since). Anything applied here is a
              temporary override — tonight&apos;s 00:15 AWST run will recalculate it.
              Use this only when the automatic sync is unavailable.
            </div>
            <div style={{ display: 'flex', gap: 10 }}>
              <BtnPrimary onClick={() => onApply(preview.map(u => ({ projId: u.projId, hrsUsed: u.current + u.hrsUsed })))}>
                APPLY UPDATES
              </BtnPrimary>
              <button onClick={onClose} style={{ fontFamily: 'Montserrat', fontSize: 12.5, padding: '9px 18px', background: 'transparent', border: '1px solid var(--bd)', color: 'var(--t2)', borderRadius: 7, cursor: 'pointer' }}>Cancel</button>
            </div>
          </div>
        )}
        {!parsed && (
          <div style={{ display: 'flex', justifyContent: 'flex-end', gap: 10, marginTop: 8 }}>
            <button onClick={() => {
              const typed = prompt('This zeroes Hours Used on EVERY project.\n\nThe nightly Time Doctor sync will restore hours for any project it tracks, so this mainly affects projects it does not. EWOs are left alone.\n\nType RESET to confirm.');
              if (typed === null) return;                       // cancelled — say nothing
              if (typed.trim().toUpperCase() !== 'RESET') {      // typo must not look like success
                setError('Hours were NOT reset — you need to type RESET exactly.');
                return;
              }
              onResetHours();
            }} style={{ fontFamily: 'Montserrat', fontSize: 12.5, padding: '9px 18px', background: 'var(--rd)', border: 'none', color: '#fff', borderRadius: 7, cursor: 'pointer', fontWeight: 700 }}>Reset All Hours</button>
            <button onClick={onClose} style={{ fontFamily: 'Montserrat', fontSize: 12.5, padding: '9px 18px', background: 'transparent', border: '1px solid var(--bd)', color: 'var(--t2)', borderRadius: 7, cursor: 'pointer' }}>Close</button>
          </div>
        )}
      </div>
    </div>
  );
};

// ── Main Component ─────────────────────────────────────────────────────────────
const ManagerDashboard: React.FC<IManagerDashboardProps> = (props) => {
  const spService = React.useRef(new SharePointService(props.siteUrl, props.spHttpClient));
  const { show: toast, Toast } = useToast();

  // ── Data
  const [projects, setProjects] = React.useState<IProject[]>([]);
  const [rfis, setRfis] = React.useState<IRfi[]>([]);
  const [teamMembers, setTeamMembers] = React.useState<ITeamMember[]>([]);
  const [spLoading, setSpLoading] = React.useState(true);

  // ── View
  const [mod, setMod] = React.useState<Mod>('projects');
  const [clock, setClock] = React.useState({ aus: '', ph: '' });
  const [role, setRole] = React.useState<Role>('manager');
  const [spMode, setSpMode] = React.useState<SpMode>('detecting');
  const [userRole, setUserRole] = React.useState<'owner' | 'member' | 'loading'>('loading');

  // ── Project filters & sort
  const [srch, setSrch] = React.useState('');
  const [yr, setYr] = React.useState('2026');
  const [stFilt, setStFilt] = React.useState('');
  const [showArchived, setShowArchived] = React.useState(false);
  const [sCol, setSCol] = React.useState('projNum');
  const [sDir, setSDir] = React.useState<SDir>('asc');

  // ── EWO expand
  const [exp, setExp] = React.useState<Record<string, boolean>>({});

  // ── RFI filters & sort
  const [rfiSrch, setRfiSrch] = React.useState('');
  const [rfiProj, setRfiProj] = React.useState('');
  const [rfiSt, setRfiSt] = React.useState('');
  const [rSCol, setRSCol] = React.useState('rfiNum');
  const [rSDir, setRSDir] = React.useState<SDir>('asc');
  const [rfiExp, setRfiExp] = React.useState<Record<string, boolean>>({});

  // ── EWO filters
  const [ewoSrch, setEwoSrch] = React.useState('');
  const [ewoParent, setEwoParent] = React.useState('');
  const [ewoStFilt, setEwoStFilt] = React.useState('');
  const [ewoExp, setEwoExp] = React.useState<Record<string, boolean>>(() => {
    try { const v = localStorage.getItem('3edge-ewo-exp'); return v ? JSON.parse(v) as Record<string, boolean> : {}; } catch { return {}; }
  });

  // ── Panel
  const [panel, setPanel] = React.useState<PanelState>({ type: null });

  // ── Delete modal
  const [del, setDel] = React.useState<DelState>({ open: false, label: '', onConfirm: () => undefined });

  // ── CRM password gate
  const [crmUnlocked, setCrmUnlocked] = React.useState(false);
  const [crmPw, setCrmPw] = React.useState('');
  const [crmPwShow, setCrmPwShow] = React.useState(false);
  const [crmPwError, setCrmPwError] = React.useState(false);
  const [crmAttempts, setCrmAttempts] = React.useState(0);
  const [crmLockedUntil, setCrmLockedUntil] = React.useState<number | null>(null);
  const [crmLockRemain, setCrmLockRemain] = React.useState(0);
  React.useEffect(() => {
    if (!crmUnlocked) return;
    const t = setTimeout(() => setCrmUnlocked(false), 30 * 60 * 1000);
    return () => clearTimeout(t);
  }, [crmUnlocked]);
  React.useEffect(() => {
    if (!crmLockedUntil) return;
    const tick = setInterval(() => {
      const rem = Math.ceil((crmLockedUntil - Date.now()) / 1000);
      if (rem <= 0) { setCrmLockedUntil(null); setCrmAttempts(0); setCrmLockRemain(0); clearInterval(tick); }
      else setCrmLockRemain(rem);
    }, 1000);
    return () => clearInterval(tick);
  }, [crmLockedUntil]);
  const crmTryUnlock = (): void => {
    if (crmLockedUntil) return;
    if (crmPw === 'Account123!@#') {
      setCrmUnlocked(true); setCrmPw(''); setCrmPwError(false); setCrmAttempts(0);
    } else {
      const next = crmAttempts + 1;
      setCrmAttempts(next); setCrmPwError(true); setCrmPw('');
      if (next >= 3) { setCrmLockedUntil(Date.now() + 15 * 60 * 1000); }
    }
  };

  // ── Time Doctor
  const [tdModal, setTdModal] = React.useState(false);
  const [lastTdImport, setLastTdImport] = React.useState<string | null>(null);
  // undefined = not read yet / read failed, null = read succeeded and there is no
  // report. getSetting returns undefined for both a missing key AND a thrown
  // request, so conflating them would paint the chip red on any transient 429.
  const [tdReport, setTdReport] = React.useState<string | null | undefined>(undefined);

  // ── Load data
  const loadData = React.useCallback(async () => {
    setSpLoading(true);
    setSpMode('detecting');
    try {
      const [p, r, tm] = await Promise.all([
        spService.current.loadProjects(),
        spService.current.loadRfis(),
        spService.current.loadTeamMembers().catch(() => [] as ITeamMember[])
      ]);
      setProjects(p);
      setRfis(r);
      setTeamMembers(tm);
      setSpMode('live');
      // Both rows are needed: lastTdImport says the sync RAN, tdSyncReport says
      // whether it SUCCEEDED. The chip is only green when they agree. Settle them
      // together — React 17 does not batch separate .then callbacks, so resolving
      // them independently guarantees a render with one set and the other not,
      // which would flash the chip red on every load.
      Promise.all([
        spService.current.getSetting('lastTdImport'),
        spService.current.getSetting('tdSyncReport'),
      ]).then(([impRaw, repRaw]) => {
        if (impRaw) setLastTdImport(impRaw);
        setTdReport(repRaw === undefined ? undefined : (repRaw || null));
      }).catch(() => undefined);
    } catch (e) {
      const msg = (e instanceof Error) ? e.message : String(e);
      toast('SharePoint unavailable — running in local mode. (' + msg + ')', 'error');
      setSpMode('local');
    } finally {
      setSpLoading(false);
    }
  }, [toast]);

  React.useEffect(() => {
    loadData().catch(() => undefined);
  }, [loadData]);

  // ── Check user role (Owner vs Member)
  React.useEffect(() => {
    (async () => {
      try {
        const hdrs = { credentials: 'include' as RequestCredentials, headers: { 'Accept': 'application/json;odata=nometadata' } };
        // 1. Check if site admin
        const uRes = await fetch(props.siteUrl + '/_api/web/currentuser', hdrs);
        if (uRes.ok) {
          const u = await uRes.json();
          if (u.IsSiteAdmin) { setUserRole('owner'); return; }
          const userId = u.Id;
          // 2. Check if user is in the site's associated owner group
          const oRes = await fetch(props.siteUrl + '/_api/web/associatedownergroup/users', hdrs);
          if (oRes.ok) {
            const oData = await oRes.json();
            const owners: Array<{ Id: number }> = oData.value || [];
            if (owners.some(o => o.Id === userId)) { setUserRole('owner'); return; }
          }
          // 3. Also check group titles as fallback
          const gRes = await fetch(props.siteUrl + '/_api/web/currentuser/groups', hdrs);
          if (gRes.ok) {
            const gData = await gRes.json();
            const groups: Array<{ Title: string }> = gData.value || [];
            if (groups.some(g => /owner/i.test(g.Title))) { setUserRole('owner'); return; }
          }
          // Not an owner
          setUserRole('member');
          setRole('staff');
        } else {
          setUserRole('owner'); // fallback
        }
      } catch (_e) {
        setUserRole('owner'); // fallback
      }
    })().catch(() => undefined);
  }, [props.siteUrl]);

  // ── Derived: is current user allowed to act as manager?
  const isManager = userRole === 'owner' && role === 'manager';

  // ── Live clock
  React.useEffect(() => {
    const tick = (): void => {
      const now = new Date();
      setClock({
        aus: now.toLocaleTimeString('en-AU', { hour: '2-digit', minute: '2-digit', second: '2-digit', timeZone: 'Australia/Sydney' }),
        ph:  now.toLocaleTimeString('en-PH', { hour: '2-digit', minute: '2-digit', second: '2-digit', timeZone: 'Asia/Manila' })
      });
    };
    tick();
    const id = setInterval(tick, 1000);
    return () => clearInterval(id);
  }, []);

  // ── Sort helpers
  const sortList = <T,>(arr: T[], col: string, dir: SDir): T[] => {
    return arr.slice().sort((a, b) => {
      const va = (a as Record<string, unknown>)[col] as string | number ?? '';
      const vb = (b as Record<string, unknown>)[col] as string | number ?? '';
      const cmp = String(va).localeCompare(String(vb), undefined, { numeric: true });
      return dir === 'asc' ? cmp : -cmp;
    });
  };

  const onSort = (col: string): void => {
    if (sCol === col) setSDir((d: SDir) => d === 'asc' ? 'desc' : 'asc');
    else { setSCol(col); setSDir('asc'); }
  };

  const onRSort = (col: string): void => {
    if (rSCol === col) setRSDir((d: SDir) => d === 'asc' ? 'desc' : 'asc');
    else { setRSCol(col); setRSDir('asc'); }
  };

  const sortArrow = (col: string, active: string, dir: SDir): string => {
    if (col !== active) return ' ↕';
    return dir === 'asc' ? ' ↑' : ' ↓';
  };

  // ── Filtered projects
  const visProjects = React.useMemo(() => {
    const list = projects.filter(p => {
      if (p.isEwo) return false; // EWOs shown as sub-rows
      if (showArchived) {
        if (p.status !== 'Archive') return false;
      } else {
        if (p.status === 'Archive') return false;
      }
      if (yr && yr !== 'all' && String(p.year) !== yr) return false;
      if (stFilt && p.status !== stFilt) return false;
      if (srch) {
        const q = srch.toLowerCase();
        return (p.projNum + p.name + p.company + p.contact + p.quoteNum).toLowerCase().indexOf(q) >= 0;
      }
      return true;
    });
    return sortList(list, sCol, sDir);
  }, [projects, yr, stFilt, srch, sCol, sDir, showArchived]);

  // ── RFI used per project
  const rfiCountByProj = React.useMemo(() => {
    const m: Record<string, number> = {};
    rfis.forEach(r => { m[r.projectId] = (m[r.projectId] || 0) + 1; });
    return m;
  }, [rfis]);

  // ── Filtered RFIs
  const visRfis = React.useMemo(() => {
    const list = rfis.filter(r => {
      if (rfiProj && r.projectId !== rfiProj) return false;
      if (rfiSt && effSt(r) !== rfiSt) return false;
      if (rfiSrch) {
        const q = rfiSrch.toLowerCase();
        return (r.rfiNum + r.description + r.submittedTo + r.projectName).toLowerCase().indexOf(q) >= 0;
      }
      return true;
    });
    return sortList(list, rSCol, rSDir);
  }, [rfis, rfiProj, rfiSt, rfiSrch, rSCol, rSDir]);

  // ── RFIs grouped by project
  const rfisByProject = React.useMemo(() => {
    const m: Record<string, IRfi[]> = {};
    visRfis.forEach(r => {
      if (!m[r.projectId]) m[r.projectId] = [];
      m[r.projectId].push(r);
    });
    return m;
  }, [visRfis]);

  // ── Filtered EWOs
  const visEwos = React.useMemo(() => {
    return projects.filter(p => {
      if (!p.isEwo) return false;
      if (ewoParent && p.parentId !== ewoParent) return false;
      if (ewoStFilt && p.status !== ewoStFilt) return false;
      if (ewoSrch) {
        const q = ewoSrch.toLowerCase();
        return (p.projNum + p.name + p.company + p.ewoNum).toLowerCase().indexOf(q) >= 0;
      }
      return true;
    });
  }, [projects, ewoParent, ewoStFilt, ewoSrch]);

  // ── EWOs grouped by parent project
  const ewosByParent = React.useMemo(() => {
    const m: Record<string, IProject[]> = {};
    visEwos.forEach(e => {
      const pid = e.parentId || 'unknown';
      if (!m[pid]) m[pid] = [];
      m[pid].push(e);
    });
    return m;
  }, [visEwos]);

  // ── Stat cards
  const mainProjects = projects.filter(p => !p.isEwo);
  const allActive = mainProjects.filter(p => p.status === 'Active').length;
  const allComplete = mainProjects.filter(isProjectDelivered).length;
  const allOverBudget = mainProjects.filter(p => p.status === 'Over Budget' || (p.hrsAllowed > 0 && p.hrsUsed > p.hrsAllowed)).length;
  const allEwos = projects.filter(p => p.isEwo).length;
  const totalHrsUsed = projects.reduce((s, p) => s + p.hrsUsed, 0);
  const totalHrsAllowed = projects.reduce((s, p) => s + p.hrsAllowed, 0);
  const rfiOpen = rfis.filter(r => effSt(r) === 'Open').length;
  const rfiOverdue = rfis.filter(r => isOD(r)).length;
  const rfiPartial = rfis.filter(r => r.status === 'Partially Open (Revise and Resend)').length;
  const rfiClosed = rfis.filter(r => r.status === 'Closed').length;
  const rfiImpact = rfis.filter(r => r.impacted === 'Yes').reduce((s, r) => s + rfiTot(r), 0);

  // ── Local-mode temp ID counter
  const localIdRef = React.useRef(1);
  const nextLocalId = (): number => { localIdRef.current -= 1; return localIdRef.current; };
  const isLocal = (): boolean => spMode === 'local';

  // ── CRUD helpers
  const saveProject = async (d: IProject, isNew: boolean, addedFiles?: File[], removedFiles?: string[]): Promise<void> => {
    try {
      const prev = isNew ? undefined : projects.find(p => p.id === d.id);
      const payload = withProjectDeliveryOnSave(d, prev);
      if (isLocal()) {
        if (isNew) {
          const tempId = nextLocalId();
          const saved: IProject = { ...payload, id: payload.projNum || String(tempId) };
          setProjects(prevList => [...prevList, saved]);
          toast('Project created (local mode — will sync when SP lists are ready).');
        } else {
          setProjects(prevList => prevList.map(p => p.id === payload.id ? { ...payload } : p));
          toast('Project saved (local mode).');
        }
        if (addedFiles && addedFiles.length > 0) {
          toast('Note: attachments are not saved in local mode.', 'error');
        }
        setPanel({ type: null });
        return;
      }
      let spId: number | undefined;
      if (isNew) {
        spId = await spService.current.addProject(payload);
        const saved: IProject = { ...payload, id: payload.projNum || String(spId), spId };
        setProjects(prevList => [...prevList, saved]);
        toast('Project created.');
      } else {
        if (!payload.spId) throw new Error('No spId on project');
        spId = payload.spId;
        await spService.current.updateProject(payload.spId, payload);
        // pBody deliberately omits hrsUsed so the sync owns it — but the sync never
        // touches EWOs, so this is the only path that can persist their hours.
        if (payload.isEwo) {
          await spService.current.updateProjectHours(payload.spId, Number(payload.hrsUsed) || 0);
        }
        setProjects(prevList => prevList.map(p => p.id === payload.id ? { ...payload } : p));
        toast('Project saved.');
      }
      const projListName = spService.current.getProjectListName();
      if (spId && removedFiles && removedFiles.length > 0) {
        for (const name of removedFiles) {
          try { await spService.current.deleteAttachment(spId, name, projListName); } catch (_e) { /* ignore */ }
        }
      }
      if (spId && addedFiles && addedFiles.length > 0) {
        for (const f of addedFiles) {
          await spService.current.uploadAttachment(spId, f, projListName);
        }
        toast(addedFiles.length + ' attachment(s) uploaded.');
      }
      setPanel({ type: null });
    } catch (e) {
      const msg = (e instanceof Error) ? e.message : String(e);
      toast('Save failed: ' + msg, 'error');
    }
  };

  const toggleArchive = async (proj: IProject): Promise<void> => {
    const updated = withProjectArchiveToggle(proj);
    try {
      if (!isLocal() && proj.spId) {
        await spService.current.updateProject(proj.spId, updated);
      }
      setProjects(prev => prev.map(p => p.id === proj.id ? updated : p));
      toast(updated.status === 'Archive' ? 'Project archived.' : 'Project restored.');
    } catch (e) {
      const msg = (e instanceof Error) ? e.message : String(e);
      toast('Failed: ' + msg, 'error');
    }
  };

  const deleteProject = async (proj: IProject): Promise<void> => {
    if (isLocal() || !proj.spId) {
      setProjects(prev => prev.filter(p => p.id !== proj.id));
      setPanel({ type: null });
      setDel({ open: false, label: '', onConfirm: () => undefined });
      toast('Project deleted' + (isLocal() ? ' (local mode).' : '.'));
      return;
    }
    try {
      await spService.current.deleteProject(proj.spId);
      setProjects(prev => prev.filter(p => p.id !== proj.id));
      setPanel({ type: null });
      setDel({ open: false, label: '', onConfirm: () => undefined });
      toast('Project deleted.');
    } catch (e) {
      const msg = (e instanceof Error) ? e.message : String(e);
      toast('Delete failed: ' + msg, 'error');
    }
  };

  const saveRfi = async (d: IRfi, isNew: boolean, files?: File[]): Promise<void> => {
    try {
      if (isLocal()) {
        if (isNew) {
          const tempId = nextLocalId();
          const saved: IRfi = { ...d, id: d.rfiNum || String(tempId) };
          setRfis(prev => [...prev, saved]);
          toast('RFI created (local mode — will sync when SP lists are ready).');
        } else {
          setRfis(prev => prev.map(r => r.id === d.id ? { ...d } : r));
          toast('RFI saved (local mode).');
        }
        setPanel({ type: null });
        return;
      }
      let spId = d.spId;
      if (isNew) {
        spId = await spService.current.addRfi(d);
        const saved: IRfi = { ...d, id: d.rfiNum || String(spId), spId };
        setRfis(prev => [...prev, saved]);
        toast('RFI created.');
      } else {
        if (!spId) throw new Error('No spId on RFI');
        await spService.current.updateRfi(spId, d);
        setRfis(prev => prev.map(r => r.id === d.id ? { ...d } : r));
        toast('RFI saved.');
      }
      // Upload pending files
      if (files && files.length > 0 && spId) {
        for (const f of files) {
          await spService.current.uploadAttachment(spId, f);
        }
        toast(files.length + ' file(s) attached.');
      }
      setPanel({ type: null });
    } catch (e) {
      const msg = (e instanceof Error) ? e.message : String(e);
      toast('Save failed: ' + msg, 'error');
    }
  };

  const deleteRfi = async (rfi: IRfi): Promise<void> => {
    if (isLocal() || !rfi.spId) {
      setRfis(prev => prev.filter(r => r.id !== rfi.id));
      setPanel({ type: null });
      setDel({ open: false, label: '', onConfirm: () => undefined });
      toast('RFI deleted' + (isLocal() ? ' (local mode).' : '.'));
      return;
    }
    try {
      await spService.current.deleteRfi(rfi.spId);
      setRfis(prev => prev.filter(r => r.id !== rfi.id));
      setPanel({ type: null });
      setDel({ open: false, label: '', onConfirm: () => undefined });
      toast('RFI deleted.');
    } catch (e) {
      const msg = (e instanceof Error) ? e.message : String(e);
      toast('Delete failed: ' + msg, 'error');
    }
  };

  const confirmDelete = (label: string, fn: () => void): void => {
    setDel({ open: true, label, onConfirm: fn });
  };

  // ── Time Doctor apply
  const applyTdUpdates = async (updates: Array<{ projId: string; hrsUsed: number }>): Promise<void> => {
    setTdModal(false);
    let success = 0;
    for (let i = 0; i < updates.length; i++) {
      const u = updates[i];
      const p = projects.filter(x => x.id === u.projId)[0];
      if (!p || !p.spId) continue;
      try {
        const updated: IProject = { ...p, hrsUsed: u.hrsUsed };
        // Hours-only write: must not carry the rest of this browser's copy of the
        // project, which may be minutes or hours out of date.
        await spService.current.updateProjectHours(p.spId, u.hrsUsed);
        setProjects(prev => prev.map(x => x.id === u.projId ? updated : x));
        success++;
      } catch (e) {
        const msg = (e instanceof Error) ? e.message : String(e);
        toast('Failed to update ' + p.projNum + ': ' + msg, 'error');
      }
    }
    const tsVal = new Date().toISOString() + '|' + props.userDisplayName;
    spService.current.setSetting('lastTdImport', tsVal).catch(() => undefined);
    setLastTdImport(tsVal);
    toast('Time Doctor import: ' + success + ' project' + (success !== 1 ? 's' : '') + ' updated.');
  };

  const resetAllHours = async (): Promise<void> => {
    setTdModal(false);
    let success = 0;
    for (const p of projects) {
      // Projects are safe to zero — tonight's run recomputes them from Time Doctor.
      // EWO hours are hand-entered and nothing would ever restore them.
      if (p.isEwo) continue;
      if (!p.spId || p.hrsUsed === 0) continue;
      try {
        const updated: IProject = { ...p, hrsUsed: 0 };
        await spService.current.updateProjectHours(p.spId, 0);
        setProjects(prev => prev.map(x => x.id === p.id ? updated : x));
        success++;
      } catch (e) {
        toast('Failed to reset ' + p.projNum + ': ' + ((e instanceof Error) ? e.message : String(e)), 'error');
      }
    }
    // Hand-set hours are not the sync's work — stamp provenance so the header chip
    // reports "Manual" rather than certifying these numbers as synced.
    const resetStamp = new Date().toISOString() + '|' + (props.userDisplayName || 'manual reset');
    spService.current.setSetting('lastTdImport', resetStamp).catch(() => undefined);
    setLastTdImport(resetStamp);
    toast('Reset hours: ' + success + ' project' + (success !== 1 ? 's' : '') + ' set to 0. EWOs were left alone.');
  };

  // ── Years for filter
  const years = React.useMemo(() => {
    const seen: Record<string, boolean> = {};
    seen['2026'] = true;
    projects.forEach(p => { seen[String(p.year)] = true; });
    const arr = Object.keys(seen).sort().reverse();
    return (['all'] as string[]).concat(arr);
  }, [projects]);

  // ── Panel helpers
  const openProjDetail = (p: IProject): void => setPanel({ type: 'projDetail', proj: p });
  const openProjForm = (p: IProject | null): void => setPanel({ type: 'projForm', proj: p });
  const openRfiDetail = (r: IRfi, parentProj?: IProject): void => setPanel({ type: 'rfiDetail', rfi: r, parentProj });
  const openRfiForm = (r: IRfi | null, parentProj?: IProject): void => setPanel({ type: 'rfiForm', rfi: r, parentProj });

  // ── EWOs for a project
  const getEwos = (parentId: string): IProject[] => projects.filter(p => p.isEwo && p.parentId === parentId);

  // ── Th helper component
  const Th: React.FC<{ col: string; label: string; rfi?: boolean; pad?: string }> = ({ col, label, rfi: isRfi, pad }) => {
    const active = isRfi ? rSCol : sCol;
    const dir = isRfi ? rSDir : sDir;
    return (
      <th onClick={() => isRfi ? onRSort(col) : onSort(col)}
        style={{ padding: pad || '8px 6px', fontFamily: 'Montserrat', fontWeight: 700, fontSize: 11, letterSpacing: '.1em', textTransform: 'uppercase', color: active === col ? 'var(--3eg)' : 'var(--t3)', cursor: 'pointer', whiteSpace: 'nowrap', borderBottom: '2px solid var(--bd)', textAlign: 'left', userSelect: 'none', background: 'var(--s2)' }}>
        {label}<span style={{ opacity: 0.6 }}>{sortArrow(col, active, dir)}</span>
      </th>
    );
  };

  // ── Plain th
  const ThPlain: React.FC<{ label: string; pad?: string }> = ({ label, pad }) => (
    <th style={{ padding: pad || '8px 6px', fontFamily: 'Montserrat', fontWeight: 700, fontSize: 11, letterSpacing: '.1em', textTransform: 'uppercase', color: 'var(--t3)', whiteSpace: 'nowrap', borderBottom: '2px solid var(--bd)', textAlign: 'left', background: 'var(--s2)' }}>{label}</th>
  );

  const headerBg: React.CSSProperties = {
    background: 'var(--hdr)', display: 'flex', alignItems: 'center',
    padding: '0 20px', height: 56, flexShrink: 0, position: 'relative', zIndex: 200,
    boxShadow: '0 2px 12px rgba(0,0,0,.18)'
  };

  // ── Render
  return (
    <div className={styles.dashboardRoot}>
      {/* ── Header ─────────────────────────────────────────────── */}
      <header style={headerBg}>
        {/* Logo */}
        <div style={{ display: 'flex', alignItems: 'center', gap: 10, marginRight: 22, flexShrink: 0 }}>
          {IMG_LOGO_DASH
            ? <img src={IMG_LOGO_DASH} alt="3 Edge" style={{ height: 84 }} />
            : (
              <div style={{ display: 'flex', flexDirection: 'column', lineHeight: 1 }}>
                <span style={{ fontFamily: 'Montserrat', fontWeight: 900, fontSize: 14, color: 'var(--3eg)', letterSpacing: '.18em' }}>3 EDGE</span>
                <span style={{ fontFamily: 'Montserrat', fontWeight: 400, fontSize: 9, color: '#8a9bb0', letterSpacing: '.2em', marginTop: 1 }}>DESIGN</span>
              </div>
            )
          }
        </div>

        {/* Nav Tabs */}
        <div style={{ display: 'flex', gap: 2, flexShrink: 0 }}>
          {(['projects', 'rfis', 'ewos', 'tasks', 'checklist', 'crm'] as Mod[]).map((m: Mod) => (
            <button key={m} onClick={() => setMod(m)} style={{
              fontFamily: 'Montserrat', fontWeight: 700, fontSize: 11, letterSpacing: '.12em',
              textTransform: 'uppercase', padding: '6px 16px', borderRadius: 4, cursor: 'pointer',
              whiteSpace: 'nowrap', flexShrink: 0,
              background: mod === m ? 'var(--3eg3)' : 'transparent',
              border: mod === m ? '1px solid var(--3eg)' : '1px solid transparent',
              color: mod === m ? 'var(--3eg)' : '#8a9bb0', transition: 'all .15s'
            }}>
              {m === 'projects' ? 'Projects' : m === 'rfis' ? 'RFIs' : m === 'ewos' ? 'EWOs' : m === 'tasks' ? 'Tasks' : m === 'checklist' ? 'Checklist' : 'CRM'}
            </button>
          ))}
        </div>

        <div style={{ flex: 1 }} />

        {/* Time Doctor sync health. Colour is derived, never hard-coded — see tdBadge.
            This is the only shrinkable item in the header, so when space runs out it
            ellipsises instead of squashing the controls to its right. */}
        {mod === 'projects' && isManager && lastTdImport && tdReport !== undefined && (() => {
          const b = tdBadge(lastTdImport, tdReport);
          const col = b.tone === 'ok' ? 'var(--gn)' : b.tone === 'warn' ? 'var(--am)' : 'var(--rd)';
          const bg = b.tone === 'ok' ? 'var(--gn2)' : b.tone === 'warn' ? 'var(--am2)' : 'var(--rd2)';
          return (
            <div style={{ display: 'flex', alignItems: 'center', marginRight: 10, minWidth: 0, flexShrink: 1, overflow: 'hidden' }}>
              <span
                title={`Hours sync from Time Doctor nightly at 00:15 AWST, covering complete days only — today's hours appear after tonight's run.\n\n${b.title}`}
                style={{
                  display: 'inline-flex', alignItems: 'baseline', minWidth: 0,
                  fontFamily: 'Montserrat', fontWeight: 600, fontSize: 11, letterSpacing: '.08em',
                  textTransform: 'uppercase', padding: '5px 12px', borderRadius: 4, cursor: 'default',
                  background: bg, border: `1px solid ${col}`, color: col,
                  whiteSpace: 'nowrap', overflow: 'hidden', maxWidth: '100%'
                }}>
                {/* The state must survive any width — only the date is allowed to go. */}
                <span style={{ flexShrink: 0 }}>{b.status}</span>
                {b.when && (
                  <span style={{ flexShrink: 1, minWidth: 0, overflow: 'hidden', textOverflow: 'ellipsis', opacity: 0.85, marginLeft: 6 }}>
                    · {b.when}
                  </span>
                )}
              </span>
            </div>
          );
        })()}

        {/* SP Status indicator */}
        <div style={{ display: 'flex', alignItems: 'center', gap: 5, marginRight: 14, flexShrink: 0 }}>
          <div style={{
            width: 8, height: 8, borderRadius: '50%',
            background: spMode === 'live' ? 'var(--gn)' : spMode === 'local' ? 'var(--am)' : '#5a6a80',
            flexShrink: 0
          }} className={spMode === 'detecting' ? styles.pulse : ''} />
          {spMode === 'local' && (
            <span style={{ fontFamily: 'Montserrat', fontWeight: 700, fontSize: 10, letterSpacing: '.08em', textTransform: 'uppercase', color: 'var(--am)', background: 'rgba(212,136,10,0.14)', border: '1px solid var(--am)', borderRadius: 3, padding: '1px 6px' }}>
              Local Mode
            </span>
          )}
        </div>

        {/* Role toggle — only visible for Owners.
            flexShrink: 0 is load-bearing. Without it the header squeezed this to
            zero width and all you saw was a green sliver: the active button's
            background with its label crushed out of existence. */}
        {userRole === 'owner' && (
          <div style={{ display: 'flex', border: '1px solid rgba(138,155,176,.3)', borderRadius: 4, overflow: 'hidden', marginRight: 14, flexShrink: 0 }}>
            {(['manager', 'staff'] as Role[]).map(r => (
              <button key={r} onClick={() => setRole(r)} style={{
                fontFamily: 'Montserrat', fontWeight: 700, fontSize: 10.5, letterSpacing: '.1em',
                textTransform: 'uppercase', padding: '4px 12px', cursor: 'pointer', border: 'none',
                whiteSpace: 'nowrap', flexShrink: 0,
                background: role === r ? (r === 'manager' ? 'var(--3eg)' : 'rgba(90,106,128,0.25)') : 'transparent',
                color: role === r ? (r === 'manager' ? '#111418' : '#fff') : '#8a9bb0',
                transition: 'all .15s'
              }}>
                {r === 'staff' ? 'Team' : r}
              </button>
            ))}
          </div>
        )}

        {/* Clock + user */}
        <div className={styles.hdrClock} style={{ display: 'flex', flexDirection: 'column', alignItems: 'flex-end', whiteSpace: 'nowrap', gap: 2, flexShrink: 0 }}>
          <div style={{ fontFamily: 'Montserrat', fontWeight: 600, fontSize: 11, color: '#8a9bb0' }}>
            <span style={{ color: '#5a7a9a', marginRight: 4 }}>AUS</span>{clock.aus}
          </div>
          <div style={{ fontFamily: 'Montserrat', fontWeight: 600, fontSize: 11, color: '#8a9bb0' }}>
            <span style={{ color: '#5a7a9a', marginRight: 4 }}>PH</span>{clock.ph}
          </div>
        </div>
        {props.userDisplayName && (
          <div className={styles.hdrUser} title={props.userDisplayName} style={{ fontFamily: 'Montserrat', fontWeight: 600, fontSize: 11, color: '#8a9bb0', marginLeft: 12, whiteSpace: 'nowrap', flexShrink: 0, maxWidth: 150, overflow: 'hidden', textOverflow: 'ellipsis' }}>{props.userDisplayName}</div>
        )}
      </header>

      {/* ── Values Banner ─────────────────────────────────────────── */}
      <div style={{ display: 'flex', justifyContent: 'center', gap: 32, padding: '6px 24px', background: 'linear-gradient(90deg, #1a1e24 0%, #232830 50%, #1a1e24 100%)', borderBottom: '1px solid rgba(42,158,42,0.25)', flexShrink: 0 }}>
        {['Trust', 'Collaboration', 'Accuracy', 'Progress'].map(v => (
          <span key={v} style={{ fontFamily: 'Montserrat', fontWeight: 700, fontSize: 11, letterSpacing: '.18em', textTransform: 'uppercase', color: 'var(--3eg)' }}>{v}</span>
        ))}
      </div>

      {/* ── Staff Banner ─────────────────────────────────────────── */}
      {role === 'staff' && (
        <div style={{ background: 'rgba(212,136,10,0.10)', borderBottom: '1px solid var(--am)', padding: '7px 24px', display: 'flex', alignItems: 'center', gap: 10, flexShrink: 0 }}>
          <span style={{ fontFamily: 'Montserrat', fontWeight: 700, fontSize: 11, letterSpacing: '.1em', textTransform: 'uppercase', color: 'var(--am)' }}>Team View</span>
          <span style={{ fontFamily: 'Montserrat', fontSize: 12, color: 'var(--t3)' }}>— Read-only access. Switch to Manager to create or edit records.</span>
        </div>
      )}

      {/* ── Body ────────────────────────────────────────────────── */}
      <div style={{ flex: 1, overflowY: 'auto', padding: '24px 24px 48px' }}>

        {/* ═══════════════ PROJECT TRACKER ═══════════════ */}
        {mod === 'projects' && (
          <div className={styles.fade}>
            <div style={{ display: 'flex', gap: 14, marginBottom: 22, flexWrap: 'wrap' }}>
              <Stat label="Total Projects" value={mainProjects.length} col="var(--bl)" sub={yr !== 'all' ? yr + ' year' : 'all time'} />
              <Stat label="Active" value={allActive} col="var(--3eg)" sub="in progress" />
              <Stat label="Over Budget" value={allOverBudget} col="var(--rd)" warn={allOverBudget > 0} sub="hrs exceeded" />
              <Stat label="Complete" value={allComplete} col="var(--bl)" sub="delivered" />
              <Stat label="EWOs" value={allEwos} col="var(--am)" sub="extra work orders" />
              <Stat label="Total Hrs Used" value={totalHrsUsed.toFixed(0) + 'h'} col={totalHrsAllowed > 0 && totalHrsUsed > totalHrsAllowed ? 'var(--rd)' : 'var(--gn)'} sub={'of ' + totalHrsAllowed.toFixed(0) + 'h allowed'} />
            </div>

            <div style={{ display: 'flex', gap: 10, marginBottom: 16, flexWrap: 'wrap', alignItems: 'center' }}>
              <input style={{ ...inp, maxWidth: 220 }} placeholder="Search projects..." value={srch} onChange={e => setSrch(e.target.value)} />
              <select style={{ ...selStyle, maxWidth: 120 }} value={yr} onChange={e => setYr(e.target.value)}>
                {years.map(y => <option key={y} value={y}>{y === 'all' ? 'All Years' : y}</option>)}
              </select>
              <select style={{ ...selStyle, maxWidth: 160 }} value={stFilt} onChange={e => setStFilt(e.target.value)}>
                <option value="">All Statuses</option>
                {PROJ_STATUSES.map(s => <option key={s} value={s}>{s}</option>)}
              </select>
              <button onClick={() => setShowArchived(!showArchived)} style={{
                fontFamily: 'Montserrat', fontWeight: 700, fontSize: 11, letterSpacing: '.06em',
                textTransform: 'uppercase', padding: '7px 14px', borderRadius: 6, cursor: 'pointer',
                background: showArchived ? 'var(--am)' : 'transparent',
                color: showArchived ? '#fff' : 'var(--t3)',
                border: `1px solid ${showArchived ? 'var(--am)' : 'var(--bd)'}`,
                transition: 'all .15s'
              }}>
                {showArchived ? 'Show Active' : 'Archived'}
              </button>
              <div style={{ flex: 1 }} />
              <button onClick={() => generateAllProjectsPdf(visProjects, rfis)} style={{
                fontFamily: 'Montserrat', fontWeight: 700, fontSize: 12, letterSpacing: '.08em',
                textTransform: 'uppercase', padding: '7px 18px', borderRadius: 6, cursor: 'pointer',
                background: 'transparent', color: 'var(--t2)', border: '1px solid var(--bd)',
                marginRight: 8
              }}>
                Export All
              </button>
              {isManager && (
                <button onClick={() => openProjForm(null)} style={{
                  fontFamily: 'Montserrat', fontWeight: 700, fontSize: 12, letterSpacing: '.08em',
                  textTransform: 'uppercase', padding: '7px 18px', borderRadius: 6, cursor: 'pointer',
                  background: 'var(--3eg)', color: '#1a2030', border: 'none',
                  boxShadow: '0 2px 8px rgba(42,158,42,.3)'
                }}>
                  + New Project
                </button>
              )}
            </div>

            <div style={{ background: 'var(--s1)', border: '1px solid var(--bd)', borderRadius: 8, boxShadow: '0 1px 6px rgba(0,0,0,.06)' }}>
              <div style={{ overflowX: 'auto', borderRadius: 8 }}>
                <table style={{ width: '100%', borderCollapse: 'collapse', minWidth: 1100, tableLayout: 'fixed' }}>
                  <colgroup>
                    <col style={{ width: 32 }} />
                    <col style={{ width: 88 }} />
                    <col style={{ width: 84 }} />
                    <col style={{ width: '20%' }} />
                    <col style={{ width: '11%' }} />
                    <col style={{ width: '10%' }} />
                    <col style={{ width: 88 }} />
                    <col style={{ width: 92 }} />
                    <col style={{ width: 92 }} />
                    <col style={{ width: 80 }} />
                    <col style={{ width: 108 }} />
                    <col style={{ width: 76 }} />
                  </colgroup>
                  <thead>
                    <tr>
                      <th style={{ width: 32, background: 'var(--s2)', borderBottom: '2px solid var(--bd)' }} />
                      <Th col="projNum" label="Project #" pad="8px 14px" />
                      <Th col="quoteNum" label="Quote #" pad="8px 14px" />
                      <Th col="name" label="Name" pad="8px 14px" />
                      <Th col="company" label="Company" pad="8px 10px" />
                      <Th col="contact" label="Contact" pad="8px 10px" />
                      <Th col="hrsUsed" label="Hours" pad="8px 10px" />
                      <Th col="startDate" label="Start" pad="8px 10px" />
                      <Th col="finishDate" label="Finish" pad="8px 10px" />
                      <ThPlain label="RFIs" pad="8px 10px" />
                      <Th col="status" label="Status" pad="8px 8px" />
                      <ThPlain label="Actions" pad="8px 10px" />
                    </tr>
                  </thead>
                  <tbody>
                    {spLoading && (
                      <tr><td colSpan={12} style={{ padding: '32px', textAlign: 'center', fontFamily: 'Montserrat', fontSize: 13, color: 'var(--t4)' }}>Loading projects...</td></tr>
                    )}
                    {!spLoading && visProjects.length === 0 && (
                      <tr><td colSpan={12} style={{ padding: '32px', textAlign: 'center', fontFamily: 'Montserrat', fontSize: 13, color: 'var(--t4)' }}>No projects found.</td></tr>
                    )}
                    {!spLoading && visProjects.map(p => {
                      const ewos = getEwos(p.id);
                      const expanded = !!exp[p.id];
                      const rfiCount = rfiCountByProj[p.id] || 0;
                      const rowBg = 'var(--s1)';
                      return (
                        <React.Fragment key={p.id}>
                          <tr style={{ background: rowBg, borderBottom: '1px solid var(--s3)' }}
                            onMouseEnter={e => { (e.currentTarget as HTMLTableRowElement).style.background = 'var(--s2)'; }}
                            onMouseLeave={e => { (e.currentTarget as HTMLTableRowElement).style.background = rowBg; }}>
                            <td style={{ padding: '0 0 0 8px', width: 32, textAlign: 'center' }}>
                              {ewos.length > 0 && (
                                <button onClick={() => setExp(prev => ({ ...prev, [p.id]: !prev[p.id] }))}
                                  style={{ background: 'var(--am2)', border: '1px solid var(--am)', borderRadius: 4, cursor: 'pointer', fontSize: 12, color: 'var(--am)', fontFamily: 'Montserrat', fontWeight: 700, padding: '2px 6px', lineHeight: 1, transition: 'all .15s' }}>
                                  {expanded ? '▾' : '▸'}
                                </button>
                              )}
                            </td>
                            <td style={{ padding: '9px 14px', fontFamily: 'Montserrat', fontWeight: 700, fontSize: 13, color: 'var(--3eg)', whiteSpace: 'nowrap', cursor: 'pointer' }} onClick={() => openProjDetail(p)}>{p.projNum}</td>
                            <td style={{ padding: '9px 14px', fontFamily: 'Montserrat', fontWeight: 600, fontSize: 12.5, color: 'var(--t3)', whiteSpace: 'nowrap' }}>{p.quoteNum || '—'}</td>
                            <td style={{ padding: '9px 14px', fontFamily: 'Montserrat', fontWeight: 600, fontSize: 12, color: 'var(--t1)', cursor: 'pointer', lineHeight: 1.35, overflow: 'hidden', overflowWrap: 'break-word', wordBreak: 'normal' }} onClick={() => openProjDetail(p)}>
                              <div style={{ display: '-webkit-box', WebkitLineClamp: 2, WebkitBoxOrient: 'vertical', overflow: 'hidden' }}>{p.name}</div>
                              {p.discipline && (
                                <div style={{ display: 'flex', gap: 3, marginTop: 3, flexWrap: 'wrap' }}>
                                  {(p.discipline === 'Steel & Concrete' ? ['Steel', 'Concrete'] : [p.discipline]).map(disc => (
                                    <span key={disc} style={{
                                      display: 'inline-block', padding: '1px 6px', borderRadius: 3,
                                      fontSize: 9, fontWeight: 700, letterSpacing: '.05em', textTransform: 'uppercase',
                                      background: disc === 'Concrete' ? 'rgba(107,79,200,0.12)' : 'rgba(37,99,235,0.12)',
                                      color: disc === 'Concrete' ? '#6b4fc8' : '#2563eb',
                                      border: `1px solid ${disc === 'Concrete' ? '#6b4fc8' : '#2563eb'}`
                                    }}>{disc}</span>
                                  ))}
                                </div>
                              )}
                            </td>
                            <td style={{ padding: '9px 10px', fontFamily: 'Montserrat', fontWeight: 600, fontSize: 11.5, color: 'var(--t2)', overflow: 'hidden', lineHeight: 1.3 }} title={p.company || ''}>
                              <div style={{ display: '-webkit-box', WebkitLineClamp: 2, WebkitBoxOrient: 'vertical', overflow: 'hidden', overflowWrap: 'break-word', wordBreak: 'normal' }}>{p.company || '—'}</div>
                            </td>
                            <td style={{ padding: '9px 10px', fontFamily: 'Montserrat', fontWeight: 600, fontSize: 11.5, color: 'var(--t2)', overflow: 'hidden', lineHeight: 1.3 }} title={p.contact || ''}>
                              <div style={{ display: '-webkit-box', WebkitLineClamp: 2, WebkitBoxOrient: 'vertical', overflow: 'hidden', overflowWrap: 'break-word', wordBreak: 'normal' }}>{p.contact || '—'}</div>
                            </td>
                            <td style={{ padding: '9px 10px', overflow: 'hidden' }}><HrsBar allowed={p.hrsAllowed} used={p.hrsUsed} compact /></td>
                            <td style={{ padding: '9px 10px', fontFamily: 'Montserrat', fontWeight: 600, fontSize: 11.5, color: 'var(--t3)', whiteSpace: 'nowrap' }}>{fmtD(p.startDate)}</td>
                            <td style={{ padding: '9px 10px', fontFamily: 'Montserrat', fontWeight: 600, fontSize: 11.5, color: 'var(--t3)', whiteSpace: 'nowrap' }}>{fmtD(p.finishDate)}</td>
                            <td style={{ padding: '9px 10px', minWidth: 0, overflow: 'hidden' }}><RfiBar allowed={p.rfisAllowed} used={rfiCount} /></td>
                            <td style={{ padding: '9px 8px', whiteSpace: 'nowrap' }}><Tag s={p.status} small /></td>
                            <td style={{ padding: '9px 10px', whiteSpace: 'nowrap' }}>
                              <div style={{ display: 'flex', flexDirection: 'column', gap: 4, alignItems: 'flex-start' }}>
                                <IBtn onClick={() => openProjDetail(p)} title="View details">View</IBtn>
                                {isManager && <IBtn onClick={() => openProjForm(p)} title="Edit project">Edit</IBtn>}
                                {isManager && <IBtn onClick={() => toggleArchive(p)} title={p.status === 'Archive' ? 'Restore project' : 'Archive project'}>{p.status === 'Archive' ? 'Restore' : 'Archive'}</IBtn>}
                              </div>
                            </td>
                          </tr>
                          {expanded && ewos.map(ewo => {
                            const ewoRfis = rfiCountByProj[ewo.id] || 0;
                            return (
                              <tr key={ewo.id} className={styles.ewoRow} style={{ background: 'rgba(42,158,42,0.035)', borderBottom: '1px solid var(--s3)' }}>
                                <td style={{ padding: '0 0 0 8px', width: 28 }} />
                                <td style={{ padding: '7px 6px 7px 20px', fontFamily: 'Montserrat', fontWeight: 700, fontSize: 12, color: 'var(--am)', whiteSpace: 'nowrap', cursor: 'pointer' }} onClick={() => openProjDetail(ewo)}>
                                  <span style={{ color: 'var(--t4)', fontWeight: 400, fontSize: 10, marginRight: 4 }}>EWO</span>{ewo.ewoNum || ewo.projNum}
                                </td>
                                <td style={{ padding: '7px 6px', fontFamily: 'Montserrat', fontSize: 12, color: 'var(--t4)' }}>{ewo.quoteNum || '—'}</td>
                                <td style={{ padding: '7px 6px', fontFamily: 'Montserrat', fontWeight: 600, fontSize: 12, color: 'var(--t2)', cursor: 'pointer', lineHeight: 1.35, overflowWrap: 'break-word', wordBreak: 'normal' }} onClick={() => openProjDetail(ewo)}>
                                  <div style={{ display: '-webkit-box', WebkitLineClamp: 2, WebkitBoxOrient: 'vertical', overflow: 'hidden' }}>{ewo.name}</div>
                                </td>
                                <td style={{ padding: '7px 6px', fontFamily: 'Montserrat', fontWeight: 600, fontSize: 12, color: 'var(--t3)', overflow: 'hidden', textOverflow: 'ellipsis', whiteSpace: 'nowrap' }}>{ewo.company || '—'}</td>
                                <td style={{ padding: '7px 6px', fontFamily: 'Montserrat', fontSize: 12, color: 'var(--t3)' }}>{ewo.contact || '—'}</td>
                                <td style={{ padding: '7px 6px' }}><HrsBar allowed={ewo.hrsAllowed} used={ewo.hrsUsed} compact /></td>
                                <td style={{ padding: '7px 6px', fontFamily: 'Montserrat', fontSize: 12, color: 'var(--t4)' }}>{fmtD(ewo.startDate)}</td>
                                <td style={{ padding: '7px 6px', fontFamily: 'Montserrat', fontSize: 12, color: 'var(--t4)' }}>{fmtD(ewo.finishDate)}</td>
                                <td style={{ padding: '7px 6px', minWidth: 110 }}><RfiBar allowed={ewo.rfisAllowed} used={ewoRfis} /></td>
                                <td style={{ padding: '7px 6px', whiteSpace: 'nowrap' }}><Tag s={ewo.status} small /></td>
                                <td style={{ padding: '7px 6px', whiteSpace: 'nowrap' }}>
                                  <div style={{ display: 'flex', flexDirection: 'column', gap: 4, alignItems: 'flex-start' }}>
                                    <IBtn onClick={() => openProjDetail(ewo)} title="View EWO">View</IBtn>
                                    {isManager && <IBtn onClick={() => openProjForm(ewo)} title="Edit EWO">Edit</IBtn>}
                                  </div>
                                </td>
                              </tr>
                            );
                          })}
                        </React.Fragment>
                      );
                    })}
                  </tbody>
                </table>
              </div>
              {!spLoading && visProjects.length > 0 && (
                <div style={{ padding: '10px 16px', borderTop: '1px solid var(--bd)', fontFamily: 'Montserrat', fontSize: 11.5, color: 'var(--t4)' }}>
                  Showing {visProjects.length} project{visProjects.length !== 1 ? 's' : ''}
                </div>
              )}
            </div>
          </div>
        )}

        {/* ═══════════════ RFI TRACKER ═══════════════ */}
        {mod === 'rfis' && (
          <div className={styles.fade}>
            <div style={{ display: 'flex', gap: 14, marginBottom: 22, flexWrap: 'wrap' }}>
              <Stat label="Total RFIs" value={rfis.length} col="var(--pu)" sub="all projects" />
              <Stat label="Open" value={rfiOpen} col="var(--gn)" sub="awaiting response" />
              <Stat label="Overdue" value={rfiOverdue} col="var(--rd)" warn={rfiOverdue > 0} sub="past due date" />
              <Stat label="Partial" value={rfiPartial} col="var(--am)" sub="revise and resend" />
              <Stat label="Closed" value={rfiClosed} col="var(--bl)" sub="resolved" />
              <Stat label="Impact Hrs" value={rfiImpact.toFixed(1) + 'h'} col="var(--pu)" sub="total tracked" />
            </div>

            <div style={{ display: 'flex', gap: 10, marginBottom: 16, flexWrap: 'wrap', alignItems: 'center' }}>
              <input style={{ ...inp, maxWidth: 220 }} placeholder="Search RFIs..." value={rfiSrch} onChange={e => setRfiSrch(e.target.value)} />
              <select style={{ ...selStyle, maxWidth: 260 }} value={rfiProj} onChange={e => setRfiProj(e.target.value)}>
                <option value="">All Projects</option>
                {projects.map(p => <option key={p.id} value={p.id}>{p.projNum} — {p.name}</option>)}
              </select>
              <select style={{ ...selStyle, maxWidth: 200 }} value={rfiSt} onChange={e => setRfiSt(e.target.value)}>
                <option value="">All Statuses</option>
                {RFI_STATUSES.map(s => <option key={s} value={s}>{s}</option>)}
                <option value="Overdue">Overdue</option>
              </select>
              <div style={{ flex: 1 }} />
              <button onClick={() => generateAllRfisPdf(visRfis, projects)} style={{
                fontFamily: 'Montserrat', fontWeight: 700, fontSize: 12, letterSpacing: '.08em',
                textTransform: 'uppercase', padding: '7px 18px', borderRadius: 6, cursor: 'pointer',
                background: 'transparent', color: 'var(--t2)', border: '1px solid var(--bd)',
                marginRight: 8
              }}>
                Export All
              </button>
              {isManager && (
                <button onClick={() => openRfiForm(null)} style={{
                  fontFamily: 'Montserrat', fontWeight: 700, fontSize: 12, letterSpacing: '.08em',
                  textTransform: 'uppercase', padding: '7px 18px', borderRadius: 6, cursor: 'pointer',
                  background: '#2563eb', color: '#fff', border: 'none',
                  boxShadow: '0 2px 8px rgba(37,99,235,.3)'
                }}>
                  + New RFI
                </button>
              )}
            </div>

            {spLoading && (
              <div style={{ padding: 32, textAlign: 'center', fontFamily: 'Montserrat', fontSize: 13, color: 'var(--t4)' }}>Loading RFIs...</div>
            )}
            {!spLoading && visRfis.length === 0 && (
              <div style={{ padding: 32, textAlign: 'center', fontFamily: 'Montserrat', fontSize: 13, color: 'var(--t4)' }}>No RFIs found.</div>
            )}
            {!spLoading && Object.keys(rfisByProject).map(projId => {
              const projRfis = rfisByProject[projId];
              const proj = projects.filter(p => p.id === projId)[0];
              const groupExpanded = rfiExp[projId] !== false;
              return (
                <div key={projId} style={{ marginBottom: 20, background: 'var(--s1)', border: '1px solid var(--bd)', borderRadius: 8, overflow: 'hidden', boxShadow: '0 1px 6px rgba(0,0,0,.06)' }}>
                  <div onClick={() => setRfiExp(prev => ({ ...prev, [projId]: !groupExpanded }))}
                    style={{ padding: '12px 18px', background: 'var(--s2)', borderBottom: '1px solid var(--bd)', display: 'flex', alignItems: 'center', gap: 12, cursor: 'pointer' }}>
                    <span style={{ fontFamily: 'Montserrat', fontWeight: 400, fontSize: 10, color: 'var(--t4)' }}>{groupExpanded ? 'v' : '>'}</span>
                    <span style={{ fontFamily: 'Montserrat', fontWeight: 800, fontSize: 13, color: 'var(--3eg)' }}>{proj ? proj.projNum : projId}</span>
                    <span style={{ fontFamily: 'Montserrat', fontWeight: 600, fontSize: 13, color: 'var(--t1)' }}>{proj ? proj.name : ''}</span>
                    <span style={{ fontFamily: 'Montserrat', fontSize: 11.5, color: 'var(--t4)', marginLeft: 4 }}>— {projRfis.length} RFI{projRfis.length !== 1 ? 's' : ''}</span>
                    {proj ? <Tag s={proj.status} /> : null}
                    <div style={{ flex: 1 }} />
                    {isManager && (
                      <button onClick={e => { e.stopPropagation(); openRfiForm(null, proj); }}
                        style={{ fontFamily: 'Montserrat', fontWeight: 600, fontSize: 11, padding: '3px 10px', borderRadius: 4, cursor: 'pointer', background: 'rgba(37,99,235,0.12)', border: '1px solid #2563eb', color: '#2563eb' }}>
                        + RFI
                      </button>
                    )}
                  </div>
                  {groupExpanded && (
                    <div style={{ overflowX: 'auto' }}>
                      <table style={{ width: '100%', borderCollapse: 'collapse' }}>
                        <thead>
                          <tr>
                            <Th col="rfiNum" label="RFI #" rfi />
                            <Th col="rfiType" label="Type" rfi />
                            <Th col="status" label="Status" rfi />
                            <Th col="dateIssued" label="Issued" rfi />
                            <Th col="dateRequired" label="Required" rfi />
                            <Th col="submittedTo" label="To" rfi />
                            <Th col="response" label="Response" rfi />
                            <Th col="impacted" label="Impact" rfi />
                            <ThPlain label="Actions" />
                          </tr>
                        </thead>
                        <tbody>
                          {projRfis.map(r => {
                            const st = effSt(r);
                            const overdue = isOD(r);
                            const rowBg = overdue ? 'rgba(204,51,51,0.03)' : 'var(--s1)';
                            return (
                              <tr key={r.id} style={{ background: rowBg, borderBottom: '1px solid var(--s3)' }}
                                onMouseEnter={ev => { (ev.currentTarget as HTMLTableRowElement).style.background = 'var(--s2)'; }}
                                onMouseLeave={ev => { (ev.currentTarget as HTMLTableRowElement).style.background = rowBg; }}>
                                <td style={{ padding: '10px 12px', fontFamily: 'Montserrat', fontWeight: 700, fontSize: 12.5, color: '#2563eb', whiteSpace: 'nowrap', cursor: 'pointer' }} onClick={() => openRfiDetail(r, proj)}>{r.rfiNum}</td>
                                <td style={{ padding: '10px 12px', fontFamily: 'Montserrat', fontWeight: 600, fontSize: 12, color: 'var(--t2)', whiteSpace: 'nowrap' }}>{r.rfiType}</td>
                                <td style={{ padding: '10px 12px', whiteSpace: 'nowrap' }}><Tag s={st} /></td>
                                <td style={{ padding: '10px 12px', fontFamily: 'Montserrat', fontWeight: 600, fontSize: 12, color: 'var(--t3)', whiteSpace: 'nowrap' }}>{fmtD(r.dateIssued)}</td>
                                <td style={{ padding: '10px 12px', fontFamily: 'Montserrat', fontSize: 12, color: overdue ? 'var(--rd)' : 'var(--t3)', fontWeight: overdue ? 700 : 600, whiteSpace: 'nowrap' }}>{fmtD(r.dateRequired)}</td>
                                <td style={{ padding: '10px 12px', fontFamily: 'Montserrat', fontWeight: 600, fontSize: 12, color: 'var(--t2)', whiteSpace: 'nowrap' }}>{r.submittedTo || '—'}</td>
                                <td style={{ padding: '10px 12px', fontFamily: 'Montserrat', fontWeight: 600, fontSize: 12, color: 'var(--t3)', whiteSpace: 'nowrap' }}>{r.response || '—'}</td>
                                <td style={{ padding: '10px 12px', fontFamily: 'Montserrat', fontWeight: 600, fontSize: 12, color: r.impacted === 'Yes' ? 'var(--am)' : 'var(--t4)', whiteSpace: 'nowrap' }}>
                                  {r.impacted === 'Yes' ? ('Yes (' + rfiTot(r).toFixed(1) + 'h)') : 'No'}
                                </td>
                                <td style={{ padding: '10px 12px', whiteSpace: 'nowrap' }}>
                                  <div style={{ display: 'flex', flexDirection: 'column', gap: 4, alignItems: 'flex-start' }}>
                                    <IBtn onClick={() => openRfiDetail(r, proj)} title="View RFI">View</IBtn>
                                    {isManager && <IBtn onClick={() => openRfiForm(r, proj)} title="Edit RFI">Edit</IBtn>}
                                  </div>
                                </td>
                              </tr>
                            );
                          })}
                        </tbody>
                      </table>
                    </div>
                  )}
                </div>
              );
            })}
          </div>
        )}

        {/* ═══════════════ EWO TRACKER ═══════════════ */}
        {mod === 'ewos' && (
          <div className={styles.fade}>
            <div style={{ display: 'flex', gap: 14, marginBottom: 22, flexWrap: 'wrap' }}>
              <Stat label="Total EWOs" value={visEwos.length} col="var(--am)" sub="extra work orders" />
              <Stat label="Active" value={visEwos.filter(e => e.status === 'Active').length} col="var(--gn)" sub="in progress" />
              <Stat label="Complete" value={visEwos.filter(e => e.status === 'Complete').length} col="var(--bl)" sub="delivered" />
              <Stat label="Total Hrs" value={visEwos.reduce((s, e) => s + e.hrsUsed, 0).toFixed(0) + 'h'} col="var(--am)" sub={'of ' + visEwos.reduce((s, e) => s + e.hrsAllowed, 0).toFixed(0) + 'h allowed'} />
            </div>

            <div style={{ display: 'flex', gap: 10, marginBottom: 16, flexWrap: 'wrap', alignItems: 'center' }}>
              <input style={{ ...inp, maxWidth: 220 }} placeholder="Search EWOs..." value={ewoSrch} onChange={e => setEwoSrch(e.target.value)} />
              <select style={{ ...selStyle, maxWidth: 260 }} value={ewoParent} onChange={e => setEwoParent(e.target.value)}>
                <option value="">All Projects</option>
                {mainProjects.map(p => <option key={p.id} value={p.id}>{p.projNum} — {p.name}</option>)}
              </select>
              <select style={{ ...selStyle, maxWidth: 200 }} value={ewoStFilt} onChange={e => setEwoStFilt(e.target.value)}>
                <option value="">All Statuses</option>
                {PROJ_STATUSES.map(s => <option key={s} value={s}>{s}</option>)}
              </select>
              <div style={{ flex: 1 }} />
              <button onClick={() => generateAllEwosPdf(visEwos, projects)} style={{
                fontFamily: 'Montserrat', fontWeight: 700, fontSize: 12, letterSpacing: '.08em',
                textTransform: 'uppercase', padding: '7px 18px', borderRadius: 6, cursor: 'pointer',
                background: 'transparent', color: 'var(--t2)', border: '1px solid var(--bd)',
                marginRight: 8
              }}>
                Export All
              </button>
              {isManager && (
                <button onClick={() => setPanel({ type: 'ewoForm', proj: null })} style={{
                  fontFamily: 'Montserrat', fontWeight: 700, fontSize: 12, letterSpacing: '.08em',
                  textTransform: 'uppercase', padding: '7px 18px', borderRadius: 6, cursor: 'pointer',
                  background: 'var(--am)', color: '#1a2030', border: 'none',
                  boxShadow: '0 2px 8px rgba(212,136,10,.3)'
                }}>
                  + New EWO
                </button>
              )}
            </div>

            {spLoading && (
              <div style={{ padding: 32, textAlign: 'center', fontFamily: 'Montserrat', fontSize: 13, color: 'var(--t4)' }}>Loading EWOs...</div>
            )}
            {!spLoading && visEwos.length === 0 && (
              <div style={{ padding: 32, textAlign: 'center', fontFamily: 'Montserrat', fontSize: 13, color: 'var(--t4)' }}>No EWOs found.</div>
            )}
            {!spLoading && Object.keys(ewosByParent).map(parentId => {
              const groupEwos = ewosByParent[parentId];
              const parent = projects.filter(p => p.id === parentId)[0];
              const groupExpanded = ewoExp[parentId] !== false;
              return (
                <div key={parentId} style={{ marginBottom: 20, background: 'var(--s1)', border: '1px solid var(--bd)', borderRadius: 8, overflow: 'hidden', boxShadow: '0 1px 6px rgba(0,0,0,.06)' }}>
                  <div onClick={() => setEwoExp(prev => { const next = { ...prev, [parentId]: !groupExpanded }; try { localStorage.setItem('3edge-ewo-exp', JSON.stringify(next)); } catch { /* ignore */ } return next; })}
                    style={{ padding: '12px 18px', background: 'var(--s2)', borderBottom: '1px solid var(--bd)', display: 'flex', alignItems: 'center', gap: 12, cursor: 'pointer' }}>
                    <span style={{ fontFamily: 'Montserrat', fontWeight: 400, fontSize: 10, color: 'var(--t4)' }}>{groupExpanded ? 'v' : '>'}</span>
                    <span style={{ fontFamily: 'Montserrat', fontWeight: 800, fontSize: 13, color: 'var(--3eg)' }}>{parent ? parent.projNum : parentId}</span>
                    <span style={{ fontFamily: 'Montserrat', fontWeight: 600, fontSize: 13, color: 'var(--t1)' }}>{parent ? parent.name : ''}</span>
                    <span style={{ fontFamily: 'Montserrat', fontSize: 11.5, color: 'var(--t4)', marginLeft: 4 }}>— {groupEwos.length} EWO{groupEwos.length !== 1 ? 's' : ''}</span>
                    {parent ? <Tag s={parent.status} /> : null}
                  </div>
                  {groupExpanded && (
                    <div style={{ overflowX: 'auto' }}>
                      <table style={{ width: '100%', borderCollapse: 'collapse' }}>
                        <thead>
                          <tr>
                            <ThPlain label="EWO #" />
                            <ThPlain label="Name" />
                            <ThPlain label="Company" />
                            <ThPlain label="Contact" />
                            <ThPlain label="Hours" />
                            <ThPlain label="Start" />
                            <ThPlain label="Status" />
                            <ThPlain label="Actions" />
                          </tr>
                        </thead>
                        <tbody>
                          {groupEwos.map(ewo => (
                            <tr key={ewo.id} style={{ background: 'var(--s1)', borderBottom: '1px solid var(--s3)' }}
                              onMouseEnter={ev => { (ev.currentTarget as HTMLTableRowElement).style.background = 'var(--s2)'; }}
                              onMouseLeave={ev => { (ev.currentTarget as HTMLTableRowElement).style.background = 'var(--s1)'; }}>
                              <td style={{ padding: '10px 12px', fontFamily: 'Montserrat', fontWeight: 700, fontSize: 12.5, color: 'var(--am)', whiteSpace: 'nowrap', cursor: 'pointer' }}
                                onClick={() => openProjDetail(ewo)}>{ewo.ewoNum || ewo.projNum}</td>
                              <td style={{ padding: '10px 12px', fontFamily: 'Montserrat', fontWeight: 600, fontSize: 12, color: 'var(--t1)', maxWidth: 180, overflow: 'hidden', textOverflow: 'ellipsis', whiteSpace: 'nowrap' }}>{ewo.name}</td>
                              <td style={{ padding: '10px 12px', fontFamily: 'Montserrat', fontWeight: 600, fontSize: 12, color: 'var(--t2)', whiteSpace: 'nowrap', maxWidth: 120, overflow: 'hidden', textOverflow: 'ellipsis' }}>{ewo.company || '—'}</td>
                              <td style={{ padding: '10px 12px', fontFamily: 'Montserrat', fontWeight: 600, fontSize: 12, color: 'var(--t2)', whiteSpace: 'nowrap' }}>{ewo.contact || '—'}</td>
                              <td style={{ padding: '10px 12px', fontFamily: 'Montserrat', fontWeight: 600, fontSize: 12, color: 'var(--t2)', whiteSpace: 'nowrap' }}>
                                {ewo.hrsUsed > 0 ? ewo.hrsUsed : '—'}{ewo.hrsAllowed > 0 ? (' / ' + ewo.hrsAllowed + 'h') : ''}
                              </td>
                              <td style={{ padding: '10px 12px', fontFamily: 'Montserrat', fontWeight: 600, fontSize: 12, color: 'var(--t3)', whiteSpace: 'nowrap' }}>{fmtD(ewo.startDate)}</td>
                              <td style={{ padding: '10px 12px', whiteSpace: 'nowrap' }}><Tag s={ewo.status} small /></td>
                              <td style={{ padding: '10px 12px', whiteSpace: 'nowrap' }}>
                                <div style={{ display: 'flex', flexDirection: 'column', gap: 4, alignItems: 'flex-start' }}>
                                  <IBtn onClick={() => openProjDetail(ewo)} title="View EWO">View</IBtn>
                                  {isManager && <IBtn onClick={() => setPanel({ type: 'ewoForm', proj: ewo })} title="Edit EWO">Edit</IBtn>}
                                </div>
                              </td>
                            </tr>
                          ))}
                        </tbody>
                      </table>
                    </div>
                  )}
                </div>
              );
            })}
          </div>
        )}

        {/* ═══════════════ TASKS ═══════════════ */}
        {mod === 'tasks' && (
          <div className={styles.fade}>
            <TaskBoard
              spService={spService.current}
              projects={projects}
              userDisplayName={props.userDisplayName}
              siteUrl={props.siteUrl}
              isManager={isManager}
              toast={toast}
            />
          </div>
        )}

        {/* ═══════════════ CHECKLIST ═══════════════ */}
        {mod === 'checklist' && (
          <div className={styles.fade}>
            <ChecklistBoard
              projects={projects}
              userDisplayName={props.userDisplayName}
              isManager={isManager}
              toast={toast}
              spService={spService.current}
            />
          </div>
        )}

        {/* ═══════════════ CRM ═══════════════ */}
        {mod === 'crm' && (
          <div className={styles.fade}>
            {!crmUnlocked ? (
              <div style={{ display: 'flex', flexDirection: 'column', alignItems: 'center', justifyContent: 'center', minHeight: 320, gap: 16 }}>
                <div style={{ fontFamily: 'Montserrat', fontWeight: 700, fontSize: 15, color: '#d3d1c7', letterSpacing: '.1em' }}>CRM ACCESS</div>
                <div style={{ fontFamily: 'Montserrat', fontSize: 12, color: '#8a9bb0' }}>
                  {crmLockedUntil ? 'Too many attempts — try again in' : 'Enter password to continue'}
                </div>
                {crmLockedUntil ? (
                  <div style={{ fontFamily: 'Montserrat', fontWeight: 700, fontSize: 22, color: '#c0392b', letterSpacing: '.05em' }}>
                    {`${String(Math.floor(crmLockRemain / 60)).padStart(2, '0')}:${String(crmLockRemain % 60).padStart(2, '0')}`}
                  </div>
                ) : (
                  <div style={{ display: 'flex', flexDirection: 'column', gap: 8, width: 280 }}>
                    <div style={{ position: 'relative' }}>
                      <input
                        type={crmPwShow ? 'text' : 'password'}
                        value={crmPw}
                        onChange={e => { setCrmPw(e.target.value); setCrmPwError(false); }}
                        onKeyDown={e => { if (e.key === 'Enter') crmTryUnlock(); }}
                        placeholder="Password"
                        autoFocus
                        style={{
                          width: '100%', boxSizing: 'border-box',
                          padding: '10px 40px 10px 14px', borderRadius: 5, fontSize: 13, fontFamily: 'Montserrat',
                          background: '#1a2030', color: '#d3d1c7',
                          border: crmPwError ? '1px solid #c0392b' : '1px solid rgba(138,155,176,.3)',
                          outline: 'none'
                        }}
                      />
                      <button
                        onClick={() => setCrmPwShow(s => !s)}
                        style={{
                          position: 'absolute', right: 10, top: '50%', transform: 'translateY(-50%)',
                          background: 'none', border: 'none', cursor: 'pointer', padding: 0,
                          color: '#8a9bb0', display: 'flex', alignItems: 'center'
                        }}
                        tabIndex={-1}
                      >
                        {crmPwShow ? (
                          <svg width="18" height="18" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round">
                            <path d="M17.94 17.94A10.07 10.07 0 0 1 12 20c-7 0-11-8-11-8a18.45 18.45 0 0 1 5.06-5.94"/>
                            <path d="M9.9 4.24A9.12 9.12 0 0 1 12 4c7 0 11 8 11 8a18.5 18.5 0 0 1-2.16 3.19"/>
                            <line x1="1" y1="1" x2="23" y2="23"/>
                          </svg>
                        ) : (
                          <svg width="18" height="18" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round">
                            <path d="M1 12s4-8 11-8 11 8 11 8-4 8-11 8-11-8-11-8z"/>
                            <circle cx="12" cy="12" r="3"/>
                          </svg>
                        )}
                      </button>
                    </div>
                    {crmPwError && (
                      <div style={{ fontFamily: 'Montserrat', fontSize: 11, color: '#c0392b' }}>
                        Incorrect password — {3 - crmAttempts} attempt{3 - crmAttempts === 1 ? '' : 's'} remaining
                      </div>
                    )}
                    <button
                      onClick={crmTryUnlock}
                      style={{
                        padding: '10px 0', borderRadius: 5, border: 'none', cursor: 'pointer',
                        background: 'var(--3eg)', color: '#fff', fontFamily: 'Montserrat', fontWeight: 700,
                        fontSize: 12, letterSpacing: '.1em', textTransform: 'uppercase'
                      }}
                    >
                      Unlock
                    </button>
                  </div>
                )}
              </div>
            ) : (
              <CrmBoard spService={spService.current} />
            )}
          </div>
        )}
      </div>

      {/* ── Slide-over Panels ──────────────────────────────────── */}

      <Panel
        open={panel.type === 'projDetail'}
        onClose={() => setPanel({ type: null })}
        title={panel.proj ? panel.proj.projNum : ''}
        subtitle={panel.proj ? panel.proj.name : ''}
        tag={panel.proj ? <Tag s={panel.proj.status} /> : undefined}
      >
        {panel.type === 'projDetail' && panel.proj && (
          <ProjDetail
            proj={panel.proj}
            rfis={rfis}
            isManager={isManager}
            onEdit={() => setPanel({ type: 'projForm', proj: panel.proj })}
            onDelete={() => {
              const pRef = panel.proj!;
              confirmDelete('Delete project "' + pRef.projNum + ' — ' + pRef.name + '"?', () => { deleteProject(pRef).catch(() => undefined); });
            }}
            onNewRfi={() => {
              const pRef = panel.proj!;
              const newRfi = emptyRfi();
              newRfi.projectId = pRef.id;
              newRfi.projectName = pRef.name;
              setPanel({ type: 'rfiForm', rfi: newRfi, parentProj: pRef });
            }}
            onViewRfi={(r) => setPanel({ type: 'rfiDetail', rfi: r, parentProj: panel.proj })}
          />
        )}
      </Panel>

      <Panel
        open={panel.type === 'projForm'}
        onClose={() => setPanel({ type: null })}
        title={panel.proj ? ('Edit Project — ' + panel.proj.projNum) : 'New Project'}
        subtitle={panel.proj ? panel.proj.name : 'Fill in the details below'}
      >
        {panel.type === 'projForm' && (
          <ProjForm
            initial={panel.proj || (() => {
              const p = emptyProj();
              const nums = projects
                .map(x => x.projNum.startsWith('3E-') ? parseInt(x.projNum.slice(3), 10) : NaN)
                .filter(n => !isNaN(n));
              p.projNum = '3E-' + (nums.length > 0 ? Math.max(...nums) + 1 : 500);
              return p;
            })()}
            isNew={!panel.proj}
            projects={projects}
            spService={spService.current}
            siteUrl={props.siteUrl}
            onSave={(d, files, removed) => { saveProject(d, !panel.proj, files, removed).catch(() => undefined); }}
            onCancel={() => setPanel({ type: null })}
          />
        )}
      </Panel>

      <Panel
        open={panel.type === 'rfiDetail'}
        onClose={() => setPanel({ type: null })}
        title={panel.rfi ? ('RFI ' + panel.rfi.rfiNum) : ''}
        subtitle={panel.rfi ? (panel.rfi.rfiType + ' — ' + panel.rfi.projectName) : ''}
        tag={panel.rfi ? <Tag s={effSt(panel.rfi)} /> : undefined}
      >
        {panel.type === 'rfiDetail' && panel.rfi && (
          <RfiDetail
            rfi={panel.rfi}
            proj={panel.parentProj || projects.filter(p => p.id === panel.rfi!.projectId)[0]}
            isManager={isManager}
            siteUrl={props.siteUrl}
            spService={spService.current}
            onNotify={toast}
            onSendEmail={async (to, cc, subject, body, _pdfFileName) => {
              try {
                const plainBody = body.replace(/<br\s*\/?>/gi, '\n').replace(/<[^>]*>/g, '');
                let mailto = 'mailto:' + encodeURIComponent(to) +
                  '?subject=' + encodeURIComponent(subject) +
                  '&body=' + encodeURIComponent(plainBody);
                if (cc.trim()) {
                  mailto += '&cc=' + encodeURIComponent(cc.trim());
                }
                const a = document.createElement('a');
                a.href = mailto;
                a.click();
                toast('PDF downloaded. Email draft opened — attach the PDF before sending.');
              } catch (e) {
                toast('Failed: ' + String(e));
              }
            }}
            onEdit={() => setPanel({ type: 'rfiForm', rfi: panel.rfi, parentProj: panel.parentProj })}
          />
        )}
      </Panel>

      <Panel
        open={panel.type === 'rfiForm'}
        onClose={() => setPanel({ type: null })}
        title={(panel.rfi && panel.rfi.rfiNum) ? ('Edit RFI — ' + panel.rfi.rfiNum) : 'New RFI'}
        subtitle={panel.parentProj ? (panel.parentProj.projNum + ' — ' + panel.parentProj.name) : 'Fill in the details below'}
      >
        {panel.type === 'rfiForm' && (
          <RfiForm
            initial={(() => {
              if (panel.rfi) return panel.rfi;
              const r = emptyRfi();
              if (panel.parentProj) { r.projectId = panel.parentProj.id; r.projectName = panel.parentProj.name; }
              return r;
            })()}
            isNew={!panel.rfi || !panel.rfi.spId}
            projects={projects}
            rfis={rfis}
            userDisplayName={props.userDisplayName}
            teamMembers={teamMembers}
            onSave={(d, files) => { saveRfi(d, !panel.rfi || !panel.rfi.spId, files).catch(() => undefined); }}
            onCancel={() => setPanel({ type: null })}
          />
        )}
      </Panel>

      <Panel
        open={panel.type === 'ewoForm'}
        onClose={() => setPanel({ type: null })}
        title={panel.proj && panel.proj.spId ? ('Edit EWO — ' + (panel.proj.ewoNum || panel.proj.projNum)) : 'New EWO'}
        subtitle="Fill in the EWO details below"
      >
        {panel.type === 'ewoForm' && (
          <EwoForm
            initial={(() => {
              if (panel.proj) return panel.proj;
              const p = emptyProj();
              p.isEwo = true;
              if (panel.parentProj) { p.parentId = panel.parentProj.id; }
              return p;
            })()}
            isNew={!panel.proj || !panel.proj.spId}
            projects={projects}
            onSave={(d) => { saveProject(d, !panel.proj || !panel.proj.spId).catch(() => undefined); }}
            onCancel={() => setPanel({ type: null })}
          />
        )}
      </Panel>

      {/* ── Delete Confirmation Modal ──────────────────────────── */}
      <DelModal
        open={del.open}
        label={del.label}
        onConfirm={() => { del.onConfirm(); }}
        onCancel={() => setDel({ open: false, label: '', onConfirm: () => undefined })}
      />

      {/* ── Time Doctor Import Modal ──────────────────────────── */}
      {tdModal && (
        <TdImportModal
          projects={projects}
          onClose={() => setTdModal(false)}
          onApply={(updates) => { applyTdUpdates(updates).catch(() => undefined); }}
          onResetHours={() => { resetAllHours().catch(() => undefined); }}
          lastImport={lastTdImport || undefined}
        />
      )}

      {/* ── Toast ─────────────────────────────────────────────── */}
      {Toast}
    </div>
  );
};

export default ManagerDashboard;
