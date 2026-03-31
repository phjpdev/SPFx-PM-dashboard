import * as React from 'react';
import {
  Tag, HrsBar, RfiBar, Stat, Panel, FF, IBtn, DelModal, useToast,
  BtnPrimary, SDiv, CcField, fmtD, rfiTot, effSt, isOD
} from '../../../shared/components/SharedComponents';
import { IProject, IRfi, PROJ_STATUSES, RFI_STATUSES, RFI_TYPES, RFI_RESPONSES } from '../../../shared/models/IProject';
import { SharePointService } from '../../../shared/services/SharePointService';
import styles from './ManagerDashboard.module.scss';
import type { IManagerDashboardProps } from './IManagerDashboardProps';

import { jsPDF } from 'jspdf';
import * as XLSX from 'xlsx';
import TaskBoard from './TaskBoard';

// ── Assets ────────────────────────────────────────────────────────────────────
import logoImg from '../assets/3edge-logo.png';
import pdfBgImg from '../assets/pdf-backgroundimage.png';
const IMG_LOGO_DASH: string = logoImg;

// ── Montserrat local fonts ─────────────────────────────────────────────────────
import _fExtraLight from '../assets/Montserrat-ExtraLight.ttf';
import _fBold from '../assets/Montserrat-Bold.ttf';
import _fBoldI from '../assets/Montserrat-BoldItalic.ttf';
import _fExtraBold from '../assets/Montserrat-ExtraBold.ttf';
import _fExtraBoldI from '../assets/Montserrat-ExtraBoldItalic.ttf';
import _fBlack from '../assets/Montserrat-Black.ttf';
import _fBlackI from '../assets/Montserrat-BlackItalic.ttf';

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
type Mod = 'projects' | 'rfis' | 'ewos' | 'tasks';
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

// ── Shared Letterhead for all PDFs (uses actual letterhead images) ────────────
// eslint-disable-next-line @typescript-eslint/no-explicit-any
function drawPdfBg(doc: any, pw: number, ph: number): void {
  doc.addImage(pdfBgImg, 'PNG', 0, 0, pw, ph);
}

// eslint-disable-next-line @typescript-eslint/no-explicit-any
function drawLetterhead(doc: any, pw: number, ph: number, title: string, subtitle: string): number {
  drawPdfBg(doc, pw, ph);
  // Title bar below header area
  const barY = 40;
  const barH = 12;
  doc.setFillColor(26, 32, 48);
  doc.rect(0, barY, pw, barH, 'F');
  doc.setFillColor(42, 158, 42);
  doc.rect(0, barY, 3, barH, 'F');
  doc.setFontSize(11); doc.setFont('helvetica', 'bold'); doc.setTextColor(255, 255, 255);
  doc.text(title, 8, barY + 7.5);
  if (subtitle) {
    doc.setFontSize(8); doc.setFont('helvetica', 'normal'); doc.setTextColor(160, 175, 195);
    doc.text(subtitle, pw - 8, barY + 7.5, { align: 'right' });
  }
  return barY + barH + 4;
}

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
  onSave: (p: IProject) => void;
  onCancel: () => void;
}

const ProjForm: React.FC<ProjFormProps> = ({ initial, isNew, projects, onSave, onCancel }) => {
  const [dupError, setDupError] = React.useState('');
  const [valError, setValError] = React.useState('');

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
    if (missing.length > 0) {
      setValError('Required: ' + missing.join(', '));
      return;
    }
    setValError('');
    if (usedNums.has(d.projNum.toUpperCase())) {
      setDupError('Project # ' + d.projNum + ' is already in use. Choose a different number.');
      return;
    }
    onSave(d);
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
          <input style={inp} type="number" step="0.5" value={d.hrsUsed} onChange={e => set('hrsUsed', Number(e.target.value))} />
        </FF>
        <FF label="RFIs Allowed">
          <input style={inp} type="number" value={d.rfisAllowed} onChange={e => set('rfisAllowed', Number(e.target.value))} />
        </FF>
      </div>

      <SDiv label="Notes" />
      <FF label="Project Notes">
        <textarea style={{ ...inp, minHeight: 70, resize: 'vertical' }} value={d.notes} onChange={e => set('notes', e.target.value)} placeholder="Add notes..." />
      </FF>

      <SDiv label="Invoices" />
      {(d.invoices.length > 0 ? d.invoices : []).map((inv, idx) => (
        <div key={idx} style={{ display: 'grid', gridTemplateColumns: '1fr 1fr auto auto', gap: '8px', alignItems: 'end', marginBottom: 8 }}>
          <FF label={idx === 0 ? 'Invoice Number' : ''}>
            <input style={inp} value={inv.invNumber} onChange={e => { const invs = [...d.invoices]; invs[idx] = { ...invs[idx], invNumber: e.target.value }; set('invoices', invs); }} placeholder="INV-001" />
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
      {d.invoices.length < 4 && (
        <button onClick={() => set('invoices', [...d.invoices, { invNumber: '', invDate: '', invPaid: false }])} style={{ fontFamily: 'Montserrat', fontSize: 11, fontWeight: 600, padding: '6px 14px', background: 'transparent', border: '1px dashed var(--bd)', color: 'var(--t3)', borderRadius: 5, cursor: 'pointer', marginTop: 4 }}>+ Add Invoice</button>
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
          <input style={inp} type="number" value={d.hrsUsed} onChange={e => set('hrsUsed', Number(e.target.value))} />
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
  onSave: (r: IRfi, files: File[]) => void;
  onCancel: () => void;
}

const RfiForm: React.FC<RfiFormProps> = ({ initial, isNew, projects, rfis, onSave, onCancel }) => {
  const [d, setD] = React.useState<IRfi>({ ...initial });
  const [rfiValError, setRfiValError] = React.useState('');
  const [pendingFiles, setPendingFiles] = React.useState<File[]>([]);
  const fileRef = React.useRef<HTMLInputElement>(null);

  const set = <K extends keyof IRfi>(k: K, v: IRfi[K]): void => {
    setD(prev => ({ ...prev, [k]: v }));
  };

  const onProjectChange = (projId: string): void => {
    const p = projects.find(x => x.id === projId);
    const updates: Partial<IRfi> = { projectId: projId, projectName: p ? p.name : '' };
    if (isNew && p) {
      const count = rfis.filter(r => r.projectId === projId).length;
      const seq = String(count + 1).padStart(3, '0');
      updates.rfiNum = `${p.projNum}-RFI-${seq}`;
      // Inherit company details from parent project
      if (p.contact) updates.submittedTo = p.contact;
      if (p.company) updates.toCompany = p.company;
      if (p.email) updates.email = p.email;
      updates.byCompany = '3 Edge Design';
    }
    setD(prev => ({ ...prev, ...updates }));
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
  onSendEmail: (to: string, cc: string, subject: string, body: string) => Promise<void>;
  onEdit: () => void;
}

const RfiDetail: React.FC<RfiDetailProps> = ({ rfi, proj, isManager, siteUrl, spService, onSendEmail, onEdit }) => {
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

    // Auto-download the PDF
    const fileName = 'RFI_' + rfi.rfiNum.replace(/[^a-zA-Z0-9_-]/g, '_') + '.pdf';
    const url = URL.createObjectURL(blob);
    const a = document.createElement('a');
    a.href = url;
    a.download = fileName;
    a.click();
    setTimeout(() => URL.revokeObjectURL(url), 3000);

    // Send email via SharePoint if email is set
    const recipients = rfi.email || '';
    if (!recipients) return;
    const subject = 'RFI ' + rfi.rfiNum + ' — ' + (proj ? proj.name : rfi.projectName);
    const body =
      'Dear ' + (rfi.submittedTo || 'Client') + ',<br><br>' +
      'Please find attached RFI <strong>' + rfi.rfiNum + '</strong> for your review and response.<br><br>' +
      '<strong>Project:</strong> ' + (proj ? proj.name : rfi.projectName) + '<br>' +
      '<strong>RFI Type:</strong> ' + rfi.rfiType + '<br>' +
      '<strong>Date Issued:</strong> ' + fmtD(rfi.dateIssued) + '<br>' +
      '<strong>Date Required:</strong> ' + fmtD(rfi.dateRequired) + '<br><br>' +
      '<strong>Description:</strong><br>' + (rfi.description || '—') + '<br><br>' +
      'Please respond by ' + fmtD(rfi.dateRequired) + '.<br><br>' +
      'Kind regards,<br>' + (rfi.by || '') + '<br>3 Edge Design';
    onSendEmail(recipients, rfi.cc || '', subject, body).catch(console.error);
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
          // Extract 3E-XXX pattern from project name like "01 - 3E-500 SAMPLE TASK"
          const projNumMatch = xlsName.match(/3E-\d+/i);
          const match = projects.find(p => {
            if (projNumMatch && p.projNum.toLowerCase() === projNumMatch[0].toLowerCase()) return true;
            if (p.name && xlsName.toLowerCase().indexOf(p.name.toLowerCase()) >= 0) return true;
            if (p.name && p.name.toLowerCase().indexOf(xlsName.toLowerCase()) >= 0) return true;
            return false;
          });
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
            <button onClick={() => { if (confirm('Reset ALL project hours to 0? This cannot be undone.')) onResetHours(); }} style={{ fontFamily: 'Montserrat', fontSize: 12.5, padding: '9px 18px', background: 'var(--rd)', border: 'none', color: '#fff', borderRadius: 7, cursor: 'pointer', fontWeight: 700 }}>Reset All Hours</button>
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
  const [ewoExp, setEwoExp] = React.useState<Record<string, boolean>>({});

  // ── Panel
  const [panel, setPanel] = React.useState<PanelState>({ type: null });

  // ── Delete modal
  const [del, setDel] = React.useState<DelState>({ open: false, label: '', onConfirm: () => undefined });

  // ── Time Doctor
  const [tdModal, setTdModal] = React.useState(false);
  const [lastTdImport, setLastTdImport] = React.useState<string | null>(null);

  // ── Load data
  const loadData = React.useCallback(async () => {
    setSpLoading(true);
    setSpMode('detecting');
    try {
      const [p, r] = await Promise.all([
        spService.current.loadProjects(),
        spService.current.loadRfis()
      ]);
      setProjects(p);
      setRfis(r);
      setSpMode('live');
      // Load last TD import timestamp
      spService.current.getSetting('lastTdImport').then(v => { if (v) setLastTdImport(v); }).catch(() => undefined);
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
  const allComplete = mainProjects.filter(p => p.status === 'Complete').length;
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
  const saveProject = async (d: IProject, isNew: boolean): Promise<void> => {
    try {
      if (isLocal()) {
        if (isNew) {
          const tempId = nextLocalId();
          const saved: IProject = { ...d, id: d.projNum || String(tempId) };
          setProjects(prev => [...prev, saved]);
          toast('Project created (local mode — will sync when SP lists are ready).');
        } else {
          setProjects(prev => prev.map(p => p.id === d.id ? { ...d } : p));
          toast('Project saved (local mode).');
        }
        setPanel({ type: null });
        return;
      }
      if (isNew) {
        const spId = await spService.current.addProject(d);
        const saved: IProject = { ...d, id: d.projNum || String(spId), spId };
        setProjects(prev => [...prev, saved]);
        toast('Project created.');
      } else {
        if (!d.spId) throw new Error('No spId on project');
        await spService.current.updateProject(d.spId, d);
        setProjects(prev => prev.map(p => p.id === d.id ? { ...d } : p));
        toast('Project saved.');
      }
      setPanel({ type: null });
    } catch (e) {
      const msg = (e instanceof Error) ? e.message : String(e);
      toast('Save failed: ' + msg, 'error');
    }
  };

  const toggleArchive = async (proj: IProject): Promise<void> => {
    const newStatus = proj.status === 'Archive' ? 'Active' : 'Archive';
    const updated = { ...proj, status: newStatus };
    try {
      if (!isLocal() && proj.spId) {
        await spService.current.updateProject(proj.spId, updated);
      }
      setProjects(prev => prev.map(p => p.id === proj.id ? updated : p));
      toast(newStatus === 'Archive' ? 'Project archived.' : 'Project restored.');
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
        await spService.current.updateProject(p.spId, updated);
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
      if (!p.spId || p.hrsUsed === 0) continue;
      try {
        const updated: IProject = { ...p, hrsUsed: 0 };
        await spService.current.updateProject(p.spId, updated);
        setProjects(prev => prev.map(x => x.id === p.id ? updated : x));
        success++;
      } catch (e) {
        toast('Failed to reset ' + p.projNum + ': ' + ((e instanceof Error) ? e.message : String(e)), 'error');
      }
    }
    toast('Reset hours: ' + success + ' project' + (success !== 1 ? 's' : '') + ' set to 0.');
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
  const Th: React.FC<{ col: string; label: string; rfi?: boolean }> = ({ col, label, rfi: isRfi }) => {
    const active = isRfi ? rSCol : sCol;
    const dir = isRfi ? rSDir : sDir;
    return (
      <th onClick={() => isRfi ? onRSort(col) : onSort(col)}
        style={{ padding: '8px 6px', fontFamily: 'Montserrat', fontWeight: 700, fontSize: 11, letterSpacing: '.1em', textTransform: 'uppercase', color: active === col ? 'var(--3eg)' : 'var(--t3)', cursor: 'pointer', whiteSpace: 'nowrap', borderBottom: '2px solid var(--bd)', textAlign: 'left', userSelect: 'none', background: 'var(--s2)' }}>
        {label}<span style={{ opacity: 0.6 }}>{sortArrow(col, active, dir)}</span>
      </th>
    );
  };

  // ── Plain th
  const ThPlain: React.FC<{ label: string }> = ({ label }) => (
    <th style={{ padding: '8px 6px', fontFamily: 'Montserrat', fontWeight: 700, fontSize: 11, letterSpacing: '.1em', textTransform: 'uppercase', color: 'var(--t3)', whiteSpace: 'nowrap', borderBottom: '2px solid var(--bd)', textAlign: 'left', background: 'var(--s2)' }}>{label}</th>
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
        <div style={{ display: 'flex', alignItems: 'center', gap: 10, marginRight: 22 }}>
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
        <div style={{ display: 'flex', gap: 2 }}>
          {(['projects', 'rfis', 'ewos', 'tasks'] as Mod[]).map((m: Mod) => (
            <button key={m} onClick={() => setMod(m)} style={{
              fontFamily: 'Montserrat', fontWeight: 700, fontSize: 11, letterSpacing: '.12em',
              textTransform: 'uppercase', padding: '6px 16px', borderRadius: 4, cursor: 'pointer',
              background: mod === m ? 'var(--3eg3)' : 'transparent',
              border: mod === m ? '1px solid var(--3eg)' : '1px solid transparent',
              color: mod === m ? 'var(--3eg)' : '#8a9bb0', transition: 'all .15s'
            }}>
              {m === 'projects' ? 'Projects' : m === 'rfis' ? 'RFIs' : m === 'ewos' ? 'EWOs' : 'Tasks'}
            </button>
          ))}
        </div>

        <div style={{ flex: 1 }} />

        {/* Time Doctor (manager only) */}
        {mod === 'projects' && isManager && (
          <div style={{ display: 'flex', alignItems: 'center', gap: 8, marginRight: 8 }}>
            {/* {lastTdImport && (
              <span style={{ fontFamily: 'Montserrat', fontSize: 10, color: 'var(--t4)', whiteSpace: 'nowrap' }}>
                Last import: {(() => { const d = new Date(lastTdImport); const days = ['Sun','Mon','Tue','Wed','Thu','Fri','Sat']; const months = ['Jan','Feb','Mar','Apr','May','Jun','Jul','Aug','Sep','Oct','Nov','Dec']; const h = d.getHours(); const ampm = h >= 12 ? 'PM' : 'AM'; const h12 = h % 12 || 12; return `${days[d.getDay()]} ${d.getDate()} ${months[d.getMonth()]}, ${h12}:${String(d.getMinutes()).padStart(2,'0')} ${ampm}`; })()}
              </span>
            )} */}
            <button onClick={() => setTdModal(true)} style={{
              fontFamily: 'Montserrat', fontWeight: 600, fontSize: 11, letterSpacing: '.1em',
              textTransform: 'uppercase', padding: '5px 14px', borderRadius: 4, cursor: 'pointer',
              background: 'rgba(212,136,10,0.14)', border: '1px solid var(--am)', color: 'var(--am)',
              whiteSpace: 'nowrap'
            }}>
              Time Doctor Import
            </button>
          </div>
        )}

        {/* SP Status indicator */}
        <div style={{ display: 'flex', alignItems: 'center', gap: 5, marginRight: 14 }}>
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

        {/* Role toggle — only visible for Owners */}
        {userRole === 'owner' && (
          <div style={{ display: 'flex', border: '1px solid rgba(138,155,176,.3)', borderRadius: 4, overflow: 'hidden', marginRight: 14 }}>
            {(['manager', 'staff'] as Role[]).map(r => (
              <button key={r} onClick={() => setRole(r)} style={{
                fontFamily: 'Montserrat', fontWeight: 700, fontSize: 10.5, letterSpacing: '.1em',
                textTransform: 'uppercase', padding: '4px 12px', cursor: 'pointer', border: 'none',
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
        <div style={{ display: 'flex', flexDirection: 'column', alignItems: 'flex-end', whiteSpace: 'nowrap', gap: 2 }}>
          <div style={{ fontFamily: 'Montserrat', fontWeight: 600, fontSize: 11, color: '#8a9bb0' }}>
            <span style={{ color: '#5a7a9a', marginRight: 4 }}>AUS</span>{clock.aus}
          </div>
          <div style={{ fontFamily: 'Montserrat', fontWeight: 600, fontSize: 11, color: '#8a9bb0' }}>
            <span style={{ color: '#5a7a9a', marginRight: 4 }}>PH</span>{clock.ph}
          </div>
        </div>
        {props.userDisplayName && (
          <div style={{ fontFamily: 'Montserrat', fontWeight: 600, fontSize: 11, color: '#8a9bb0', marginLeft: 12, whiteSpace: 'nowrap' }}>{props.userDisplayName}</div>
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
                <table style={{ width: '100%', borderCollapse: 'collapse', minWidth: 900 }}>
                  <thead>
                    <tr>
                      <th style={{ width: 20, background: 'var(--s2)', borderBottom: '2px solid var(--bd)' }} />
                      <Th col="projNum" label="Project #" />
                      <Th col="quoteNum" label="Quote #" />
                      <Th col="name" label="Name" />
                      <Th col="company" label="Company" />
                      <Th col="contact" label="Contact" />
                      <Th col="hrsUsed" label="Hours" />
                      <Th col="startDate" label="Start" />
                      <Th col="finishDate" label="Finish" />
                      <ThPlain label="RFIs" />
                      <Th col="status" label="Status" />
                      <ThPlain label="Actions" />
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
                            <td style={{ padding: '0 0 0 8px', width: 28, textAlign: 'center' }}>
                              {ewos.length > 0 && (
                                <button onClick={() => setExp(prev => ({ ...prev, [p.id]: !prev[p.id] }))}
                                  style={{ background: 'var(--am2)', border: '1px solid var(--am)', borderRadius: 4, cursor: 'pointer', fontSize: 12, color: 'var(--am)', fontFamily: 'Montserrat', fontWeight: 700, padding: '2px 6px', lineHeight: 1, transition: 'all .15s' }}>
                                  {expanded ? '▾' : '▸'}
                                </button>
                              )}
                            </td>
                            <td style={{ padding: '9px 6px', fontFamily: 'Montserrat', fontWeight: 700, fontSize: 13, color: 'var(--3eg)', whiteSpace: 'nowrap', cursor: 'pointer' }} onClick={() => openProjDetail(p)}>{p.projNum}</td>
                            <td style={{ padding: '9px 6px', fontFamily: 'Montserrat', fontWeight: 600, fontSize: 12.5, color: 'var(--t3)', whiteSpace: 'nowrap' }}>{p.quoteNum || '—'}</td>
                            <td style={{ padding: '9px 6px', fontFamily: 'Montserrat', fontWeight: 600, fontSize: 12, color: 'var(--t1)', cursor: 'pointer', maxWidth: 160, wordBreak: 'break-word' }} onClick={() => openProjDetail(p)}>
                              <div>{p.name}</div>
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
                            <td style={{ padding: '9px 6px', fontFamily: 'Montserrat', fontWeight: 600, fontSize: 11.5, color: 'var(--t2)', maxWidth: 120, wordBreak: 'break-word' }}>{p.company || '—'}</td>
                            <td style={{ padding: '9px 6px', fontFamily: 'Montserrat', fontWeight: 600, fontSize: 11.5, color: 'var(--t2)', whiteSpace: 'nowrap' }}>{p.contact || '—'}</td>
                            <td style={{ padding: '9px 6px', minWidth: 90 }}><HrsBar allowed={p.hrsAllowed} used={p.hrsUsed} /></td>
                            <td style={{ padding: '9px 6px', fontFamily: 'Montserrat', fontWeight: 600, fontSize: 11.5, color: 'var(--t3)', whiteSpace: 'nowrap' }}>{fmtD(p.startDate)}</td>
                            <td style={{ padding: '9px 6px', fontFamily: 'Montserrat', fontWeight: 600, fontSize: 11.5, color: 'var(--t3)', whiteSpace: 'nowrap' }}>{fmtD(p.finishDate)}</td>
                            <td style={{ padding: '9px 6px', minWidth: 80 }}><RfiBar allowed={p.rfisAllowed} used={rfiCount} /></td>
                            <td style={{ padding: '9px 4px', whiteSpace: 'nowrap' }}><Tag s={p.status} small /></td>
                            <td style={{ padding: '9px 6px', whiteSpace: 'nowrap' }}>
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
                                <td style={{ padding: '7px 6px', fontFamily: 'Montserrat', fontWeight: 600, fontSize: 12, color: 'var(--t2)', cursor: 'pointer', overflow: 'hidden', textOverflow: 'ellipsis', whiteSpace: 'nowrap' }} onClick={() => openProjDetail(ewo)}>{ewo.name}</td>
                                <td style={{ padding: '7px 6px', fontFamily: 'Montserrat', fontWeight: 600, fontSize: 12, color: 'var(--t3)', overflow: 'hidden', textOverflow: 'ellipsis', whiteSpace: 'nowrap' }}>{ewo.company || '—'}</td>
                                <td style={{ padding: '7px 6px', fontFamily: 'Montserrat', fontSize: 12, color: 'var(--t3)' }}>{ewo.contact || '—'}</td>
                                <td style={{ padding: '7px 6px', minWidth: 130 }}><HrsBar allowed={ewo.hrsAllowed} used={ewo.hrsUsed} /></td>
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
                  <div onClick={() => setEwoExp(prev => ({ ...prev, [parentId]: !groupExpanded }))}
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
            onSave={(d) => { saveProject(d, !panel.proj).catch(() => undefined); }}
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
            onSendEmail={async (to, _cc, subject, body) => {
              try {
                // Open email client with pre-filled content
                const plainBody = body.replace(/<br\s*\/?>/gi, '\n').replace(/<[^>]*>/g, '');
                const mailto = 'mailto:' + encodeURIComponent(to) +
                  '?subject=' + encodeURIComponent(subject) +
                  '&body=' + encodeURIComponent(plainBody);
                const a = document.createElement('a');
                a.href = mailto;
                a.click();
                // Record sent date on the RFI
                if (panel.rfi && panel.rfi.spId) {
                  const sentDate = new Date().toISOString().substring(0, 10);
                  const updated = { ...panel.rfi, emailSentDate: sentDate };
                  await spService.current.updateRfi(panel.rfi.spId, updated);
                  setRfis(prev => prev.map(r => r.id === panel.rfi!.id ? updated : r));
                  setPanel(prev => ({ ...prev, rfi: updated }));
                }
                toast('Email client opened. PDF downloaded. Sent date recorded.');
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
