import { jsPDF } from 'jspdf';
import { IProject } from '../../../shared/models/IProject';
import type { ChecklistItemState, ChecklistOverrideLog } from '../../../shared/services/SharePointService';
import { drawLetterhead, drawPdfBg } from '../../../shared/utils/pdfLetterhead';
import {
  CHECKLIST, IItemState, itemIdOf, ProjectType, SectionType,
} from './checklistData';

export interface ChecklistPhasePdfParams {
  project: IProject;
  phaseId: string;
  projectType: ProjectType;
  items: Record<string, ChecklistItemState | IItemState>;
  overrides: ChecklistOverrideLog[];
  exportedBy: string;
}

const isSectionVisible = (t: SectionType, projectType: ProjectType): boolean => {
  if (projectType === 'both') return true;
  if (t === 'both') return true;
  return t === projectType;
};

const isResolved = (s: ChecklistItemState | IItemState): boolean =>
  !!s.override || s.c2 === 'cleared' || s.c2 === 'na';

const statusLabel = (st: ChecklistItemState | IItemState): string => {
  if (st.override) return 'PM CLEARED';
  if (st.c2 === 'cleared') return 'CLEARED';
  if (st.c2 === 'na') return 'N/A';
  if (st.c2 === 'incorrect') return 'FIX REQUIRED';
  if (st.c1) return 'READY FOR C2';
  return 'WAITING C1';
};

const c1Label = (st: ChecklistItemState | IItemState): string => {
  if (st.c1 && st.c1By) return `✓ ${st.c1By.split(' ')[0]} · ${st.c1At || ''}`;
  if (st.override) return 'skipped';
  return '—';
};

const c2Label = (st: ChecklistItemState | IItemState): string => {
  if (st.override) return `PM override · ${st.overrideAt || ''}`;
  if (st.c2 === 'cleared' && st.c2By) return `Cleared · ${st.c2By.split(' ')[0]} · ${st.c2At || ''}`;
  if (st.c2 === 'na' && st.c2By) return `N/A · ${st.c2By.split(' ')[0]} · ${st.c2At || ''}`;
  if (st.c2 === 'incorrect' && st.c2By) return `Flagged · ${st.c2By.split(' ')[0]} · ${st.c2At || ''}`;
  if (st.c1) return 'awaiting checker';
  return 'locked';
};

const projectTypeLabel = (t: ProjectType): string => {
  if (t === 'both') return 'Steel & Concrete';
  return t.charAt(0).toUpperCase() + t.slice(1);
};

const slugify = (s: string): string =>
  s.replace(/[^a-zA-Z0-9]+/g, '-').replace(/^-|-$/g, '');

export function checklistPhaseExportFilename(project: IProject, phaseId: string, phaseName: string): string {
  const date = new Date().toISOString().substring(0, 10);
  const projNum = (project.projNum || 'Project').replace(/[^a-zA-Z0-9_-]/g, '_');
  return `${projNum}_Checklist_Phase${phaseId}_${slugify(phaseName)}_${date}.pdf`;
}

export function generateChecklistPhasePdf(params: ChecklistPhasePdfParams): Blob | undefined {
  const { project, phaseId, projectType, items, overrides, exportedBy } = params;
  const phaseIdx = CHECKLIST.findIndex(p => p.id === phaseId);
  if (phaseIdx < 0) return undefined;

  const phase = CHECKLIST[phaseIdx];
  const visibleSections = phase.sections.filter(s => isSectionVisible(s.type, projectType));
  if (visibleSections.length === 0) return undefined;

  const phaseItemIds = new Set<string>();
  visibleSections.forEach((section) => {
    const si = phase.sections.indexOf(section);
    section.items.forEach((_, ii) => phaseItemIds.add(itemIdOf(phaseIdx, si, ii)));
  });

  const getSt = (id: string): ChecklistItemState | IItemState => items[id] || { c1: false, c2: null };

  let totalItems = 0;
  let clearedItems = 0;
  visibleSections.forEach((section) => {
    const si = phase.sections.indexOf(section);
    section.items.forEach((_, ii) => {
      totalItems++;
      if (isResolved(getSt(itemIdOf(phaseIdx, si, ii)))) clearedItems++;
    });
  });

  const phaseOverrides = overrides.filter(ov => phaseItemIds.has(ov.itemId));
  const exportTs = new Date().toLocaleString('en-AU', {
    day: '2-digit', month: 'short', year: 'numeric', hour: '2-digit', minute: '2-digit', hour12: true,
  });

  try {
    // eslint-disable-next-line @typescript-eslint/no-explicit-any
    const doc: any = new jsPDF({ orientation: 'portrait', unit: 'mm', format: 'a4' });
    const pw = 210;
    const ph = 297;
    const ml = 10;
    const mr = 10;
    const tw = pw - ml - mr;
    const bottom = ph - 14;

    let y = drawLetterhead(
      doc,
      pw,
      ph,
      `QA CHECKLIST — Phase ${phase.id} · ${phase.name}`,
      `${project.projNum || '—'}  |  ${new Date().toLocaleDateString('en-AU')}`,
    );

    const ensureSpace = (need: number): void => {
      if (y + need <= bottom) return;
      doc.addPage();
      drawPdfBg(doc, pw, ph);
      y = 14;
    };

    const sectionHeader = (title: string): void => {
      ensureSpace(10);
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

    const row2 = (l1: string, v1: string, l2: string, v2: string): void => {
      ensureSpace(7);
      const cw = tw / 2;
      doc.setFontSize(7.5);
      doc.setFont('helvetica', 'bold');
      doc.setTextColor(90, 110, 136);
      doc.text(l1.toUpperCase(), ml, y + 3.5);
      doc.setFont('helvetica', 'normal');
      doc.setTextColor(26, 32, 48);
      doc.text(v1 || '—', ml + 28, y + 3.5, { maxWidth: cw - 30 });
      doc.setFont('helvetica', 'bold');
      doc.setTextColor(90, 110, 136);
      doc.text(l2.toUpperCase(), ml + cw, y + 3.5);
      doc.setFont('helvetica', 'normal');
      doc.setTextColor(26, 32, 48);
      doc.text(v2 || '—', ml + cw + 28, y + 3.5, { maxWidth: cw - 30 });
      doc.setDrawColor(208, 213, 222);
      doc.line(ml, y + 6, ml + tw, y + 6);
      y += 7;
    };

    sectionHeader('PROJECT DETAILS');
    row2('Project #', project.projNum, 'Project Name', project.name);
    row2('Company', project.company, 'Status', project.status);
    row2('Project Type', projectTypeLabel(projectType), 'Discipline', project.discipline || '—');
    row2('Detailer', project.detailers || '—', 'PM', project.teamLead || '—');
    row2('Phase Progress', `${clearedItems} of ${totalItems} items cleared`, 'Exported By', exportedBy);
    y += 2;

    sectionHeader('CHECKLIST ITEMS');

    const colW = { c1: 24, c2: 34, task: 14, status: 22 };
    const itemW = tw - colW.c1 - colW.c2 - colW.task - colW.status;

    const drawTableHeader = (): void => {
      ensureSpace(8);
      doc.setFillColor(240, 242, 245);
      doc.rect(ml, y, tw, 7, 'F');
      doc.setDrawColor(208, 213, 222);
      doc.rect(ml, y, tw, 7, 'S');
      doc.setFontSize(6.5);
      doc.setFont('helvetica', 'bold');
      doc.setTextColor(90, 110, 136);
      let x = ml + 2;
      doc.text('CHECK 1', x, y + 4.5);
      x += colW.c1;
      doc.text('CHECK 2', x, y + 4.5);
      x += colW.c2;
      doc.text('ITEM', x, y + 4.5);
      x += itemW;
      doc.text('TASK', x, y + 4.5);
      x += colW.task;
      doc.text('STATUS', x, y + 4.5);
      y += 8;
    };

    drawTableHeader();

    visibleSections.forEach((section) => {
      const si = phase.sections.indexOf(section);
      ensureSpace(8);
      doc.setFontSize(7.5);
      doc.setFont('helvetica', 'bold');
      doc.setTextColor(42, 158, 42);
      doc.text(section.title, ml, y + 3);
      y += 5;

      section.items.forEach((item, ii) => {
        const st = getSt(itemIdOf(phaseIdx, si, ii));
        const itemLines = doc.splitTextToSize(item[0], itemW - 2);
        const c1Lines = doc.splitTextToSize(c1Label(st), colW.c1 - 2);
        const c2Lines = doc.splitTextToSize(c2Label(st), colW.c2 - 2);
        const rowH = Math.max(itemLines.length, c1Lines.length, c2Lines.length, 1) * 4 + 3;

        if (y + rowH > bottom) {
          doc.addPage();
          drawPdfBg(doc, pw, ph);
          y = 14;
          drawTableHeader();
        }

        doc.setDrawColor(220, 225, 230);
        doc.line(ml, y, ml + tw, y);

        doc.setFontSize(6.5);
        doc.setFont('helvetica', 'normal');
        doc.setTextColor(26, 32, 48);

        let x = ml + 1;
        doc.text(c1Lines, x, y + 3.5);
        x += colW.c1;
        doc.text(c2Lines, x, y + 3.5);
        x += colW.c2;
        doc.text(itemLines, x, y + 3.5);
        x += itemW;
        doc.setFont('helvetica', 'bold');
        doc.text(item[1], x, y + 3.5);
        x += colW.task;
        doc.setFont('helvetica', 'normal');
        doc.text(statusLabel(st), x, y + 3.5, { maxWidth: colW.status - 1 });

        y += rowH;
      });

      y += 2;
    });

    if (phaseOverrides.length > 0) {
      sectionHeader('PM OVERRIDE LOG');
      phaseOverrides.forEach((ov) => {
        const lines = doc.splitTextToSize(
          `${ov.by} cleared Check 2 on "${ov.itemText}" (${ov.taskCode}) at ${ov.at}. Reason: "${ov.reason}"`,
          tw - 4,
        );
        ensureSpace(lines.length * 4 + 3);
        doc.setFontSize(7.5);
        doc.setFont('helvetica', 'normal');
        doc.setTextColor(74, 47, 156);
        doc.text(lines, ml + 2, y + 3);
        y += lines.length * 4 + 2;
        doc.setDrawColor(208, 213, 222);
        doc.line(ml, y, ml + tw, y);
        y += 2;
      });
    }

    ensureSpace(8);
    doc.setFontSize(7);
    doc.setFont('helvetica', 'italic');
    doc.setTextColor(90, 110, 136);
    doc.text(`Internal QA record — 3 Edge Design · Exported by ${exportedBy} at ${exportTs}`, ml, y + 3);

    const pages = doc.internal.getNumberOfPages();
    for (let i = 2; i <= pages; i++) {
      doc.setPage(i);
      drawPdfBg(doc, pw, ph);
    }

    return doc.output('blob') as Blob;
  } catch (e) {
    console.error('Checklist phase PDF generation error:', e);
    return undefined;
  }
}

export function downloadChecklistPhasePdf(blob: Blob, filename: string): void {
  const url = URL.createObjectURL(blob);
  const a = document.createElement('a');
  a.href = url;
  a.download = filename;
  a.click();
  setTimeout(() => URL.revokeObjectURL(url), 3000);
}
