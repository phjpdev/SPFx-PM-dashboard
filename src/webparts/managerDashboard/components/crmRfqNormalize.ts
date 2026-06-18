import type { CrmQuote, CrmQuoteStatus, CrmRfq, CrmRfqDiscipline, CrmRfqStage } from './crmTypes';

const QUOTE_STATUSES: CrmQuoteStatus[] = ['Draft', 'Sent', 'Pending', 'Follow up', 'Lost'];

const STAGES: CrmRfqStage[] = ['New Enquiry', 'Under Review', 'Ready to Quote', 'Won', 'Declined'];
const DISCIPLINES: CrmRfqDiscipline[] = ['Steel', 'Concrete', 'Both'];

const uid = (): string => `${Date.now()}-${Math.random().toString(36).substr(2, 9)}`;

export const normalizeRfqNum = (raw: string): string => {
  const v = (raw || '').trim();
  if (!v) return '';
  if (v.startsWith('RFQ-')) return v;
  if (/^\d{2}-\d{3}$/.test(v)) return `RFQ-${v}`;
  if (/^RFQ\d/i.test(v)) return v.replace(/^RFQ/i, 'RFQ-');
  return v;
};

/** Accept legacy MM-DD / DD-MM and ISO dates for year filtering. */
export const normalizeCrmDate = (raw: string, fallbackYear = new Date().getFullYear()): string => {
  const v = (raw || '').trim();
  if (!v) return '';
  if (/^\d{4}-\d{2}-\d{2}/.test(v)) return v.substring(0, 10);

  const slash = v.match(/^(\d{1,2})\/(\d{1,2})(?:\/(\d{2,4}))?$/);
  if (slash) {
    const a = parseInt(slash[1], 10);
    const b = parseInt(slash[2], 10);
    const y = slash[3]
      ? (slash[3].length === 2 ? 2000 + parseInt(slash[3], 10) : parseInt(slash[3], 10))
      : fallbackYear;
    const month = a <= 12 ? a : b;
    const day = a <= 12 ? b : a;
    return `${y}-${String(month).padStart(2, '0')}-${String(day).padStart(2, '0')}`;
  }

  const dash = v.match(/^(\d{1,2})-(\d{1,2})(?:-(\d{2,4}))?$/);
  if (dash) {
    const a = parseInt(dash[1], 10);
    const b = parseInt(dash[2], 10);
    const y = dash[3]
      ? (dash[3].length === 2 ? 2000 + parseInt(dash[3], 10) : parseInt(dash[3], 10))
      : fallbackYear;
    const month = a <= 12 ? a : b;
    const day = a <= 12 ? b : a;
    return `${y}-${String(month).padStart(2, '0')}-${String(day).padStart(2, '0')}`;
  }

  return v;
};

const asStage = (v: unknown): CrmRfqStage => {
  const s = String(v || 'New Enquiry') as CrmRfqStage;
  return STAGES.includes(s) ? s : 'New Enquiry';
};

const asDiscipline = (v: unknown): CrmRfqDiscipline => {
  const s = String(v || 'Steel') as CrmRfqDiscipline;
  return DISCIPLINES.includes(s) ? s : 'Steel';
};

/** Normalize current + legacy SharePoint RFQ JSON into one shape. */
export function normalizeCrmRfq(raw: Partial<CrmRfq> & Record<string, unknown>): CrmRfq {
  const rfqNum = normalizeRfqNum(String(raw.rfqNum || raw.rfqNumber || ''));
  const year = new Date().getFullYear();
  return {
    id: String(raw.id || uid()),
    rfqNum,
    dateReceived: normalizeCrmDate(String(raw.dateReceived || ''), year),
    personId: String(raw.personId || ''),
    organizationId: String(raw.organizationId || raw.companyId || ''),
    companyAddress: String(raw.companyAddress || ''),
    projectTitle: String(raw.projectTitle || ''),
    projectAddress: String(raw.projectAddress || ''),
    discipline: asDiscipline(raw.discipline),
    quoteRequiredBy: normalizeCrmDate(String(raw.quoteRequiredBy || ''), year),
    projectValue: typeof raw.projectValue === 'number' ? raw.projectValue : Number(raw.projectValue) || 0,
    approximateHours: typeof raw.approximateHours === 'number'
      ? raw.approximateHours
      : (typeof raw.estHours === 'number' ? raw.estHours : Number(raw.estHours) || 0),
    engineerDrawingReceived: !!raw.engineerDrawingReceived,
    engineerDrawingDate: normalizeCrmDate(String(raw.engineerDrawingDate || ''), year),
    revisionVersionEng: String(raw.revisionVersionEng || ''),
    architectDrawingReceived: !!raw.architectDrawingReceived,
    architectDrawingDate: normalizeCrmDate(String(raw.architectDrawingDate || ''), year),
    revisionVersionArch: String(raw.revisionVersionArch || ''),
    rfiAllowed: typeof raw.rfiAllowed === 'number' ? raw.rfiAllowed : (raw.rfiAllowed ? 1 : 0),
    createQuoteXero: !!raw.createQuoteXero,
    relatedRfqId: String(raw.relatedRfqId || ''),
    notes: String(raw.notes || ''),
    source: String(raw.source || 'Email'),
    stage: asStage(raw.stage),
    assignedTo: String(raw.assignedTo || 'MK'),
  };
}

export const normalizeQuoteNum = (raw: string): string => {
  const v = (raw || '').trim();
  if (!v) return '';
  if (v.startsWith('QU-') || v.startsWith('QT-')) return v.startsWith('QU-') ? v : v.replace(/^QT-/i, 'QU-');
  if (/^\d{4}$/.test(v)) return `QU-${v}`;
  return v;
};

const asQuoteStatus = (v: unknown): CrmQuoteStatus => {
  const s = String(v || 'Draft');
  if (s === 'Accepted') return 'Draft';
  if (s === 'Declined') return 'Lost';
  if (s === 'Follow Up' || s === 'Follow-up') return 'Follow up';
  return QUOTE_STATUSES.includes(s as CrmQuoteStatus) ? (s as CrmQuoteStatus) : 'Draft';
};

/** Normalize current + legacy SharePoint quote JSON into one shape. */
export function normalizeCrmQuote(raw: Partial<CrmQuote> & Record<string, unknown>): CrmQuote {
  const year = new Date().getFullYear();
  const rfqNum = normalizeRfqNum(String(raw.rfqNum || raw.rfqNumber || ''));
  return {
    id: String(raw.id || uid()),
    quoteNum: normalizeQuoteNum(String(raw.quoteNum || raw.quoteNumber || '')),
    rfqId: String(raw.rfqId || ''),
    rfqNum,
    quotedDate: normalizeCrmDate(String(raw.quotedDate || raw.dateQuoted || ''), year),
    dateReceived: normalizeCrmDate(String(raw.dateReceived || ''), year),
    personId: String(raw.personId || ''),
    organizationId: String(raw.organizationId || raw.companyId || ''),
    companyAddress: String(raw.companyAddress || ''),
    projectTitle: String(raw.projectTitle || ''),
    projectAddress: String(raw.projectAddress || ''),
    discipline: asDiscipline(raw.discipline),
    quoteRequiredBy: normalizeCrmDate(String(raw.quoteRequiredBy || ''), year),
    projectValue: typeof raw.projectValue === 'number' ? raw.projectValue : Number(raw.projectValue) || 0,
    approximateHours: typeof raw.approximateHours === 'number'
      ? raw.approximateHours
      : (typeof raw.estHours === 'number' ? raw.estHours : Number(raw.estHours) || 0),
    engineerDrawingReceived: !!raw.engineerDrawingReceived,
    engineerDrawingDate: normalizeCrmDate(String(raw.engineerDrawingDate || ''), year),
    revisionVersionEng: String(raw.revisionVersionEng || ''),
    architectDrawingReceived: !!raw.architectDrawingReceived,
    architectDrawingDate: normalizeCrmDate(String(raw.architectDrawingDate || ''), year),
    revisionVersionArch: String(raw.revisionVersionArch || ''),
    rfiAllowed: typeof raw.rfiAllowed === 'number' ? raw.rfiAllowed : (raw.rfiAllowed ? 1 : 0),
    assignedTo: String(raw.assignedTo || 'MK'),
    source: String(raw.source || 'Email'),
    notes: String(raw.notes || ''),
    createQuoteXero: !!raw.createQuoteXero,
    relatedRfqId: String(raw.relatedRfqId || ''),
    lostReason: String(raw.lostReason || ''),
    status: asQuoteStatus(raw.status),
    lostAt: normalizeCrmDate(String(raw.lostAt || ''), year) || undefined,
    archived: !!raw.archived,
    archivedAt: normalizeCrmDate(String(raw.archivedAt || ''), year) || undefined,
    ...(typeof raw.spId === 'number' ? { spId: raw.spId } : {}),
  };
}

/** Parse one legacy settings Value cell into a single RFQ object, or null if not an RFQ row. */
export const tryParseLegacyRfqValue = (value: string): CrmRfq | null => {
  const t = (value || '').trim();
  if (!t || t.startsWith('[')) return null;
  try {
    const obj = JSON.parse(t) as Record<string, unknown>;
    if (!obj || typeof obj !== 'object' || Array.isArray(obj)) return null;
    if (obj.rfqNum || obj.rfqNumber || (obj.projectTitle && (obj.dateReceived || obj.stage))) {
      return normalizeCrmRfq(obj);
    }
  } catch { /* partial chunk or invalid */ }
  return null;
};

/** Parse one legacy settings Value cell into a single quote object, or null. */
export const tryParseLegacyQuoteValue = (value: string): CrmQuote | null => {
  const t = (value || '').trim();
  if (!t || t.startsWith('[')) return null;
  try {
    const obj = JSON.parse(t) as Record<string, unknown>;
    if (!obj || typeof obj !== 'object' || Array.isArray(obj)) return null;
    if (obj.quoteNum || obj.quoteNumber || (obj.rfqNum || obj.rfqNumber)) {
      return normalizeCrmQuote(obj);
    }
  } catch { /* partial chunk or invalid */ }
  return null;
};
