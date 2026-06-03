import * as XLSX from 'xlsx';
import type { CrmCompany, CrmPerson, CrmPhone, CrmEmail } from './crmTypes';

const DEFAULT_CC = '+61';

const uid = (): string => `${Date.now()}-${Math.random().toString(36).substr(2, 9)}`;

const splitList = (s: string): string[] =>
  String(s || '')
    .split(/[,;]/)
    .map(x => x.trim())
    .filter(Boolean);

const normName = (s: string): string => s.trim().toLowerCase();

const findHeaderRow = (rows: unknown[][], matchers: ((cell: string) => boolean)[]): number => {
  for (let i = 0; i < Math.min(rows.length, 15); i++) {
    const row = rows[i] as unknown[];
    const hits = matchers.map(fn => row.some(c => fn(String(c || '').toLowerCase().trim())));
    if (hits.every(Boolean)) return i;
  }
  return -1;
};

const colIndex = (rows: unknown[][], headerRow: number, ...needles: string[]): number => {
  const row = rows[headerRow] as unknown[];
  for (let j = 0; j < row.length; j++) {
    const cell = String(row[j] || '').toLowerCase().trim();
    if (needles.every(n => cell.includes(n))) return j;
  }
  return -1;
};

const emptyPhone = (): CrmPhone => ({ value: '', type: 'Work', cc: DEFAULT_CC });

const parsePhones = (work: string, mobile: string): CrmPhone[] => {
  const phones: CrmPhone[] = [];
  splitList(work).forEach(v => phones.push({ value: v, type: 'Work', cc: DEFAULT_CC }));
  splitList(mobile).forEach(v => phones.push({ value: v, type: 'Mobile', cc: DEFAULT_CC }));
  return phones.length ? phones : [emptyPhone()];
};

const parseEmails = (raw: string): CrmEmail[] => {
  const vals = splitList(raw);
  return vals.length ? vals.map(v => ({ value: v, type: 'Work' as const })) : [{ value: '', type: 'Work' }];
};

export function parseOrganizationsFile(data: ArrayBuffer): CrmCompany[] {
  const wb = XLSX.read(data, { type: 'array' });
  const ws = wb.Sheets[wb.SheetNames[0]];
  const rows: unknown[][] = XLSX.utils.sheet_to_json(ws, { header: 1, defval: '' });
  const hRow = findHeaderRow(rows, [
    c => c.includes('organization') && c.includes('name'),
    c => c.includes('address'),
  ]);
  if (hRow < 0) throw new Error('Could not find organization columns (Organization - Name / Address).');
  const nameCol = colIndex(rows, hRow, 'organization', 'name');
  const addrCol = colIndex(rows, hRow, 'address');
  if (nameCol < 0) throw new Error('Missing "Organization - Name" column.');

  const out: CrmCompany[] = [];
  for (let i = hRow + 1; i < rows.length; i++) {
    const row = rows[i] as unknown[];
    const name = String(row[nameCol] || '').trim();
    if (!name) continue;
    const address = addrCol >= 0 ? String(row[addrCol] || '').trim() : '';
    out.push({
      id: uid(),
      name,
      labels: '',
      address,
      phones: [emptyPhone()],
      emails: [{ value: '', type: 'Work' }],
    });
  }
  return out;
}

/** First company per normalized name wins (for linking when Excel has duplicate org names). */
export function buildCompanyMap(companies: CrmCompany[]): Map<string, CrmCompany> {
  const m = new Map<string, CrmCompany>();
  companies.forEach(c => {
    const k = normName(c.name);
    if (!m.has(k)) m.set(k, c);
  });
  return m;
}

export function parsePeopleFile(data: ArrayBuffer, companyByName: Map<string, CrmCompany>): CrmPerson[] {
  const wb = XLSX.read(data, { type: 'array' });
  const ws = wb.Sheets[wb.SheetNames[0]];
  const rows: unknown[][] = XLSX.utils.sheet_to_json(ws, { header: 1, defval: '' });
  const hRow = findHeaderRow(rows, [
    c => c.includes('person') && c.includes('name'),
    c => c.includes('organization'),
  ]);
  if (hRow < 0) throw new Error('Could not find person columns (Person - Name / Organization).');
  const nameCol = colIndex(rows, hRow, 'person', 'name');
  const posCol  = colIndex(rows, hRow, 'position');
  const orgCol  = colIndex(rows, hRow, 'organization');
  const emailCol = colIndex(rows, hRow, 'email', 'work');
  const phoneWorkCol = colIndex(rows, hRow, 'phone', 'work');
  const phoneMobCol  = colIndex(rows, hRow, 'phone', 'mobile');
  if (nameCol < 0) throw new Error('Missing "Person - Name" column.');

  const out: CrmPerson[] = [];
  for (let i = hRow + 1; i < rows.length; i++) {
    const row = rows[i] as unknown[];
    const name = String(row[nameCol] || '').trim();
    if (!name) continue;
    const orgName = orgCol >= 0 ? String(row[orgCol] || '').trim() : '';
    const company = orgName ? companyByName.get(normName(orgName)) : undefined;
    out.push({
      id: uid(),
      name,
      organizationId: company?.id || '',
      position: posCol >= 0 ? String(row[posCol] || '').trim() : '',
      phones: parsePhones(
        phoneWorkCol >= 0 ? String(row[phoneWorkCol] || '') : '',
        phoneMobCol >= 0 ? String(row[phoneMobCol] || '') : '',
      ),
      emails: parseEmails(emailCol >= 0 ? String(row[emailCol] || '') : ''),
    });
  }
  return out;
}

export interface CrmImportStats {
  companyRows: number;
  personRows: number;
}

/**
 * Import every Excel row as its own record (no deduplication).
 * Selected files replace that section of CRM data entirely.
 */
export function runCrmImport(
  existingCompanies: CrmCompany[],
  existingPersons: CrmPerson[],
  orgBuffer: ArrayBuffer | null,
  peopleBuffer: ArrayBuffer | null,
): { companies: CrmCompany[]; persons: CrmPerson[]; stats: CrmImportStats } {
  if (!orgBuffer && !peopleBuffer) {
    throw new Error('Select at least one file to import.');
  }

  let companies = existingCompanies;
  let persons = existingPersons;

  if (orgBuffer) {
    companies = parseOrganizationsFile(orgBuffer);
  }

  if (peopleBuffer) {
    persons = parsePeopleFile(peopleBuffer, buildCompanyMap(companies));
  }

  return {
    companies,
    persons,
    stats: {
      companyRows: companies.length,
      personRows: persons.length,
    },
  };
}
