import * as XLSX from 'xlsx';
import type { CrmCompany, CrmPerson, CrmPhone, CrmEmail } from './crmTypes';

const DEFAULT_CC = '+61';

const uid = (): string => `${Date.now()}-${Math.random().toString(36).substr(2, 5)}`;

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

export interface CrmImportPreview {
  companies: CrmCompany[];
  persons: CrmPerson[];
  companyCount: number;
  personCount: number;
  skippedCompanies: number;
  skippedPersons: number;
}

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

/** Merge imported rows into existing CRM data (match by name / name+org). */
export function mergeCrmImport(
  existingCompanies: CrmCompany[],
  existingPersons: CrmPerson[],
  importedCompanies: CrmCompany[],
  importedPersons: CrmPerson[],
): { companies: CrmCompany[]; persons: CrmPerson[]; preview: CrmImportPreview } {
  const companies = [...existingCompanies];
  const companyByName = new Map<string, CrmCompany>();
  companies.forEach(c => companyByName.set(normName(c.name), c));

  let skippedCompanies = 0;
  for (const imp of importedCompanies) {
    const key = normName(imp.name);
    const prev = companyByName.get(key);
    if (prev) {
      prev.address = imp.address || prev.address;
      skippedCompanies++;
    } else {
      companies.push(imp);
      companyByName.set(key, imp);
    }
  }

  // Re-link persons to merged company ids
  const relinkedPersons = importedPersons.map(p => {
    if (!p.organizationId) return p;
    const impCo = importedCompanies.find(c => c.id === p.organizationId);
    if (!impCo) return p;
    const merged = companyByName.get(normName(impCo.name));
    return merged ? { ...p, organizationId: merged.id } : p;
  });

  // Fix org links by name from import file (organizationId may point to temp ids)
  relinkedPersons.forEach(p => {
    const imp = importedPersons.find(x => x.id === p.id);
    if (!imp?.organizationId) return;
    const impCo = importedCompanies.find(c => c.id === imp.organizationId);
    if (impCo) {
      const merged = companyByName.get(normName(impCo.name));
      if (merged) p.organizationId = merged.id;
    }
  });

  const persons = [...existingPersons];
  let skippedPersons = 0;
  for (const imp of relinkedPersons) {
    const orgKey = imp.organizationId || '';
    const dup = persons.find(
      p => normName(p.name) === normName(imp.name) && p.organizationId === orgKey,
    );
    if (dup) {
      dup.position = imp.position || dup.position;
      dup.phones = imp.phones.some(x => x.value) ? imp.phones : dup.phones;
      dup.emails = imp.emails.some(x => x.value) ? imp.emails : dup.emails;
      dup.organizationId = imp.organizationId || dup.organizationId;
      skippedPersons++;
    } else {
      persons.push(imp);
    }
  }

  return {
    companies,
    persons,
    preview: {
      companies,
      persons,
      companyCount: importedCompanies.length,
      personCount: importedPersons.length,
      skippedCompanies,
      skippedPersons,
    },
  };
}

export function buildCompanyMap(companies: CrmCompany[]): Map<string, CrmCompany> {
  const m = new Map<string, CrmCompany>();
  companies.forEach(c => m.set(normName(c.name), c));
  return m;
}

export interface CrmImportStats {
  importedCompanies: number;
  importedPersons: number;
  updatedCompanies: number;
  updatedPersons: number;
}

export function runCrmImport(
  existingCompanies: CrmCompany[],
  existingPersons: CrmPerson[],
  orgBuffer: ArrayBuffer | null,
  peopleBuffer: ArrayBuffer | null,
): { companies: CrmCompany[]; persons: CrmPerson[]; stats: CrmImportStats } {
  if (!orgBuffer && !peopleBuffer) {
    throw new Error('Select at least one file to import.');
  }

  let importedCompanies: CrmCompany[] = [];
  let importedPersons: CrmPerson[] = [];

  if (orgBuffer) importedCompanies = parseOrganizationsFile(orgBuffer);

  const { companies: mergedCompanies } = mergeCrmImport(
    existingCompanies,
    [],
    importedCompanies,
    [],
  );

  if (peopleBuffer) {
    importedPersons = parsePeopleFile(peopleBuffer, buildCompanyMap(mergedCompanies));
  }

  const { companies, persons, preview } = mergeCrmImport(
    existingCompanies,
    existingPersons,
    importedCompanies,
    importedPersons,
  );

  return {
    companies,
    persons,
    stats: {
      importedCompanies: preview.companyCount,
      importedPersons: preview.personCount,
      updatedCompanies: preview.skippedCompanies,
      updatedPersons: preview.skippedPersons,
    },
  };
}
