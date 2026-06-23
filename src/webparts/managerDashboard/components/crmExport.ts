import * as XLSX from 'xlsx';
import type { CrmCompany, CrmEmail, CrmPerson, CrmPhone } from './crmTypes';

const phoneValue = (arr: CrmPhone[]): string => {
  const p = arr.find(x => x.value);
  return p ? `${p.cc || ''} ${p.value}`.trim() : '';
};

const emailValue = (arr: CrmEmail[]): string => arr.find(x => x.value)?.value || '';

const exportDate = (): string => {
  const t = new Date();
  return `${t.getFullYear()}-${String(t.getMonth() + 1).padStart(2, '0')}-${String(t.getDate()).padStart(2, '0')}`;
};

const downloadWorkbook = (filename: string, sheetName: string, rows: Record<string, string>[]): void => {
  const ws = XLSX.utils.json_to_sheet(rows);
  const wb = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(wb, ws, sheetName);
  XLSX.writeFile(wb, filename);
};

export function exportPersonsExcel(
  persons: CrmPerson[],
  companyName: (id: string) => string,
  companyAddress: (id: string) => string,
): void {
  const rows = persons.map((p, idx) => ({
    '#': String(idx + 1),
    'Name': p.name,
    'Company': p.organizationId ? companyName(p.organizationId) : '',
    'Position': p.position,
    'Location': p.organizationId ? companyAddress(p.organizationId) : '',
    'Phone': phoneValue(p.phones),
    'Email': emailValue(p.emails),
  }));
  downloadWorkbook(`CRM-Persons-${exportDate()}.xlsx`, 'Persons', rows);
}

export function exportCompaniesExcel(companies: CrmCompany[]): void {
  const rows = companies.map((c, idx) => ({
    '#': String(idx + 1),
    'Company': c.name,
    'Labels': c.labels,
    'Phone': phoneValue(c.phones),
    'Email': emailValue(c.emails),
    'Address': c.address,
  }));
  downloadWorkbook(`CRM-Companies-${exportDate()}.xlsx`, 'Companies', rows);
}
