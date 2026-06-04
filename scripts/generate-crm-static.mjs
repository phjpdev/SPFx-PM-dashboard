/**
 * One-time / on-demand: build static CRM JSON from Pipedrive Excel exports.
 * Run: node scripts/generate-crm-static.mjs
 */
import * as fs from 'fs';
import * as path from 'path';
import { fileURLToPath } from 'url';
import XLSX from 'xlsx';

const __dirname = path.dirname(fileURLToPath(import.meta.url));
const root = path.join(__dirname, '..');
const outDir = path.join(root, 'src/webparts/managerDashboard/data');

const ORG_FILE = path.join(root, 'organizations-14239720-8updated2.6.2026.xlsx');
const PEOPLE_FILE = path.join(root, 'people-14239720-9updated 2.6.26.xlsx');

const DEFAULT_CC = '+61';

const splitList = (s) =>
  String(s || '')
    .split(/[,;]/)
    .map((x) => x.trim())
    .filter(Boolean);

const normName = (s) => s.trim().toLowerCase();

const findHeaderRow = (rows, matchers) => {
  for (let i = 0; i < Math.min(rows.length, 15); i++) {
    const row = rows[i];
    const hits = matchers.map((fn) => row.some((c) => fn(String(c || '').toLowerCase().trim())));
    if (hits.every(Boolean)) return i;
  }
  return -1;
};

const colIndex = (rows, headerRow, ...needles) => {
  const row = rows[headerRow];
  for (let j = 0; j < row.length; j++) {
    const cell = String(row[j] || '').toLowerCase().trim();
    if (needles.every((n) => cell.includes(n))) return j;
  }
  return -1;
};

const emptyPhone = () => ({ value: '', type: 'Work', cc: DEFAULT_CC });

const parsePhones = (work, mobile) => {
  const phones = [];
  splitList(work).forEach((v) => phones.push({ value: v, type: 'Work', cc: DEFAULT_CC }));
  splitList(mobile).forEach((v) => phones.push({ value: v, type: 'Mobile', cc: DEFAULT_CC }));
  return phones.length ? phones : [emptyPhone()];
};

const parseEmails = (raw) => {
  const vals = splitList(raw);
  return vals.length ? vals.map((v) => ({ value: v, type: 'Work' })) : [{ value: '', type: 'Work' }];
};

function parseOrganizations(buffer) {
  const wb = XLSX.read(buffer, { type: 'buffer' });
  const ws = wb.Sheets[wb.SheetNames[0]];
  const rows = XLSX.utils.sheet_to_json(ws, { header: 1, defval: '' });
  const hRow = findHeaderRow(rows, [
    (c) => c.includes('organization') && c.includes('name'),
    (c) => c.includes('address'),
  ]);
  if (hRow < 0) throw new Error('Could not find organization columns.');
  const nameCol = colIndex(rows, hRow, 'organization', 'name');
  const addrCol = colIndex(rows, hRow, 'address');

  const out = [];
  let idx = 0;
  for (let i = hRow + 1; i < rows.length; i++) {
    const row = rows[i];
    const name = String(row[nameCol] || '').trim();
    if (!name) continue;
    idx += 1;
    out.push({
      id: `co-${String(idx).padStart(5, '0')}`,
      name,
      labels: '',
      address: addrCol >= 0 ? String(row[addrCol] || '').trim() : '',
      phones: [emptyPhone()],
      emails: [{ value: '', type: 'Work' }],
    });
  }
  return out;
}

function buildCompanyMap(companies) {
  const m = new Map();
  companies.forEach((c) => {
    const k = normName(c.name);
    if (!m.has(k)) m.set(k, c);
  });
  return m;
}

function parsePeople(buffer, companyByName) {
  const wb = XLSX.read(buffer, { type: 'buffer' });
  const ws = wb.Sheets[wb.SheetNames[0]];
  const rows = XLSX.utils.sheet_to_json(ws, { header: 1, defval: '' });
  const hRow = findHeaderRow(rows, [
    (c) => c.includes('person') && c.includes('name'),
    (c) => c.includes('organization'),
  ]);
  if (hRow < 0) throw new Error('Could not find person columns.');
  const nameCol = colIndex(rows, hRow, 'person', 'name');
  const posCol = colIndex(rows, hRow, 'position');
  const orgCol = colIndex(rows, hRow, 'organization');
  const emailCol = colIndex(rows, hRow, 'email', 'work');
  const phoneWorkCol = colIndex(rows, hRow, 'phone', 'work');
  const phoneMobCol = colIndex(rows, hRow, 'phone', 'mobile');

  const out = [];
  let idx = 0;
  for (let i = hRow + 1; i < rows.length; i++) {
    const row = rows[i];
    const name = String(row[nameCol] || '').trim();
    if (!name) continue;
    idx += 1;
    const orgName = orgCol >= 0 ? String(row[orgCol] || '').trim() : '';
    const company = orgName ? companyByName.get(normName(orgName)) : undefined;
    out.push({
      id: `pe-${String(idx).padStart(5, '0')}`,
      name,
      organizationId: company?.id || '',
      position: posCol >= 0 ? String(row[posCol] || '').trim() : '',
      phones: parsePhones(
        phoneWorkCol >= 0 ? String(row[phoneWorkCol] || '') : '',
        phoneMobCol >= 0 ? String(row[phoneMobCol] || '') : '',
      ),
      emails: parseEmails(emailCol >= 0 ? String(row[emailCol] || '') : ''),
      activities: [],
      attachments: [],
    });
  }
  return out;
}

function main() {
  if (!fs.existsSync(ORG_FILE)) {
    console.error('Missing:', ORG_FILE);
    process.exit(1);
  }
  if (!fs.existsSync(PEOPLE_FILE)) {
    console.error('Missing:', PEOPLE_FILE);
    process.exit(1);
  }

  const companies = parseOrganizations(fs.readFileSync(ORG_FILE));
  const companyMap = buildCompanyMap(companies);
  const persons = parsePeople(fs.readFileSync(PEOPLE_FILE), companyMap);

  fs.mkdirSync(outDir, { recursive: true });
  fs.writeFileSync(path.join(outDir, 'crm-companies.json'), JSON.stringify(companies));
  fs.writeFileSync(path.join(outDir, 'crm-persons.json'), JSON.stringify(persons));

  const meta = {
    generatedAt: new Date().toISOString(),
    companyCount: companies.length,
    personCount: persons.length,
    sourceFiles: [
      'organizations-14239720-8updated2.6.2026.xlsx',
      'people-14239720-9updated 2.6.26.xlsx',
    ],
  };
  fs.writeFileSync(path.join(outDir, 'crm-static-meta.json'), JSON.stringify(meta, null, 2));

  console.log(`Wrote ${companies.length} companies, ${persons.length} persons to ${outDir}`);
}

main();
