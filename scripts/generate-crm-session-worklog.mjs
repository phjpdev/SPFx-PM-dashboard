/**
 * Generates CRM session work log — same format as CRM_Work_Log.xlsx
 * Run: node scripts/generate-crm-session-worklog.mjs
 */
import * as XLSX from 'xlsx';
import { join, dirname } from 'path';
import { fileURLToPath } from 'url';

const __dirname = dirname(fileURLToPath(import.meta.url));
const root = join(__dirname, '..');
const outPath = join(root, process.argv[2] || 'CRM-Session-Work-Log-June-2026.xlsx');

/** @type {[string, string, string][]} [Work Item, Description, duration] */
const items = [
  ['Pipedrive Excel Import', 'Parse organizations + people .xlsx (Pipedrive export); Import modal; merge by name; link persons to companies', '3 hr'],
  ['Static CRM Baseline (645 / 512)', 'generate-crm-static.mjs bundles persons/companies JSON into SPFx; same data for all users without re-import', '2 hr'],
  ['Persons Table — Location Column', 'Location shows company address as green Google Maps link; visible to all users from bundled data', '1 hr'],
  ['Table Layout — No Horizontal Scroll', 'Fixed overflow on Persons, Companies, RFQ tables; column widths and wrapping', '0.75 hr'],
  ['RFQ Actions Column Fix', 'RFQ # and Edit/Del buttons stacked vertically; no cut-off; min-width on action cells', '0.5 hr'],
  ['CRM Toolbar Alignment', 'Search, Import, + Person / + Company aligned on one line in Persons & Companies tabs', '0.25 hr'],
  ['RFQ Toolbar — + RFQ Button', 'Stage filter width reduced; + RFQ button aligned right; search left', '0.5 hr'],
  ['Person Modal — Activity Log', 'Date (default today), type dropdown (Phone call/Email/Text/In person), Notes, Follow-up date, Mark as done checkbox; + Activity rows', '2 hr'],
  ['Person Modal — 2-Column Layout', 'Wide modal: contact details left, Activity Log + attachments right; View mode read-only cards', '1.5 hr'],
  ['Person Modal — Note Attachments', 'Image/screenshot attachments on person record; thumbnail grid (browser storage)', '1 hr'],
  ['Position Dropdown Update', 'Added Draftee, Drafting Manager; all positions sorted A–Z', '0.25 hr'],
  ['View / Edit Modal Modes', 'View, Edit, Delete, Save, Cancel footer on Person and Company modals', '1 hr'],
  ['Activity Log — Edit Saved Entries', 'All saved activities editable in Edit mode (not read-only after save)', '0.5 hr'],
  ['Modal Close on Save', 'Person and Company modals close immediately after Save', '0.25 hr'],
  ['Alphabetical Table Sort', 'Click column headers to sort Persons and Companies tables A–Z / Z–A', '1 hr'],
  ['Company → Person Links', 'Linked person name chips on company rows open Person View modal', '0.5 hr'],
  ['SharePoint CRM Sync — Delta Storage', 'Only changed persons/companies saved to 3Edge_Settings (crm_delta_meta + per-record keys); 12s poll; debounced save', '4 hr'],
  ['SharePoint HTTP 500 Fix', 'Chunk size fallback (200 char); delta sync avoids full 645-person payload; local + SP merge', '2.5 hr'],
  ['Person Edit — View Data Fix', 'Edited person shows in View modal; skip remote overwrite 15s after local edit', '1 hr'],
  ['Checklist SharePoint Fallback', '403 on document library → 3Edge_Settings chunks; grey info banner when local save OK', '1.5 hr'],
  ['RFQ & Quotes Tab Buttons', 'CRM sub-tabs: Persons, Companies, RFQ, Quotes', '0.25 hr'],
  ['RFQ Tab — Full Module', 'KPI cards, search, stage filter, add/edit modal, RFQ-26-### auto numbering, table, SharePoint sync (crm_rfqs)', '4 hr'],
  ['RFQ — Approximate / Est Hours', 'Est Hours field on RFQ form and Hours column in table', '0.5 hr'],
  ['RFQ — Discipline Badges', 'Steel blue / Concrete purple outline badges matching Projects page; Both = two badges', '0.5 hr'],
  ['RFQ — Drawing & RFI Fields', 'Engineer drawing received (checkbox + date), Architect drawing received (checkbox + date), RFI allowed checkbox; grid layout per mockup', '1 hr'],
  ['Quotes Tab — Move from RFQ', 'Quote button replaces Del when stage = Ready to Quote; creates QT-26-###; removes RFQ from pipeline', '2.5 hr'],
  ['Quotes Tab — Full Module', 'KPI cards, search, status filter, table, edit modal (status/value/hours/Xero/notes), SharePoint sync (crm_quotes)', '2 hr'],
  ['Quotes — Display After Move Fix', 'Quote shows immediately after OK; seed data to Quotes tab; localStorage before SP load', '1 hr'],
  ['Production Builds & Fixes', 'heft build --production; TS Map iteration, lint fixes, deploy-ready package', '2 hr'],
];

const parseHours = (s) => {
  const m = String(s).match(/([\d.]+)/);
  return m ? parseFloat(m[1]) : 0;
};

const total = items.reduce((sum, row) => sum + parseHours(row[2]), 0);

const sheet = [
  ['#', 'Work Item', 'Description', 'Est. Duration'],
  ...items.map((row, i) => [i + 1, row[0], row[1], row[2]]),
  ['', '', '', ''],
  ['', '', 'TOTAL', `${total} hrs`],
  ['', '', '', ''],
];

const wb = XLSX.utils.book_new();
XLSX.utils.book_append_sheet(wb, XLSX.utils.aoa_to_sheet(sheet), 'Work Log');
XLSX.writeFile(wb, outPath);

console.log('Written:', outPath);
console.log('Items:', items.length, '| Total:', total, 'hrs');
