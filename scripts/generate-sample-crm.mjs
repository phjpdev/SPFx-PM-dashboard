/**
 * Builds the sample CRM dataset that ships with the manager web part.
 *
 * The real deployment seeds its CRM from a Pipedrive export (see
 * generate-crm-static.mjs). That export holds live client contacts, so this
 * repository ships a synthetic set of the same shape and size instead: 512
 * companies and 645 people, using ACMA fiction-reserved phone numbers and
 * .example.com addresses so nothing here can reach a real person.
 *
 * Run: node scripts/generate-sample-crm.mjs
 */
import * as fs from 'fs';
import * as path from 'path';
import { fileURLToPath } from 'url';

const __dirname = path.dirname(fileURLToPath(import.meta.url));
const outDir = path.join(__dirname, '..', 'src/webparts/managerDashboard/data');

const COMPANY_COUNT = 512;
const PERSON_COUNT = 645;
const DEFAULT_CC = '+61';

/** Deterministic PRNG so regenerating does not churn the diff. */
let seed = 20260604;
const rnd = () => {
  seed = (seed * 1103515245 + 12345) & 0x7fffffff;
  return seed / 0x7fffffff;
};
const pick = (arr) => arr[Math.floor(rnd() * arr.length)];
const chance = (pct) => rnd() * 100 < pct;

const COMPANY_HEAD = [
  'Arkwright', 'Balwarra', 'Cootamundra', 'Darling Range', 'Eastbrook', 'Fernhill', 'Gundaroo',
  'Hawkesbury', 'Illawarra', 'Jerrabomberra', 'Kurrajong', 'Lachlan', 'Mologa', 'Nepean',
  'Orana', 'Pindari', 'Quambone', 'Riverina', 'Southgate', 'Tarana', 'Ulmarra', 'Verona',
  'Wentworth', 'Yarrawonga', 'Alderton', 'Brightwater', 'Clearspan', 'Deepwater', 'Everton',
  'Foxhall', 'Granville', 'Highmark', 'Ironbark', 'Jasper', 'Kingsley', 'Longreach', 'Marchmont',
  'Northcliff', 'Oakridge', 'Penwood', 'Redgum', 'Stoneleigh', 'Thornbury', 'Uplands', 'Westmead',
];
const COMPANY_MID = [
  'Steel', 'Structural', 'Fabrication', 'Engineering', 'Construction', 'Industries', 'Metalworks',
  'Building', 'Civil', 'Precast', 'Rigging', 'Projects', 'Detailing', 'Site Services', 'Concrete',
];
const COMPANY_TAIL = ['Pty Ltd', 'Group', 'Co', 'Services', 'Australia', '', '', ''];

const STREETS = [
  'Acacia St', 'Bridge Rd', 'Canning Dr', 'Dampier Ave', 'Elder St', 'Forge Pl', 'Grange Rd',
  'Hume Hwy', 'Ivy Ln', 'Jubilee Ave', 'Kestrel Cr', 'Lachlan St', 'Merino Rd', 'Newcomen St',
  'Orchard Way', 'Parkes St', 'Quarry Rd', 'Riverside Dr', 'Stanley St', 'Tumut Pl', 'Wattle Gr',
];
const TOWNS = [
  ['BATHURST', '2795'], ['ORANGE', '2800'], ['DUBBO', '2830'], ['WAGGA WAGGA', '2650'],
  ['ALBURY', '2640'], ['GRIFFITH', '2680'], ['TAMWORTH', '2340'], ['ARMIDALE', '2350'],
  ['GOULBURN', '2580'], ['NOWRA', '2541'], ['LITHGOW', '2790'], ['PARKES', '2870'],
  ['MUDGEE', '2850'], ['COWRA', '2794'], ['YOUNG', '2594'], ['FORBES', '2871'],
  ['PENRITH', '2750'], ['BLACKTOWN', '2148'], ['WETHERILL PARK', '2164'], ['SMITHFIELD', '2164'],
];

const FIRST = [
  'Adam', 'Alan', 'Alex', 'Amy', 'Andrew', 'Anna', 'Ben', 'Bianca', 'Brad', 'Brett', 'Cameron',
  'Carla', 'Chris', 'Claire', 'Craig', 'Dale', 'Dan', 'Dave', 'Dean', 'Elise', 'Emma', 'Eric',
  'Fiona', 'Gareth', 'Gavin', 'Georgia', 'Glen', 'Grant', 'Hayley', 'Heath', 'Ian', 'Jacinta',
  'Jack', 'Jake', 'James', 'Jane', 'Jason', 'Jenny', 'Joel', 'Jordan', 'Josh', 'Julia', 'Karl',
  'Kate', 'Keith', 'Kelly', 'Kurt', 'Lachlan', 'Lara', 'Leigh', 'Liam', 'Lisa', 'Luke', 'Marcus',
  'Maria', 'Mark', 'Matt', 'Megan', 'Michael', 'Nadia', 'Nathan', 'Neil', 'Nick', 'Olivia',
  'Owen', 'Paul', 'Peta', 'Peter', 'Phil', 'Rachel', 'Reece', 'Rhys', 'Rob', 'Rose', 'Ryan',
  'Sam', 'Sarah', 'Scott', 'Sean', 'Shane', 'Simon', 'Sonia', 'Stuart', 'Tania', 'Ted', 'Tim',
  'Todd', 'Tom', 'Travis', 'Trent', 'Vanessa', 'Wade', 'Wayne', 'Will', 'Zac',
];
const LAST = [
  'Abbott', 'Ashby', 'Bannister', 'Barlow', 'Beasley', 'Bourke', 'Bradshaw', 'Callaghan',
  'Carmody', 'Chalmers', 'Colquhoun', 'Cousins', 'Crabtree', 'Dalgleish', 'Deakin', 'Dunmore',
  'Eastwood', 'Ellery', 'Fairbairn', 'Farrelly', 'Fitzhugh', 'Gallagher', 'Garrick', 'Gillespie',
  'Halloran', 'Hardwick', 'Hathaway', 'Hennessy', 'Hollis', 'Ingram', 'Jarrett', 'Kavanagh',
  'Kerrigan', 'Lansdowne', 'Larkin', 'Lockyer', 'Mahoney', 'Marchetti', 'Merrick', 'Mulcahy',
  'Nairn', 'Norquay', 'Oakley', 'Ormsby', 'Pemberton', 'Pendlebury', 'Quigley', 'Radcliffe',
  'Rathbone', 'Redmond', 'Rickards', 'Sandford', 'Sheridan', 'Somerville', 'Stapleton',
  'Tarrant', 'Thackeray', 'Tregear', 'Underwood', 'Vasey', 'Wakefield', 'Whitlock', 'Wyndham',
];

/** Mirrors the position spread in the real dataset — most rows are blank. */
const POSITIONS = [
  ...Array(60).fill(''),
  'Project Manager', 'Project Manager', 'Project Manager', 'Project Manager', 'Project Manager',
  'Estimator', 'Estimator', 'Estimator', 'Director', 'Director', 'Director',
  'Design Manager', 'Construction Manager', 'Operations Manager', 'Owner',
  'Drafting Manager', 'Drafting', 'Operations', 'Engineer', 'Architect', 'Accounts', 'Admin',
];

/** ACMA reserves 02 5550 xxxx for fiction — it can never be connected. */
const fictionLandline = (n) => `02 5550 ${String(1000 + (n % 9000))}`;
/** ACMA reserves 0491 570 006 – 0491 570 156 for fiction. */
const fictionMobile = (n) => `0491 570 ${String(6 + (n % 151)).padStart(3, '0')}`;

const slug = (name) =>
  name
    .toLowerCase()
    .replace(/[^a-z0-9]+/g, '')
    .slice(0, 18) || 'client';

const emptyPhone = () => ({ value: '', type: 'Work', cc: DEFAULT_CC });
const emptyEmail = () => ({ value: '', type: 'Work' });

// ── Companies ────────────────────────────────────────────────────────────────
const usedNames = new Set();
const companies = [];
for (let i = 0; i < COMPANY_COUNT; i++) {
  let name;
  do {
    const tail = pick(COMPANY_TAIL);
    name = `${pick(COMPANY_HEAD)} ${pick(COMPANY_MID)}${tail ? ` ${tail}` : ''}`;
  } while (usedNames.has(name));
  usedNames.add(name);

  const [town, postcode] = pick(TOWNS);
  // 9 of 512 rows in the source export have no address; keep that gap.
  const address = i % 57 === 0 ? '' : `${1 + Math.floor(rnd() * 240)} ${pick(STREETS)}, ${town}, NSW, ${postcode}, Australia`;

  companies.push({
    id: `co-${String(i + 1).padStart(5, '0')}`,
    name,
    labels: '',
    address,
    phones: [emptyPhone()],
    emails: [emptyEmail()],
  });
}

// ── People ───────────────────────────────────────────────────────────────────
const persons = [];
for (let i = 0; i < PERSON_COUNT; i++) {
  const first = pick(FIRST);
  const last = pick(LAST);
  const name = `${first} ${last}`;

  // 79 of 645 people in the source export are not linked to a company.
  const linked = !chance(12);
  const company = linked ? companies[Math.floor(rnd() * companies.length)] : null;
  const domain = company ? `${slug(company.name)}.example.com` : 'example.com';

  const phones = [];
  if (chance(43)) {
    phones.push({ value: fictionLandline(i * 7), type: 'Work', cc: DEFAULT_CC });
    if (chance(6)) phones.push({ value: fictionMobile(i * 3), type: 'Mobile', cc: DEFAULT_CC });
  } else {
    phones.push(emptyPhone());
  }

  const emails = [];
  if (chance(83)) {
    emails.push({ value: `${first.toLowerCase()}.${last.toLowerCase()}@${domain}`, type: 'Work' });
    if (chance(3)) emails.push({ value: `accounts@${domain}`, type: 'Work' });
  } else {
    emails.push(emptyEmail());
  }

  persons.push({
    id: `pe-${String(i + 1).padStart(5, '0')}`,
    name,
    organizationId: company ? company.id : '',
    position: pick(POSITIONS),
    phones,
    emails,
    activities: [],
    attachments: [],
  });
}

fs.writeFileSync(path.join(outDir, 'crm-companies.json'), JSON.stringify(companies));
fs.writeFileSync(path.join(outDir, 'crm-persons.json'), JSON.stringify(persons));
fs.writeFileSync(
  path.join(outDir, 'crm-static-meta.json'),
  `${JSON.stringify(
    {
      generatedAt: '2026-06-04T14:52:50.506Z',
      companyCount: companies.length,
      personCount: persons.length,
      source: 'scripts/generate-sample-crm.mjs (synthetic sample data)',
    },
    null,
    2,
  )}\n`,
);

console.log(`Wrote ${companies.length} companies and ${persons.length} people to ${outDir}`);
