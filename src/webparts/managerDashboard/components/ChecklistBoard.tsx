import * as React from 'react';
import { IProject } from '../../../shared/models/IProject';

// ───────────────────────────── Types ─────────────────────────────
type SectionType = 'steel' | 'concrete' | 'both';
type ProjectType = 'steel' | 'concrete' | 'both';
type Role = 'detailer' | 'checker' | 'pm';
type C2Action = 'cleared' | 'na' | 'incorrect';

interface ISection {
  title: string;
  type: SectionType;
  items: Array<[string, string]>; // [text, taskCode]
}
interface IPhase {
  id: string;
  name: string;
  sections: ISection[];
}
interface IItemState {
  c1?: boolean;
  c2?: C2Action | null;
  c1By?: string;
  c2By?: string;
  c1At?: string;
  c2At?: string;
  override?: boolean;
  overrideBy?: string;
  overrideAt?: string;
  overrideReason?: string;
}
interface IOverrideLog {
  itemId: string;
  by: string;
  at: string;
  reason: string;
  itemText: string;
  taskCode: string;
  action: C2Action;
}

// ───────────────────────────── Data ─────────────────────────────
const CHECKLIST: IPhase[] = [
  { id: '01', name: 'Pre-Project Setup', sections: [
    { title: 'Project creation & dashboard setup', type: 'both', items: [
      ['Create project in 3 Edge dashboard — assign 3E project number', '01a'],
      ['Client company, primary contact and project name entered correctly', '01a'],
      ['Quote number linked to project record', '01a'],
      ['Project type confirmed: Steel / Concrete / Steel & Concrete', '01a'],
      ['Start date, finish date and hours allowed set (required fields)', '01a'],
      ['Assigned staff selected — detailer(s) and project manager confirmed', '01a'],
      ['Project status set to ACTIVE', '01a'],
    ]},
    { title: 'Pre-project client checklist', type: 'both', items: [
      ['Project details: name, number, site address, client contacts confirmed', '01a'],
      ['Design documentation: engineer + architect drawings received, CAD files with grids provided, IFC model received (if applicable), all docs reviewed', '01a'],
      ['Project timeline: 3 Edge start date, IFA date, IFC date, site start date, key milestones confirmed', '01a'],
      ['Previous meetings: relevant meeting notes / minutes attached', '01a'],
      ['Communication: preferred email subject line set, all client + architect + engineer contacts listed with email and mobile', '01a'],
    ]},
    { title: 'Software & file setup', type: 'both', items: [
      ['Tekla model folder structure created per 3 Edge standard', '01d'],
      ['Project properties set: units (mm), grid, levels, north point', '01d'],
      ['Correct profile catalogue loaded — verify AS/NZS sections before modelling', '01d'],
      ['3 Edge company template with standard components loaded', '01d'],
      ['Model shared to team / server location confirmed and accessible', '01d'],
      ['Backup / auto-save interval confirmed and active', '01d'],
      ['IFC import of reference model from architect completed (if provided)', '01d'],
      ['Reference model revision matches latest issued drawings', '01d'],
    ]},
  ]},
  { id: '02', name: 'Grid Setup', sections: [
    { title: 'Grid & reference geometry', type: 'both', items: [
      ['Check Architectural grid against Structural grid for discrepancies', '02a'],
      ['Check Tekla Model grid complies with design information', '02a'],
      ['Check model orientation correctly set and aligned with Project North', '02a'],
      ['Levels / storeys set correctly (top of steel, not finished floor unless specified)', '02a'],
      ['Reference planes created for complex geometry (rakers, roof pitches) where required', '02a'],
      ['Setting-out points for column bases confirmed against structural drawings', '02a'],
    ]},
  ]},
  { id: '03', name: 'Stick Model', sections: [
    { title: 'Coordination & engineering inputs', type: 'both', items: [
      ['Camber requirements confirmed for beams (if any)', '03a'],
      ['Hold-down bolt templates and anchor bolt layouts received from engineer', '03a'],
      ['Embedded steel / cast-in plates coordinated with concrete contractor', '03a'],
      ['Fire rating requirements noted — affects intumescent spec and section selection', '03a/NA'],
      ['Revision log created — track drawing issue dates and superseded sheets', '01j'],
    ]},
    { title: 'Primary steel members', type: 'steel', items: [
      ['All columns correctly oriented / located per design drawings', '03c'],
      ['All columns correctly modelled (start point at bottom)', '03c'],
      ['All beams correctly modelled (location and elevation)', '03c'],
      ['All beams modelled with top face up', '03c'],
      ['All purlins correct spacing per Eng specs', '03c'],
      ['All purlins correct sizing per Eng specs', '03c'],
      ['All purlins correct class per Eng specs', '03c'],
      ['All girts correct spacing per Eng specs', '03c'],
      ['All girts correct sizing per Eng specs', '03c'],
      ['All girts correct class per Eng specs', '03c'],
      ['Bridging as per Eng specs', '03c'],
      ['All bracing correctly faced / orientated / located', '03c'],
      ['Fly bracing as per Eng specs', '03c'],
      ['Beams spanning across columns modelled to final extent', '03c'],
      ['Beams at different elevations extended to match design', '03c'],
      ['Cantilever beams extended to support incoming beam', '03c'],
      ['All skewed / sloping members modelled with enough length', '03c'],
      ['Transfer beams and cranked beams modelled with correct geometry', '03c'],
      ['Door opening — all members as per Eng specs', '03c'],
      ['Window opening — all members as per Eng specs', '03c'],
      ['Roof penetrations — all members as per Eng specs', '03c'],
      ['Door opening — all sizing, location and clearances as per Arch specs', '03c'],
      ['Window opening — all sizing, location & clearances as per Arch specs', '03c'],
      ['Roof penetration — all sizing, location and clearances as per Arch specs', '03c'],
    ]},
    { title: 'Member properties, naming & attributes', type: 'both', items: [
      ['All section sizes match IFC drawings — no assumptions from similar projects', '03a'],
      ['All members in the correct Sequence or Phase', '03a'],
      ['All members named correctly (refer to Client Manual)', '03a'],
      ['All members numbered and prefixed correctly per 3 Edge standard', '03a'],
      ['Main parts and secondary parts correctly assigned', '03a'],
      ['No duplicate assemblies where different geometry exists', '03a'],
      ['All members have the correct material grade (verify zones)', '03a'],
      ['All cambers correct and with correct format', '03a'],
      ['Interface members modelled correctly per design drawings', '03a'],
    ]},
    { title: 'Member colour, finish & notes', type: 'both', items: [
      ['All members coloured correctly per 3 Edge colour standard', '06b'],
      ["All 'on hold' members coloured with communication info populated", '06b'],
      ['Remarks (BOM) field utilised correctly', '06b'],
      ['Special Notes field utilised correctly', '06b'],
      ["UDA (User Defined Attributes) 'Notes' tab utilised correctly", '06b'],
      ['All specified finish / paint notes correct', '06b'],
      ['Correct surface treatment / paint system and fire protection noted', '06b'],
    ]},
    { title: 'Fabrication & shipping', type: 'both', items: [
      ['Check stock length of the steel', '04d'],
      ['All galvanised assemblies within size limit', '04d'],
      ['All assemblies are shippable', '04d'],
    ]},
  ]},
  { id: '04', name: 'Connections', sections: [
    { title: 'Custom Components & Connection Check — STEEL', type: 'steel', items: [
      ['Connections comply with design drawings', '04d'],
      ['Connection material is the correct profile', '04d'],
      ['Connection material is the correct grade', '04d'],
      ['Parts are named correctly', '04d'],
      ['Parts have correct prefix and start numbers', '04d'],
      ['Parts are phased and sequenced correctly', '04d'],
      ['Connections have been Clash Checked', '04d'],
      ['Connection materials welded or bolted to assembly', '04d'],
      ['Connections checked for erectability', '04d'],
      ['Connections checked for bolt clearances', '04d'],
      ['Connections checked for weld access', '04d'],
      ['Check washer and bolt requirements', '04d'],
      ['Bolt grade and AS 4100 category confirmed', '04d'],
      ['Thread exclusion noted where required (X-type)', '04d'],
      ['Check hole tolerances are correct', '04d'],
      ['Check bolts standard edge distance', '04d'],
      ['Min bolt edge/end distances comply with AS 4100 Table 9.5.1', '04d'],
      ['Check shear tabs on correct side', '04d'],
      ['Check fabrication and erectability of ALL assemblies', '04d'],
      ['Check main part of welded assemblies', '04d'],
      ['All welds checked per Welds Modelling guidelines', '04d'],
      ['Weld category (SP or GP) per AS/NZS 1554 confirmed', '04d'],
      ['All site welds correctly designated with field weld flag', '04d'],
      ['Provide vent holes for galvanised steel', '04d'],
      ['Provide 2mm gap for stiffeners', '04d'],
      ['Moment connections: stiffeners and plates modelled per design', '04d'],
    ]},
    { title: 'Column bases & holding down — STEEL', type: 'steel', items: [
      ['Base plate size, thickness and grade confirmed per design drawing', '04d'],
      ['Holding down bolt pattern matches structural drawing exactly', '04d'],
      ['HD bolt diameter, grade and projection above grout confirmed', '04d'],
      ['Oversize / slotted holes in base plate per erection tolerance spec', '04d'],
      ['Grout depth and packing arrangement confirmed with engineer', '04d'],
      ['Shear key or shear stud detail confirmed where applicable', '04d'],
      ['Foundation recess / pocket detail coordinated with concrete contractor', '04d'],
      ['Backing bars / run-off tabs specified on CJP (complete joint penetration) welds', '04d'],
    ]},
    { title: 'Custom Components & Connections Check — CONCRETE', type: 'concrete', items: [
      ['Connections comply with design drawings', '04c'],
      ['Connection material is the correct profile', '04c'],
      ['Connection material is the correct grade', '04c'],
      ['Parts are named correctly', '04c'],
      ['Parts have correct prefix and start numbers', '04c'],
      ['Parts are phased and sequenced correctly', '04c'],
      ['Connections have been Clash Checked', '04c'],
      ['Connection materials welded or bolted to assembly', '04c'],
      ['Connections checked for erectability', '04c'],
      ['Connections checked for bolt clearances', '04c'],
      ['Connections checked for weld access', '04c'],
      ['Check washer and bolt requirements', '04c'],
      ['Bolt grade and AS 4100 category confirmed', '04c'],
      ['Thread exclusion noted where required (X-type)', '04c'],
      ['Check hole tolerances are correct', '04c'],
      ['Check bolts standard edge distance', '04c'],
      ['Min bolt edge/end distances comply with AS 4100 Table 9.5.1', '04c'],
      ['Check shear tabs on correct side', '04c'],
      ['Check fabrication and erectability of ALL assemblies', '04c'],
      ['Check main part of welded assemblies', '04c'],
      ['All welds checked per Welds Modelling guidelines', '04c'],
      ['Weld category (SP or GP) per AS/NZS 1554 confirmed', '04c'],
      ['All site welds correctly designated with field weld flag', '04c'],
      ['Moment connections: stiffeners and plates modelled per design', '04c'],
    ]},
  ]},
  { id: '05', name: 'Drawings Production', sections: [
    { title: 'General drawing QA — STEEL', type: 'steel', items: [
      ['Title block complete and correct', '05a'],
      ['Revision information correct and matches transmittal', '05a'],
      ['All design drawing references show correct latest revision', '05a'],
      ['Drawing statuses correct (IFA / IFC) and consistent', '05a'],
      ['Standard / general notes block present with required info', '05a'],
      ['All changes from previous revision correctly clouded', '05a'],
      ['Transmittal / drawing register prepared and reviewed', '05a'],
    ]},
    { title: 'General drawing QA — CONCRETE', type: 'concrete', items: [
      ['Title block complete and correct', '05a'],
      ['Revision information correct and matches transmittal', '05a'],
      ['All design drawing references show correct latest revision', '05a'],
      ['Drawing statuses correct (IFA / IFC) and consistent', '05a'],
      ['Standard / general notes block present with required info', '05a'],
      ['All changes from previous revision correctly clouded', '05a'],
      ['Transmittal / drawing register prepared and reviewed', '05a'],
    ]},
    { title: 'Assembly drawing — STEEL', type: 'steel', items: [
      ['Overall dimensions correct', '05b'],
      ['Running dimensions on holes correct', '05b'],
      ['Running dimensions for cleat positions correct', '05b'],
      ['All necessary sections / views present', '05b'],
      ['All quantities match the BOM', '05b'],
      ['BOM matches dimensions on drawings', '05b'],
      ['Special welding shown where required', '05b'],
      ['PFC web near side correct (holes only)', '05b'],
      ['PFC near side correct (with cleats)', '05b'],
      ['Weld symbols comply with AS 1101.3', '05b'],
      ['Section cut markers correctly reference sheet/detail', '05b'],
      ['Assembly marks consistent across all views and BOM', '05b'],
      ['Bolt specification schedule present', '05b'],
      ['Copes and notches fully dimensioned with radius at re-entrant corners', '05b'],
      ['Weight of member shown in title block', '05b'],
      ['All client standards met/reflected in drawings', '05b'],
    ]},
    { title: 'Assembly drawing — CONCRETE', type: 'concrete', items: [
      ['Overall dimensions correct', '05c'],
      ['Running dimensions on holes correct', '05c'],
      ['Running dimensions for cleat positions correct', '05c'],
      ['All necessary sections / views present', '05c'],
      ['All quantities match the BOM', '05c'],
      ['BOM matches dimensions on drawings', '05c'],
      ['Special welding shown where required', '05c'],
      ['PFC web near side correct (holes only)', '05c'],
      ['PFC near side correct (with cleats)', '05c'],
      ['Weld symbols comply with AS 1101.3', '05c'],
      ['Section cut markers correctly reference sheet/detail', '05c'],
      ['Assembly marks consistent across all views and BOM', '05c'],
      ['Bolt specification schedule present', '05c'],
      ['Copes and notches fully dimensioned with radius at re-entrant corners', '05c'],
      ['Weight of member shown in title block', '05c'],
      ['All client standards met/reflected in drawings', '05c'],
    ]},
    { title: 'Marking plan — column set-out — STEEL', type: 'steel', items: [
      ['Location of each column provided', '05d'],
      ['Mark for each column provided (BP type, elevations)', '05d'],
      ['Details for each base plate type provided', '05d'],
      ['Holding down bolt layouts shown on foundation plan with coordinates or offsets', '05d'],
    ]},
    { title: 'Marking plan — framing plan — STEEL', type: 'steel', items: [
      ['All assembly marks shown', '05d'],
      ['Drawing scale correct', '05d'],
      ['North arrow correct', '05d'],
      ['Location of all steel members provided', '05d'],
      ['All necessary sections / views present', '05d'],
      ['Details/location for site welds and site brackets provided', '05d'],
      ['Interface dimensions shown and agree with design', '05d'],
    ]},
    { title: 'Marking plan — framing plan — CONCRETE', type: 'concrete', items: [
      ['All assembly marks shown', '05d'],
      ['Drawing scale correct', '05d'],
      ['North arrow correct', '05d'],
      ['Location of all steel members provided', '05d'],
      ['All necessary sections / views present', '05d'],
      ['Details/location for site welds and site brackets provided', '05d'],
      ['Interface dimensions shown and agree with design', '05d'],
    ]},
    { title: 'Marking plan — purlin plan — STEEL', type: 'steel', items: [
      ['All purlin marks shown', '05d'],
      ['All bridging marks shown', '05d'],
      ['Purlin bundle number provided', '05d'],
      ['Purlin plan N/A confirmed (if applicable)', '05d'],
    ]},
    { title: 'Elevations and sections — STEEL', type: 'steel', items: [
      ['RLs for holes and top of steel provided', '05d'],
      ['All assembly marks shown in elevations / sections — IFA drawings need to show profile', '05d'],
      ['Drawing set reflects current approved model state', '05d'],
    ]},
    { title: 'Elevations and sections — CONCRETE', type: 'concrete', items: [
      ['RLs for holes and top of steel provided', '05d'],
      ['All assembly marks shown in elevations / sections — IFA drawings need to show profile', '05d'],
      ['Drawing set reflects current approved model state', '05d'],
    ]},
    { title: 'Erection & assembly drawings — STEEL', type: 'steel', items: [
      ['Erection sequence / zone breakdown clear to site team', '05d'],
      ['Temporary bracing positions noted (if required by engineer)', '05d'],
      ['Shear stud layout shown on composite beams', '05d'],
      ['Pre-cambered members flagged for erector', '05d'],
      ['Slotted holes and adjustment range noted for erection tolerance', '05d'],
    ]},
  ]},
  { id: '06', name: 'Model Review', sections: [
    { title: 'Model QA — final verification — STEEL', type: 'steel', items: [
      ['All TOS elevations for internal members correct', '06b'],
      ['All members accounted for per design drawings', '06b'],
      ['All members have the correct profile', '06b'],
      ['All shear stud layouts shown and correct', '06b'],
      ['Connection locations and types comply with design intent', '04d'],
      ['No hard clashes between steel–steel and steel–concrete', '04d'],
      ['Erection clearance adequate at connections', '04d'],
      ['Holes, notches and web penetrations coordinated', '04d'],
      ['All open RFIs reviewed and model updated', '01f'],
      ['Issues logged in project issue register — RFI register', '01f'],
      ['Weld objects placed for all connections', '04d'],
      ['No zero-length members or duplicate members in model', '06b'],
      ['Steel model clashed against MEP / services model (if provided)', '06b'],
      ['Cladding rails / support steelwork clearance to structure confirmed', '06b'],
    ]},
    { title: 'Model QA — final verification — CONCRETE', type: 'concrete', items: [
      ['All TOS elevations for internal members correct', '06b'],
      ['All members accounted for per design drawings', '06b'],
      ['All members have the correct profile', '06b'],
      ['All shear stud layouts shown and correct', '06b'],
      ['Connection locations and types comply with design intent', '04c'],
      ['No hard clashes between steel–steel and steel–concrete', '04c'],
      ['Erection clearance adequate at connections', '04c'],
      ['Holes, notches and web penetrations coordinated', '04c'],
      ['All open RFIs reviewed and model updated', '01f'],
      ['Issues logged in project issue register — RFI register', '01f'],
      ['Weld objects placed for all connections', '04c'],
      ['No zero-length members or duplicate members in model', '06b'],
      ['Steel model clashed against MEP / services model (if provided)', '06b'],
      ['Cladding rails / support steelwork clearance to structure confirmed', '06b'],
    ]},
  ]},
  { id: '07', name: 'Post-IFC', sections: [
    { title: 'Drawing issue & transmittal — STEEL', type: 'steel', items: [
      ['Transmittal prepared — all drawing numbers and revisions listed', '07b'],
      ['Correct issue status on all drawings (IFA / IFC / AFC / As-Built)', '07b'],
      ['PDF and native files issued as per contract requirement', '07b'],
      ['Drawing register updated immediately on issue', '07b'],
      ['Engineer / consultant approval obtained for IFC issue (if required by contract)', '07c'],
      ['Superseded drawings marked as void in register', '01j'],
      ['Client / builder confirmation of IFA/IFC package receipt obtained', '07b'],
    ]},
    { title: 'Drawing issue & transmittal — CONCRETE', type: 'concrete', items: [
      ['Transmittal prepared — all drawing numbers and revisions listed', '07b'],
      ['Correct issue status on all drawings (IFA / IFC / AFC / As-Built)', '07b'],
      ['PDF and native files issued as per contract requirement', '07b'],
      ['Drawing register updated immediately on issue', '07b'],
      ['Engineer / consultant approval obtained for IFC issue (if required by contract)', '07c'],
      ['Superseded drawings marked as void in register', '01j'],
      ['Client / builder confirmation of IFA/IFC package receipt obtained', '07b'],
    ]},
    { title: 'Fabrication support — IFA', type: 'both', items: [
      ['RFI log current — all fabrication queries have unique number and response', '07b'],
      ['All verbal instructions confirmed in writing (email minimum)', '07b'],
      ['Revised drawings issued with revision cloud highlighting all changes', '07b'],
      ['Revision history table on drawing updated with change descriptions', '07b'],
      ['Hold items tracked — no drawings issued while holds remain unresolved', '07b'],
      ['Material test certificates (MTCs) requested and filed for traceability', '07b'],
    ]},
    { title: 'Close Out', type: 'both', items: [
      ['As-built drawings prepared — all site changes incorporated and noted', '07c'],
      ['Final Tekla model issued to client / BIM manager in agreed format', '07c'],
      ['IFC export of final model completed (if required by BIM protocol) ', '07c'],
      ['All drawing files archived per 3 Edge document control procedure', '08'],
      ['All project RFIs closed or formally handed over to client', '08'],
      ['Project lessons learned documented — hours variance, RFI patterns, rework causes', '08'],
      ['Internal quality review completed — checker and PM sign-off', '08'],
      ['Project folder audit — confirm all deliverables accounted for and filed', '08'],
    ]},
  ]},
];

// ───────────────────────────── Helpers ─────────────────────────────
const itemIdOf = (pi: number, si: number, ii: number): string => `p${pi}s${si}i${ii}`;
const nowString = (): string => {
  const d = new Date();
  return d.toLocaleString('en-AU', { day: '2-digit', month: 'short', hour: '2-digit', minute: '2-digit', hour12: true });
};
const disciplineToType = (d: string): ProjectType => {
  const t = (d || '').toLowerCase();
  if (t.indexOf('steel') >= 0 && t.indexOf('concrete') >= 0) return 'both';
  if (t.indexOf('concrete') >= 0) return 'concrete';
  return 'steel';
};
const storageKey = (projId: string): string => `3edge_checklist_v1_${projId}`;

interface IPersisted {
  items: Record<string, IItemState>;
  overrides: IOverrideLog[];
  projectType: ProjectType;
  role: Role;
  currentPhase: string;
}

// ───────────────────────────── Props ─────────────────────────────
interface ChecklistBoardProps {
  projects: IProject[];
  userDisplayName: string;
  isManager: boolean;
  toast: (msg: string) => void;
}

// ───────────────────────────── Sub-components ─────────────────────────────
interface StatCardProps { label: string; value: string; sub: string; col: string; valueCol?: string; }
const StatCard: React.FC<StatCardProps> = ({ label, value, sub, col, valueCol }) => (
  <div style={{ background: 'var(--s1)', padding: '12px 14px', borderRadius: 8, borderTop: `3px solid ${col}`, borderLeft: '1px solid var(--bd)', borderRight: '1px solid var(--bd)', borderBottom: '1px solid var(--bd)' }}>
    <div style={{ fontSize: 10, color: 'var(--t2)', letterSpacing: '.08em', fontWeight: 700, marginBottom: 4, textTransform: 'uppercase' }}>{label}</div>
    <div style={{ fontSize: 22, fontWeight: 700, color: valueCol || 'var(--t1)' }}>{value}</div>
    <div style={{ fontSize: 11, color: 'var(--t3)', marginTop: 2, fontWeight: 600 }}>{sub}</div>
  </div>
);

const StatusPill: React.FC<{ st: IItemState }> = ({ st }) => {
  let bg = 'var(--s2)', color = 'var(--t4)', text = 'WAITING C1';
  if (st.override) { bg = 'rgba(107,79,200,.15)'; color = '#4a2f9c'; text = 'PM CLEARED'; }
  else if (st.c2 === 'cleared') { bg = 'rgba(42,158,42,.15)'; color = '#157a15'; text = 'CLEARED'; }
  else if (st.c2 === 'na') { bg = 'rgba(90,110,136,.15)'; color = 'var(--t2)'; text = 'N/A'; }
  else if (st.c2 === 'incorrect') { bg = 'rgba(204,51,51,.15)'; color = '#b82020'; text = 'FIX REQUIRED'; }
  else if (st.c1) { bg = 'rgba(46,109,180,.15)'; color = '#1d5ec0'; text = 'READY FOR C2'; }
  return <span style={{ fontSize: 10, padding: '3px 8px', borderRadius: 4, letterSpacing: '.05em', fontWeight: 700, background: bg, color, display: 'inline-block' }}>{text}</span>;
};

// ───────────────────────────── Component ─────────────────────────────
const ChecklistBoard: React.FC<ChecklistBoardProps> = ({ projects, userDisplayName, isManager, toast }) => {
  const mainProjects = React.useMemo(() => projects.filter(p => !p.isEwo), [projects]);
  const [selProjId, setSelProjId] = React.useState<string>('');
  const [role, setRole] = React.useState<Role>(isManager ? 'pm' : 'detailer');
  const [projectType, setProjectType] = React.useState<ProjectType>('steel');
  const [currentPhase, setCurrentPhase] = React.useState<string>('01');
  const [items, setItems] = React.useState<Record<string, IItemState>>({});
  const [overrides, setOverrides] = React.useState<IOverrideLog[]>([]);

  // Override modal state
  const [ovModal, setOvModal] = React.useState<{ id: string | null; action: C2Action; reason: string }>({ id: null, action: 'cleared', reason: '' });

  // Auto-select first project when loaded
  React.useEffect(() => {
    if (!selProjId && mainProjects.length > 0) {
      const p = mainProjects[0];
      setSelProjId(p.id);
    }
  }, [mainProjects, selProjId]);

  // Load persisted state when project changes
  React.useEffect(() => {
    if (!selProjId) return;
    try {
      const raw = window.localStorage.getItem(storageKey(selProjId));
      if (raw) {
        const d: IPersisted = JSON.parse(raw);
        setItems(d.items || {});
        setOverrides(d.overrides || []);
        setProjectType(d.projectType || 'steel');
        setCurrentPhase(d.currentPhase || '01');
        if (d.role) setRole(d.role);
        return;
      }
    } catch (e) { /* ignore */ }
    // new project — reset and pick type from discipline
    const proj = mainProjects.filter(p => p.id === selProjId)[0];
    setItems({});
    setOverrides([]);
    setCurrentPhase('01');
    setProjectType(proj ? disciplineToType(proj.discipline) : 'steel');
  // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [selProjId]);

  // Persist whenever state changes
  React.useEffect(() => {
    if (!selProjId) return;
    try {
      const payload: IPersisted = { items, overrides, projectType, role, currentPhase };
      window.localStorage.setItem(storageKey(selProjId), JSON.stringify(payload));
    } catch (e) { /* ignore */ }
  }, [selProjId, items, overrides, projectType, role, currentPhase]);

  // ─── Derived lists ───
  const allFlat = React.useMemo(() => {
    const list: Array<{ id: string; phaseIdx: number; sectionIdx: number; itemIdx: number; text: string; taskCode: string; sectionType: SectionType; phaseId: string }> = [];
    CHECKLIST.forEach((phase, pi) => {
      phase.sections.forEach((section, si) => {
        section.items.forEach((item, ii) => {
          list.push({ id: itemIdOf(pi, si, ii), phaseIdx: pi, sectionIdx: si, itemIdx: ii, text: item[0], taskCode: item[1], sectionType: section.type, phaseId: phase.id });
        });
      });
    });
    return list;
  }, []);

  const isSectionVisible = (t: SectionType): boolean => {
    if (projectType === 'both') return true;
    if (t === 'both') return true;
    return t === projectType;
  };

  const visible = React.useMemo(() => allFlat.filter(it => isSectionVisible(it.sectionType)), [allFlat, projectType]);

  const getSt = (id: string): IItemState => items[id] || { c1: false, c2: null };

  const isResolved = (s: IItemState): boolean => !!s.override || s.c2 === 'cleared' || s.c2 === 'na';

  // ─── Stats ───
  const stats = React.useMemo(() => {
    const total = visible.length;
    const c1 = visible.filter(it => { const s = getSt(it.id); return s.c1 || s.override; }).length;
    const cleared = visible.filter(it => { const s = getSt(it.id); return s.c2 === 'cleared' || s.override; }).length;
    const na = visible.filter(it => getSt(it.id).c2 === 'na').length;
    const incorrect = visible.filter(it => getSt(it.id).c2 === 'incorrect').length;
    const ovCount = visible.filter(it => getSt(it.id).override).length;
    const pct = (n: number): number => total === 0 ? 0 : Math.round((n / total) * 100);
    return { total, c1, cleared, na, incorrect, overrides: ovCount, pct };
  // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [visible, items]);

  // ─── Actions ───
  const users = React.useMemo(() => ({
    detailer: userDisplayName || 'Detailer',
    checker: userDisplayName || 'Checker',
    pm: userDisplayName || 'PM',
  }), [userDisplayName]);

  const toggleC1 = (id: string): void => {
    if (role !== 'detailer') { toast(`Only the Detailer can tick Check 1. You're viewing as ${role.toUpperCase()}.`); return; }
    const st = items[id] || {};
    if (st.c2 === 'cleared' || st.c2 === 'na') { toast(`Check 2 already actioned — checker must clear the status first.`); return; }
    if (st.override) { toast('Item cleared by PM override — contact PM to reset.'); return; }
    setItems(prev => {
      const cur = prev[id] || {};
      if (cur.c1) return { ...prev, [id]: { ...cur, c1: false, c1By: undefined, c1At: undefined } };
      return { ...prev, [id]: { ...cur, c1: true, c1By: users.detailer, c1At: nowString() } };
    });
  };

  const handleC2Click = (id: string, action: C2Action): void => {
    const st = items[id] || {};
    if (role === 'detailer') { toast("Detailers can't action Check 2. Switch to Checker or PM role."); return; }

    if (st.override) {
      if (role === 'pm') {
        if (window.confirm('Revert this PM override? The item will return to its original state.')) {
          setItems(prev => { const p = { ...prev }; delete p[id]; return p; });
          setOverrides(prev => prev.filter(o => o.itemId !== id));
          toast('Override reverted.');
        }
        return;
      }
      toast('Only the PM can revert an override.');
      return;
    }

    // click same button → clear to neutral
    if (st.c2 === action) {
      setItems(prev => ({ ...prev, [id]: { ...(prev[id] || {}), c2: null, c2By: undefined, c2At: undefined } }));
      return;
    }

    // locked state — need C1 first (or PM override)
    if (!st.c1) {
      if (role === 'pm') {
        setOvModal({ id, action, reason: '' });
      } else {
        toast('Check 2 is locked — Check 1 must be ticked first. Switch to PM role to override.');
      }
      return;
    }

    setItems(prev => ({ ...prev, [id]: { ...(prev[id] || {}), c2: action, c2By: users[role], c2At: nowString() } }));
    if (action === 'incorrect') toast('Item flagged as Incorrect — detailer will be notified to fix.');
  };

  const confirmOverride = (): void => {
    if (!ovModal.id || !ovModal.reason.trim()) return;
    const meta = allFlat.filter(it => it.id === ovModal.id)[0];
    if (!meta) return;
    const now = nowString();
    const action = ovModal.action;
    setItems(prev => {
      if (action === 'cleared') {
        return { ...prev, [ovModal.id!]: { c1: false, c2: null, override: true, overrideBy: users.pm, overrideAt: now, overrideReason: ovModal.reason } };
      }
      return { ...prev, [ovModal.id!]: { c1: false, c2: action, c2By: users.pm, c2At: now } };
    });
    setOverrides(prev => [...prev, { itemId: ovModal.id!, by: users.pm, at: now, reason: ovModal.reason, itemText: meta.text, taskCode: meta.taskCode, action }]);
    const label = action === 'na' ? 'N/A' : action === 'incorrect' ? 'Incorrect' : 'Cleared';
    toast(`PM override applied (${label}) — logged to audit trail.`);
    setOvModal({ id: null, action: 'cleared', reason: '' });
  };

  const resetAll = (): void => {
    if (!window.confirm('Reset all ticks and overrides on this project? This cannot be undone.')) return;
    setItems({});
    setOverrides([]);
    toast('All ticks cleared.');
  };

  // ─── Render ───
  const selProj = mainProjects.filter(p => p.id === selProjId)[0];

  const roleHintText = role === 'detailer'
    ? <>You&rsquo;re viewing as <strong>Detailer</strong> — you can tick <strong>Check 1</strong> only. Check 2 column is read-only for you.</>
    : role === 'checker'
      ? <>You&rsquo;re viewing as <strong>Checker</strong> — action Check 2 with one of three states: <strong>&#10003; Cleared</strong>, <strong>N/A</strong>, or <strong>&#10007; Incorrect</strong>. Check 1 must be ticked first.</>
      : <>You&rsquo;re viewing as <strong>PM</strong> — you can clear locked Check 2 items via override. A reason is required and all overrides are audit-logged.</>;

  const hintBg = role === 'pm' ? 'rgba(107,79,200,.08)' : role === 'detailer' ? 'rgba(42,158,42,.08)' : 'rgba(46,109,180,.08)';
  const hintBd = role === 'pm' ? 'var(--pu)' : role === 'detailer' ? 'var(--3eg)' : 'var(--bl)';
  const hintColor = role === 'pm' ? '#4a2f9c' : role === 'detailer' ? '#157a15' : '#1d5ec0';

  const badgeStyle = (t: ProjectType): React.CSSProperties => ({
    fontFamily: 'Montserrat', fontWeight: 700, fontSize: 10, letterSpacing: '.08em', padding: '3px 8px', borderRadius: 4, textTransform: 'uppercase',
    background: t === 'steel' ? 'rgba(46,109,180,.12)' : t === 'concrete' ? 'rgba(107,79,200,.12)' : 'rgba(20,150,120,.12)',
    color: t === 'steel' ? '#1d5ec0' : t === 'concrete' ? '#4a2f9c' : '#0d7a56',
    border: `1px solid ${t === 'steel' ? '#4a90d9' : t === 'concrete' ? '#9b7fe8' : '#10b981'}`,
  });

  return (
    <div style={{ fontFamily: 'Montserrat' }}>
      {/* ── Project Context Bar ── */}
      <div style={{ background: 'var(--s1)', border: '1px solid var(--bd)', borderRadius: 8, padding: '12px 16px', marginBottom: 14, display: 'flex', alignItems: 'center', gap: 14, flexWrap: 'wrap' }}>
        <span style={{ fontSize: 11, color: 'var(--t4)', letterSpacing: '.06em', fontWeight: 700, textTransform: 'uppercase' }}>Project</span>
        <select
          value={selProjId}
          onChange={e => setSelProjId(e.target.value)}
          style={{ fontFamily: 'Montserrat', fontSize: 13, fontWeight: 600, padding: '7px 10px', border: '1px solid var(--bd)', borderRadius: 4, background: 'var(--s2)', color: 'var(--t1)', minWidth: 280, cursor: 'pointer', outline: 'none' }}
        >
          {mainProjects.length === 0 && <option value="">No projects available</option>}
          {mainProjects.map(p => (
            <option key={p.id} value={p.id}>{p.projNum} — {p.name}</option>
          ))}
        </select>
        {selProj && (
          <>
            <span style={badgeStyle(projectType)}>{projectType === 'both' ? 'Steel & Concrete' : projectType}</span>
            <span style={{ fontFamily: 'Montserrat', fontWeight: 700, fontSize: 10, letterSpacing: '.08em', padding: '3px 8px', borderRadius: 4, textTransform: 'uppercase', background: 'rgba(42,158,42,.12)', color: '#157a15', border: '1px solid #3db63d' }}>
              {selProj.status}
            </span>
            <span style={{ marginLeft: 'auto', fontSize: 12, color: 'var(--t3)' }}>
              {selProj.detailers && <>Detailer: <span style={{ color: 'var(--t1)', fontWeight: 700 }}>{selProj.detailers}</span> · </>}
              {selProj.teamLead && <>PM: <span style={{ color: 'var(--t1)', fontWeight: 700 }}>{selProj.teamLead}</span></>}
            </span>
          </>
        )}
      </div>

      {!selProjId && (
        <div style={{ padding: 32, textAlign: 'center', color: 'var(--t4)', fontSize: 13 }}>Select a project to view its checklist.</div>
      )}

      {selProjId && (
        <>
          {/* ── Controls ── */}
          <div style={{ display: 'flex', gap: 10, marginBottom: 14, flexWrap: 'wrap', alignItems: 'center' }}>
            <div style={{ background: 'var(--s1)', border: '1px solid var(--bd)', borderRadius: 8, padding: '6px 10px', display: 'flex', alignItems: 'center', gap: 8 }}>
              <span style={{ fontSize: 10, color: 'var(--t4)', letterSpacing: '.08em', fontWeight: 700, textTransform: 'uppercase' }}>Project Type</span>
              <div style={{ display: 'flex', gap: 3 }}>
                {(['steel', 'concrete', 'both'] as ProjectType[]).map(t => (
                  <button key={t} onClick={() => { setProjectType(t); const firstWithItems = CHECKLIST.filter(p => p.sections.some(s => t === 'both' || s.type === 'both' || s.type === t))[0]; if (firstWithItems && !CHECKLIST.filter(p => p.id === currentPhase)[0].sections.some(s => t === 'both' || s.type === 'both' || s.type === t)) setCurrentPhase(firstWithItems.id); }} style={{
                    fontFamily: 'Montserrat', fontSize: 12, fontWeight: 600, padding: '5px 11px', borderRadius: 5, cursor: 'pointer',
                    border: '1px solid ' + (projectType === t ? 'var(--t1)' : 'transparent'),
                    background: projectType === t ? 'var(--t1)' : 'transparent',
                    color: projectType === t ? '#fff' : 'var(--t3)'
                  }}>
                    {t === 'steel' ? 'Steel only' : t === 'concrete' ? 'Concrete only' : 'Steel & Concrete'}
                  </button>
                ))}
              </div>
            </div>

            <div style={{ background: 'var(--s1)', border: '1px solid var(--bd)', borderRadius: 8, padding: '6px 10px', display: 'flex', alignItems: 'center', gap: 8 }}>
              <span style={{ fontSize: 10, color: 'var(--t4)', letterSpacing: '.08em', fontWeight: 700, textTransform: 'uppercase' }}>View as</span>
              <div style={{ display: 'flex', gap: 3 }}>
                {(['detailer', 'checker', 'pm'] as Role[]).map(r => (
                  <button key={r} onClick={() => setRole(r)} style={{
                    fontFamily: 'Montserrat', fontSize: 12, fontWeight: 600, padding: '5px 11px', borderRadius: 5, cursor: 'pointer',
                    border: '1px solid ' + (role === r ? 'var(--t1)' : 'transparent'),
                    background: role === r ? 'var(--t1)' : 'transparent',
                    color: role === r ? '#fff' : 'var(--t3)'
                  }}>
                    {r === 'pm' ? 'PM' : r.charAt(0).toUpperCase() + r.slice(1)}
                  </button>
                ))}
              </div>
            </div>

            <div style={{ flex: 1 }} />

            <button onClick={resetAll} style={{ fontFamily: 'Montserrat', fontSize: 12, fontWeight: 600, padding: '7px 14px', borderRadius: 6, cursor: 'pointer', background: 'transparent', color: 'var(--t3)', border: '1px solid var(--bd)' }}>
              Reset all ticks
            </button>
          </div>

          {/* ── Role hint banner ── */}
          <div style={{ background: hintBg, border: `1px solid ${hintBd}`, borderRadius: 8, padding: '9px 14px', fontSize: 12, color: hintColor, marginBottom: 12, display: 'flex', alignItems: 'center', gap: 10 }}>
            <span style={{ width: 18, height: 18, borderRadius: '50%', background: hintBd, color: '#fff', display: 'inline-flex', alignItems: 'center', justifyContent: 'center', fontSize: 10, fontWeight: 700, flexShrink: 0 }}>
              {role === 'pm' ? 'PM' : role === 'checker' ? 'C' : 'D'}
            </span>
            <span>{roleHintText}</span>
          </div>

          {/* ── Stats ── */}
          <div style={{ display: 'grid', gridTemplateColumns: 'repeat(5, 1fr)', gap: 10, marginBottom: 16 }}>
            <StatCard label="Total Items" value={String(stats.total)} sub={`across ${CHECKLIST.length} phases`} col="#5a6e88" />
            <StatCard label="Check 1 · Detailer" value={`${stats.c1} / ${stats.total}`} sub={`${stats.pct(stats.c1)}% detailer passed`} col="var(--3eg)" valueCol="#157a15" />
            <StatCard label="Check 2 · Cleared" value={`${stats.cleared} / ${stats.total}`} sub={`${stats.pct(stats.cleared)}% cleared · ${stats.na} N/A`} col="var(--bl)" valueCol="#1d5ec0" />
            <StatCard label="Incorrect · Fix" value={String(stats.incorrect)} sub="flagged for rework" col="var(--rd)" valueCol="#b82020" />
            <StatCard label="PM Overrides" value={String(stats.overrides)} sub="audit logged" col="var(--pu)" valueCol="#4a2f9c" />
          </div>

          {/* ── Phase Nav ── */}
          <div style={{ display: 'flex', gap: 4, marginBottom: 16, borderBottom: '1px solid var(--bd)', overflowX: 'auto' }}>
            {CHECKLIST.map(phase => {
              const pItems = visible.filter(it => it.phaseId === phase.id);
              const done = pItems.filter(it => isResolved(getSt(it.id))).length;
              const pct = pItems.length === 0 ? 0 : Math.round((done / pItems.length) * 100);
              const active = currentPhase === phase.id;
              const pillText = pItems.length === 0 ? 'N/A' : (pct === 100 ? '✓' : `${pct}%`);
              const pillBg = pct === 100 && pItems.length > 0 ? 'rgba(42,158,42,.15)' : pct > 0 ? 'rgba(212,136,10,.15)' : 'var(--s2)';
              const pillColor = pct === 100 && pItems.length > 0 ? '#157a15' : pct > 0 ? '#a06808' : 'var(--t4)';
              return (
                <button key={phase.id} onClick={() => setCurrentPhase(phase.id)} style={{
                  padding: '9px 14px', fontSize: 11, letterSpacing: '.05em', fontFamily: 'Montserrat',
                  fontWeight: active ? 700 : 600,
                  color: active ? 'var(--t1)' : 'var(--t4)',
                  background: 'none',
                  border: 'none',
                  borderBottom: active ? '2px solid var(--3eg)' : '2px solid transparent',
                  whiteSpace: 'nowrap', cursor: 'pointer', textTransform: 'uppercase'
                }}>
                  {phase.id} · {phase.name}
                  <span style={{ background: pillBg, color: pillColor, padding: '1px 6px', borderRadius: 3, marginLeft: 6, fontSize: 10, fontWeight: 700 }}>{pillText}</span>
                </button>
              );
            })}
          </div>

          {/* ── Filter banner ── */}
          {projectType !== 'both' && (() => {
            const hidden = allFlat.length - visible.length;
            if (hidden === 0) return null;
            const otherType = projectType === 'steel' ? 'concrete-only' : 'steel-only';
            return (
              <div style={{ background: 'rgba(212,136,10,.10)', border: '1px solid var(--am)', padding: '9px 14px', borderRadius: 8, fontSize: 12, color: '#a06808', marginBottom: 12, display: 'flex', alignItems: 'center', gap: 10 }}>
                <span style={{ width: 16, height: 16, borderRadius: '50%', background: 'var(--am)', color: '#fff', display: 'inline-flex', alignItems: 'center', justifyContent: 'center', fontSize: 10, fontWeight: 700 }}>i</span>
                <span>{hidden} {otherType} items hidden — project type is <strong>{projectType.toUpperCase()}</strong>.</span>
              </div>
            );
          })()}

          {/* ── Phase content ── */}
          {/* eslint-disable-next-line @typescript-eslint/no-use-before-define */}
          <PhaseContent
            phaseId={currentPhase}
            projectType={projectType}
            items={items}
            role={role}
            onToggleC1={toggleC1}
            onC2Click={handleC2Click}
          />

          {/* ── Override log ── */}
          <div style={{ background: 'rgba(107,79,200,.06)', border: '1px solid rgba(107,79,200,.3)', borderRadius: 8, padding: '12px 14px', marginTop: 16 }}>
            <h3 style={{ fontSize: 11, letterSpacing: '.1em', color: '#4a2f9c', fontWeight: 700, marginBottom: 8, textTransform: 'uppercase' }}>PM Override Log</h3>
            {overrides.length === 0 ? (
              <div style={{ fontSize: 12, color: 'var(--t4)', fontStyle: 'italic' }}>No PM overrides yet on this project.</div>
            ) : (
              [...overrides].reverse().map((ov, i) => (
                <div key={i} style={{ fontSize: 12, color: 'var(--t2)', padding: '6px 0', borderTop: i === 0 ? 'none' : '1px solid rgba(107,79,200,.25)', lineHeight: 1.5 }}>
                  <span style={{ color: 'var(--t1)', fontWeight: 700 }}>{ov.by}</span> cleared Check 2 on <em>&ldquo;{ov.itemText}&rdquo;</em> ({ov.taskCode}) at {ov.at}.
                  <span style={{ color: '#4a2f9c', fontStyle: 'italic', display: 'block', marginTop: 2 }}>Reason: &ldquo;{ov.reason}&rdquo;</span>
                </div>
              ))
            )}
          </div>
        </>
      )}

      {/* ── Override modal ── */}
      {ovModal.id && (
        <div onClick={(e) => { if (e.target === e.currentTarget) setOvModal({ id: null, action: 'cleared', reason: '' }); }}
          style={{ position: 'fixed', inset: 0, background: 'rgba(0,0,0,.5)', zIndex: 100, display: 'flex', alignItems: 'center', justifyContent: 'center', padding: 20 }}>
          <div style={{ background: '#fff', borderRadius: 10, padding: 22, maxWidth: 480, width: '100%', boxShadow: '0 20px 50px rgba(0,0,0,.25)', fontFamily: 'Montserrat' }}>
            <h3 style={{ fontSize: 15, fontWeight: 700, marginBottom: 6, color: 'var(--t1)' }}>PM Override — Check 2 without Check 1</h3>
            <div style={{ fontSize: 12, color: 'var(--t3)', marginBottom: 14 }}>Clearing Check 2 when the Detailer hasn&rsquo;t ticked Check 1 requires a reason. This action is logged to the audit trail.</div>
            {(() => {
              const it = allFlat.filter(x => x.id === ovModal.id)[0];
              return it ? (
                <div style={{ background: 'var(--s2)', padding: '10px 12px', borderRadius: 6, fontSize: 13, marginBottom: 14, borderLeft: '3px solid var(--pu)', color: 'var(--t1)' }}>
                  {it.text} ({it.taskCode})
                </div>
              ) : null;
            })()}
            <label style={{ fontSize: 11, letterSpacing: '.05em', color: 'var(--t4)', fontWeight: 700, display: 'block', marginBottom: 6, textTransform: 'uppercase' }}>Reason (required)</label>
            <textarea
              value={ovModal.reason}
              onChange={e => setOvModal(prev => ({ ...prev, reason: e.target.value }))}
              placeholder="e.g. Detailer on site call — item visually verified by PM, Craig to backfill Check 1 before EOD"
              style={{ width: '100%', padding: 10, border: '1px solid var(--bd)', borderRadius: 6, fontFamily: 'inherit', fontSize: 13, resize: 'vertical', minHeight: 80, outline: 'none', boxSizing: 'border-box' }}
            />
            <div style={{ display: 'flex', gap: 8, marginTop: 16, justifyContent: 'flex-end' }}>
              <button onClick={() => setOvModal({ id: null, action: 'cleared', reason: '' })}
                style={{ padding: '8px 14px', borderRadius: 6, fontSize: 13, background: 'transparent', color: 'var(--t2)', border: '1px solid var(--bd)', cursor: 'pointer', fontFamily: 'Montserrat' }}>
                Cancel
              </button>
              <button onClick={confirmOverride} disabled={!ovModal.reason.trim()}
                style={{ padding: '8px 14px', borderRadius: 6, fontSize: 13, background: ovModal.reason.trim() ? 'var(--pu)' : '#c4b5fd', color: '#fff', border: 'none', cursor: ovModal.reason.trim() ? 'pointer' : 'not-allowed', fontFamily: 'Montserrat', fontWeight: 700 }}>
                Override &amp; mark {ovModal.action === 'na' ? 'N/A' : ovModal.action === 'incorrect' ? 'Incorrect' : 'Cleared'}
              </button>
            </div>
          </div>
        </div>
      )}
    </div>
  );
};

interface PhaseContentProps {
  phaseId: string;
  projectType: ProjectType;
  items: Record<string, IItemState>;
  role: Role;
  onToggleC1: (id: string) => void;
  onC2Click: (id: string, action: C2Action) => void;
}
const PhaseContent: React.FC<PhaseContentProps> = ({ phaseId, projectType, items, role, onToggleC1, onC2Click }) => {
  const phaseIdx = CHECKLIST.map(p => p.id).indexOf(phaseId);
  if (phaseIdx < 0) return null;
  const phase = CHECKLIST[phaseIdx];
  const isSecVisible = (t: SectionType): boolean => projectType === 'both' || t === 'both' || t === projectType;
  const visibleSections = phase.sections.filter(s => isSecVisible(s.type));
  const getSt = (id: string): IItemState => items[id] || { c1: false, c2: null };
  const isResolved = (s: IItemState): boolean => !!s.override || s.c2 === 'cleared' || s.c2 === 'na';

  if (visibleSections.length === 0) {
    return (
      <div>
        <h2 style={{ fontSize: 18, fontWeight: 700, marginBottom: 6, color: 'var(--t1)' }}>Phase {phase.id} · {phase.name}</h2>
        <p style={{ fontSize: 13, color: 'var(--t4)', marginBottom: 14 }}>No applicable items for <strong>{projectType.toUpperCase()}</strong> projects — this phase will unlock if the project type changes.</p>
      </div>
    );
  }

  const allPhaseItems = visibleSections.reduce((acc, s, si) => acc + s.items.length, 0);
  const clearedPhase = visibleSections.reduce((acc, s, si) => {
    return acc + s.items.filter((_, ii) => isResolved(getSt(itemIdOf(phaseIdx, phase.sections.indexOf(s), ii)))).length;
  }, 0);

  return (
    <div>
      <h2 style={{ fontSize: 18, fontWeight: 700, marginBottom: 6, color: 'var(--t1)' }}>Phase {phase.id} · {phase.name}</h2>
      <p style={{ fontSize: 13, color: 'var(--t2)', marginBottom: 14 }}>{clearedPhase} of {allPhaseItems} items cleared by Checker.</p>

      {visibleSections.map((section) => {
        const si = phase.sections.indexOf(section);
        const sectionIds = section.items.map((_, ii) => itemIdOf(phaseIdx, si, ii));
        const sectionCleared = sectionIds.filter(id => isResolved(getSt(id))).length;

        return (
          <div key={si} style={{ background: 'var(--s1)', borderRadius: 8, border: '1px solid var(--bd)', marginBottom: 12, overflow: 'hidden' }}>
            <div style={{ background: 'var(--s2)', padding: '10px 16px', fontSize: 12, fontWeight: 700, color: 'var(--t2)', borderBottom: '1px solid var(--bd)', display: 'flex', alignItems: 'center', gap: 10 }}>
              {section.title}
              <span style={{ fontSize: 9, padding: '2px 7px', borderRadius: 3, letterSpacing: '.05em', fontWeight: 700, textTransform: 'uppercase',
                background: section.type === 'steel' ? 'rgba(46,109,180,.12)' : section.type === 'concrete' ? 'rgba(107,79,200,.12)' : 'rgba(20,150,120,.12)',
                color: section.type === 'steel' ? '#1d5ec0' : section.type === 'concrete' ? '#4a2f9c' : '#0d7a56'
              }}>{section.type}</span>
              <span style={{ marginLeft: 'auto', fontSize: 11, color: 'var(--t2)', fontWeight: 700 }}>{sectionCleared} / {sectionIds.length} cleared</span>
            </div>

            {/* Column header */}
            <div style={{ display: 'grid', gridTemplateColumns: '88px 130px 1fr 80px 120px', background: 'var(--s2)', padding: '9px 16px', fontSize: 10, letterSpacing: '.08em', color: 'var(--t2)', fontWeight: 700, borderTop: '1px solid var(--bd)', textTransform: 'uppercase' }}>
              <div style={{ textAlign: 'center' }}>Check 1<div style={{ fontSize: 9, color: '#157a15', marginTop: 1, fontWeight: 600 }}>DETAILER</div></div>
              <div style={{ textAlign: 'center' }}>Check 2<div style={{ fontSize: 9, color: '#1d5ec0', marginTop: 1, fontWeight: 600 }}>✓ &nbsp; N/A &nbsp; ✗</div></div>
              <div>Checklist Item</div>
              <div>Task</div>
              <div>Status</div>
            </div>

            {section.items.map((item, ii) => {
              const id = itemIdOf(phaseIdx, si, ii);
              const st = getSt(id);
              const rowBg = st.override ? 'rgba(107,79,200,.06)'
                : st.c2 === 'incorrect' ? 'rgba(204,51,51,.06)'
                : st.c2 === 'na' ? 'var(--s2)'
                : st.c2 === 'cleared' ? 'rgba(42,158,42,.06)'
                : st.c1 ? 'rgba(46,109,180,.06)' : 'transparent';
              const rowBorderLeft = st.c2 === 'incorrect' ? '3px solid var(--rd)' : 'none';
              const c2Locked = !st.c1 && !st.override && !st.c2;
              const canAction = role === 'checker' || role === 'pm';

              return (
                <div key={ii} style={{ display: 'grid', gridTemplateColumns: '88px 130px 1fr 80px 120px', padding: '10px 16px', fontSize: 13, borderTop: ii === 0 ? 'none' : '1px solid var(--s2)', alignItems: 'center', background: rowBg, borderLeft: rowBorderLeft }}>
                  {/* Check 1 */}
                  <div style={{ textAlign: 'center' }}>
                    <span onClick={() => onToggleC1(id)} style={{
                      width: 20, height: 20, border: `2px solid ${st.c1 ? 'var(--3eg)' : st.override ? '#c4b5fd' : 'var(--bd)'}`, borderStyle: st.override && !st.c1 ? 'dashed' : 'solid', borderRadius: 4,
                      background: st.c1 ? 'var(--3eg)' : '#fff', display: 'inline-block', position: 'relative', verticalAlign: 'middle',
                      cursor: role === 'detailer' ? 'pointer' : 'not-allowed', opacity: role === 'detailer' ? 1 : 0.55,
                    }}>
                      {st.c1 && <span style={{ position: 'absolute', left: 5, top: 1, width: 5, height: 10, border: 'solid #fff', borderWidth: '0 2px 2px 0', transform: 'rotate(45deg)', display: 'block' }} />}
                      {st.override && !st.c1 && <span style={{ position: 'absolute', top: '50%', left: '50%', transform: 'translate(-50%,-50%)', fontSize: 11, color: '#4a2f9c' }}>−</span>}
                    </span>
                    <div style={{ fontSize: 10, color: st.override ? '#4a2f9c' : 'var(--t2)', marginTop: 3, lineHeight: 1.2, fontWeight: 600 }}>
                      {st.c1 && st.c1By ? `${st.c1By.split(' ')[0]} · ${st.c1At}` : st.override ? 'skipped' : ''}
                    </div>
                  </div>

                  {/* Check 2 */}
                  <div style={{ textAlign: 'center' }}>
                    <div style={{ display: 'inline-flex', gap: 3 }}>
                      {(['cleared', 'na', 'incorrect'] as C2Action[]).map(action => {
                        const glyph = action === 'cleared' ? '✓' : action === 'na' ? 'N/A' : '✗';
                        const isActive = st.override ? action === 'cleared' : st.c2 === action;
                        const activeBg = action === 'cleared' ? 'var(--3eg)' : action === 'na' ? '#5a6e88' : 'var(--rd)';
                        const overrideBg = 'var(--pu)';
                        const disabled = st.override ? role !== 'pm' : !canAction || (c2Locked && role === 'checker');
                        const bgColor = isActive ? (st.override ? overrideBg : activeBg) : '#fff';
                        const bdColor = isActive ? (st.override ? overrideBg : activeBg) : (c2Locked && role === 'checker') ? 'var(--bd)' : 'var(--bd)';
                        const bdStyle = c2Locked && role === 'checker' && !isActive ? 'dashed' : 'solid';
                        return (
                          <button key={action} onClick={() => onC2Click(id, action)} disabled={disabled} title={action === 'cleared' ? 'Cleared — complete and correct' : action === 'na' ? "Doesn't apply to this project" : 'Incorrect — needs to be fixed'}
                            style={{
                              width: action === 'na' ? 30 : 24, height: 24, borderRadius: 5, border: `2px solid ${bdColor}`, borderStyle: bdStyle,
                              background: bgColor, display: 'inline-flex', alignItems: 'center', justifyContent: 'center',
                              fontSize: action === 'na' ? 9 : 11, fontWeight: 700, color: isActive ? '#fff' : 'var(--t4)',
                              padding: 0, cursor: disabled ? 'not-allowed' : 'pointer', opacity: disabled && !isActive ? 0.5 : 1,
                              fontFamily: 'Montserrat',
                            }}>
                            {glyph}
                          </button>
                        );
                      })}
                    </div>
                    <div style={{ fontSize: 10, marginTop: 3, lineHeight: 1.2,
                      color: st.override ? '#4a2f9c'
                        : st.c2 === 'incorrect' ? '#b82020'
                        : st.c2 === 'na' ? 'var(--t2)'
                        : st.c1 ? '#1d5ec0' : 'var(--t2)',
                      fontWeight: st.c2 === 'incorrect' ? 700 : 600
                    }}>
                      {st.override ? `PM override · ${st.overrideAt || ''}`
                        : st.c2 === 'cleared' && st.c2By ? `${st.c2By.split(' ')[0]} · ${st.c2At}`
                        : st.c2 === 'na' && st.c2By ? `N/A · ${st.c2By.split(' ')[0]}`
                        : st.c2 === 'incorrect' && st.c2By ? `Flagged · ${st.c2By.split(' ')[0]}`
                        : st.c1 ? 'awaiting checker'
                        : c2Locked ? 'locked — need Check 1' : ''}
                    </div>
                  </div>

                  {/* Item text */}
                  <div style={{ color: c2Locked ? 'var(--t2)' : 'var(--t1)', lineHeight: 1.4 }}>
                    {item[0]}
                    {st.override && <span style={{ background: 'rgba(107,79,200,.12)', color: '#4a2f9c', padding: '2px 6px', borderRadius: 3, fontSize: 10, fontWeight: 700, marginLeft: 6 }}>PM OVERRIDE</span>}
                    {st.c2 === 'incorrect' && <span style={{ fontSize: 11, marginLeft: 6, color: '#b82020', fontWeight: 600 }}>· needs fixing</span>}
                    {st.c2 === 'na' && <span style={{ fontSize: 11, marginLeft: 6, color: 'var(--t4)' }}>· not applicable</span>}
                    {st.c1 && !st.c2 && !st.override && <span style={{ fontSize: 11, marginLeft: 6, color: '#1d5ec0' }}>· ready for Check 2</span>}
                  </div>

                  {/* Task code */}
                  <div>
                    <span style={{ background: 'rgba(83,74,183,.12)', color: '#3730a3', padding: '2px 8px', borderRadius: 3, fontSize: 11, fontWeight: 700, letterSpacing: '.03em' }}>{item[1]}</span>
                  </div>

                  {/* Status */}
                  <div>
                    <StatusPill st={st} />
                  </div>
                </div>
              );
            })}
          </div>
        );
      })}
    </div>
  );
};

export default ChecklistBoard;
