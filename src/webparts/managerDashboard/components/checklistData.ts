// ───────────────────────────── Types ─────────────────────────────
export type SectionType = 'steel' | 'concrete' | 'both';
export type ProjectType = 'steel' | 'concrete' | 'both';
export type Role = 'detailer' | 'checker' | 'pm';
export type C2Action = 'cleared' | 'na' | 'incorrect';

export interface ISection {
  title: string;
  type: SectionType;
  items: Array<[string, string]>; // [text, taskCode]
}
export interface IPhase {
  id: string;
  name: string;
  sections: ISection[];
}
export interface IItemState {
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
export interface IOverrideLog {
  itemId: string;
  by: string;
  at: string;
  reason: string;
  itemText: string;
  taskCode: string;
  action: C2Action;
}

// ───────────────────────────── Data ─────────────────────────────
export const CHECKLIST: IPhase[] = [
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
export const itemIdOf = (pi: number, si: number, ii: number): string => `p${pi}s${si}i${ii}`;
export const nowString = (): string => {
  const d = new Date();
  return d.toLocaleString('en-AU', { day: '2-digit', month: 'short', hour: '2-digit', minute: '2-digit', hour12: true });
};
export const disciplineToType = (d: string): ProjectType => {
  const t = (d || '').toLowerCase();
  if (t.indexOf('steel') >= 0 && t.indexOf('concrete') >= 0) return 'both';
  if (t.indexOf('concrete') >= 0) return 'concrete';
  return 'steel';
};
