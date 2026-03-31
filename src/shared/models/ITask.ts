export interface ITeamMember {
  id: string;
  spId?: number;
  initials: string;
  fullName: string;
  roleType: string;
  email: string;
  totalHrsPerWeek: number;
  productionPct: number;
  prodHrsPerWeek: number;
  startDate: string;
  endDate: string;
  isActive: boolean;
}

export interface ITaskHistory {
  action: string;
  user: string;
  ts: string;
  detail: string;
}

export interface ITask {
  id: string;
  spId?: number;
  project: string;
  taskCode: string;
  description: string;
  assignee: string;
  day: number;
  weekStartDate: string;
  hoursPlanned: number;
  hoursActual: number;
  wipPct: number;
  status: string;
  priority: string;
  completedBy: string;
  completedAt: string;
  completionNote: string;
  reviewedBy: string;
  reviewStatus: string;
  history: ITaskHistory[];
}

export const TASK_STATUSES = ['not_started', 'wip', 'complete', 'blocked', 'rework'];
export const TASK_PRIORITIES = ['high', 'medium', 'low'];
export const DAYS = ['Monday', 'Tuesday', 'Wednesday', 'Thursday', 'Friday'];

export const PROD_TASK_CODES = [
  { group: '01 — Project Management', codes: [
    { id: '01a', label: 'Client Job Setup' }, { id: '01b', label: 'Client Drawing Review' },
    { id: '01c', label: 'General Notes' }, { id: '01d', label: 'Tekla Model Setup' },
    { id: '01e', label: 'Client Meetings' }, { id: '01f', label: 'RFI / TQ / CQ' },
    { id: '01g', label: 'EWO-00 / CCN' }, { id: '01h', label: 'Progress Reports' },
    { id: '01j', label: 'PM / Document Control' },
  ]},
  { group: '02 — Grid', codes: [{ id: '02a', label: 'Grid' }, { id: '02b', label: 'Grid Check' }] },
  { group: '03 — Stick Model', codes: [
    { id: '03a', label: 'Stick Model / Portal Frame Set Up' }, { id: '03b', label: 'Concrete Modelling' },
    { id: '03c', label: 'Stick Model / Portal Frame Check' },
  ]},
  { group: '04 — Connections', codes: [
    { id: '04a', label: 'Macros / Custom Components' }, { id: '04b', label: 'Connection Application' },
    { id: '04c', label: 'Concrete Re-Bar Connection' }, { id: '04d', label: 'Connection Check' },
  ]},
  { group: '05 — Drawings', codes: [
    { id: '05a', label: 'Drawing Setup / Creation' }, { id: '05b', label: 'Assembly Editing' },
    { id: '05c', label: 'Concrete Editing' }, { id: '05d', label: 'GA Editing' },
    { id: '05e', label: 'Part Editing' }, { id: '05f', label: 'Drawing Checking' },
    { id: '05g', label: 'Back Drafting / Rework' },
  ]},
  { group: '06 — Model Review', codes: [{ id: '06a', label: 'Model Clean Up' }, { id: '06b', label: 'High Level Review' }] },
  { group: '07 — Submittals', codes: [
    { id: '07a', label: 'ABM Submittal' }, { id: '07b', label: 'IFA Submittal' },
    { id: '07c', label: 'IFC Submittal' }, { id: '07d', label: 'BFA Review' }, { id: '07e', label: 'BFA Execution' },
  ]},
  { group: '08 — QA Internal Review', codes: [{ id: '08a', label: 'QA Internal Review' }] },
];

export const NON_PROD_TASK_CODES = [
  { group: '00 — General / Internal', codes: [
    { id: '00a', label: 'Estimating' }, { id: '00b', label: 'Training' }, { id: '00c', label: 'Team Meetings' },
    { id: '00d', label: 'Technical Issues' }, { id: '00e', label: 'Tekla Development' },
    { id: '00f', label: '3E Standards Development' }, { id: '00g', label: 'Resources & Staffing' }, { id: '00h', label: 'General' },
  ]},
  { group: '02 — Admin', codes: [
    { id: 'A02a', label: 'General Admin' }, { id: 'A02b', label: 'Marketing' }, { id: 'A02c', label: 'Pipedrive Updating' },
    { id: 'A02d', label: 'Emails' }, { id: 'A02e', label: 'Invoicing' }, { id: 'A02f', label: 'Accounts Receivable' },
    { id: 'A02g', label: 'Accounts Payable' }, { id: 'A02h', label: 'Dashboard' }, { id: 'A02i', label: 'Contractor Payments' },
  ]},
];

export const ALL_CODES_FLAT: { id: string; label: string }[] = ([] as { id: string; label: string }[]).concat(
  ...PROD_TASK_CODES.map(g => g.codes), ...NON_PROD_TASK_CODES.map(g => g.codes)
);
export const getCodeLabel = (id: string): string => { const c = ALL_CODES_FLAT.filter(x => x.id === id)[0]; return c ? c.label : id; };
export const CHECK_CODES = ['03b', '04c', '05e'];
export const isNonProd = (project: string): boolean => project === '3E-INT';
