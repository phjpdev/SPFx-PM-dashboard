export interface IProject {
  id: string;
  spId?: number;
  projNum: string;
  name: string;
  discipline: string;
  status: string;
  year: number;
  hrsAllowed: number;
  hrsUsed: number;
  rfisAllowed: number;
  quoteNum: string;
  contact: string;
  company: string;
  email: string;
  mobile: string;
  clientNum: string;
  clientp0: string;
  startDate: string;
  finishDate: string;
  ifaDate: string;
  ifcDate: string;
  detailers: string;
  teamLead: string;
  teamMembers: string;
  notes: string;
  invNumber: string;
  invDate: string;
  invSent: boolean;
  invPaid: boolean;
  isEwo: boolean;
  ewoNum: string;
  parentId: string | null;
}

export interface IRfi {
  id: string;
  spId?: number;
  rfiNum: string;
  rfiSeq: number;
  projectId: string;
  projectName: string;
  rfiType: string;
  status: string;
  submittedTo: string;
  toCompany: string;
  by: string;
  byCompany: string;
  cc: string;
  dateIssued: string;
  dateRequired: string;
  description: string;
  attachments?: string;
  clientRfi: string;
  dateReceived: string;
  response: string;
  responseDesc?: string;
  sentBy: string;
  sentByCompany: string;
  impacted: string;
  ewoRef?: string;
  ewoCcn: string;
  tracked: boolean;
  model: number;
  connections: number;
  checking: number;
  drawings: number;
  admin: number;
  revision?: string;
  email?: string;
  emailSentDate?: string;
  total?: number;
}

export const PROJ_STATUSES = ['Active', 'Complete', 'Project Won', 'Waiting PO', 'Sent for RFI', 'Sent for IFA', 'Waiting on Eng/Arch', 'Ready to Inv', 'Archive', 'On Hold', 'Waiting on Dwgs', 'Waiting EWO Approval', 'Cancelled'];
export const RFI_STATUSES = ['Open', 'Closed', 'Partially Open (Revise and Resend)', 'On Hold', 'Overdue'];
export const RFI_TYPES = ['Specifications', 'Drawings', 'Shop Drawings', 'Design', 'General', 'Coordination', 'Material', 'Other'];
export const RFI_RESPONSES = ['Pending', 'Approved', 'Approved with Comments', 'Rejected', 'For Information Only', 'Revise and Resubmit'];

export const STATUS_CFG: Record<string, { bg: string; color: string; bd: string }> = {
  'Active': { bg: 'rgba(61,182,61,0.13)', color: '#1e8a1e', bd: '#3db63d' },
  'Complete': { bg: 'rgba(74,144,217,0.12)', color: '#2065b0', bd: '#4a90d9' },
  'Project Won': { bg: 'rgba(42,158,42,0.15)', color: '#157a15', bd: '#2a9e2a' },
  'Waiting PO': { bg: 'rgba(212,136,10,0.12)', color: '#a06808', bd: '#d4880a' },
  'Sent for RFI': { bg: 'rgba(37,99,235,0.12)', color: '#1d5ec0', bd: '#2563eb' },
  'Sent for IFA': { bg: 'rgba(37,99,235,0.12)', color: '#1d5ec0', bd: '#2563eb' },
  'Waiting on Eng/Arch': { bg: 'rgba(212,136,10,0.12)', color: '#a06808', bd: '#d4880a' },
  'Ready to Inv': { bg: 'rgba(16,185,129,0.12)', color: '#0d7a56', bd: '#10b981' },
  'Archive': { bg: 'rgba(90,110,136,.10)', color: '#4a5e78', bd: '#5a6e88' },
  'On Hold': { bg: 'rgba(155,127,232,0.12)', color: '#5838b8', bd: '#9b7fe8' },
  'Waiting on Dwgs': { bg: 'rgba(212,136,10,0.12)', color: '#a06808', bd: '#d4880a' },
  'Waiting EWO Approval': { bg: 'rgba(249,115,22,0.12)', color: '#b84a10', bd: '#f97316' },
  'Over Budget': { bg: 'rgba(232,69,69,0.13)', color: '#b82020', bd: '#e84545' },
  'Cancelled': { bg: 'rgba(90,110,136,.10)', color: '#4a5e78', bd: '#5a6e88' },
  'Open': { bg: 'rgba(61,182,61,0.13)', color: '#1e8a1e', bd: '#3db63d' },
  'Closed': { bg: 'rgba(74,144,217,0.12)', color: '#2065b0', bd: '#4a90d9' },
  'Overdue': { bg: 'rgba(232,69,69,0.13)', color: '#b82020', bd: '#e84545' },
  'Partially Open (Revise and Resend)': { bg: 'rgba(249,115,22,0.12)', color: '#b84a10', bd: '#f97316' },
};
