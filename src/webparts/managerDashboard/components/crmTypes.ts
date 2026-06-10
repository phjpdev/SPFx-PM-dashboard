export interface CrmPhone { value: string; type: string; cc?: string; }
export interface CrmEmail { value: string; type: string; }

export type CrmActivityType = 'Phone call' | 'Email' | 'Text' | 'In person';

export interface CrmActivity {
  id: string;
  date: string;
  type: CrmActivityType;
  notes: string;
  followUpDate: string;
  done: boolean;
}

export interface CrmAttachment {
  id: string;
  name: string;
  dataUrl: string;
}

export interface CrmPerson {
  id: string;
  name: string;
  organizationId: string;
  position: string;
  phones: CrmPhone[];
  emails: CrmEmail[];
  activities?: CrmActivity[];
  attachments?: CrmAttachment[];
}

export interface CrmCompany {
  id: string;
  name: string;
  labels: string;
  address: string;
  phones: CrmPhone[];
  emails: CrmEmail[];
}

export type CrmRfqDiscipline = 'Steel' | 'Concrete' | 'Both';
export type CrmRfqStage =
  | 'New Enquiry'
  | 'Under Review'
  | 'Ready to Quote'
  | 'Won'
  | 'Declined';

/** Shared enquiry fields carried RFQ → Quote → Project. */
export interface CrmEnquiryCore {
  dateReceived: string;
  personId: string;
  organizationId: string;
  companyAddress: string;
  projectTitle: string;
  projectAddress: string;
  discipline: CrmRfqDiscipline;
  quoteRequiredBy: string;
  projectValue: number;
  approximateHours: number;
  engineerDrawingReceived: boolean;
  engineerDrawingDate: string;
  revisionVersionEng: string;
  architectDrawingReceived: boolean;
  architectDrawingDate: string;
  revisionVersionArch: string;
  rfiAllowed: number;
  assignedTo: string;
  source: string;
  notes: string;
}

export interface CrmRfq extends CrmEnquiryCore {
  id: string;
  rfqNum: string;
  createQuoteXero: boolean;
  relatedRfqId: string;
  stage: CrmRfqStage;
}

export type CrmQuoteStatus = 'Draft' | 'Sent' | 'Lost';

export type CrmLostReason =
  | 'Price was too high'
  | 'Price was too low'
  | 'Client lost project'
  | 'Inexperienced'
  | 'Project been cancelled'
  | 'Other';

/** Quote created from an RFQ at Ready to Quote stage. */
export interface CrmQuote extends CrmEnquiryCore {
  id: string;
  quoteNum: string;
  rfqId: string;
  rfqNum: string;
  quotedDate: string;
  createQuoteXero: boolean;
  relatedRfqId: string;
  lostReason: string;
  status: CrmQuoteStatus;
}

/** Won quote promoted to an active 3EDGE project. */
export interface CrmProject extends CrmEnquiryCore {
  id: string;
  projNum: string;
  spProjectId?: number;
  quoteId: string;
  quoteNum: string;
  rfqId: string;
  rfqNum: string;
  wonDate: string;
}
