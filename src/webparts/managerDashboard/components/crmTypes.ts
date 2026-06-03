export interface CrmPhone { value: string; type: string; cc?: string; }
export interface CrmEmail { value: string; type: string; }

export interface CrmPerson {
  id: string;
  name: string;
  organizationId: string;
  position: string;
  phones: CrmPhone[];
  emails: CrmEmail[];
}

export interface CrmCompany {
  id: string;
  name: string;
  labels: string;
  address: string;
  phones: CrmPhone[];
  emails: CrmEmail[];
}
