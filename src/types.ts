export type ProviderType = "sharepoint" | "onedrive" | "shared-folder";
export type CaseState = "circulating" | "returned" | "completed";
export type MemberStatus = "pending" | "current" | "approved" | "stamped";

export interface Approver {
  id: string;
  name: string;
  email: string;
  department: string;
  entraUserId?: string;
  active: boolean;
  createdAt: string;
  updatedAt: string;
}

export interface RouteMember {
  id: string;
  approverId: string;
  name: string;
  email: string;
  department: string;
  sequence: number;
  requiresStamp: boolean;
  isInitiator?: boolean;
  status: MemberStatus;
  completedAt?: string;
}

export interface RouteTemplate {
  id: string;
  name: string;
  description: string;
  createdBy: string;
  createdAt: string;
  updatedAt: string;
  members: Omit<RouteMember, "status" | "completedAt">[];
}

export interface HistoryEntry {
  id: string;
  caseId: string;
  actionUserId: string;
  actionMemberId?: string;
  actionUserName: string;
  action: string;
  comment?: string;
  returnToUserId?: string;
  returnToMemberId?: string;
  returnToUserName?: string;
  previousState: string;
  newState: string;
  createdAt: string;
  routeSnapshot?: RouteMember[];
}

export interface CaseDocument {
  id: string;
  fileId: string;
  name: string;
  type: string;
  location: string;
  fileUrl: string;
  requiresStamp: boolean;
}

export interface CirculationCase {
  id: string;
  accessKey: string;
  title?: string;
  provider: ProviderType;
  fileId: string;
  fileName: string;
  fileUrl: string;
  documents: CaseDocument[];
  initiatorId: string;
  initiatorName: string;
  startedAt: string;
  updatedAt: string;
  state: CaseState;
  currentMemberId?: string;
  returnedToStarter?: boolean;
  members: RouteMember[];
  history: HistoryEntry[];
}

export interface AppData {
  approvers: Approver[];
  templates: RouteTemplate[];
  cases: CirculationCase[];
  currentUserId: string;
}

export interface MockFile {
  id: string;
  provider: ProviderType;
  name: string;
  type: string;
  location: string;
  updatedAt: string;
  updatedBy: string;
}
