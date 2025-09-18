export interface IRequest {
  Id?: number;
  Title: string;
  RequestType: string;
  Description: string;
  Department: string;
  Requestor: IUser;
  Manager?: IUser;
  Status: RequestStatus;
  ApprovalOutcome?: ApprovalOutcome;
  ApproverComments?: string;
  Created: string;
  Modified: string;
  Author: IUser;
  Editor: IUser;
  Attachments?: IAttachment[];
}

export interface IUser {
  Id: number;
  Title: string;
  Email: string;
  Department?: string;
  DisplayName: string;
}

export interface IAttachment {
  FileName: string;
  ServerRelativeUrl: string;
  ContentType: string;
  Length: number;
}

export enum RequestStatus {
  Draft = 'Draft',
  Submitted = 'Submitted',
  PendingApproval = 'Pending Approval',
  Approved = 'Approved',
  Rejected = 'Rejected',
  Completed = 'Completed'
}

export enum ApprovalOutcome {
  Approved = 'Approved',
  Rejected = 'Rejected'
}

export interface IRequestFormData {
  title: string;
  requestType: string;
  description: string;
  department: string;
  manager?: IUser;
  attachment?: File;
}

export interface IRequestFilters {
  status?: RequestStatus[];
  department?: string;
  requestType?: string;
  dateFrom?: Date;
  dateTo?: Date;
} 