export interface IRole {
  Id: number;
  Name: string;
  Description: string;
  BasePermissions: IPermission;
}

export interface IPermission {
  High: number;
  Low: number;
}

export interface IPrincipal {
  Id: number;
  Title: string;
  Email: string;
  PrincipalType: PrincipalType;
}

export enum PrincipalType {
  User = 1,
  DistributionList = 2,
  SecurityGroup = 4,
  SharePointGroup = 8
}

export interface IRoleAssignment {
  PrincipalId: number;
  RoleDefId: number;
  Member: IPrincipal;
  RoleDefinition: IRole;
}

export interface IPermissionRequest {
  principalId: number;
  roleDefId: number;
}

export interface IUserPermissions {
  canEdit: boolean;
  canDelete: boolean;
  canApprove: boolean;
  canView: boolean;
  isOwner: boolean;
  isHRAdmin: boolean;
} 