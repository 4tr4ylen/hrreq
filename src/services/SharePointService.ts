import { SPHttpClient, SPHttpClientResponse, ISPHttpClientOptions } from '@microsoft/sp-http';
import { WebPartContext } from '@microsoft/sp-webpart-base';
import { IRequest, IUser, RequestStatus, IRequestFilters } from '../models/IRequest';
import { IUserPermissions } from '../models/IRole';
import { IPermissionRequest } from '../models/IRole';

export class SharePointService {
  private context: WebPartContext;
  private listTitle: string = 'HR Requests';
  private siteUrl: string;

  constructor(context: WebPartContext) {
    this.context = context;
    this.siteUrl = context.pageContext.web.absoluteUrl;
  }

  // Get current user information
  public async getCurrentUser(): Promise<IUser> {
    const response: SPHttpClientResponse = await this.context.spHttpClient.get(
      `${this.siteUrl}/_api/web/currentuser`,
      SPHttpClient.configurations.v1
    );

    if (!response.ok) {
      throw new Error(`Failed to get current user: ${response.statusText}`);
    }

    const userData = await response.json();
    return {
      Id: userData.Id,
      Title: userData.Title,
      Email: userData.Email,
      DisplayName: userData.Title
    };
  }

  // Check if current user is HR admin
  public async isHRAdmin(): Promise<boolean> {
    try {
      const response: SPHttpClientResponse = await this.context.spHttpClient.get(
        `${this.siteUrl}/_api/web/sitegroups/getbyname('HR Admins')/users?$filter=Email eq '${this.context.pageContext.user.email}'`,
        SPHttpClient.configurations.v1
      );

      if (!response.ok) {
        return false;
      }

      const users = await response.json();
      return users.value && users.value.length > 0;
    } catch (error) {
      console.error('Error checking HR admin status:', error);
      return false;
    }
  }

  // Create a new request
  public async createRequest(requestData: Partial<IRequest>): Promise<IRequest> {
    const itemData = {
      Title: requestData.Title,
      RequestType: requestData.RequestType,
      Description: requestData.Description,
      Department: requestData.Department,
      Status: RequestStatus.Submitted,
      RequestorId: requestData.Requestor?.Id,
      ManagerId: requestData.Manager?.Id
    };

    const options: ISPHttpClientOptions = {
      headers: {
        'Accept': 'application/json;odata=nometadata',
        'Content-type': 'application/json;odata=nometadata',
        'odata-version': ''
      },
      body: JSON.stringify(itemData)
    };

    const response: SPHttpClientResponse = await this.context.spHttpClient.post(
      `${this.siteUrl}/_api/web/lists/getbytitle('${this.listTitle}')/items`,
      SPHttpClient.configurations.v1,
      options
    );

    if (!response.ok) {
      throw new Error(`Failed to create request: ${response.statusText}`);
    }

    return await response.json();
  }

  // Upload attachment to a request
  public async uploadAttachment(itemId: number, fileName: string, fileContent: ArrayBuffer): Promise<void> {
    const options: ISPHttpClientOptions = {
      headers: {
        'Accept': 'application/json;odata=nometadata',
        'Content-type': 'application/octet-stream',
        'odata-version': ''
      },
      body: fileContent
    };

    const response: SPHttpClientResponse = await this.context.spHttpClient.post(
      `${this.siteUrl}/_api/web/lists/getbytitle('${this.listTitle}')/items(${itemId})/AttachmentFiles/add(FileName='${fileName}')`,
      SPHttpClient.configurations.v1,
      options
    );

    if (!response.ok) {
      throw new Error(`Failed to upload attachment: ${response.statusText}`);
    }
  }

  // Get requests with filtering
  public async getRequests(filters?: IRequestFilters, top?: number): Promise<IRequest[]> {
    let filterQuery = '';
    
    if (filters) {
      const filterConditions: string[] = [];
      
      if (filters.status && filters.status.length > 0) {
        const statusFilter = filters.status.map(s => `Status eq '${s}'`).join(' or ');
        filterConditions.push(`(${statusFilter})`);
      }
      
      if (filters.department) {
        filterConditions.push(`Department eq '${filters.department}'`);
      }
      
      if (filters.requestType) {
        filterConditions.push(`RequestType eq '${filters.requestType}'`);
      }
      
      if (filters.dateFrom) {
        filterConditions.push(`Created ge datetime'${filters.dateFrom.toISOString()}'`);
      }
      
      if (filters.dateTo) {
        filterConditions.push(`Created le datetime'${filters.dateTo.toISOString()}'`);
      }
      
      if (filterConditions.length > 0) {
        filterQuery = `&$filter=${filterConditions.join(' and ')}`;
      }
    }

    const selectQuery = '&$select=Id,Title,RequestType,Description,Department,Status,ApprovalOutcome,ApproverComments,Created,Modified,Author,Editor,RequestorId,ManagerId';
    const expandQuery = '&$expand=Author,Editor,Attachments';
    const orderQuery = '&$orderby=Created desc';
    const topQuery = top ? `&$top=${top}` : '';

    const response: SPHttpClientResponse = await this.context.spHttpClient.get(
      `${this.siteUrl}/_api/web/lists/getbytitle('${this.listTitle}')/items?${selectQuery}${expandQuery}${orderQuery}${topQuery}${filterQuery}`,
      SPHttpClient.configurations.v1
    );

    if (!response.ok) {
      throw new Error(`Failed to get requests: ${response.statusText}`);
    }

    const data = await response.json();
    return this.mapRequests(data.value);
  }

  // Get requests for current user's department
  public async getDepartmentRequests(): Promise<IRequest[]> {
    const currentUser = await this.getCurrentUser();
    const isHRAdmin = await this.isHRAdmin();
    
    if (isHRAdmin) {
      return this.getRequests();
    }

    // For regular users, filter by their department
    return this.getRequests({ department: currentUser.Department });
  }

  // Get a single request by ID
  public async getRequestById(id: number): Promise<IRequest> {
    const response: SPHttpClientResponse = await this.context.spHttpClient.get(
      `${this.siteUrl}/_api/web/lists/getbytitle('${this.listTitle}')/items(${id})?$select=Id,Title,RequestType,Description,Department,Status,ApprovalOutcome,ApproverComments,Created,Modified,Author,Editor,RequestorId,ManagerId&$expand=Author,Editor,Attachments`,
      SPHttpClient.configurations.v1
    );

    if (!response.ok) {
      throw new Error(`Failed to get request: ${response.statusText}`);
    }

    const request = await response.json();
    return this.mapRequest(request);
  }

  // Update a request
  public async updateRequest(id: number, updates: Partial<IRequest>): Promise<void> {
    const options: ISPHttpClientOptions = {
      headers: {
        'Accept': 'application/json;odata=nometadata',
        'Content-type': 'application/json;odata=nometadata',
        'odata-version': '',
        'X-HTTP-Method': 'MERGE',
        'If-Match': '*'
      },
      body: JSON.stringify(updates)
    };

    const response: SPHttpClientResponse = await this.context.spHttpClient.post(
      `${this.siteUrl}/_api/web/lists/getbytitle('${this.listTitle}')/items(${id})`,
      SPHttpClient.configurations.v1,
      options
    );

    if (!response.ok) {
      throw new Error(`Failed to update request: ${response.statusText}`);
    }
  }

  // Check user permissions for a request
  public async getUserPermissions(request: IRequest): Promise<IUserPermissions> {
    const currentUser = await this.getCurrentUser();
    const isHRAdmin = await this.isHRAdmin();
    const isOwner = request.Author.Email === currentUser.Email;
    const terminalStatuses: RequestStatus[] = [RequestStatus.Approved, RequestStatus.Rejected, RequestStatus.Completed];
    const isEditableStatus = terminalStatuses.indexOf(request.Status) === -1;

    return {
      canEdit: isHRAdmin || (isOwner && isEditableStatus),
      canDelete: isHRAdmin,
      canApprove: isHRAdmin,
      canView: isHRAdmin || isOwner || request.Department === currentUser.Department,
      isOwner,
      isHRAdmin
    };
  }

  // Break inheritance and set permissions
  public async setItemPermissions(itemId: number, permissions: IPermissionRequest[]): Promise<void> {
    // First break inheritance
    const breakInheritanceOptions: ISPHttpClientOptions = {
      headers: {
        'Accept': 'application/json;odata=nometadata',
        'Content-type': 'application/json;odata=nometadata',
        'odata-version': ''
      }
    };

    await this.context.spHttpClient.post(
      `${this.siteUrl}/_api/web/lists/getbytitle('${this.listTitle}')/items(${itemId})/breakroleinheritance(copyRoleAssignments=false,clearSubscopes=true)`,
      SPHttpClient.configurations.v1,
      breakInheritanceOptions
    );

    // Then add role assignments
    for (const permission of permissions) {
      const addRoleOptions: ISPHttpClientOptions = {
        headers: {
          'Accept': 'application/json;odata=nometadata',
          'Content-type': 'application/json;odata=nometadata',
          'odata-version': ''
        }
      };

      await this.context.spHttpClient.post(
        `${this.siteUrl}/_api/web/lists/getbytitle('${this.listTitle}')/items(${itemId})/roleassignments/addroleassignment(principalid=${permission.principalId},roledefid=${permission.roleDefId})`,
        SPHttpClient.configurations.v1,
        addRoleOptions
      );
    }
  }

  // Helper method to map SharePoint response to IRequest
  private mapRequest(item: any): IRequest {
    return {
      Id: item.Id,
      Title: item.Title,
      RequestType: item.RequestType,
      Description: item.Description,
      Department: item.Department,
      Requestor: this.mapUser(item.Author),
      Status: item.Status,
      ApprovalOutcome: item.ApprovalOutcome,
      ApproverComments: item.ApproverComments,
      Created: item.Created,
      Modified: item.Modified,
      Author: this.mapUser(item.Author),
      Editor: this.mapUser(item.Editor),
      Attachments: item.Attachments ? item.Attachments.map((a: any) => this.mapAttachment(a)) : []
    };
  }

  // Helper method to map multiple requests
  private mapRequests(items: any[]): IRequest[] {
    return items.map(item => this.mapRequest(item));
  }

  // Helper method to map user data
  private mapUser(userData: any): IUser {
    return {
      Id: userData.Id,
      Title: userData.Title,
      Email: userData.EMail || userData.Email,
      DisplayName: userData.Title
    };
  }

  // Helper method to map attachment data
  private mapAttachment(attachmentData: any): any {
    return {
      FileName: attachmentData.FileName,
      ServerRelativeUrl: attachmentData.ServerRelativeUrl,
      ContentType: attachmentData.ContentType,
      Length: attachmentData.Length
    };
  }
} 