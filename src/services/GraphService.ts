import { MSGraphClientV3 } from '@microsoft/sp-http';
import { WebPartContext } from '@microsoft/sp-webpart-base';
import { IUser } from '../models/IRequest';

export interface IGraphUser {
  id: string;
  displayName: string;
  mail: string;
  department: string;
  jobTitle: string;
  userPrincipalName: string;
}

export class GraphService {
  private context: WebPartContext;

  constructor(context: WebPartContext) {
    this.context = context;
  }

  // Get current user's department
  public async getCurrentUserDepartment(): Promise<string> {
    try {
      const graphClient: MSGraphClientV3 = await this.context.msGraphClientFactory.getClient('3');
      const user = await graphClient.api('/me').select('department').get();
      return user.department || '';
    } catch (error) {
      console.error('Error getting current user department:', error);
      return '';
    }
  }

  // Get current user's full profile
  public async getCurrentUserProfile(): Promise<IGraphUser> {
    try {
      const graphClient: MSGraphClientV3 = await this.context.msGraphClientFactory.getClient('3');
      const user = await graphClient.api('/me').select('id,displayName,mail,department,jobTitle,userPrincipalName').get();
      
      return {
        id: user.id,
        displayName: user.displayName,
        mail: user.mail,
        department: user.department || '',
        jobTitle: user.jobTitle || '',
        userPrincipalName: user.userPrincipalName
      };
    } catch (error) {
      console.error('Error getting current user profile:', error);
      throw error;
    }
  }

  // Get users by department
  public async getUsersByDepartment(department: string): Promise<IGraphUser[]> {
    try {
      const graphClient: MSGraphClientV3 = await this.context.msGraphClientFactory.getClient('3');
      const filter = department ? `department eq '${department}'` : '';
      
      const response = await graphClient.api('/users')
        .select('id,displayName,mail,department,jobTitle,userPrincipalName')
        .filter(filter)
        .top(999)
        .get();

      return response.value.map((user: any) => ({
        id: user.id,
        displayName: user.displayName,
        mail: user.mail,
        department: user.department || '',
        jobTitle: user.jobTitle || '',
        userPrincipalName: user.userPrincipalName
      }));
    } catch (error) {
      console.error('Error getting users by department:', error);
      return [];
    }
  }

  // Search users by name or email
  public async searchUsers(searchTerm: string, department?: string): Promise<IGraphUser[]> {
    try {
      const graphClient: MSGraphClientV3 = await this.context.msGraphClientFactory.getClient('3');
      
      let filter = `(startswith(displayName,'${searchTerm}') or startswith(mail,'${searchTerm}'))`;
      if (department) {
        filter += ` and department eq '${department}'`;
      }

      const response = await graphClient.api('/users')
        .select('id,displayName,mail,department,jobTitle,userPrincipalName')
        .filter(filter)
        .top(50)
        .get();

      return response.value.map((user: any) => ({
        id: user.id,
        displayName: user.displayName,
        mail: user.mail,
        department: user.department || '',
        jobTitle: user.jobTitle || '',
        userPrincipalName: user.userPrincipalName
      }));
    } catch (error) {
      console.error('Error searching users:', error);
      return [];
    }
  }

  // Get all departments
  public async getAllDepartments(): Promise<string[]> {
    try {
      const graphClient: MSGraphClientV3 = await this.context.msGraphClientFactory.getClient('3');
      
      const response = await graphClient.api('/users')
        .select('department')
        .filter('department ne null')
        .top(999)
        .get();

      const departments = new Set<string>();
      response.value.forEach((user: any) => {
        if (user.department) {
          departments.add(user.department);
        }
      });

      return Array.from(departments).sort();
    } catch (error) {
      console.error('Error getting departments:', error);
      return [];
    }
  }

  // Get user by ID
  public async getUserById(userId: string): Promise<IGraphUser | null> {
    try {
      const graphClient: MSGraphClientV3 = await this.context.msGraphClientFactory.getClient('3');
      const user = await graphClient.api(`/users/${userId}`)
        .select('id,displayName,mail,department,jobTitle,userPrincipalName')
        .get();

      return {
        id: user.id,
        displayName: user.displayName,
        mail: user.mail,
        department: user.department || '',
        jobTitle: user.jobTitle || '',
        userPrincipalName: user.userPrincipalName
      };
    } catch (error) {
      console.error('Error getting user by ID:', error);
      return null;
    }
  }

  // Convert Graph user to IUser format
  public graphUserToIUser(graphUser: IGraphUser): IUser {
    return {
      Id: parseInt(graphUser.id) || 0,
      Title: graphUser.displayName,
      Email: graphUser.mail,
      Department: graphUser.department,
      DisplayName: graphUser.displayName
    };
  }
} 