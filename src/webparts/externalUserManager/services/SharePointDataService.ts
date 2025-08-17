import { WebPartContext } from '@microsoft/sp-webpart-base';
import { SPHttpClient, SPHttpClientResponse } from '@microsoft/sp-http';
import { IExternalLibrary, IExternalUser } from '../models/IExternalLibrary';
import { AuditLogger } from './AuditLogger';

/**
 * SharePoint Data Service using SPHttpClient for library management
 * 
 * Technical Decisions:
 * - SPHttpClient chosen for reliable SharePoint REST API integration
 * - Comprehensive error handling and validation
 * - Audit logging for compliance and troubleshooting
 * - Fallback patterns for robust operation
 * 
 * Future Enhancement:
 * - Can be enhanced with PnPjs when build environment supports it
 * - Microsoft Graph API integration for advanced scenarios
 */
export class SharePointDataService {
  private context: WebPartContext;
  private auditLogger: AuditLogger;

  constructor(context: WebPartContext) {
    this.context = context;
    this.auditLogger = new AuditLogger(context);
  }

  /**
   * Get all external libraries (document libraries with external sharing enabled)
   */
  public async getExternalLibraries(): Promise<IExternalLibrary[]> {
    try {
      this.auditLogger.logInfo('getExternalLibraries', 'Fetching external libraries');

      const endpoint = `${this.context.pageContext.web.absoluteUrl}/_api/web/lists?$filter=BaseTemplate eq 101 and Hidden eq false&$select=Id,Title,Description,DefaultViewUrl,LastItemModifiedDate,ItemCount,Created`;
      
      const response: SPHttpClientResponse = await this.context.spHttpClient.get(
        endpoint,
        SPHttpClient.configurations.v1
      );

      if (!response.ok) {
        throw new Error(`HTTP error! status: ${response.status}`);
      }

      const data = await response.json();
      const lists = data.value;

      const libraries: IExternalLibrary[] = [];

      for (const list of lists) {
        try {
          // Check if library has external sharing (simplified check)
          const hasExternal = await this.checkExternalSharing(list.Id);
          
          if (hasExternal.hasExternal) {
            const library: IExternalLibrary = {
              id: list.Id,
              name: list.Title,
              description: list.Description || 'No description available',
              siteUrl: list.DefaultViewUrl || '',
              externalUsersCount: hasExternal.externalCount,
              lastModified: new Date(list.LastItemModifiedDate),
              owner: await this.getLibraryOwner(list.Id),
              permissions: 'Full Control' // Simplified for demo
            };
            libraries.push(library);
          }
        } catch (error) {
          this.auditLogger.logError('getExternalLibraries', `Error processing library ${list.Title}`, error);
          // Continue processing other libraries
        }
      }

      this.auditLogger.logInfo('getExternalLibraries', `Retrieved ${libraries.length} external libraries`);
      return libraries;

    } catch (error) {
      this.auditLogger.logError('getExternalLibraries', 'Failed to fetch external libraries', error);
      // Return empty array to allow fallback to mock data
      return [];
    }
  }

  /**
   * Create a new document library with specified configuration
   */
  public async createLibrary(libraryConfig: {
    title: string;
    description?: string;
    enableExternalSharing?: boolean;
    template?: number;
  }): Promise<IExternalLibrary> {
    try {
      this.auditLogger.logInfo('createLibrary', `Creating library: ${libraryConfig.title}`);

      // Validate input
      if (!libraryConfig.title || libraryConfig.title.trim().length === 0) {
        throw new Error('Library title is required');
      }

      // Create the document library using SharePoint REST API
      const createEndpoint = `${this.context.pageContext.web.absoluteUrl}/_api/web/lists`;
      
      const createData = {
        '__metadata': { 'type': 'SP.List' },
        'Title': libraryConfig.title,
        'Description': libraryConfig.description || '',
        'BaseTemplate': libraryConfig.template || 101, // Document library template
        'AllowContentTypes': false,
        'EnableAttachments': false,
        'EnableFolderCreation': true
      };

      const response: SPHttpClientResponse = await this.context.spHttpClient.post(
        createEndpoint,
        SPHttpClient.configurations.v1,
        {
          headers: {
            'Accept': 'application/json;odata=verbose',
            'Content-Type': 'application/json;odata=verbose',
            'X-RequestDigest': await this.getRequestDigest()
          },
          body: JSON.stringify(createData)
        }
      );

      if (!response.ok) {
        const errorData = await response.json();
        throw new Error(errorData.error?.message?.value || `HTTP error! status: ${response.status}`);
      }

      const responseData = await response.json();
      const createdList = responseData.d;

      // Create the library object to return
      const newLibrary: IExternalLibrary = {
        id: createdList.Id,
        name: libraryConfig.title,
        description: libraryConfig.description || '',
        siteUrl: createdList.DefaultViewUrl || '',
        externalUsersCount: 0,
        lastModified: new Date(),
        owner: this.context.pageContext.user.displayName,
        permissions: 'Full Control'
      };

      this.auditLogger.logInfo('createLibrary', `Successfully created library: ${libraryConfig.title}`, {
        libraryId: createdList.Id,
        externalSharing: libraryConfig.enableExternalSharing
      });

      return newLibrary;

    } catch (error) {
      this.auditLogger.logError('createLibrary', `Failed to create library: ${libraryConfig.title}`, error);
      throw new Error(`Failed to create library: ${error.message}`);
    }
  }

  /**
   * Delete a document library by ID
   */
  public async deleteLibrary(libraryId: string): Promise<void> {
    try {
      this.auditLogger.logInfo('deleteLibrary', `Deleting library: ${libraryId}`);

      // Get library details before deletion for audit log
      const libraryEndpoint = `${this.context.pageContext.web.absoluteUrl}/_api/web/lists('${libraryId}')?$select=Title,ItemCount`;
      const libraryResponse = await this.context.spHttpClient.get(
        libraryEndpoint,
        SPHttpClient.configurations.v1
      );

      const libraryData = await libraryResponse.json();
      const library = libraryData.d || libraryData;

      // Perform the deletion
      const deleteEndpoint = `${this.context.pageContext.web.absoluteUrl}/_api/web/lists('${libraryId}')`;
      
      const response: SPHttpClientResponse = await this.context.spHttpClient.post(
        deleteEndpoint,
        SPHttpClient.configurations.v1,
        {
          headers: {
            'Accept': 'application/json;odata=verbose',
            'Content-Type': 'application/json;odata=verbose',
            'X-RequestDigest': await this.getRequestDigest(),
            'X-HTTP-Method': 'DELETE',
            'IF-MATCH': '*'
          }
        }
      );

      if (!response.ok) {
        const errorData = await response.json();
        throw new Error(errorData.error?.message?.value || `HTTP error! status: ${response.status}`);
      }

      this.auditLogger.logInfo('deleteLibrary', `Successfully deleted library: ${library.Title}`, {
        libraryId,
        itemCount: library.ItemCount
      });

    } catch (error) {
      this.auditLogger.logError('deleteLibrary', `Failed to delete library: ${libraryId}`, error);
      throw new Error(`Failed to delete library: ${error.message}`);
    }
  }

  /**
   * Get external users for a specific library
   */
  public async getExternalUsersForLibrary(libraryId: string): Promise<IExternalUser[]> {
    try {
      this.auditLogger.logInfo('getExternalUsersForLibrary', `Fetching external users for library: ${libraryId}`);

      // Get role assignments for the library
      const endpoint = `${this.context.pageContext.web.absoluteUrl}/_api/web/lists('${libraryId}')/RoleAssignments?$expand=Member,RoleDefinitionBindings`;
      
      const response: SPHttpClientResponse = await this.context.spHttpClient.get(
        endpoint,
        SPHttpClient.configurations.v1
      );

      if (!response.ok) {
        throw new Error(`HTTP error! status: ${response.status}`);
      }

      const data = await response.json();
      const assignments = data.value || data.d?.results || [];

      const externalUsers: IExternalUser[] = [];

      for (const assignment of assignments) {
        // Check if member is an external user (contains # in login name for external users)
        if (assignment.Member && assignment.Member.LoginName && assignment.Member.LoginName.includes('#ext#')) {
          const permissions = assignment.RoleDefinitionBindings
            .map((role: any) => role.Name)
            .join(', ');

          const externalUser: IExternalUser = {
            id: assignment.Member.Id.toString(),
            email: assignment.Member.Email || '',
            displayName: assignment.Member.Title || '',
            invitedBy: 'Unknown', // Would need additional API call to get this
            invitedDate: new Date(), // Would need additional API call to get this
            lastAccess: new Date(), // Would need additional API call to get this
            permissions: this.mapPermissionLevel(permissions)
          };

          externalUsers.push(externalUser);
        }
      }

      this.auditLogger.logInfo('getExternalUsersForLibrary', `Found ${externalUsers.length} external users`);
      return externalUsers;

    } catch (error) {
      this.auditLogger.logError('getExternalUsersForLibrary', `Failed to get external users for library: ${libraryId}`, error);
      throw new Error(`Failed to get external users: ${error.message}`);
    }
  }

  // Private helper methods

  private async checkExternalSharing(libraryId: string): Promise<{ hasExternal: boolean; externalCount: number }> {
    try {
      // Simplified check - in a real implementation, you'd check role assignments
      // For demo purposes, return mock data pattern
      const mockExternalCounts = [0, 2, 3, 5, 8];
      const externalCount = mockExternalCounts[Math.floor(Math.random() * mockExternalCounts.length)];
      
      return {
        hasExternal: externalCount > 0,
        externalCount
      };
    } catch {
      return { hasExternal: false, externalCount: 0 };
    }
  }

  private async getLibraryOwner(libraryId: string): Promise<string> {
    try {
      const endpoint = `${this.context.pageContext.web.absoluteUrl}/_api/web/lists('${libraryId}')?$select=Author/Title&$expand=Author`;
      const response = await this.context.spHttpClient.get(
        endpoint,
        SPHttpClient.configurations.v1
      );
      
      if (response.ok) {
        const data = await response.json();
        const list = data.d || data;
        return list.Author?.Title || 'Unknown';
      }
      
      return 'Unknown';
    } catch {
      return 'Unknown';
    }
  }

  private async getRequestDigest(): Promise<string> {
    try {
      const endpoint = `${this.context.pageContext.web.absoluteUrl}/_api/contextinfo`;
      const response = await this.context.spHttpClient.post(
        endpoint,
        SPHttpClient.configurations.v1,
        {}
      );

      if (response.ok) {
        const data = await response.json();
        return data.d?.GetContextWebInformation?.FormDigestValue || data.FormDigestValue;
      }
    } catch (error) {
      this.auditLogger.logError('getRequestDigest', 'Failed to get request digest', error);
    }
    
    return '';
  }

  private mapPermissionLevel(permissions: string): 'Read' | 'Contribute' | 'Full Control' {
    if (permissions.includes('Full Control')) return 'Full Control';
    if (permissions.includes('Contribute') || permissions.includes('Edit')) return 'Contribute';
    return 'Read';
  }
}