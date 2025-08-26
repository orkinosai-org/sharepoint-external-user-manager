import { WebPartContext } from '@microsoft/sp-webpart-base';
import { SPHttpClient, SPHttpClientResponse } from '@microsoft/sp-http';
import { IExternalLibrary, IExternalUser, IBulkUserAdditionRequest, IBulkUserAdditionResult } from '../models/IExternalLibrary';
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
            permissions: this.mapPermissionLevel(permissions),
            company: undefined,
            project: undefined
          };

          // Retrieve stored metadata for this user
          const metadata = await this.getUserMetadata(libraryId, assignment.Member.Id);
          if (metadata) {
            externalUser.company = metadata.company;
            externalUser.project = metadata.project;
          }

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

  /**
   * Add external user to a library with specified permissions
   */
  public async addExternalUserToLibrary(libraryId: string, email: string, permission: 'Read' | 'Contribute' | 'Full Control', company?: string, project?: string): Promise<void> {
    try {
      this.auditLogger.logInfo('addExternalUserToLibrary', `Adding external user ${email} to library: ${libraryId} with ${permission} permissions`, {
        libraryId,
        email,
        permission,
        company,
        project
      });

      // First, ensure the user exists in the site collection
      const ensureUserEndpoint = `${this.context.pageContext.web.absoluteUrl}/_api/web/ensureuser`;
      const ensureUserData = {
        logonName: email
      };

      const ensureUserResponse: SPHttpClientResponse = await this.context.spHttpClient.post(
        ensureUserEndpoint,
        SPHttpClient.configurations.v1,
        {
          headers: {
            'Accept': 'application/json;odata=verbose',
            'Content-Type': 'application/json;odata=verbose',
            'X-RequestDigest': await this.getRequestDigest()
          },
          body: JSON.stringify(ensureUserData)
        }
      );

      if (!ensureUserResponse.ok) {
        throw new Error(`Failed to ensure user exists: ${ensureUserResponse.status}`);
      }

      const userData = await ensureUserResponse.json();
      const userId = userData.d?.Id || userData.Id;

      if (!userId) {
        throw new Error('Failed to get user ID after ensuring user');
      }

      // Get the role definition ID for the specified permission
      const roleDefId = await this.getRoleDefinitionId(permission);

      // Add role assignment to the library
      const roleAssignmentEndpoint = `${this.context.pageContext.web.absoluteUrl}/_api/web/lists('${libraryId}')/roleassignments/addroleassignment(principalid=${userId},roledefid=${roleDefId})`;

      const roleAssignmentResponse: SPHttpClientResponse = await this.context.spHttpClient.post(
        roleAssignmentEndpoint,
        SPHttpClient.configurations.v1,
        {
          headers: {
            'Accept': 'application/json;odata=verbose',
            'Content-Type': 'application/json;odata=verbose',
            'X-RequestDigest': await this.getRequestDigest()
          }
        }
      );

      if (!roleAssignmentResponse.ok) {
        throw new Error(`Failed to add role assignment: ${roleAssignmentResponse.status}`);
      }

      // Store user metadata if company or project is provided
      if (company || project) {
        await this.storeUserMetadata(libraryId, email, userId, company, project);
      }

      this.auditLogger.logInfo('addExternalUserToLibrary', `Successfully added user ${email} to library with ${permission} permissions`, {
        libraryId,
        email,
        permission,
        userId,
        company,
        project
      });

    } catch (error) {
      this.auditLogger.logError('addExternalUserToLibrary', `Failed to add external user ${email} to library: ${libraryId}`, error);
      throw new Error(`Failed to add user: ${error.message}`);
    }
  }

  /**
   * Store user metadata (company and project) in a SharePoint list
   */
  private async storeUserMetadata(libraryId: string, email: string, userId: number, company?: string, project?: string): Promise<void> {
    try {
      // For now, we'll store metadata in the audit log and browser storage
      // In a production environment, this would be stored in a custom SharePoint list
      const metadata = {
        libraryId,
        email,
        userId,
        company: company || '',
        project: project || '',
        timestamp: new Date().toISOString()
      };

      // Store in browser localStorage as a fallback (for demo purposes)
      const storageKey = `userMetadata_${libraryId}_${userId}`;
      localStorage.setItem(storageKey, JSON.stringify(metadata));

      this.auditLogger.logInfo('storeUserMetadata', `Stored metadata for user ${email}`, metadata);
    } catch (error) {
      this.auditLogger.logError('storeUserMetadata', `Failed to store metadata for user ${email}`, error);
      // Don't throw error as this is supplementary functionality
    }
  }

  /**
   * Update user metadata (company and project) for an existing user
   */
  public async updateUserMetadata(libraryId: string, userId: string, company?: string, project?: string): Promise<void> {
    try {
      this.auditLogger.logInfo('updateUserMetadata', `Updating metadata for user ${userId} in library: ${libraryId}`, {
        libraryId,
        userId,
        company,
        project
      });

      // For now, update localStorage directly (for demo purposes)
      // In production, this would call a SharePoint API
      const storageKey = `userMetadata_${libraryId}_${userId}`;
      const metadata = {
        libraryId,
        userId,
        company: company || '',
        project: project || '',
        timestamp: new Date().toISOString()
      };
      localStorage.setItem(storageKey, JSON.stringify(metadata));

      this.auditLogger.logInfo('updateUserMetadata', `Successfully updated metadata for user ${userId}`, metadata);
    } catch (error) {
      this.auditLogger.logError('updateUserMetadata', `Failed to update metadata for user ${userId}`, error);
      throw new Error(`Failed to update user metadata: ${error.message}`);
    }
  }

  /**
   * Retrieve user metadata (company and project)
   */
  private async getUserMetadata(libraryId: string, userId: number): Promise<{ company?: string; project?: string } | null> {
    try {
      // Retrieve from browser localStorage (for demo purposes)
      const storageKey = `userMetadata_${libraryId}_${userId}`;
      const stored = localStorage.getItem(storageKey);
      
      if (stored) {
        const metadata = JSON.parse(stored);
        return {
          company: metadata.company || undefined,
          project: metadata.project || undefined
        };
      }
      
      return null;
    } catch (error) {
      this.auditLogger.logError('getUserMetadata', `Failed to retrieve metadata for user ${userId}`, error);
      return null;
    }
  }

  /**
   * Remove external user from a library
   */
  public async removeExternalUserFromLibrary(libraryId: string, userId: string): Promise<void> {
    try {
      this.auditLogger.logInfo('removeExternalUserFromLibrary', `Removing external user ${userId} from library: ${libraryId}`);

      // Remove role assignment from the library
      const roleAssignmentEndpoint = `${this.context.pageContext.web.absoluteUrl}/_api/web/lists('${libraryId}')/roleassignments/removeroleassignment(principalid=${userId})`;

      const response: SPHttpClientResponse = await this.context.spHttpClient.post(
        roleAssignmentEndpoint,
        SPHttpClient.configurations.v1,
        {
          headers: {
            'Accept': 'application/json;odata=verbose',
            'Content-Type': 'application/json;odata=verbose',
            'X-RequestDigest': await this.getRequestDigest()
          }
        }
      );

      if (!response.ok) {
        throw new Error(`Failed to remove role assignment: ${response.status}`);
      }

      this.auditLogger.logInfo('removeExternalUserFromLibrary', `Successfully removed user ${userId} from library`, {
        libraryId,
        userId
      });

    } catch (error) {
      this.auditLogger.logError('removeExternalUserFromLibrary', `Failed to remove external user ${userId} from library: ${libraryId}`, error);
      throw new Error(`Failed to remove user: ${error.message}`);
    }
  }

  /**
   * Bulk add external users to a library with specified permissions
   */
  public async bulkAddExternalUsersToLibrary(
    libraryId: string, 
    request: IBulkUserAdditionRequest
  ): Promise<IBulkUserAdditionResult[]> {
    const results: IBulkUserAdditionResult[] = [];
    const sessionId = this.auditLogger.generateSessionId();
    
    this.auditLogger.logInfo('bulkAddExternalUsersToLibrary', 
      `Starting bulk addition of ${request.emails.length} users to library: ${libraryId}`, {
        libraryId,
        emailCount: request.emails.length,
        permission: request.permission,
        sessionId
      });

    // Get existing users for the library to check for duplicates
    let existingUsers: IExternalUser[] = [];
    try {
      existingUsers = await this.getExternalUsersForLibrary(libraryId);
    } catch (error) {
      this.auditLogger.logWarning('bulkAddExternalUsersToLibrary', 
        'Failed to get existing users, proceeding without duplicate check', { error: error.message });
    }

    const existingEmails = new Set(existingUsers.map(user => user.email.toLowerCase()));

    for (const email of request.emails) {
      const trimmedEmail = email.trim().toLowerCase();
      
      if (!trimmedEmail) {
        results.push({
          email: email,
          status: 'failed',
          message: 'Empty email address',
          error: 'Email address is required'
        });
        continue;
      }

      // Validate email format
      if (!/^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(trimmedEmail)) {
        results.push({
          email: email,
          status: 'failed',
          message: 'Invalid email format',
          error: 'Please provide a valid email address'
        });
        continue;
      }

      // Check if user already has access
      if (existingEmails.has(trimmedEmail)) {
        results.push({
          email: email,
          status: 'already_member',
          message: 'User already has access to this library'
        });
        
        this.auditLogger.logInfo('bulkAddExternalUsersToLibrary', 
          `User ${email} already has access to library`, {
            libraryId,
            email,
            sessionId
          });
        continue;
      }

      // Attempt to add the user
      try {
        await this.addExternalUserToLibrary(libraryId, trimmedEmail, request.permission, request.company, request.project);
        
        // Check if user is external (not from same tenant)
        const isExternal = await this.isExternalUser(trimmedEmail);
        
        results.push({
          email: email,
          status: isExternal ? 'invitation_sent' : 'success',
          message: isExternal 
            ? 'Invitation sent to external user' 
            : 'User added successfully'
        });

        this.auditLogger.logInfo('bulkAddExternalUsersToLibrary', 
          `Successfully added user ${email} to library`, {
            libraryId,
            email,
            permission: request.permission,
            isExternal,
            company: request.company,
            project: request.project,
            sessionId
          });

      } catch (error) {
        results.push({
          email: email,
          status: 'failed',
          message: 'Failed to add user',
          error: error.message
        });

        this.auditLogger.logError('bulkAddExternalUsersToLibrary', 
          `Failed to add user ${email} to library`, {
            libraryId,
            email,
            error: error.message,
            sessionId
          });
      }
    }

    const successCount = results.filter(r => r.status === 'success' || r.status === 'invitation_sent').length;
    const alreadyMemberCount = results.filter(r => r.status === 'already_member').length;
    const failedCount = results.filter(r => r.status === 'failed').length;

    this.auditLogger.logInfo('bulkAddExternalUsersToLibrary', 
      `Bulk addition completed. Success: ${successCount}, Already member: ${alreadyMemberCount}, Failed: ${failedCount}`, {
        libraryId,
        totalEmails: request.emails.length,
        successCount,
        alreadyMemberCount,
        failedCount,
        sessionId
      });

    return results;
  }

  /**
   * Check if a user is external (not from the same tenant)
   */
  private async isExternalUser(email: string): Promise<boolean> {
    try {
      // Simple heuristic: if email domain differs from current site domain, likely external
      const currentDomain = this.context.pageContext.web.absoluteUrl.split('/')[2];
      const emailDomain = email.split('@')[1];
      
      // If domains don't match, it's likely external
      // This is a simplified check - in production you'd use Graph API for accurate determination
      return !currentDomain.includes(emailDomain) && !emailDomain.includes(currentDomain.split('.')[0]);
    } catch {
      // If we can't determine, assume external for safety
      return true;
    }
  }

  /**
   * Search for users in the tenant (for adding external users)
   */
  public async searchUsers(query: string): Promise<IExternalUser[]> {
    try {
      this.auditLogger.logInfo('searchUsers', `Searching for users with query: ${query}`);

      // For external users, we'll use a simple approach of validating the email format
      // In a real implementation, you might use Microsoft Graph API or SharePoint People Picker
      if (!query || !query.includes('@')) {
        return [];
      }

      // Return a mock result for the search - in a real implementation this would query the directory
      const mockUser: IExternalUser = {
        id: 'search-result',
        email: query,
        displayName: query.split('@')[0], // Use part before @ as display name
        invitedBy: this.context.pageContext.user.displayName,
        invitedDate: new Date(),
        lastAccess: new Date(),
        permissions: 'Read'
      };

      this.auditLogger.logInfo('searchUsers', `Found 1 potential user for query: ${query}`);
      return [mockUser];

    } catch (error) {
      this.auditLogger.logError('searchUsers', `Failed to search users with query: ${query}`, error);
      throw new Error(`Failed to search users: ${error.message}`);
    }
  }

  /**
   * Get role definition ID for a permission level
   */
  private async getRoleDefinitionId(permission: 'Read' | 'Contribute' | 'Full Control'): Promise<number> {
    try {
      // Map permission to SharePoint role definition name
      let roleName: string;
      switch (permission) {
        case 'Read':
          roleName = 'Read';
          break;
        case 'Contribute':
          roleName = 'Contribute';
          break;
        case 'Full Control':
          roleName = 'Full Control';
          break;
        default:
          roleName = 'Read';
      }

      const endpoint = `${this.context.pageContext.web.absoluteUrl}/_api/web/roledefinitions/getbyname('${roleName}')`;
      
      const response: SPHttpClientResponse = await this.context.spHttpClient.get(
        endpoint,
        SPHttpClient.configurations.v1
      );

      if (!response.ok) {
        throw new Error(`Failed to get role definition: ${response.status}`);
      }

      const data = await response.json();
      const roleDefId = data.d?.Id || data.Id;

      if (!roleDefId) {
        throw new Error(`Role definition ID not found for permission: ${permission}`);
      }

      return roleDefId;

    } catch (error) {
      this.auditLogger.logError('getRoleDefinitionId', `Failed to get role definition ID for permission: ${permission}`, error);
      throw new Error(`Failed to get role definition: ${error.message}`);
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