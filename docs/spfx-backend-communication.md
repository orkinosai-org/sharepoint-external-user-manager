# SPFx Frontend to Backend Communication

## Overview

This document outlines how the SharePoint Framework (SPFx) frontend communicates with the backend API service, including authentication, data flow, error handling, and integration patterns.

## Communication Architecture

```
┌─────────────────────────────────────────────────────────────────────────────┐
│                           SPFx Frontend Architecture                        │
├─────────────────────────────────────────────────────────────────────────────┤
│                                                                             │
│  ┌─────────────────────┐    ┌─────────────────────┐    ┌─────────────────┐ │
│  │   React Components  │    │   Service Layer     │    │   Data Models   │ │
│  │                     │    │                     │    │                 │ │
│  │  ├─ ExternalUser    │    │  ├─ BackendApi      │    │  ├─ ILibrary    │ │
│  │  │  Manager         │◄──►│  │  Service         │◄──►│  ├─ IUser       │ │
│  │  ├─ LibraryList     │    │  ├─ Authentication  │    │  ├─ IPermission │ │
│  │  ├─ UserList        │    │  │  Service         │    │  └─ IResponse   │ │
│  │  └─ Permissions     │    │  └─ Cache Service   │    │                 │ │
│  │                     │    │                     │    │                 │ │
│  └─────────────────────┘    └─────────────────────┘    └─────────────────┘ │
│                                        │                                   │
└────────────────────────────────────────┼───────────────────────────────────┘
                                         │
                                         │ HTTPS/REST API
                                         │
┌────────────────────────────────────────▼───────────────────────────────────┐
│                           Backend API Architecture                          │
├─────────────────────────────────────────────────────────────────────────────┤
│                                                                             │
│  ┌─────────────────────┐    ┌─────────────────────┐    ┌─────────────────┐ │
│  │   Azure API         │    │   Azure Functions   │    │   Data Layer    │ │
│  │   Management        │    │                     │    │                 │ │
│  │                     │    │  ├─ Library Mgmt    │    │  ├─ SQL Database│ │
│  │  ├─ Authentication  │◄──►│  ├─ User Mgmt       │◄──►│  ├─ Cosmos DB   │ │
│  │  ├─ Rate Limiting   │    │  ├─ Tenant Mgmt     │    │  ├─ Graph API   │ │
│  │  ├─ Request/Response│    │  └─ Audit Logging   │    │  └─ SharePoint  │ │
│  │  └─ Monitoring      │    │                     │    │                 │ │
│  │                     │    │                     │    │                 │ │
│  └─────────────────────┘    └─────────────────────┘    └─────────────────┘ │
└─────────────────────────────────────────────────────────────────────────────┘
```

## Backend API Service Implementation

### 1. Service Interface

```typescript
// services/IBackendApiService.ts
export interface IBackendApiService {
  // Library Management
  getLibraries(page?: number, pageSize?: number): Promise<IApiResponse<IExternalLibrary[]>>;
  createLibrary(library: ICreateLibraryRequest): Promise<IApiResponse<IExternalLibrary>>;
  deleteLibrary(libraryId: string): Promise<IApiResponse<void>>;
  updateLibrary(libraryId: string, updates: IUpdateLibraryRequest): Promise<IApiResponse<IExternalLibrary>>;

  // User Management  
  getExternalUsers(libraryId: string, page?: number, pageSize?: number): Promise<IApiResponse<IExternalUser[]>>;
  inviteUser(libraryId: string, invitation: IUserInvitationRequest): Promise<IApiResponse<IExternalUser>>;
  updateUserPermissions(libraryId: string, userId: string, permissions: IPermissionUpdate): Promise<IApiResponse<IExternalUser>>;
  revokeUserAccess(libraryId: string, userId: string): Promise<IApiResponse<void>>;

  // Tenant Management
  getTenantSettings(): Promise<IApiResponse<ITenantSettings>>;
  updateTenantSettings(settings: ITenantSettingsUpdate): Promise<IApiResponse<ITenantSettings>>;
}
```

### 2. Service Implementation

```typescript
// services/BackendApiService.ts
import { MSGraphClientV3 } from '@microsoft/sp-http';
import { WebPartContext } from '@microsoft/sp-webpart-base';

export class BackendApiService implements IBackendApiService {
  private context: WebPartContext;
  private baseUrl: string;
  private authService: AuthenticationService;

  constructor(context: WebPartContext, baseUrl: string) {
    this.context = context;
    this.baseUrl = baseUrl;
    this.authService = new AuthenticationService(context);
  }

  private async makeRequest<T>(
    endpoint: string, 
    method: 'GET' | 'POST' | 'PUT' | 'DELETE' = 'GET',
    body?: any
  ): Promise<IApiResponse<T>> {
    try {
      // Get authentication token
      const token = await this.authService.getAccessToken();
      const tenantId = await this.authService.getTenantId();

      // Prepare request
      const url = `${this.baseUrl}${endpoint}`;
      const headers: HeadersInit = {
        'Authorization': `Bearer ${token}`,
        'X-Tenant-ID': tenantId,
        'Content-Type': 'application/json',
        'Accept': 'application/json'
      };

      const requestInit: RequestInit = {
        method,
        headers,
        body: body ? JSON.stringify(body) : undefined
      };

      // Make request
      const response = await fetch(url, requestInit);
      
      // Handle response
      if (!response.ok) {
        const errorData = await response.json().catch(() => ({}));
        throw new ApiError(
          response.status,
          errorData.error?.code || 'UNKNOWN_ERROR',
          errorData.error?.message || 'An unknown error occurred'
        );
      }

      const data = await response.json();
      return data as IApiResponse<T>;

    } catch (error) {
      if (error instanceof ApiError) {
        throw error;
      }
      
      // Handle network errors
      throw new ApiError(
        0,
        'NETWORK_ERROR',
        'Failed to communicate with backend service'
      );
    }
  }

  // Library Management Implementation
  public async getLibraries(page = 1, pageSize = 50): Promise<IApiResponse<IExternalLibrary[]>> {
    const endpoint = `/libraries?page=${page}&pageSize=${pageSize}`;
    return this.makeRequest<IExternalLibrary[]>(endpoint);
  }

  public async createLibrary(library: ICreateLibraryRequest): Promise<IApiResponse<IExternalLibrary>> {
    return this.makeRequest<IExternalLibrary>('/libraries', 'POST', library);
  }

  public async deleteLibrary(libraryId: string): Promise<IApiResponse<void>> {
    return this.makeRequest<void>(`/libraries/${libraryId}`, 'DELETE');
  }

  // User Management Implementation
  public async getExternalUsers(
    libraryId: string, 
    page = 1, 
    pageSize = 50
  ): Promise<IApiResponse<IExternalUser[]>> {
    const endpoint = `/libraries/${libraryId}/users?page=${page}&pageSize=${pageSize}`;
    return this.makeRequest<IExternalUser[]>(endpoint);
  }

  public async inviteUser(
    libraryId: string, 
    invitation: IUserInvitationRequest
  ): Promise<IApiResponse<IExternalUser>> {
    return this.makeRequest<IExternalUser>(
      `/libraries/${libraryId}/users`, 
      'POST', 
      invitation
    );
  }

  public async updateUserPermissions(
    libraryId: string,
    userId: string,
    permissions: IPermissionUpdate
  ): Promise<IApiResponse<IExternalUser>> {
    return this.makeRequest<IExternalUser>(
      `/libraries/${libraryId}/users/${userId}`,
      'PUT',
      permissions
    );
  }

  public async revokeUserAccess(
    libraryId: string,
    userId: string
  ): Promise<IApiResponse<void>> {
    return this.makeRequest<void>(`/libraries/${libraryId}/users/${userId}`, 'DELETE');
  }

  // Tenant Management Implementation
  public async getTenantSettings(): Promise<IApiResponse<ITenantSettings>> {
    return this.makeRequest<ITenantSettings>('/tenant/settings');
  }

  public async updateTenantSettings(
    settings: ITenantSettingsUpdate
  ): Promise<IApiResponse<ITenantSettings>> {
    return this.makeRequest<ITenantSettings>('/tenant/settings', 'PUT', settings);
  }
}
```

### 3. Authentication Service

```typescript
// services/AuthenticationService.ts
export class AuthenticationService {
  private context: WebPartContext;
  private tokenCache: Map<string, { token: string; expiry: number }> = new Map();

  constructor(context: WebPartContext) {
    this.context = context;
  }

  public async getAccessToken(scopes: string[] = ['https://api.spexternal.com/.default']): Promise<string> {
    const cacheKey = scopes.join(',');
    const cached = this.tokenCache.get(cacheKey);
    
    // Check if cached token is still valid (with 5-minute buffer)
    if (cached && cached.expiry > Date.now() + 300000) {
      return cached.token;
    }

    try {
      // Get token using MSGraph client factory
      const graphClient = await this.context.msGraphClientFactory.getClient('3');
      
      // For custom API, we need to use the AadHttpClient
      const aadClient = await this.context.aadHttpClientFactory.getClient('https://api.spexternal.com');
      
      // Extract token from client (implementation depends on Microsoft's client)
      const token = await this.extractTokenFromClient(aadClient);
      
      // Cache token (assuming 1-hour expiry)
      this.tokenCache.set(cacheKey, {
        token,
        expiry: Date.now() + 3600000 // 1 hour
      });

      return token;
    } catch (error) {
      throw new Error(`Failed to acquire access token: ${error.message}`);
    }
  }

  public async getTenantId(): Promise<string> {
    return this.context.pageContext.aadInfo.tenantId._guid;
  }

  private async extractTokenFromClient(client: any): Promise<string> {
    // This is a simplified example - actual implementation would depend
    // on the specific client interface provided by Microsoft
    // In practice, you might need to make a dummy request to extract the token
    // or use undocumented client properties
    
    // Alternative: Use MSAL directly
    return await this.getMsalToken();
  }

  private async getMsalToken(): Promise<string> {
    // Direct MSAL usage if needed
    // This would require configuring MSAL in the web part
    throw new Error('MSAL direct usage not implemented - use AadHttpClient');
  }
}
```

### 4. Data Models

```typescript
// models/IApiModels.ts
export interface IApiResponse<T> {
  success: boolean;
  data?: T;
  error?: {
    code: string;
    message: string;
    details?: string;
  };
  pagination?: {
    page: number;
    pageSize: number;
    total: number;
    hasNext: boolean;
  };
}

export interface ICreateLibraryRequest {
  name: string;
  description?: string;
  siteUrl: string;
  enableExternalSharing?: boolean;
  template?: number;
}

export interface IUpdateLibraryRequest {
  name?: string;
  description?: string;
  enableExternalSharing?: boolean;
}

export interface IUserInvitationRequest {
  email: string;
  displayName: string;
  permissions: 'Read' | 'Contribute' | 'Full Control';
  message?: string;
}

export interface IPermissionUpdate {
  permissions: 'Read' | 'Contribute' | 'Full Control';
}

export interface ITenantSettings {
  externalSharingEnabled: boolean;
  allowAnonymousLinks: boolean;
  defaultLinkPermission: 'View' | 'Edit';
  externalUserExpirationDays: number;
  requireApprovalForExternalSharing: boolean;
}

export interface ITenantSettingsUpdate {
  externalSharingEnabled?: boolean;
  allowAnonymousLinks?: boolean;
  defaultLinkPermission?: 'View' | 'Edit';
  externalUserExpirationDays?: number;
  requireApprovalForExternalSharing?: boolean;
}
```

### 5. Error Handling

```typescript
// models/ApiError.ts
export class ApiError extends Error {
  public readonly status: number;
  public readonly code: string;
  public readonly details?: string;

  constructor(status: number, code: string, message: string, details?: string) {
    super(message);
    this.status = status;
    this.code = code;
    this.details = details;
    this.name = 'ApiError';
  }

  public isNetworkError(): boolean {
    return this.status === 0;
  }

  public isAuthenticationError(): boolean {
    return this.status === 401;
  }

  public isAuthorizationError(): boolean {
    return this.status === 403;
  }

  public isNotFoundError(): boolean {
    return this.status === 404;
  }

  public isValidationError(): boolean {
    return this.status === 400;
  }

  public isRateLimitError(): boolean {
    return this.status === 429;
  }

  public isServerError(): boolean {
    return this.status >= 500;
  }
}

// Error handling utility
export class ErrorHandler {
  public static handleApiError(error: ApiError): string {
    switch (error.code) {
      case 'UNAUTHORIZED':
        return 'Please sign in to access this feature.';
      case 'FORBIDDEN':
        return 'You do not have permission to perform this action.';
      case 'LIBRARY_NOT_FOUND':
        return 'The requested library could not be found.';
      case 'USER_NOT_FOUND':
        return 'The requested user could not be found.';
      case 'VALIDATION_ERROR':
        return `Invalid input: ${error.details || error.message}`;
      case 'EXTERNAL_SHARING_DISABLED':
        return 'External sharing is disabled for your organization.';
      case 'RATE_LIMIT_EXCEEDED':
        return 'Too many requests. Please wait a moment and try again.';
      case 'NETWORK_ERROR':
        return 'Unable to connect to the service. Please check your internet connection.';
      default:
        return error.message || 'An unexpected error occurred.';
    }
  }
}
```

## Integration with React Components

### 1. Service Provider Pattern

```typescript
// components/ExternalUserManager.tsx
import React, { createContext, useContext, useMemo } from 'react';

interface IServiceContext {
  backendApi: IBackendApiService;
  cacheService: ICacheService;
}

const ServiceContext = createContext<IServiceContext | undefined>(undefined);

export const useServices = (): IServiceContext => {
  const context = useContext(ServiceContext);
  if (!context) {
    throw new Error('useServices must be used within a ServiceProvider');
  }
  return context;
};

export const ServiceProvider: React.FC<{ context: WebPartContext; children: React.ReactNode }> = ({
  context,
  children
}) => {
  const services = useMemo(() => ({
    backendApi: new BackendApiService(context, 'https://api.spexternal.com/v1'),
    cacheService: new CacheService()
  }), [context]);

  return (
    <ServiceContext.Provider value={services}>
      {children}
    </ServiceContext.Provider>
  );
};
```

### 2. Custom Hooks for API Operations

```typescript
// hooks/useLibraries.ts
import { useState, useEffect, useCallback } from 'react';
import { useServices } from '../components/ExternalUserManager';

export const useLibraries = () => {
  const { backendApi } = useServices();
  const [libraries, setLibraries] = useState<IExternalLibrary[]>([]);
  const [loading, setLoading] = useState(false);
  const [error, setError] = useState<string | null>(null);

  const loadLibraries = useCallback(async () => {
    try {
      setLoading(true);
      setError(null);
      
      const response = await backendApi.getLibraries();
      if (response.success && response.data) {
        setLibraries(response.data);
      } else {
        throw new Error(response.error?.message || 'Failed to load libraries');
      }
    } catch (err) {
      const errorMessage = err instanceof ApiError 
        ? ErrorHandler.handleApiError(err)
        : 'Failed to load libraries';
      setError(errorMessage);
    } finally {
      setLoading(false);
    }
  }, [backendApi]);

  const createLibrary = useCallback(async (library: ICreateLibraryRequest) => {
    try {
      setLoading(true);
      const response = await backendApi.createLibrary(library);
      if (response.success && response.data) {
        setLibraries(prev => [...prev, response.data!]);
        return response.data;
      } else {
        throw new Error(response.error?.message || 'Failed to create library');
      }
    } catch (err) {
      const errorMessage = err instanceof ApiError 
        ? ErrorHandler.handleApiError(err)
        : 'Failed to create library';
      setError(errorMessage);
      throw new Error(errorMessage);
    } finally {
      setLoading(false);
    }
  }, [backendApi]);

  const deleteLibrary = useCallback(async (libraryId: string) => {
    try {
      setLoading(true);
      const response = await backendApi.deleteLibrary(libraryId);
      if (response.success) {
        setLibraries(prev => prev.filter(lib => lib.id !== libraryId));
      } else {
        throw new Error(response.error?.message || 'Failed to delete library');
      }
    } catch (err) {
      const errorMessage = err instanceof ApiError 
        ? ErrorHandler.handleApiError(err)
        : 'Failed to delete library';
      setError(errorMessage);
      throw new Error(errorMessage);
    } finally {
      setLoading(false);
    }
  }, [backendApi]);

  useEffect(() => {
    loadLibraries();
  }, [loadLibraries]);

  return {
    libraries,
    loading,
    error,
    loadLibraries,
    createLibrary,
    deleteLibrary
  };
};
```

### 3. Caching Strategy

```typescript
// services/CacheService.ts
export interface ICacheService {
  get<T>(key: string): T | null;
  set<T>(key: string, value: T, ttlMs?: number): void;
  remove(key: string): void;
  clear(): void;
}

export class CacheService implements ICacheService {
  private cache = new Map<string, { value: any; expiry: number }>();

  public get<T>(key: string): T | null {
    const item = this.cache.get(key);
    if (!item) {
      return null;
    }

    if (Date.now() > item.expiry) {
      this.cache.delete(key);
      return null;
    }

    return item.value as T;
  }

  public set<T>(key: string, value: T, ttlMs = 300000): void { // Default 5 minutes
    const expiry = Date.now() + ttlMs;
    this.cache.set(key, { value, expiry });
  }

  public remove(key: string): void {
    this.cache.delete(key);
  }

  public clear(): void {
    this.cache.clear();
  }
}
```

## Configuration and Environment Management

### 1. Environment Configuration

```typescript
// config/EnvironmentConfig.ts
export interface IEnvironmentConfig {
  backendApiUrl: string;
  authScopes: string[];
  cacheSettings: {
    defaultTtlMs: number;
    maxCacheSize: number;
  };
  retrySettings: {
    maxRetries: number;
    retryDelayMs: number;
  };
}

export class EnvironmentConfig {
  public static getConfig(): IEnvironmentConfig {
    // In a real implementation, this would come from web part properties
    // or SharePoint tenant properties
    return {
      backendApiUrl: this.getBackendApiUrl(),
      authScopes: ['https://api.spexternal.com/.default'],
      cacheSettings: {
        defaultTtlMs: 300000, // 5 minutes
        maxCacheSize: 100
      },
      retrySettings: {
        maxRetries: 3,
        retryDelayMs: 1000
      }
    };
  }

  private static getBackendApiUrl(): string {
    // Environment-specific URLs
    const hostname = window.location.hostname;
    
    if (hostname.includes('localhost')) {
      return 'http://localhost:7071/api'; // Local development
    } else if (hostname.includes('.sharepoint.com')) {
      return 'https://api.spexternal.com/v1'; // Production
    } else {
      return 'https://staging-api.spexternal.com/v1'; // Staging
    }
  }
}
```

### 2. Web Part Property Configuration

```typescript
// ExternalUserManagerWebPart.ts
export interface IExternalUserManagerWebPartProps {
  description: string;
  backendApiUrl?: string;
  enableCaching?: boolean;
  cacheTimeoutMinutes?: number;
}

export default class ExternalUserManagerWebPart extends BaseClientSideWebPart<IExternalUserManagerWebPartProps> {
  
  protected getPropertyPaneConfiguration(): IPropertyPaneConfiguration {
    return {
      pages: [
        {
          header: {
            description: strings.PropertyPaneDescription
          },
          groups: [
            {
              groupName: strings.BasicGroupName,
              groupFields: [
                PropertyPaneTextField('description', {
                  label: strings.DescriptionFieldLabel
                }),
                PropertyPaneTextField('backendApiUrl', {
                  label: 'Backend API URL',
                  description: 'Override the default backend API URL'
                }),
                PropertyPaneToggle('enableCaching', {
                  label: 'Enable Caching',
                  checked: true
                }),
                PropertyPaneSlider('cacheTimeoutMinutes', {
                  label: 'Cache Timeout (minutes)',
                  min: 1,
                  max: 60,
                  value: 5,
                  showValue: true
                })
              ]
            }
          ]
        }
      ]
    };
  }

  public render(): void {
    const element: React.ReactElement<IExternalUserManagerProps> = React.createElement(
      ExternalUserManager,
      {
        description: this.properties.description,
        context: this.context,
        backendApiUrl: this.properties.backendApiUrl,
        enableCaching: this.properties.enableCaching,
        cacheTimeoutMinutes: this.properties.cacheTimeoutMinutes
      }
    );

    ReactDom.render(element, this.domElement);
  }
}
```

This comprehensive communication architecture ensures robust, secure, and efficient communication between the SPFx frontend and the backend API service while maintaining proper separation of concerns and providing excellent error handling and caching capabilities.