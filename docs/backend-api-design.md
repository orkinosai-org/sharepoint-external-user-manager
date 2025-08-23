# Backend API Design

## Overview

This document defines the RESTful API design for the SharePoint External User Manager backend service. The API provides library and external user management capabilities for the SPFx frontend in a multi-tenant SaaS environment.

## Base URL

```
https://api.spexternal.com/v1
```

## Authentication

All API endpoints require authentication via Azure AD Bearer tokens:

```http
Authorization: Bearer {azure_ad_token}
X-Tenant-ID: {tenant_id}
```

## API Endpoints

### 1. Library Management

#### GET /libraries
Fetch all libraries for the authenticated tenant.

**Request:**
```http
GET /libraries
Authorization: Bearer {token}
X-Tenant-ID: {tenant_id}
```

**Response:**
```json
{
  "success": true,
  "data": [
    {
      "id": "lib-001",
      "name": "External Projects",
      "description": "Documents shared with external partners",
      "siteUrl": "https://contoso.sharepoint.com/sites/external-projects",
      "externalUsersCount": 5,
      "lastModified": "2024-01-15T10:30:00Z",
      "owner": "john.doe@contoso.com",
      "permissions": "Contribute",
      "tenantId": "tenant-123"
    }
  ],
  "pagination": {
    "page": 1,
    "pageSize": 50,
    "total": 25,
    "hasNext": false
  }
}
```

**Query Parameters:**
- `page` (optional): Page number (default: 1)
- `pageSize` (optional): Items per page (default: 50, max: 100)
- `search` (optional): Search term for library name/description
- `owner` (optional): Filter by library owner

#### POST /libraries
Create a new library.

**Request:**
```http
POST /libraries
Authorization: Bearer {token}
X-Tenant-ID: {tenant_id}
Content-Type: application/json

{
  "name": "Partner Collaboration",
  "description": "Shared workspace for external partners",
  "siteUrl": "https://contoso.sharepoint.com/sites/partners",
  "enableExternalSharing": true,
  "template": 101
}
```

**Response:**
```json
{
  "success": true,
  "data": {
    "id": "lib-002",
    "name": "Partner Collaboration",
    "description": "Shared workspace for external partners",
    "siteUrl": "https://contoso.sharepoint.com/sites/partners",
    "externalUsersCount": 0,
    "lastModified": "2024-01-15T11:00:00Z",
    "owner": "john.doe@contoso.com",
    "permissions": "Full Control",
    "tenantId": "tenant-123"
  }
}
```

#### DELETE /libraries/{libraryId}
Remove a library.

**Request:**
```http
DELETE /libraries/lib-002
Authorization: Bearer {token}
X-Tenant-ID: {tenant_id}
```

**Response:**
```json
{
  "success": true,
  "message": "Library successfully deleted"
}
```

### 2. External User Management

#### GET /libraries/{libraryId}/users
Get external users for a specific library.

**Request:**
```http
GET /libraries/lib-001/users
Authorization: Bearer {token}
X-Tenant-ID: {tenant_id}
```

**Response:**
```json
{
  "success": true,
  "data": [
    {
      "id": "user-001",
      "email": "partner@external.com",
      "displayName": "Jane Partner",
      "invitedBy": "john.doe@contoso.com",
      "invitedDate": "2024-01-10T09:15:00Z",
      "lastAccess": "2024-01-14T16:45:00Z",
      "permissions": "Read",
      "status": "Active"
    }
  ],
  "pagination": {
    "page": 1,
    "pageSize": 50,
    "total": 5,
    "hasNext": false
  }
}
```

#### POST /libraries/{libraryId}/users
Invite external user to a library.

**Request:**
```http
POST /libraries/lib-001/users
Authorization: Bearer {token}
X-Tenant-ID: {tenant_id}
Content-Type: application/json

{
  "email": "newpartner@external.com",
  "displayName": "New Partner",
  "permissions": "Contribute",
  "message": "Welcome to our collaboration workspace"
}
```

**Response:**
```json
{
  "success": true,
  "data": {
    "id": "user-002",
    "email": "newpartner@external.com",
    "displayName": "New Partner",
    "invitedBy": "john.doe@contoso.com",
    "invitedDate": "2024-01-15T11:30:00Z",
    "lastAccess": null,
    "permissions": "Contribute",
    "status": "Invited"
  }
}
```

#### PUT /libraries/{libraryId}/users/{userId}
Update external user permissions.

**Request:**
```http
PUT /libraries/lib-001/users/user-001
Authorization: Bearer {token}
X-Tenant-ID: {tenant_id}
Content-Type: application/json

{
  "permissions": "Contribute"
}
```

**Response:**
```json
{
  "success": true,
  "data": {
    "id": "user-001",
    "email": "partner@external.com",
    "displayName": "Jane Partner",
    "invitedBy": "john.doe@contoso.com",
    "invitedDate": "2024-01-10T09:15:00Z",
    "lastAccess": "2024-01-14T16:45:00Z",
    "permissions": "Contribute",
    "status": "Active"
  }
}
```

#### DELETE /libraries/{libraryId}/users/{userId}
Remove external user from library.

**Request:**
```http
DELETE /libraries/lib-001/users/user-001
Authorization: Bearer {token}
X-Tenant-ID: {tenant_id}
```

**Response:**
```json
{
  "success": true,
  "message": "User access revoked successfully"
}
```

### 3. Tenant Management

#### GET /tenant/settings
Get tenant-level settings for external sharing.

**Request:**
```http
GET /tenant/settings
Authorization: Bearer {token}
X-Tenant-ID: {tenant_id}
```

**Response:**
```json
{
  "success": true,
  "data": {
    "externalSharingEnabled": true,
    "allowAnonymousLinks": false,
    "defaultLinkPermission": "View",
    "externalUserExpirationDays": 90,
    "requireApprovalForExternalSharing": true
  }
}
```

## Error Handling

All errors follow a consistent format:

```json
{
  "success": false,
  "error": {
    "code": "LIBRARY_NOT_FOUND",
    "message": "The specified library could not be found",
    "details": "Library ID 'lib-999' does not exist in tenant 'tenant-123'"
  }
}
```

### Error Codes

| Code | HTTP Status | Description |
|------|-------------|-------------|
| `UNAUTHORIZED` | 401 | Invalid or missing authentication token |
| `FORBIDDEN` | 403 | Insufficient permissions for operation |
| `LIBRARY_NOT_FOUND` | 404 | Library does not exist |
| `USER_NOT_FOUND` | 404 | External user does not exist |
| `VALIDATION_ERROR` | 400 | Request validation failed |
| `TENANT_NOT_FOUND` | 404 | Tenant does not exist or not configured |
| `EXTERNAL_SHARING_DISABLED` | 403 | External sharing is disabled for tenant |
| `RATE_LIMIT_EXCEEDED` | 429 | Too many requests |
| `INTERNAL_ERROR` | 500 | Internal server error |

## Rate Limiting

- **Standard Operations**: 100 requests per minute per tenant
- **Bulk Operations**: 10 requests per minute per tenant
- **User Invitations**: 50 invitations per hour per tenant

Rate limit headers are included in all responses:

```http
X-RateLimit-Limit: 100
X-RateLimit-Remaining: 95
X-RateLimit-Reset: 1705320000
```

## Pagination

All list endpoints support pagination:

- `page`: Page number (1-based)
- `pageSize`: Items per page (max 100)
- `total`: Total number of items
- `hasNext`: Whether there are more pages

## API Versioning

The API uses URL-based versioning (`/v1/`). Breaking changes will increment the major version.

## Security Considerations

1. **Authentication**: Azure AD Bearer tokens required for all endpoints
2. **Authorization**: Role-based access control (RBAC) based on SharePoint permissions
3. **Multi-tenancy**: Strict tenant isolation using X-Tenant-ID header
4. **Input Validation**: All inputs validated and sanitized
5. **Audit Logging**: All operations logged for compliance
6. **Rate Limiting**: Prevents abuse and ensures fair usage