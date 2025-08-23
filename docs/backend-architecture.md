# Backend Architecture

## Overview

This document outlines the backend architecture for the SharePoint External User Manager, designed as a scalable, multi-tenant SaaS solution using Azure cloud services.

## Architecture Diagram

```
┌─────────────────────────────────────────────────────────────────────────────────┐
│                               Azure Cloud Environment                           │
├─────────────────────────────────────────────────────────────────────────────────┤
│                                                                                 │
│  ┌─────────────────────┐    ┌─────────────────────┐    ┌─────────────────────┐ │
│  │   Azure Front Door  │    │   Azure API         │    │   Azure Functions   │ │
│  │   (Global CDN)      │◄──►│   Management        │◄──►│   (Serverless API)  │ │
│  │   - SSL/TLS         │    │   - Rate Limiting   │    │   - Auto-scaling    │ │
│  │   - WAF Protection  │    │   - API Gateway     │    │   - Event-driven    │ │
│  └─────────────────────┘    │   - Authentication  │    └─────────────────────┘ │
│                              └─────────────────────┘                            │
│                                                                                 │
│  ┌─────────────────────────────────────────────────────────────────────────────┐ │
│  │                         Authentication & Authorization                       │ │
│  │  ┌─────────────────┐    ┌─────────────────────┐    ┌─────────────────────┐ │ │
│  │  │   Azure AD B2B  │    │   Azure AD          │    │   Key Vault         │ │ │
│  │  │   Multi-tenant  │◄──►│   App Registration  │◄──►│   - Secrets         │ │ │
│  │  │   - External    │    │   - RBAC            │    │   - Certificates    │ │ │
│  │  │   - Guest Users │    │   - Scopes/Claims   │    │   - Connection Str  │ │ │
│  │  └─────────────────┘    └─────────────────────┘    └─────────────────────┘ │ │
│  └─────────────────────────────────────────────────────────────────────────────┘ │
│                                                                                 │
│  ┌─────────────────────────────────────────────────────────────────────────────┐ │
│  │                              Data Layer                                     │ │
│  │  ┌─────────────────┐    ┌─────────────────────┐    ┌─────────────────────┐ │ │
│  │  │   Azure Cosmos  │    │   Azure SQL         │    │   Azure Storage     │ │ │
│  │  │   DB (NoSQL)    │    │   Database          │    │   Account           │ │ │
│  │  │   - Audit Logs  │    │   - Tenant Data     │    │   - File Storage    │ │ │
│  │  │   - Cache       │    │   - User Sessions   │    │   - Blob Storage    │ │ │
│  │  │   - Metadata    │    │   - Configuration   │    │   - Static Assets   │ │ │
│  │  └─────────────────┘    └─────────────────────┘    └─────────────────────┘ │ │
│  └─────────────────────────────────────────────────────────────────────────────┘ │
│                                                                                 │
│  ┌─────────────────────────────────────────────────────────────────────────────┐ │
│  │                         External Integrations                               │ │
│  │  ┌─────────────────┐    ┌─────────────────────┐    ┌─────────────────────┐ │ │
│  │  │   Microsoft     │    │   SharePoint        │    │   Microsoft Graph   │ │ │
│  │  │   Graph API     │◄──►│   Online            │◄──►│   API               │ │ │
│  │  │   - User Mgmt   │    │   - Libraries       │    │   - Tenant Ops      │ │ │
│  │  │   - Permissions │    │   - Sites           │    │   - Admin Consent   │ │ │
│  │  └─────────────────┘    └─────────────────────┘    └─────────────────────┘ │ │
│  └─────────────────────────────────────────────────────────────────────────────┘ │
│                                                                                 │
│  ┌─────────────────────────────────────────────────────────────────────────────┐ │
│  │                        Monitoring & Analytics                               │ │
│  │  ┌─────────────────┐    ┌─────────────────────┐    ┌─────────────────────┐ │ │
│  │  │   Application   │    │   Log Analytics     │    │   Azure Monitor     │ │ │
│  │  │   Insights      │    │   Workspace         │    │   - Alerts          │ │ │
│  │  │   - Performance │    │   - Centralized     │    │   - Dashboards      │ │ │
│  │  │   - Telemetry   │    │   - Query Engine    │    │   - Metrics         │ │ │
│  │  └─────────────────┘    └─────────────────────┘    └─────────────────────┘ │ │
│  └─────────────────────────────────────────────────────────────────────────────┘ │
└─────────────────────────────────────────────────────────────────────────────────┘

┌─────────────────────────────────────────────────────────────────────────────────┐
│                              Client Applications                                 │
├─────────────────────────────────────────────────────────────────────────────────┤
│  ┌─────────────────┐    ┌─────────────────────┐    ┌─────────────────────┐     │
│  │   SPFx Web Part │    │   Power Platform    │    │   Custom Apps       │     │
│  │   - React UI    │◄──►│   - Power Apps      │◄──►│   - Mobile Apps     │     │
│  │   - Fluent UI   │    │   - Power Automate  │    │   - Desktop Apps    │     │
│  │   - TypeScript  │    │   - Power BI        │    │   - 3rd Party       │     │
│  └─────────────────┘    └─────────────────────┘    └─────────────────────┘     │
└─────────────────────────────────────────────────────────────────────────────────┘
```

## Core Components

### 1. API Gateway (Azure API Management)

**Purpose**: Central entry point for all API requests with security, throttling, and monitoring.

**Features**:
- Rate limiting and throttling
- API versioning and routing
- Authentication validation
- Request/response transformation
- Analytics and monitoring
- Developer portal for API documentation

**Configuration**:
```json
{
  "policies": {
    "rateLimit": "100 requests per minute per tenant",
    "authentication": "Azure AD Bearer token required",
    "cors": "Configured for SharePoint domains",
    "throttling": "Tenant-based quotas"
  }
}
```

### 2. Serverless API (Azure Functions)

**Purpose**: Business logic implementation using serverless compute for cost-effectiveness and auto-scaling.

**Function App Structure**:
```
SharePointExternalUserManager/
├── LibraryManagement/
│   ├── GetLibraries.cs
│   ├── CreateLibrary.cs
│   ├── DeleteLibrary.cs
│   └── UpdateLibrary.cs
├── UserManagement/
│   ├── GetExternalUsers.cs
│   ├── InviteUser.cs
│   ├── UpdateUserPermissions.cs
│   └── RevokeUserAccess.cs
├── TenantManagement/
│   ├── GetTenantSettings.cs
│   ├── UpdateTenantSettings.cs
│   └── ValidateTenantAccess.cs
└── Shared/
    ├── Authentication.cs
    ├── GraphApiClient.cs
    ├── SharePointClient.cs
    └── TenantResolver.cs
```

**Function Configuration**:
```json
{
  "runtime": ".NET 6 (LTS)",
  "hosting": {
    "plan": "Consumption (Y1)",
    "scaling": "Auto-scale based on demand"
  },
  "triggers": "HTTP trigger with Azure AD authentication",
  "bindings": {
    "input": ["HTTP Request", "Azure AD Token"],
    "output": ["HTTP Response", "Storage Queue", "Cosmos DB"]
  }
}
```

### 3. Multi-Tenant Data Architecture

**Tenant Isolation Strategy**: Database-per-tenant with shared infrastructure.

#### 3.1 Azure SQL Database (Tenant Data)
```sql
-- Tenant-specific databases
CREATE DATABASE [tenant_{tenantId}_spexternal]

-- Core tables per tenant
Tables:
├── Libraries
├── ExternalUsers  
├── UserPermissions
├── InvitationHistory
├── AuditLogs
└── TenantConfiguration
```

#### 3.2 Azure Cosmos DB (Shared Metadata)
```json
{
  "databases": {
    "SharedMetadata": {
      "containers": [
        "Tenants",
        "GlobalAuditLogs", 
        "SystemConfiguration",
        "UsageMetrics",
        "PerformanceMetrics"
      ]
    }
  },
  "partitioning": "By tenant ID for optimal performance"
}
```

### 4. Authentication & Authorization

#### 4.1 Azure AD Multi-Tenant Application

**Registration Configuration**:
```json
{
  "appRegistration": {
    "name": "SharePoint External User Manager",
    "signInAudience": "AzureADMultipleOrgs",
    "requiredResourceAccess": [
      {
        "resourceAppId": "00000003-0000-0000-c000-000000000000",
        "resourceAccess": [
          {
            "id": "e1fe6dd8-ba31-4d61-89e7-88639da4683d",
            "type": "Scope"
          }
        ]
      }
    ]
  }
}
```

**Required Permissions**:
- `Sites.Read.All` - Read all sites
- `Sites.Manage.All` - Manage libraries
- `Sites.FullControl.All` - Full library control
- `User.Read.All` - Read user profiles
- `Directory.Read.All` - Read directory data

#### 4.2 Role-Based Access Control (RBAC)

```json
{
  "roles": {
    "TenantAdmin": [
      "libraries:read", "libraries:write", "libraries:delete",
      "users:read", "users:write", "users:delete",
      "settings:read", "settings:write"
    ],
    "LibraryOwner": [
      "libraries:read", "libraries:write",
      "users:read", "users:write", "users:delete"
    ],
    "LibraryContributor": [
      "libraries:read",
      "users:read", "users:write"
    ],
    "LibraryReader": [
      "libraries:read",
      "users:read"
    ]
  }
}
```

## Scalability Design

### 1. Horizontal Scaling

- **Azure Functions**: Auto-scale based on demand (0-200 instances)
- **API Management**: Premium tier with multiple regions
- **Cosmos DB**: Auto-scale throughput (400-4000 RU/s)
- **SQL Database**: Elastic pool for tenant databases

### 2. Performance Optimization

- **Caching**: Redis Cache for frequently accessed data
- **CDN**: Azure Front Door for global content delivery
- **Connection Pooling**: Optimized database connections
- **Async Processing**: Message queues for heavy operations

### 3. Multi-Region Deployment

```json
{
  "regions": {
    "primary": "East US 2",
    "secondary": "West Europe", 
    "tertiary": "Southeast Asia"
  },
  "failover": "Automatic with 99.9% uptime SLA",
  "dataReplication": "Geo-redundant storage"
}
```

## Security Architecture

### 1. Network Security

- **Azure Front Door**: WAF protection and DDoS mitigation
- **Virtual Network**: Isolated network for backend services
- **Private Endpoints**: Secure database connections
- **Network Security Groups**: Traffic filtering rules

### 2. Data Security

- **Encryption at Rest**: Azure Storage Service Encryption
- **Encryption in Transit**: TLS 1.2+ for all communications
- **Key Management**: Azure Key Vault for secrets
- **Data Classification**: Sensitivity labels and protection

### 3. Identity Security

- **Azure AD**: Multi-factor authentication required
- **Conditional Access**: Risk-based authentication
- **Privileged Identity Management**: Just-in-time access
- **Identity Protection**: Anomaly detection

## Monitoring & Observability

### 1. Application Performance Monitoring

```json
{
  "applicationInsights": {
    "telemetry": [
      "Request/response times",
      "Dependency calls",
      "Exception tracking",
      "Custom events"
    ],
    "alerts": [
      "High response times (>2s)",
      "Error rate threshold (>5%)",
      "Dependency failures"
    ]
  }
}
```

### 2. Infrastructure Monitoring

- **Azure Monitor**: Resource health and metrics
- **Log Analytics**: Centralized logging and querying
- **Service Health**: Azure service status monitoring
- **Resource Health**: Individual resource status

### 3. Business Metrics

- **Usage Analytics**: API call patterns and frequency
- **Tenant Metrics**: Growth and usage per tenant
- **Performance KPIs**: SLA compliance and response times
- **Cost Analytics**: Resource utilization and optimization

## Deployment Strategy

### 1. Infrastructure as Code (IaC)

```yaml
# Azure Resource Manager Templates
Resources:
  - Resource Group
  - Function App
  - API Management
  - SQL Server/Databases
  - Cosmos DB Account
  - Storage Account
  - Key Vault
  - Application Insights
```

### 2. CI/CD Pipeline

```yaml
stages:
  - Development:
      environment: dev
      deployment: Automatic on PR merge
  - Staging:
      environment: staging  
      deployment: Manual approval required
  - Production:
      environment: prod
      deployment: Manual approval + change management
```

### 3. Deployment Slots

- **Blue-Green Deployment**: Zero-downtime deployments
- **Staged Rollouts**: Gradual traffic shifting
- **Rollback Capability**: Instant rollback on issues

## Cost Optimization

### 1. Consumption-Based Pricing

- **Azure Functions**: Pay per execution
- **Cosmos DB**: Auto-scale based on usage
- **API Management**: Pay per call
- **Application Insights**: Pay per data ingestion

### 2. Resource Optimization

- **Reserved Instances**: For predictable workloads
- **Auto-shutdown**: Development environments
- **Storage Tiering**: Archive old audit logs
- **Database Optimization**: Right-size based on usage

## Disaster Recovery

### 1. Backup Strategy

- **Database Backups**: Point-in-time recovery (35 days)
- **Application Code**: Source control with multiple branches
- **Configuration**: Azure DevOps variable groups
- **Secrets**: Key Vault with geo-replication

### 2. Recovery Procedures

- **RTO (Recovery Time Objective)**: 4 hours
- **RPO (Recovery Point Objective)**: 1 hour
- **Automated Failover**: Regional failure detection
- **Manual Procedures**: Documented step-by-step recovery

This architecture provides a solid foundation for a scalable, secure, and maintainable SaaS solution that can grow with customer demand while maintaining high availability and performance standards.