# SharePoint External User Manager - Deployment Instructions

## Prerequisites
- Node.js version 16.13.0+ or 18.17.0+ (but <19.0.0)
- SharePoint Online tenant with App Catalog
- SPFx development environment

## Build and Package
```bash
# Install dependencies
npm install

# Build the solution
npm run build

# Package for deployment
npm run package-solution
```

## Deploy to SharePoint

### Automated Deployment (Recommended)
The project includes a GitHub Actions workflow for automated deployment:

1. **Configure Repository Secrets** (one-time setup):
   - `SPO_URL`: Your SharePoint tenant URL (e.g., `https://contoso.sharepoint.com`)
   - `SPO_USERNAME`: SharePoint admin username (e.g., `admin@contoso.onmicrosoft.com`)
   - `SPO_PASSWORD`: SharePoint admin password

2. **Deploy**: Push changes to the `main` branch or manually trigger the workflow
   - The workflow automatically builds, packages, and deploys the solution
   - View deployment status in the GitHub Actions tab

See [GitHub Actions Workflow Documentation](.github/workflows/README.md) for detailed setup instructions.

### Manual Deployment
1. Navigate to your SharePoint App Catalog
2. Upload the generated `.sppkg` file from `solution/` folder
3. Trust the solution when prompted
4. Add the web part to a SharePoint page

## Features Included
- ✅ External library management with DetailsList
- ✅ Create new libraries with validation
- ✅ Delete libraries with confirmation
- ✅ Manage external users with bulk operations
- ✅ Permission-based UI and validation
- ✅ Mock data service for testing
- ✅ SharePoint API integration ready
- ✅ Responsive Fluent UI design

## Development Mode
```bash
# Start development server
npm run serve

# Access via SharePoint Workbench
https://yourtenant.sharepoint.com/_layouts/workbench.aspx
```

The web part is fully functional with placeholder data and ready for production deployment.
