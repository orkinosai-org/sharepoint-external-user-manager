# SharePoint External User Manager - Deployment Instructions

## Prerequisites
- Node.js version 18.17.1+ (but <19.0.0) - **IMPORTANT: SPFx 1.18.2 does NOT support Node.js 20.x or 21.x**
- SharePoint Online tenant with App Catalog
- SPFx development environment

### Node.js Version Requirements
This project uses SharePoint Framework (SPFx) 1.18.2 which has strict Node.js version requirements:
- ✅ Supported: Node.js 18.17.1 to 18.x.x
- ❌ Not Supported: Node.js 19.x, 20.x, 21.x

**Recommended setup:**
```bash
# Using nvm (Node Version Manager)
nvm install 18.19.0
nvm use 18.19.0

# Verify version
node --version  # Should output v18.19.0
```

**Troubleshooting:**
If you see an error like "Your dev environment is running NodeJS version vX.X.X which does not meet the requirements", ensure you're using Node.js 18.x before running any build commands.

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
