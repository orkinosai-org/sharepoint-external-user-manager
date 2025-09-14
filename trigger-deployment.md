# Triggering SPFx Solution Deployment

This guide explains how to trigger the GitHub Actions workflow that builds, packages, and deploys the SharePoint Framework solution to the SharePoint App Catalog.

## Prerequisites

Ensure the following repository secrets are configured in GitHub:

1. **SPO_URL**: `https://turqoisecms-admin.sharepoint.com`
2. **SPO_USERNAME**: SharePoint admin username (e.g., `admin@turqoisecms-admin.onmicrosoft.com`)
3. **SPO_PASSWORD**: SharePoint admin password (use app password if MFA is enabled)

## Deployment Methods

### Method 1: Automatic Deployment (Push to Main)

The deployment workflow automatically triggers when code is pushed to the `main` branch:

```bash
git checkout main
git pull origin main
# Make any necessary changes
git add .
git commit -m "Trigger deployment: Update SPFx solution"
git push origin main
```

### Method 2: Manual Deployment (Workflow Dispatch)

You can manually trigger the deployment through GitHub's interface:

1. Go to your repository on GitHub
2. Click on the **Actions** tab
3. Select **Deploy SPFx Solution to SharePoint** workflow
4. Click **Run workflow** button
5. Select the `main` branch
6. Click **Run workflow**

### Method 3: API Trigger (Advanced)

Use GitHub's REST API to trigger the deployment:

```bash
# Replace with your GitHub token and repository details
curl -X POST \
  -H "Accept: application/vnd.github.v3+json" \
  -H "Authorization: token YOUR_GITHUB_TOKEN" \
  https://api.github.com/repos/orkinosai-org/sharepoint-external-user-manager/actions/workflows/deploy-spfx.yml/dispatches \
  -d '{"ref":"main"}'
```

## Monitoring Deployment

### 1. GitHub Actions Logs

- Navigate to **Actions** tab in your repository
- Click on the running workflow
- Monitor both **Build SPFx Solution** and **Deploy to SharePoint App Catalog** jobs
- Check the detailed logs for each step

### 2. Key Log Sections to Monitor

#### Build Phase:
- ✅ Dependencies installation
- ✅ SPFx solution build
- ✅ Package creation
- ✅ Artifact upload

#### Deployment Phase:
- ✅ Prerequisites validation
- ✅ SharePoint connection
- ✅ Package upload to App Catalog
- ✅ Solution publishing
- ✅ Deployment summary

### 3. Deployment Summary

The workflow provides a comprehensive deployment summary including:
- Package details (name, size, version)
- App Catalog information (App ID, title)
- Timing information
- Next steps for using the deployed solution

## Verifying Successful Deployment

### 1. Check SharePoint App Catalog

1. Navigate to: `https://turqoisecms-admin.sharepoint.com/_layouts/15/tenantAppCatalog.aspx`
2. Look for **sharepoint-external-user-manager-client-side-solution**
3. Verify the app is **Available** and **Published**

### 2. Test Web Part Availability

1. Go to any SharePoint site
2. Edit a page or create a new page
3. Add a web part
4. Search for "External User Manager" (or other web parts from the solution)
5. Verify the web part can be added and functions correctly

## Troubleshooting

### Common Issues:

1. **Authentication Failures**
   - Verify SPO_URL, SPO_USERNAME, and SPO_PASSWORD secrets
   - Ensure the account has SharePoint administrator privileges
   - Check if MFA is enabled (use app password instead)

2. **Package Not Found**
   - Check build logs for compilation errors
   - Verify all dependencies are correctly installed
   - Ensure Node.js version compatibility (18.x in workflow)

3. **App Catalog Access Issues**
   - Confirm the SharePoint tenant has an App Catalog provisioned
   - Verify the user has permissions to upload to App Catalog
   - Check network connectivity to SharePoint Online

### Getting Help:

- Check the detailed workflow logs in GitHub Actions
- Review the deployment summary for specific error information
- Verify SharePoint tenant configuration and permissions

## Expected Output

Upon successful deployment, you should see:

```
🚀 SharePoint Framework Solution Deployment
==============================================
✅ DEPLOYMENT SUCCESSFUL
==============================================
Package: sharepoint-external-user-manager-client-side-solution.sppkg
App ID: [Generated GUID]
App Title: sharepoint-external-user-manager-client-side-solution
App Version: 1.0.0.0
Deployed to: https://turqoisecms-admin.sharepoint.com
==============================================
```

The solution includes multiple web parts:
- 📋 External User Manager
- 🏢 Meeting Room Booking  
- 📦 Inventory Product Catalog
- ⏰ Timesheet Management
- 🤖 AI-Powered FAQ

All web parts will be available for use across your SharePoint tenant after successful deployment.