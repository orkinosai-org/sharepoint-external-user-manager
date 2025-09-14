# GitHub Actions Workflows

This directory contains GitHub Actions workflows for the SharePoint External User Manager project.

## Deploy SPFx Solution (`deploy-spfx.yml`)

This workflow automatically builds and deploys the SharePoint Framework (SPFx) solution to your SharePoint tenant's App Catalog whenever code is pushed to the `main` branch.

### Workflow Overview

The deployment process consists of two main jobs:

1. **Build Job**:
   - Sets up Node.js 18.x environment
   - Installs npm dependencies
   - Builds the SPFx solution
   - Packages the solution into a `.sppkg` file
   - Uploads the package as a build artifact

2. **Deploy Job**:
   - Downloads the build artifact
   - Installs PnP PowerShell module
   - Connects to SharePoint Online
   - Uploads the `.sppkg` file to the App Catalog
   - Publishes the solution
   - Provides deployment summary

### Required Repository Secrets

Before using this workflow, you must configure the following secrets in your GitHub repository:

| Secret Name | Description | Example |
|-------------|-------------|---------|
| `SPO_URL` | SharePoint tenant URL | `https://contoso.sharepoint.com` |
| `SPO_USERNAME` | SharePoint admin username | `admin@contoso.onmicrosoft.com` |
| `SPO_PASSWORD` | SharePoint admin password | `YourSecurePassword123!` |

### Setting Up Repository Secrets

1. Navigate to your GitHub repository
2. Go to **Settings** → **Secrets and variables** → **Actions**
3. Click **New repository secret**
4. Add each of the required secrets listed above

### Authentication Requirements

The deployment account must have:
- SharePoint Administrator role
- Access to the tenant App Catalog
- Appropriate permissions to upload and publish apps

**Note**: If your organization uses Multi-Factor Authentication (MFA), you may need to:
- Use an app password instead of the regular password
- Create a dedicated service account without MFA
- Use certificate-based authentication (requires workflow modification)

### Triggering the Workflow

The workflow can be triggered in two ways:

1. **Automatic**: Push changes to the `main` branch
2. **Manual**: Use the "Run workflow" button in the GitHub Actions tab

### Environment Protection

The workflow uses a `production` environment for the deploy job, which can be configured to:
- Require manual approval before deployment
- Restrict which branches can deploy
- Add deployment protection rules

To configure environment protection:
1. Go to **Settings** → **Environments**
2. Click on the `production` environment
3. Configure protection rules as needed

### Monitoring Deployments

Each deployment provides:
- **Build artifacts**: The `.sppkg` file is stored for 30 days
- **Deployment summary**: Success/failure status with details
- **Logs**: Detailed execution logs for troubleshooting

### Troubleshooting

#### Common Issues

1. **Node.js Version Mismatch**
   - The workflow uses Node.js 18.x as required by SPFx 1.18.2
   - Local development should use the same version

2. **Authentication Failures**
   - Verify the SPO_USERNAME and SPO_PASSWORD secrets
   - Check if the account has proper permissions
   - Consider using app passwords for MFA-enabled accounts

3. **Package Upload Failures**
   - Ensure the App Catalog is accessible
   - Check if the package name conflicts with existing apps
   - Verify sufficient storage space in the App Catalog

4. **Build Failures**
   - Check for TypeScript compilation errors
   - Verify all dependencies are properly installed
   - Review ESLint warnings that might block the build

#### Viewing Logs

To view detailed logs:
1. Go to the **Actions** tab in your repository
2. Click on the failed workflow run
3. Expand the failed step to see detailed error messages

#### Common Troubleshooting Steps

1. **Test Connection Manually**:
   ```powershell
   # Test PowerShell connection locally
   Install-Module -Name PnP.PowerShell -Force
   Connect-PnPOnline -Url "https://yourtenant.sharepoint.com" -Interactive
   Get-PnPApp
   ```

2. **Verify App Catalog Access**:
   - Ensure the service account can access `https://yourtenant.sharepoint.com/sites/appcatalog`
   - Check that the App Catalog is properly configured
   - Verify sufficient storage space in the App Catalog

3. **Check Package File**:
   - Verify the `.sppkg` file is created correctly in the build step
   - Check the package size (should be > 0 bytes)
   - Ensure the package isn't corrupted

4. **Authentication Debugging**:
   ```powershell
   # Check account status
   Get-PnPTenantServicePrincipal
   Get-PnPUser -Identity "admin@tenant.onmicrosoft.com"
   ```

### Security Considerations

- **Secrets Management**: Never commit credentials to the repository
- **Least Privilege**: Use accounts with minimum required permissions
- **Regular Rotation**: Periodically update passwords and secrets
- **Audit Trail**: Monitor deployment logs for security events

### Customization

You can customize the workflow by:
- Modifying the trigger conditions (branches, paths)
- Adding additional build steps (testing, linting)
- Changing the deployment environment
- Adding notifications (Slack, Teams, email)
- Including additional validation steps

### Related Documentation

- [SharePoint Framework Development](../DEVELOPER_GUIDE.md)
- [Deployment Instructions](../deployment-instructions.md)
- [PnP PowerShell Documentation](https://pnp.github.io/powershell/)
- [GitHub Actions Documentation](https://docs.github.com/en/actions)