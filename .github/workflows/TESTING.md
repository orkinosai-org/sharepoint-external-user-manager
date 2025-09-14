# Testing the GitHub Actions Workflow

This guide helps you test the GitHub Actions workflow for SharePoint deployment.

## Pre-Deployment Testing

### 1. Test Build Workflow

The `test-build.yml` workflow validates that your solution builds correctly:

```bash
# This workflow runs automatically on:
# - Pull requests to main branch
# - Manual dispatch from GitHub Actions tab
```

**To test manually:**
1. Go to **Actions** tab in GitHub
2. Select "Test Build SPFx Solution" workflow
3. Click "Run workflow" button
4. Monitor the build process

### 2. Local Build Verification

Before deploying, verify the build works locally:

```bash
# Use Node.js 18.x (same as the workflow)
nvm use 18

# Install dependencies
npm ci

# Build the solution
npm run build

# Package for deployment
npm run package-solution

# Verify package creation
ls -la sharepoint/solution/*.sppkg
```

## Deployment Testing

### 1. Staging Environment (Recommended)

Before deploying to production, test with a staging SharePoint tenant:

1. **Setup staging secrets:**
   - `SPO_URL_STAGING`: Staging tenant URL
   - `SPO_USERNAME_STAGING`: Staging admin account
   - `SPO_PASSWORD_STAGING`: Staging admin password

2. **Create staging workflow:**
   - Copy `deploy-spfx.yml` to `deploy-staging.yml`
   - Change environment from `production` to `staging`
   - Use staging secrets instead of production ones

3. **Test deployment:**
   - Push to a staging branch
   - Monitor deployment process
   - Verify app appears in staging App Catalog

### 2. Production Deployment

Once staging is verified:

1. **Configure production secrets** (if not done already)
2. **Create pull request** to main branch
3. **Review and approve** the PR
4. **Merge to main** - this triggers automatic deployment
5. **Monitor workflow** in Actions tab
6. **Verify deployment** in SharePoint App Catalog

## Manual Testing Checklist

### Before Deployment
- [ ] Node.js version matches workflow (18.x)
- [ ] Local build completes successfully
- [ ] Package file (*.sppkg) is created
- [ ] Package size is reasonable (not 0 bytes)
- [ ] ESLint passes (if configured)

### Repository Configuration
- [ ] Required secrets are configured
- [ ] Production environment is set up (if using environment protection)
- [ ] Repository permissions allow workflow execution

### SharePoint Prerequisites
- [ ] Service account has SharePoint Admin role
- [ ] App Catalog is accessible and has sufficient space
- [ ] No conflicting apps with same name/ID

### After Deployment
- [ ] Workflow completes successfully
- [ ] App appears in SharePoint App Catalog
- [ ] App can be added to SharePoint pages
- [ ] Web part functions correctly

## Troubleshooting Test Failures

### Build Failures

```bash
# Check Node.js version compatibility
node --version
npm --version

# Clear npm cache if needed
npm cache clean --force

# Reinstall dependencies
rm -rf node_modules package-lock.json
npm install
```

### Deployment Failures

1. **Authentication Issues:**
   ```powershell
   # Test connection manually
   Connect-PnPOnline -Url "https://yourtenant.sharepoint.com" -Interactive
   Get-PnPTenantServicePrincipal
   ```

2. **Package Issues:**
   ```bash
   # Verify package structure
   unzip -l sharepoint/solution/*.sppkg
   
   # Check package manifest
   cat config/package-solution.json
   ```

3. **Permission Issues:**
   - Verify admin account permissions
   - Check App Catalog access
   - Review SharePoint admin center logs

## Rollback Procedures

If deployment fails or causes issues:

### 1. Immediate Rollback
```powershell
# Connect to SharePoint
Connect-PnPOnline -Url "https://yourtenant.sharepoint.com"

# Remove the problematic app
Remove-PnPApp -Identity "App-Name" -Confirm:$false

# Or retract if needed
Unpublish-PnPApp -Identity "App-Name"
```

### 2. Restore Previous Version
```powershell
# Get app history
Get-PnPApp | Where-Object {$_.Title -eq "Your-App-Name"}

# Restore from backup if available
# (This requires maintaining app version backups)
```

### 3. Workflow Rollback
- Revert the problematic commit in Git
- Push to main branch to trigger re-deployment
- Monitor new deployment

## Performance Testing

### Load Testing the Deployed App
```javascript
// Test web part performance
console.time('WebPartLoad');
// Load web part
console.timeEnd('WebPartLoad');

// Test API calls
console.time('APICall');
// Make SharePoint API call
console.timeEnd('APICall');
```

### Monitoring Deployment Performance
- Track deployment duration in GitHub Actions
- Monitor SharePoint App Catalog response times
- Check for any performance degradation after deployment

## Security Testing

### 1. Secret Management Test
```bash
# Verify secrets are not exposed in logs
# Check GitHub Actions logs for any credential leaks
grep -i "password\|secret\|token" workflow-logs.txt
```

### 2. Permission Validation
```powershell
# Verify app permissions match expectations
Get-PnPApp -Identity "Your-App" | Select-Object -ExpandProperty AppPermissions
```

### 3. Code Security Scan
```bash
# Run security audit on dependencies
npm audit

# Check for sensitive data in package
grep -r "password\|secret\|key" src/
```

## Automated Testing Integration

### Adding Tests to Workflow

You can enhance the workflow with automated tests:

```yaml
- name: Run Unit Tests
  run: |
    npm test
    
- name: Run E2E Tests
  run: |
    npm run test:e2e
    
- name: Security Scan
  run: |
    npm audit --audit-level moderate
```

### Test Coverage Reports
```yaml
- name: Generate Coverage Report
  run: |
    npm run test:coverage
    
- name: Upload Coverage
  uses: codecov/codecov-action@v3
```

This testing approach ensures reliable, secure deployments of your SharePoint Framework solution.