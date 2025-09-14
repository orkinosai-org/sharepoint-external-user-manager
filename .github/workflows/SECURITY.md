# Security Best Practices for GitHub Actions Workflows

This document outlines security considerations and best practices for the SharePoint deployment workflow.

## Repository Secrets Management

### Required Secrets
- `SPO_URL`: SharePoint tenant URL
- `SPO_USERNAME`: Service account username
- `SPO_PASSWORD`: Service account password

### Best Practices

1. **Use Dedicated Service Accounts**
   - Create a dedicated service account for automation
   - Don't use personal admin accounts
   - Apply principle of least privilege

2. **Secure Password Management**
   - Use strong, unique passwords
   - Consider app passwords for MFA-enabled accounts
   - Rotate passwords regularly (quarterly recommended)

3. **Secret Rotation**
   ```bash
   # Steps to rotate secrets:
   # 1. Update password in Azure AD/SharePoint
   # 2. Update GitHub repository secret
   # 3. Test workflow with new credentials
   # 4. Monitor for any authentication failures
   ```

## Authentication Alternatives

### Option 1: Certificate-Based Authentication
More secure than password-based authentication:

```yaml
- name: Connect with Certificate
  shell: pwsh
  run: |
    Connect-PnPOnline -Url $env:SPO_URL -CertificatePath "cert.pfx" -CertificatePassword $env:CERT_PASSWORD
```

### Option 2: Azure App Registration
Use Azure AD app registration with application permissions:

```yaml
- name: Connect with App Registration
  shell: pwsh
  run: |
    Connect-PnPOnline -Url $env:SPO_URL -ClientId $env:CLIENT_ID -ClientSecret $env:CLIENT_SECRET
```

## Environment Protection

### Production Environment
The workflow uses a `production` environment with these recommended settings:

1. **Required Reviewers**: At least one admin must approve deployments
2. **Branch Restrictions**: Only `main` branch can deploy to production
3. **Wait Timer**: Optional delay before deployment
4. **Custom Protection Rules**: Additional security checks

### Setup Instructions
1. Go to Repository Settings → Environments
2. Create/edit `production` environment
3. Configure protection rules:
   - Add required reviewers
   - Restrict deployment branches
   - Enable other security controls

## Audit and Monitoring

### Deployment Logging
- All deployment activities are logged in GitHub Actions
- SharePoint also maintains audit logs for app catalog activities
- Monitor both systems for security events

### Recommended Monitoring
1. **Failed Authentication Attempts**: Watch for unusual login failures
2. **Unexpected Deployments**: Monitor for deployments outside business hours
3. **Permission Changes**: Track changes to service account permissions
4. **Secret Access**: Monitor secret usage patterns

## Incident Response

### If Credentials are Compromised
1. **Immediate Actions**:
   - Disable the service account
   - Rotate all related passwords/secrets
   - Review recent deployment history
   - Check SharePoint audit logs

2. **Investigation**:
   - Review GitHub Actions logs
   - Check for unauthorized deployments
   - Verify app catalog contents
   - Assess potential data exposure

3. **Recovery**:
   - Create new service account with proper permissions
   - Update repository secrets
   - Test deployment workflow
   - Document lessons learned

## Compliance Considerations

### Data Protection
- No sensitive data should be included in the solution package
- Ensure compliance with organizational data handling policies
- Review solution permissions and data access patterns

### Change Management
- All production deployments should follow change management processes
- Use pull requests for code review before merging to main
- Maintain deployment documentation and rollback procedures

### Regulatory Requirements
- Some organizations require additional approvals for automated deployments
- Consider integrating with existing ITSM tools
- Maintain audit trails for compliance reporting

## Additional Security Measures

### Code Scanning
Consider adding security scanning to the workflow:

```yaml
- name: Run Security Scan
  uses: github/super-linter@v4
  env:
    DEFAULT_BRANCH: main
    GITHUB_TOKEN: ${{ secrets.GITHUB_TOKEN }}
```

### Dependency Scanning
Monitor for vulnerabilities in npm packages:

```yaml
- name: Run npm audit
  run: npm audit --audit-level moderate
```

### Secret Scanning
Enable GitHub's secret scanning features:
- Secret scanning for pushes
- Push protection to prevent accidental secret commits
- Custom patterns for organization-specific secrets

## Support and Questions

For security-related questions or to report security issues:
1. Contact your organization's security team
2. Review GitHub's security documentation
3. Consult SharePoint security best practices
4. Consider engaging with Microsoft support for tenant-level security guidance