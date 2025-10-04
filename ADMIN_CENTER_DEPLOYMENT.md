# Microsoft 365 Admin Center Deployment Guide

## Overview

This guide provides step-by-step instructions for deploying the Financial Report Generator add-in through Microsoft 365 Admin Center (admin.cloud.microsoft).

## Prerequisites

- ✅ Microsoft 365 Administrator account
- ✅ Access to Microsoft 365 Admin Center
- ✅ Production web hosting for add-in files (Azure, AWS, or your own server)
- ✅ HTTPS-enabled hosting (required for Office Add-ins)
- ✅ Built and tested add-in files

## Deployment Files

Two XML manifest files are provided:

1. **manifest.xml** - Development version (localhost:3000)
2. **manifest-production.xml** - Production template (requires URL updates)

## Step 1: Prepare Production Environment

### 1.1 Host Your Add-in Files

Upload the following files to your production web server:

#### Required Files:
```
/taskpane.html
/taskpane.js
/commands.html
/commands.js
/assets/icon-16.png
/assets/icon-32.png
/assets/icon-80.png
/assets/icon-128.png
/assets/logo-filled.png
/assets/color.png
/assets/outline.png
```

#### Build Files:
```bash
# Build for production
npm run build

# Files will be in the dist/ folder
ls -la dist/
```

#### Upload to Server:
- Option A: **Azure Blob Storage** (Recommended)
  - Create a storage account
  - Create a blob container
  - Enable static website hosting
  - Upload files from dist/ folder
  - Note the public URL (e.g., https://mystorageaccount.z13.web.core.windows.net/)

- Option B: **Your Web Server**
  - Ensure HTTPS is enabled (SSL certificate installed)
  - Upload files to public directory
  - Test file accessibility
  - Note the base URL (e.g., https://your-domain.com/)

- Option C: **Azure App Service** or **AWS S3**
  - Follow respective hosting provider instructions
  - Ensure static file hosting is configured
  - Enable HTTPS

### 1.2 Verify File Accessibility

Test that files are accessible:
```bash
# Test HTML files
curl https://your-domain.com/taskpane.html
curl https://your-domain.com/commands.html

# Test JS files
curl https://your-domain.com/taskpane.js
curl https://your-domain.com/commands.js

# Test images
curl https://your-domain.com/assets/icon-32.png
```

All files should return HTTP 200 status.

## Step 2: Update Manifest File

### 2.1 Edit manifest-production.xml

Replace all instances of `https://your-domain.com` with your actual production URL:

```xml
<!-- Example: If your hosting URL is https://financialreports.blob.core.windows.net/ -->

<!-- Update Icon URLs -->
<IconUrl DefaultValue="https://financialreports.blob.core.windows.net/assets/icon-32.png"/>
<HighResolutionIconUrl DefaultValue="https://financialreports.blob.core.windows.net/assets/icon-80.png"/>

<!-- Update Support URL -->
<SupportUrl DefaultValue="https://your-company.com/support"/>

<!-- Update App Domains -->
<AppDomains>
  <AppDomain>https://financialreports.blob.core.windows.net</AppDomain>
  <AppDomain>https://cdnjs.cloudflare.com</AppDomain>
</AppDomains>

<!-- Update Source Location -->
<SourceLocation DefaultValue="https://financialreports.blob.core.windows.net/taskpane.html"/>

<!-- Update Resources URLs -->
<bt:Image id="Icon.16x16" DefaultValue="https://financialreports.blob.core.windows.net/assets/icon-16.png"/>
<bt:Image id="Icon.32x32" DefaultValue="https://financialreports.blob.core.windows.net/assets/icon-32.png"/>
<bt:Image id="Icon.80x80" DefaultValue="https://financialreports.blob.core.windows.net/assets/icon-80.png"/>
<bt:Url id="Commands.Url" DefaultValue="https://financialreports.blob.core.windows.net/commands.html"/>
<bt:Url id="Taskpane.Url" DefaultValue="https://financialreports.blob.core.windows.net/taskpane.html"/>
```

### 2.2 Validate Manifest

```bash
# Install Office Add-in Validator (if not already installed)
npm install -g office-addin-manifest

# Validate the manifest
office-addin-manifest validate manifest-production.xml
```

Fix any validation errors before proceeding.

## Step 3: Deploy via Microsoft 365 Admin Center

### 3.1 Access Admin Center

1. Go to [https://admin.cloud.microsoft.com](https://admin.cloud.microsoft.com)
2. Sign in with your Microsoft 365 administrator credentials
3. Navigate to **Settings** > **Integrated apps**

### 3.2 Upload Custom App

1. Click **"Get apps"** or **"Upload custom apps"**
2. Select **"Upload custom app"**
3. Choose **"Provide link to manifest file"** OR **"Upload manifest file"**

#### Option A: Upload Manifest File
- Click **"Upload manifest file (.xml)"**
- Browse and select your updated `manifest-production.xml`
- Click **"Upload"**

#### Option B: Provide Manifest Link
- Host the manifest file on your web server
- Enter the URL: `https://your-domain.com/manifest-production.xml`
- Click **"Validate"**

### 3.3 Configure Deployment

After upload, you'll see the app configuration page:

1. **Review App Information**
   - Name: Financial Report Generator
   - Description: (as specified in manifest)
   - Version: 1.0.0.0

2. **Select Users**
   - **Option A**: Deploy to everyone in your organization
   - **Option B**: Deploy to specific users
   - **Option C**: Deploy to specific groups

3. **Choose Deployment Type**
   - **Recommended**: "Deploy to all users automatically"
   - **Alternative**: "Make available to users" (users install manually)

4. **Consent and Permissions**
   - Review required permissions: ReadWriteMailbox
   - Accept on behalf of organization (if you have authority)

5. **Click "Deploy"**

### 3.4 Monitor Deployment Status

- Deployment typically takes 1-24 hours
- Status will show "Deploying" then "Deployed"
- Users will see the add-in in their Outlook ribbon

## Step 4: Verify Deployment

### 4.1 Test in Outlook Desktop

1. Open Outlook Desktop
2. Open any email message
3. Look for "Financial Reports" group in the ribbon
4. Click "Generate Financial Report"
5. Verify task pane opens and functions correctly

### 4.2 Test in Outlook on the Web

1. Go to [https://outlook.office.com](https://outlook.office.com)
2. Open any email message
3. Click the three dots (...) in the ribbon
4. Select "Financial Report Generator"
5. Verify functionality

### 4.3 Test in Outlook Mobile (if applicable)

1. Open Outlook mobile app
2. Open an email
3. Check if add-in appears in available actions
4. Test basic functionality

## Step 5: User Communication

### 5.1 Notify Users

Send an email to all users:

**Subject**: New Financial Report Generator Add-in Now Available

**Body**:
```
Hello,

A new Outlook add-in is now available to help you generate financial reports directly from your inbox.

What it does:
- Scans your inbox for financial data
- Extracts information from emails and PDF attachments
- Generates quarterly, semi-annual, and annual financial statements

How to use:
1. Open any email in Outlook
2. Look for "Financial Reports" in the ribbon
3. Click "Generate Financial Report"
4. Follow the on-screen instructions

For support, contact [your-support-email] or visit [support-url]

Best regards,
IT Team
```

### 5.2 Provide Training Materials

- Link to QUICKSTART.md
- Video tutorial (if available)
- FAQ document
- Support contact information

## Troubleshooting

### Issue: Add-in Not Appearing

**Solutions**:
1. Wait 24 hours for deployment to complete
2. Restart Outlook
3. Clear Outlook cache:
   - Close Outlook
   - Delete: `%LocalAppData%\Microsoft\Office\16.0\Wef\`
   - Restart Outlook
4. Verify user is in deployed group
5. Check deployment status in Admin Center

### Issue: Add-in Loads but Doesn't Work

**Solutions**:
1. Check browser console for errors (F12)
2. Verify all file URLs are accessible
3. Check HTTPS is properly configured
4. Verify CORS headers if hosting on different domain
5. Check manifest URLs are correct

### Issue: Permission Errors

**Solutions**:
1. Verify admin consent was granted
2. Check user has appropriate Outlook license
3. Verify ReadWriteMailbox permission is accepted
4. Contact Microsoft support if needed

## Monitoring and Maintenance

### Monitor Usage

1. Check deployment status regularly
2. Monitor support requests
3. Collect user feedback
4. Track any error reports

### Update Process

To update the add-in:

1. Build new version
2. Upload updated files to hosting
3. Update version in manifest (e.g., 1.0.0.0 → 1.1.0.0)
4. In Admin Center, go to the app
5. Click "Update" and upload new manifest
6. Changes propagate to users automatically

### Support

For issues:
- Check [Microsoft 365 Admin Center Help](https://docs.microsoft.com/office365/admin/)
- Review [Office Add-ins Documentation](https://docs.microsoft.com/office/dev/add-ins/)
- Contact Microsoft Support
- Check application logs

## Security Checklist

Before deployment:
- [ ] HTTPS enabled on hosting
- [ ] Valid SSL certificate
- [ ] All URLs use HTTPS (no HTTP)
- [ ] Manifest validated successfully
- [ ] Tested in multiple Outlook versions
- [ ] Privacy policy available
- [ ] Terms of use documented
- [ ] Support contact provided
- [ ] CORS configured if needed
- [ ] Content Security Policy reviewed

## Rollback Plan

If issues occur:

1. In Admin Center, go to "Integrated apps"
2. Find "Financial Report Generator"
3. Click "..." menu
4. Select "Remove" or "Disable"
5. Confirm removal
6. Add-in will be removed from users within 24 hours

## Advanced Deployment Options

### Deploy to Specific Groups

1. In Azure AD, create a security group
2. Add users to the group
3. In Admin Center deployment, select "Specific groups"
4. Choose your security group

### Staged Rollout

1. **Phase 1**: Deploy to pilot group (5-10 users)
2. **Phase 2**: Deploy to department (50-100 users)
3. **Phase 3**: Deploy organization-wide

### Centralized Deployment API

For automated deployment:
```powershell
# Connect to Exchange Online
Connect-ExchangeOnline

# Deploy add-in
New-App -FileData ([System.IO.File]::ReadAllBytes("manifest-production.xml")) -DefaultStateForUser Enabled -Enabled $true
```

## Compliance and Governance

### Data Residency
- Add-in files hosted in your chosen region
- No external data transmission
- All processing happens locally

### Compliance
- GDPR compliant (local processing only)
- HIPAA compatible (no PHI transmitted)
- SOC 2 considerations met

### Audit
- Track deployment date
- Record users deployed to
- Log any updates
- Maintain version history

## Cost Considerations

### Hosting Costs
- **Azure Blob Storage**: ~$0.02/GB/month
- **Azure App Service**: Starting at $13/month
- **AWS S3**: ~$0.023/GB/month

### Licensing
- No additional Office 365 license required
- Users must have Outlook access
- Admin must have Global Admin or Exchange Admin role

## Support Resources

- **Microsoft 365 Admin Center**: [admin.cloud.microsoft.com](https://admin.cloud.microsoft.com)
- **Office Add-ins Documentation**: [docs.microsoft.com/office/dev/add-ins](https://docs.microsoft.com/office/dev/add-ins)
- **Azure Hosting Guide**: [docs.microsoft.com/azure/storage/blobs/storage-blob-static-website](https://docs.microsoft.com/azure/storage/blobs/storage-blob-static-website)

## Summary Checklist

- [ ] Production files built and uploaded
- [ ] Hosting configured with HTTPS
- [ ] Manifest file updated with production URLs
- [ ] Manifest validated successfully
- [ ] Uploaded to Microsoft 365 Admin Center
- [ ] Deployment configured (users/groups selected)
- [ ] App deployed successfully
- [ ] Tested in Outlook Desktop
- [ ] Tested in Outlook on the Web
- [ ] Users notified
- [ ] Training materials provided
- [ ] Support process established
- [ ] Monitoring in place

---

**Deployment Guide Version**: 1.0  
**Last Updated**: October 4, 2025  
**Contact**: [Your IT Support Email]
