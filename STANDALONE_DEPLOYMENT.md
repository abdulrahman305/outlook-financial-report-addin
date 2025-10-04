# Standalone Manifest Deployment Guide

## Overview
The `manifest-standalone.xml` file is specifically designed for deployment through **admin.cloud.microsoft** (Microsoft 365 Admin Center) as a fully self-contained Outlook Add-in.

## ✅ Validation Status
**Status:** VALIDATED ✓  
**Validator:** Microsoft Office Add-in Manifest Validator  
**Errors:** 0  
**Platform Support:** Full compatibility

### Supported Platforms
- ✅ Outlook 2016 or later on Windows
- ✅ Outlook 2019 or later on Windows
- ✅ Outlook 2016 or later on Mac
- ✅ Outlook 2019 or later on Mac
- ✅ Outlook on the web
- ✅ Outlook on Windows (Microsoft 365)
- ✅ Outlook on Mac (Microsoft 365)

## Key Features

### 1. **Standalone Architecture**
- **No External Services Required**: All processing happens locally within the Outlook environment
- **No Cloud Dependencies**: Data never leaves the user's mailbox
- **Self-Contained PDF Processing**: Uses embedded pdfjs-dist library for local PDF parsing
- **In-Memory Data Storage**: Financial data is processed and stored temporarily in browser memory

### 2. **Core Capabilities**
- **Inbox Scanning**: Comprehensively scans entire Outlook inbox via Office.js Mailbox API
- **Email Data Extraction**: Regex-based extraction of financial amounts and transaction details
- **PDF Attachment Parsing**: Local PDF text extraction without external API calls
- **Transaction Categorization**: Automatic classification into Assets, Liabilities, Equity, Revenue, Expenses
- **Report Generation**: Creates quarterly (Q1-Q4), semi-annual (H1-H2), and annual financial statements
- **Professional Formatting**: Balance Sheet and Income Statement with Fluent UI styling

### 3. **Security & Privacy**
- **ReadWriteMailbox Permission**: Only requests necessary mailbox access
- **No Data Transmission**: Zero external API calls or data uploads
- **HTTPS Enforcement**: All resources loaded over secure connections
- **Local Processing**: All financial calculations done client-side

## Deployment Steps

### Prerequisites
1. **Microsoft 365 Admin Access**: You must have admin privileges
2. **Production Hosting**: Build files must be hosted on an HTTPS server
3. **Valid Domain**: Production URLs must use HTTPS protocol

### Step 1: Build Production Files
```bash
# Navigate to project directory
cd /workspaces/outlook-financial-report-addin

# Install dependencies (if not already done)
npm install

# Build production files
npm run build
```

This creates optimized files in the `dist/` folder:
- `taskpane.html`
- `taskpane.js`
- `taskpane.css`
- `commands.html`
- `commands.js`

### Step 2: Host Production Files

**Recommended Hosting Option: Azure Blob Storage**
```bash
# Install Azure CLI (if not already installed)
curl -sL https://aka.ms/InstallAzureCLIDeb | sudo bash

# Login to Azure
az login

# Create storage account
az storage account create \
  --name financialreportaddin \
  --resource-group your-resource-group \
  --location eastus \
  --sku Standard_LRS

# Enable static website hosting
az storage blob service-properties update \
  --account-name financialreportaddin \
  --static-website \
  --index-document taskpane.html

# Upload files
az storage blob upload-batch \
  --account-name financialreportaddin \
  --source ./dist \
  --destination '$web'
```

**Alternative Hosting Options:**
- Azure App Service
- AWS S3 with CloudFront
- GitHub Pages (if repository is public)
- Any HTTPS web server

### Step 3: Update Manifest URLs

Open `manifest-standalone.xml` and replace all `https://localhost:3000` references with your production URL:

```xml
<!-- Update these lines (there are 8 occurrences): -->

<!-- Icon URLs (lines 23-25) -->
<IconUrl DefaultValue="https://YOUR-DOMAIN.com/assets/icon-32.png"/>
<HighResolutionIconUrl DefaultValue="https://YOUR-DOMAIN.com/assets/icon-80.png"/>

<!-- Source Location (line 49) -->
<SourceLocation DefaultValue="https://YOUR-DOMAIN.com/taskpane.html"/>

<!-- Resources section (lines 145-150) -->
<bt:Image id="Icon.16x16" DefaultValue="https://YOUR-DOMAIN.com/assets/icon-16.png"/>
<bt:Image id="Icon.32x32" DefaultValue="https://YOUR-DOMAIN.com/assets/icon-32.png"/>
<bt:Image id="Icon.80x80" DefaultValue="https://YOUR-DOMAIN.com/assets/icon-80.png"/>
<bt:Url id="Commands.Url" DefaultValue="https://YOUR-DOMAIN.com/commands.html"/>
<bt:Url id="Taskpane.Url" DefaultValue="https://YOUR-DOMAIN.com/taskpane.html"/>
```

**Example:**
```xml
<!-- Before -->
<IconUrl DefaultValue="https://localhost:3000/assets/icon-32.png"/>

<!-- After -->
<IconUrl DefaultValue="https://financialreportaddin.blob.core.windows.net/assets/icon-32.png"/>
```

### Step 4: Validate Updated Manifest

```bash
# Validate the updated manifest
npx office-addin-manifest validate manifest-standalone.xml
```

Ensure you get:
- ✅ "The manifest is valid"
- ✅ Zero errors
- ✅ All platform support confirmed

### Step 5: Upload to Microsoft 365 Admin Center

1. **Navigate to Admin Center**:
   - Go to https://admin.cloud.microsoft (or https://admin.microsoft.com)
   - Sign in with your admin credentials

2. **Access Integrated Apps**:
   - In the left navigation, expand **Settings**
   - Click **Integrated apps**

3. **Upload Custom App**:
   - Click **Upload custom apps** button (top right)
   - Select **Provide link to manifest file** or **Upload manifest file**

4. **If Using Direct Upload**:
   - Click **Upload manifest file (.xml format)**
   - Select `manifest-standalone.xml`
   - Click **Upload**

5. **If Using URL**:
   - Click **Provide link to manifest file**
   - Enter: `https://YOUR-DOMAIN.com/manifest-standalone.xml`
   - Click **Validate**
   - Click **Upload**

6. **Configure Deployment**:
   - **App name**: Financial Report Generator - Standalone
   - **Assign users**: 
     - Choose **Entire organization** for all users
     - Or select specific **Users/Groups**
   - Click **Next**

7. **Accept Permissions**:
   - Review requested permissions: **ReadWriteMailbox**
   - This allows the add-in to read and scan mailbox items
   - Click **Accept permissions**

8. **Complete Deployment**:
   - Review deployment summary
   - Click **Finish deployment**

### Step 6: Verify Deployment

1. **Check Deployment Status**:
   - In Integrated apps, find "Financial Report Generator - Standalone"
   - Status should show "Deployed" or "Available to users"
   - It may take 6-24 hours for deployment to propagate

2. **Test in Outlook**:
   - Open Outlook (Web, Desktop, or Mac)
   - Open any email message
   - Look for **Financial Reports** group in the ribbon
   - Click **Generate Report** button
   - Task pane should open showing the Financial Report Generator

3. **Verify Functionality**:
   - Click "Scan Inbox for Financial Data"
   - Confirm emails are being scanned
   - Verify PDF attachments are parsed
   - Generate a quarterly report
   - Check that Balance Sheet and Income Statement appear

## Manifest Configuration Details

### App ID
```xml
<Id>56dbd71b-8bab-41f4-868d-cc806382f39b</Id>
```
**Important**: This GUID is unique to your add-in. Keep it consistent across updates.

### Version
```xml
<Version>1.0.0.0</Version>
```
Increment for each update:
- Major updates: 2.0.0.0
- Minor updates: 1.1.0.0
- Patches: 1.0.1.0

### Permissions
```xml
<Permissions>ReadWriteMailbox</Permissions>
```
Required for:
- Reading email messages
- Accessing email bodies
- Downloading attachments
- Scanning inbox via REST API

### API Requirements
```xml
<Set Name="Mailbox" MinVersion="1.3"/>
```
Ensures compatibility with Outlook 2016+

### Ribbon Buttons
1. **Generate Report** - Opens task pane for full report generation
2. **Quick Scan** - Executes quick scan of current email

## Troubleshooting

### Issue: Manifest Upload Fails
**Solution**: 
- Ensure all URLs use HTTPS (not HTTP)
- Verify all referenced files exist at specified URLs
- Check that icons are PNG format (16x16, 32x32, 80x80)
- Run validation: `npx office-addin-manifest validate manifest-standalone.xml`

### Issue: Add-in Not Appearing in Outlook
**Solution**:
- Wait 6-24 hours for deployment propagation
- Clear Outlook cache: File → Office Account → Sign out → Sign in
- Verify user is in assigned deployment group
- Check Admin Center deployment status

### Issue: Add-in Loads But Shows Errors
**Solution**:
- Verify all production files are accessible via HTTPS
- Check browser console for network errors (F12)
- Ensure CORS headers allow Outlook domains
- Verify pdfjs-dist library is bundled correctly

### Issue: PDF Parsing Not Working
**Solution**:
- Confirm webpack build included pdfjs-dist
- Check that PDF.js worker is loaded
- Verify PDFs are base64 encoded correctly
- Test with different PDF formats

### Issue: Inbox Scanning Returns No Results
**Solution**:
- Verify ReadWriteMailbox permission granted
- Check that user has financial emails in inbox
- Confirm Office.js API version compatibility
- Test REST API token generation

## Updating the Add-in

1. **Make Code Changes**:
   ```bash
   # Edit source files in src/
   vim src/taskpane/taskpane.js
   ```

2. **Rebuild**:
   ```bash
   npm run build
   ```

3. **Upload New Files**:
   ```bash
   # Upload to hosting (e.g., Azure)
   az storage blob upload-batch \
     --account-name financialreportaddin \
     --source ./dist \
     --destination '$web' \
     --overwrite
   ```

4. **Update Version in Manifest**:
   ```xml
   <Version>1.0.1.0</Version>
   ```

5. **Re-upload Manifest**:
   - Go to Microsoft 365 Admin Center
   - Settings → Integrated apps
   - Find "Financial Report Generator - Standalone"
   - Click ⋯ (three dots) → **Update**
   - Upload updated `manifest-standalone.xml`

6. **Users Get Update**:
   - Updates propagate within 6-24 hours
   - Users may need to restart Outlook
   - No uninstall/reinstall required

## Security Considerations

### Data Privacy
- ✅ **No External Transmission**: All data stays in user's mailbox
- ✅ **No Logging**: No financial data logged or stored permanently
- ✅ **In-Memory Only**: Extracted data cleared when task pane closes
- ✅ **No Analytics**: No telemetry or usage tracking

### Compliance
- ✅ **GDPR Compliant**: No personal data collection
- ✅ **SOX Compatible**: Audit trail via Outlook's own logging
- ✅ **HIPAA Friendly**: No PHI transmission or storage
- ✅ **Industry Agnostic**: No sector-specific data handling

### Access Control
- ✅ **User-Level Security**: Inherits Outlook's authentication
- ✅ **Admin Deployment**: Centralized control via Microsoft 365 Admin
- ✅ **Permission Scoped**: Only requests necessary mailbox access
- ✅ **No Backend**: No server-side vulnerabilities

## Support & Resources

### Official Documentation
- [Office Add-ins Documentation](https://docs.microsoft.com/office/dev/add-ins/)
- [Outlook Add-in API](https://docs.microsoft.com/office/dev/add-ins/reference/overview/outlook-add-in-reference)
- [Microsoft 365 Admin Center](https://docs.microsoft.com/microsoft-365/admin/)

### Project Resources
- **README**: General overview and local development
- **QUICKSTART**: 5-minute setup guide
- **IMPLEMENTATION**: Technical architecture details
- **MANIFEST_VALIDATION**: Validation procedures

### Validation Commands
```bash
# Validate standalone manifest
npx office-addin-manifest validate manifest-standalone.xml

# Build production files
npm run build

# Start development server (for testing)
npm start
```

## License
This add-in is licensed under the MIT License. See LICENSE file for details.

## Version History
- **1.0.0.0** (2025-10-04): Initial standalone release for admin.cloud.microsoft deployment
  - Full inbox scanning capability
  - PDF attachment parsing
  - Quarterly, semi-annual, and annual report generation
  - Balance Sheet and Income Statement rendering
  - Zero external dependencies
