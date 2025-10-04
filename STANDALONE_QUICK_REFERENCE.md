# 🎯 Standalone Manifest - Quick Reference

## ✅ Validation Status
**File:** `manifest-standalone.xml`  
**Status:** VALIDATED ✓  
**Errors:** 0  
**Ready for:** admin.cloud.microsoft deployment

## 📋 Quick Deploy Checklist

### 1. Build Production Files
```bash
npm run build
```
Output: `dist/` folder with taskpane.html, taskpane.js, commands.html, etc.

### 2. Host Files on HTTPS Server
Example (Azure):
```bash
az storage blob upload-batch \
  --account-name YOUR-STORAGE \
  --source ./dist \
  --destination '$web'
```

### 3. Update URLs in Manifest
Replace all `https://localhost:3000` with your production URL:
- Line 23: `<IconUrl>`
- Line 24: `<HighResolutionIconUrl>`
- Line 49: `<SourceLocation>`
- Lines 145-150: Resources section (5 URLs)

**Total: 8 URLs to update**

### 4. Validate Updated Manifest
```bash
npm run validate:standalone
```
Must show: "The manifest is valid" with 0 errors

### 5. Upload to Admin Center
1. Go to https://admin.cloud.microsoft
2. Settings → Integrated apps
3. Upload custom apps
4. Select `manifest-standalone.xml`
5. Assign to users/organization
6. Accept ReadWriteMailbox permission

### 6. Wait for Deployment
- Propagation: 6-24 hours
- Users see add-in in Outlook ribbon
- Button: "Generate Report" under "Financial Reports" group

## 🔑 Key Features

### Self-Contained Architecture
- ✅ No external APIs
- ✅ No cloud services
- ✅ All processing local
- ✅ Data stays in mailbox

### Core Capabilities
- 📧 Inbox scanning via Office.js
- 📄 PDF attachment parsing (pdfjs-dist)
- 💰 Financial data extraction (regex)
- 📊 Quarterly/Semi-annual/Annual reports
- 📈 Balance Sheet + Income Statement
- 🎨 Professional Fluent UI styling

### Security
- 🔒 ReadWriteMailbox permission only
- 🔒 No data transmission
- 🔒 In-memory processing
- 🔒 HTTPS enforced

## 📱 Platform Support
- ✅ Outlook 2016+ Windows
- ✅ Outlook 2019+ Windows
- ✅ Outlook 2016+ Mac
- ✅ Outlook 2019+ Mac
- ✅ Outlook Web
- ✅ Outlook Microsoft 365

## 🛠️ Validation Commands
```bash
# Validate standalone manifest
npm run validate:standalone

# Validate all manifests
npm run validate:all

# Build for production
npm run build

# Start dev server (testing)
npm start
```

## 📚 Documentation
- **STANDALONE_DEPLOYMENT.md** - Complete deployment guide (step-by-step)
- **README.md** - Project overview
- **QUICKSTART.md** - 5-minute local setup
- **IMPLEMENTATION.md** - Technical architecture
- **MANIFEST_VALIDATION.md** - Validation procedures

## 🆔 App Identity
- **ID:** `56dbd71b-8bab-41f4-868d-cc806382f39b`
- **Version:** `1.0.0.0`
- **Name:** Financial Report Generator - Standalone
- **Permission:** ReadWriteMailbox
- **API Version:** Mailbox 1.3+

## ⚠️ Important Notes

### Before Upload
- ✅ Build files: `npm run build`
- ✅ Host on HTTPS server
- ✅ Update all 8 URLs in manifest
- ✅ Validate: `npm run validate:standalone`
- ✅ Verify all assets accessible (icons, HTML, JS)

### After Upload
- ⏱️ Wait 6-24 hours for propagation
- 👥 Check user assignments in Admin Center
- 🔄 Users may need to restart Outlook
- ✅ Verify add-in appears in Outlook ribbon

### Troubleshooting
- **Not appearing?** Check deployment status in Admin Center
- **Shows errors?** Verify all URLs accessible via HTTPS
- **PDF not parsing?** Confirm pdfjs-dist bundled in webpack
- **No scan results?** Verify ReadWriteMailbox permission granted

## 🚀 Quick Start URLs

### Hosting Options
- **Azure Blob Storage**: https://docs.microsoft.com/azure/storage/blobs/storage-blob-static-website
- **Azure App Service**: https://docs.microsoft.com/azure/app-service/
- **GitHub Pages**: https://pages.github.com/

### Admin Center
- **Microsoft 365 Admin**: https://admin.cloud.microsoft
- **Integrated Apps**: https://admin.microsoft.com/adminportal/home#/Settings/IntegratedApps

### Office Add-ins Docs
- **Overview**: https://docs.microsoft.com/office/dev/add-ins/
- **Outlook API**: https://docs.microsoft.com/office/dev/add-ins/outlook/
- **Deployment**: https://docs.microsoft.com/microsoft-365/admin/manage/test-and-deploy-microsoft-365-apps

## 📞 Support
For deployment assistance, refer to **STANDALONE_DEPLOYMENT.md** for comprehensive step-by-step instructions.

---
**Last Updated:** October 4, 2025  
**Manifest Version:** 1.0.0.0  
**Validation Status:** PASSED ✓
