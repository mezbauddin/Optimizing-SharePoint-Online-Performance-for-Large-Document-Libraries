# Optimizing SharePoint Online Performance for Large Document Libraries 

A comprehensive PowerShell automation toolkit for bulk metadata updates in SharePoint Online document libraries. This project provides multiple approaches to efficiently update file properties across large document collections using both PnP PowerShell and Microsoft Graph API.

## Overview

This project contains three PowerShell scripts designed for different scenarios and requirements when performing bulk metadata updates in SharePoint Online. Each script offers unique advantages depending on your specific needs, from simple operations to enterprise-grade batch processing with retry logic.

## Features

- **Multiple Update Approaches**: Choose between PnP PowerShell and Microsoft Graph API
- **Batch Processing**: Handle hundreds of files efficiently
- **Retry Logic**: Built-in error recovery for network issues
- **Progress Tracking**: Real-time feedback on processing status
- **Comprehensive Logging**: Detailed operation logs for troubleshooting
- **CSV-Driven**: Simple spreadsheet format for defining updates
- **Enterprise Ready**: Tested for production environments

## Scripts Included

### 1. Sharepoint Files Bulk Status Update.ps1
**Basic PnP PowerShell approach - Best for small to medium datasets**

- Simple, straightforward implementation
- Interactive authentication using Microsoft Entra ID
- Sequential processing with immediate feedback
- Perfect for testing and smaller operations (< 100 files)

### 2. Sharepoint_Files_Bulk_Update_WithRetry.ps1
**Advanced PnP PowerShell with batching and retry logic**

- Batch processing for improved performance
- Configurable retry mechanism for reliability
- Comprehensive error handling and logging
- Suitable for larger datasets (100-1000 files)
- Progress tracking and detailed reporting

### 3. Bulk Update with Graph API.ps1
**Microsoft Graph API implementation with advanced batching**

- Uses Microsoft Graph API for maximum performance
- OAuth2 client credentials authentication
- Advanced batch processing with Graph $batch endpoint
- Optimized for large-scale operations (1000+ files)
- Enterprise-grade reliability and performance

## Prerequisites

### Software Requirements
- PowerShell 5.1 or later
- PnP.PowerShell module (for scripts 1 & 2)
- SharePoint Online access
- Microsoft Entra ID app registration

### PowerShell Module Installation
```powershell
# Install PnP PowerShell module
Install-Module -Name PnP.PowerShell -Force -AllowClobber

# Set execution policy if needed
Set-ExecutionPolicy -ExecutionPolicy RemoteSigned -Scope CurrentUser
```

## Setup

### Microsoft Entra ID App Registration

#### For PnP PowerShell Scripts (1 & 2)
Create an Entra ID app registration using the PnP PowerShell built-in cmdlet:

```powershell
Register-PnPAzureADApp -ApplicationName "SharePoint Bulk Update" -Tenant yourtenant.onmicrosoft.com -Store CurrentUser
```

This automatically:
- Creates a new Entra ID application
- Generates and installs a self-signed certificate
- Requests necessary SharePoint permissions
- Opens browser for admin consent

#### For Graph API Script (3)
1. Go to [Microsoft Entra admin center](https://entra.microsoft.com/)
2. Navigate to **Applications > App registrations**
3. Click **New registration**
4. Configure:
   - Name: "SharePoint Graph API Bulk Update"
   - Supported account types: Single tenant
   - Redirect URI: Not needed for client credentials flow

5. After creation:
   - Note the **Application (client) ID**
   - Create a **client secret** under **Certificates & secrets**
   - Add API permissions under **API permissions**:
     - Microsoft Graph > Application permissions > Sites.FullControl.All
     - Grant admin consent

### Required Permissions

**For PnP PowerShell Scripts:**
- Sites.FullControl.All
- Group.ReadWrite.All
- User.Read.All

**For Graph API Script:**
- Sites.FullControl.All (Microsoft Graph Application Permission)

## Configuration

### Script 1 & 2 Configuration
Update these variables in the script files:

```powershell
$SiteUrl      = "https://yourtenant.sharepoint.com/sites/yoursite"
$LibraryName  = "Practice Statistics"  # Your document library name
$appId        = "your-app-id-here"     # From Entra ID app registration
$CsvPath      = "./documents-to-update.csv"
```

### Script 3 Configuration
Update these variables in `Bulk Update with Graph API.ps1`:

```powershell
$tenantId     = "<your-tenant-id>"
$clientId     = "<your-app-client-id>"
$clientSecret = "<your-app-client-secret>"
$siteHost     = "yourtenant.sharepoint.com"
$sitePath     = "/sites/yoursite"
$listName     = "Practice Statistics"
```

## Usage

### 1. Prepare Your Data File

Create `documents-to-update.csv` with this structure:

```csv
FileName,NewStatus
DummyFile_001.docx,Reviewed
DummyFile_002.docx,Reviewed
DummyFile_003.docx,Archived
DummyFile_004.docx,Reviewed
```

**Important Notes:**
- File names must match exactly as they appear in SharePoint (case-sensitive)
- The `Status` field must exist in your SharePoint library
- Ensure custom columns are created in SharePoint before running updates

### 2. Choose Your Script

#### For Small Operations (< 100 files):
```powershell
.\Sharepoint Files Bulk Status Update.ps1
```

#### For Medium Operations (100-1000 files):
```powershell
.\Sharepoint_Files_Bulk_Update_WithRetry.ps1
```

#### For Large Operations (1000+ files):
```powershell
.\Bulk Update with Graph API.ps1
```

### 3. Monitor Progress

Each script provides real-time feedback:

```
Processing 4 items...
[1/4] Updated: DummyFile_001.docx -> Status: Reviewed
[2/4] Updated: DummyFile_002.docx -> Status: Reviewed
[3/4] Updated: DummyFile_003.docx -> Status: Archived
[4/4] Updated: DummyFile_004.docx -> Status: Reviewed
Bulk update complete.
```

## Advanced Configuration

### Multiple Field Updates

Modify the CSV and script to update multiple fields:

**CSV Format:**
```csv
FileName,NewStatus,Department,ReviewDate
DummyFile_001.docx,Reviewed,IT,2024-12-29
DummyFile_002.docx,Approved,Marketing,2024-12-30
```

**Script Modification (PnP PowerShell):**
```powershell
Set-PnPListItem -List $List -Identity $ListItem.Id -Values @{ 
    "Status" = $Row.NewStatus
    "Department" = $Row.Department
    "ReviewDate" = $Row.ReviewDate
}
```

### Batch Size Tuning

For the retry script, adjust batch size based on your environment:

```powershell
# Small batches for reliability
.\Sharepoint_Files_Bulk_Update_WithRetry.ps1 -BatchSize 5

# Larger batches for speed (if network is stable)
.\Sharepoint_Files_Bulk_Update_WithRetry.ps1 -BatchSize 20
```

## Troubleshooting

### Common Issues

| Issue | Cause | Solution |
|-------|-------|----------|
| "Authentication failed" | Invalid app registration | Verify client ID and ensure admin consent granted |
| "File not found: filename.ext" | Filename mismatch | Check exact filename in SharePoint (case-sensitive) |
| "Access denied" | Insufficient permissions | Ensure app has Sites.FullControl.All permission |
| "CSV file empty or invalid" | File path or format issue | Verify CSV file exists and has proper structure |
| "The current connection holds no SharePoint context" | Connection dropped | Re-run script to re-authenticate |
| "Throttling detected" | Too many API calls | Use retry script or reduce batch size |

### Error Logging

Scripts 2 and 3 generate detailed log files:
- **Script 2**: `BulkUpdate_yyyyMMdd_HHmmss.log`
- **Script 3**: `GraphAPI_BulkUpdate_yyyyMMdd_HHmmss.log`

Review these logs for detailed error information and troubleshooting.

### Performance Optimization

**Script Selection Guidelines:**
- **< 100 files**: Use Script 1 (simple, fast setup)
- **100-1000 files**: Use Script 2 (balanced performance and reliability)
- **1000+ files**: Use Script 3 (maximum performance with Graph API)

**Network Considerations:**
- Run during off-peak hours for large operations
- Consider SharePoint Online throttling limits
- Use smaller batch sizes on unstable connections

## File Structure

```
Sharepoint Bulk Update project/
├── Bulk Update with Graph API.ps1              # Graph API implementation
├── Sharepoint Files Bulk Status Update.ps1     # Basic PnP PowerShell script
├── Sharepoint_Files_Bulk_Update_WithRetry.ps1  # Advanced PnP PowerShell with retry
├── documents-to-update.csv                     # Sample data file
└── README.md                                   # This documentation
```

## Security Best Practices

- **Client Secrets**: Store securely, rotate regularly
- **Least Privilege**: Grant minimum required permissions
- **Certificate Authentication**: Prefer certificate over client secret when possible
- **App Registration**: Use dedicated service accounts for production
- **Audit Trail**: Monitor app usage through Entra ID audit logs
- **Network Security**: Run from trusted networks only

## Best Practices

### Before Running
- **Test First**: Run with 2-3 files before bulk operations
- **Backup Metadata**: Export current field values before changes
- **Verify Permissions**: Confirm app and user permissions
- **Validate CSV**: Check file names and field values

### During Operations
- **Monitor Progress**: Watch for error messages
- **Check Logs**: Review log files for issues
- **Network Stability**: Ensure stable internet connection
- **Resource Usage**: Monitor system performance

### After Completion
- **Verify Results**: Spot-check updated files in SharePoint
- **Archive Logs**: Save log files for future reference
- **Document Changes**: Record what was updated and when
- **Clean Up**: Remove temporary files and credentials

## Version History

- **v1.0** - Initial release with basic PnP PowerShell script
- **v1.1** - Added retry logic and batch processing script
- **v1.2** - Introduced Microsoft Graph API implementation
- **v1.3** - Updated for Microsoft Entra ID terminology and latest APIs

## Support and Troubleshooting

1. **Check Prerequisites**: Ensure all requirements are met
2. **Review Logs**: Examine detailed log files for specific errors
3. **Test Connectivity**: Verify SharePoint Online access
4. **Validate App Registration**: Confirm permissions and consent
5. **Start Small**: Test with minimal dataset to isolate issues

## Contributing

When modifying these scripts:
- Follow PowerShell best practices
- Add appropriate error handling
- Update logging for new operations
- Test thoroughly before production use
- Document any changes or customizations

---

*This project is designed for enterprise SharePoint Online environments. Always follow your organization's change management and security procedures when deploying to production.*
