# SharePoint Files Bulk Status Update

A PowerShell automation tool for updating metadata fields across multiple files in SharePoint Online document libraries. Built for enterprise environments requiring efficient batch processing of file properties.

## Overview

This script provides a reliable solution for bulk metadata updates in SharePoint Online, eliminating the need for manual file-by-file editing through the web interface. Designed with production environments in mind, it handles authentication, error reporting, and progress tracking automatically.

## Features

- **Batch Processing**: Update hundreds of files in a single operation
- **Secure Authentication**: Uses Azure AD app registration with interactive login
- **Progress Tracking**: Real-time feedback on processing status
- **Error Handling**: Comprehensive logging of failed operations
- **CSV-Driven**: Simple spreadsheet format for defining updates
- **Production Ready**: Tested and optimized for enterprise use

## Prerequisites

- PowerShell 5.1 or later
- PnP.PowerShell module
- SharePoint Online access
- Azure AD app registration (see Setup section)

## Installation

1. Install the PnP PowerShell module:
```powershell
Install-Module -Name PnP.PowerShell -Force -AllowClobber
```

2. Clone or download this repository to your local machine

3. Ensure your execution policy allows script execution:
```powershell
Set-ExecutionPolicy -ExecutionPolicy RemoteSigned -Scope CurrentUser
```

## Setup

### Azure AD App Registration

Create an Azure AD app registration using the PnP PowerShell built-in cmdlet:

```powershell
Register-PnPAzureADApp -ApplicationName "SharePoint Bulk Update" -Tenant yourtenant.onmicrosoft.com -Store CurrentUser
```

This command will:
- Create a new Azure AD application
- Generate and install a self-signed certificate
- Request necessary SharePoint permissions
- Open a browser for admin consent

**Required Permissions:**
- Sites.FullControl.All
- Group.ReadWrite.All
- User.Read.All

### Script Configuration

Update the following variables in `Sharepoint Files Bulk Status Update.ps1`:

```powershell
$SiteUrl      = "https://yourtenant.sharepoint.com/sites/yoursite"
$LibraryName  = "Documents"  # Or your target library name
$appId        = "your-app-id-here"
```

## Usage

### 1. Prepare Your Data

Create a CSV file named `documents-to-update.csv` with the following structure:

```csv
FileName,NewStatus
Document1.docx,Reviewed
Presentation.pptx,Approved
Spreadsheet.xlsx,Draft
```

**Important Notes:**
- File names must match exactly as they appear in SharePoint (case-sensitive)
- The `Status` field must exist in your SharePoint library
- Add custom columns to SharePoint before running updates

### 2. Run the Script

Execute the PowerShell script:

```powershell
.\Sharepoint Files Bulk Status Update.ps1
```

The script will:
1. Authenticate using your Azure AD app
2. Load the CSV file
3. Process each file sequentially
4. Display progress and results
5. Disconnect automatically when complete

### 3. Monitor Progress

The script provides real-time feedback:

```
Processing 3 items...
[1/3] Updated: Document1.docx -> Status: Reviewed
[2/3] Updated: Presentation.pptx -> Status: Approved
[3/3] Updated: Spreadsheet.xlsx -> Status: Draft
Bulk update complete.
```

## Customization

### Updating Different Fields

To update fields other than "Status", modify the script:

```powershell
Set-PnPListItem -List $List -Identity $ListItem.Id -Values @{ 
    "YourFieldName" = $NewValue 
}
```

### Adding Multiple Fields

Update multiple fields simultaneously:

```powershell
Set-PnPListItem -List $List -Identity $ListItem.Id -Values @{ 
    "Status" = $Row.NewStatus
    "Department" = $Row.Department
    "ReviewDate" = $Row.ReviewDate
}
```

### CSV Format for Multiple Fields

```csv
FileName,NewStatus,Department,ReviewDate
Document1.docx,Reviewed,IT,2025-01-15
Presentation.pptx,Approved,Marketing,2025-01-20
```

## Troubleshooting

### Common Issues

**Authentication Failures:**
- Verify your app ID is correct
- Ensure admin consent was granted
- Check that certificates are properly installed

**File Not Found Errors:**
- Confirm file names match exactly (including extensions)
- Verify files exist in the specified library
- Check for special characters or encoding issues

**Permission Denied:**
- Ensure your app has sufficient SharePoint permissions
- Verify you have access to the target site and library
- Check that the Status field exists and is editable

### Error Messages

| Error | Solution |
|-------|----------|
| "CSV file empty or invalid" | Check CSV file path and format |
| "File not found: filename.ext" | Verify exact filename in SharePoint |
| "The current connection holds no SharePoint context" | Re-run authentication |

## File Structure

```
Sharepoint_bulk_update/
├── Sharepoint Files Bulk Status Update.ps1    # Main script
├── documents-to-update.csv                     # Data file
├── app-registration-details.json              # App configuration
└── README.md                                   # This file
```

## Best Practices

- **Test First**: Run with a small subset of files before bulk operations
- **Backup Data**: Export current metadata before making changes
- **Verify Permissions**: Ensure proper access rights before execution
- **Monitor Progress**: Watch for error messages during execution
- **Document Changes**: Keep records of what was updated and when

## Performance Considerations

- The script processes files sequentially for reliability
- Large batches (1000+ files) may take considerable time
- Network latency affects processing speed
- Consider running during off-peak hours for large operations

## Security Notes

- App credentials are stored securely in the Windows certificate store
- Interactive authentication prevents credential exposure
- All connections use modern authentication protocols
- No passwords or secrets are stored in plain text

## Support

For issues or questions:
1. Check the troubleshooting section above
2. Verify your SharePoint and Azure AD permissions
3. Review the PnP PowerShell documentation
4. Test with a minimal dataset to isolate problems

## Version History

- **v1.0** - Initial release with basic bulk update functionality
- **v1.1** - Added error handling and progress tracking
- **v1.2** - Simplified authentication and improved reliability

---

*This script was developed for enterprise SharePoint environments and has been tested with SharePoint Online. Use in production environments should follow your organization's change management procedures.*
