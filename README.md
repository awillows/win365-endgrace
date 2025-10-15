# Windows 365 Grace Period Manager

A PowerShell GUI application for managing Windows 365 Business Cloud PCs with a focus on ending grace periods. This tool provides a visual interface to view and manage Cloud PCs, with special highlighting for those in grace period.

## Features

- 📊 **Visual Cloud PC List**: Display all Windows 365 Business Cloud PCs in your tenant
- 🔵 **Grace Period Highlighting**: Cloud PCs in grace period are highlighted in blue with bold text
- 🖱️ **Context Menu Actions**: Right-click to access deprovision options
- ⚠️ **Safety Warnings**: Confirmation dialog before deprovisioning
- 🔒 **Smart Controls**: Deprovision option is disabled for Cloud PCs not in grace period
- 🔄 **Refresh Capability**: Update the list at any time with the Refresh button

## Display Columns

The application shows the following information for each Cloud PC:

- **Cloud PC Name**: The display name of the Cloud PC
- **Assigned User**: The user principal name (UPN) of the assigned user
- **Status**: Current status (e.g., provisioned, inGracePeriod, deprovisioning)
- **Grace End Date**: When the grace period will end (if applicable)

## Prerequisites

### Required Modules

Install the Microsoft Graph PowerShell SDK:

```powershell
Install-Module Microsoft.Graph -Scope CurrentUser
```

Or install specific modules:

```powershell
Install-Module Microsoft.Graph.Authentication -Scope CurrentUser
Install-Module Microsoft.Graph.DeviceManagement.Actions -Scope CurrentUser
```

### Required Permissions

The application requires the following Microsoft Graph API permissions:

- `CloudPC.ReadWrite.All` - To read and manage Cloud PCs
- `CloudPC.Read.All` - To read Cloud PC information

### Azure AD Requirements

- An Azure AD account with permissions to manage Windows 365 Cloud PCs
- Typically requires one of the following roles:
  - Global Administrator
  - Cloud PC Administrator
  - Intune Administrator

## Installation

1. Clone or download this repository to your local machine
2. Ensure you have PowerShell 5.1 or later installed
3. Install the required Microsoft Graph PowerShell modules (see Prerequisites)

## Usage

### Running the Application

1. Open PowerShell
2. Navigate to the directory containing `Win365GraceManager.ps1`
3. Run the script:

```powershell
.\Win365GraceManager.ps1
```

4. The application will prompt you to sign in to Microsoft Graph (first run only)
5. Grant the requested permissions when prompted

### Using the Interface

#### Viewing Cloud PCs

- All Cloud PCs in your tenant are displayed in the main list
- Cloud PCs in grace period appear in **blue text** with bold formatting
- The status bar shows total Cloud PCs and how many are in grace period

#### Deprovisioning a Cloud PC

1. **Right-click** on a Cloud PC in the list
2. Select **"Deprovision Now (End Grace Period)"** from the context menu
   - This option is only enabled for Cloud PCs in grace period
   - The option appears grayed out for other Cloud PCs
3. A warning dialog will appear asking for confirmation
4. Click **OK** to proceed or **Cancel** to abort
5. Upon successful deprovision, the list will automatically refresh

#### Refreshing the List

- Click the **"Refresh"** button in the status bar to reload Cloud PC data
- The list automatically refreshes after a successful deprovision operation

## Understanding Cloud PC Status

Common status values you may see:

- **provisioned**: Cloud PC is active and running
- **inGracePeriod**: Cloud PC is in grace period (user assignment removed, pending deprovision)
- **deprovisioning**: Cloud PC is being removed
- **provisionedWithWarnings**: Cloud PC is active but has warnings
- **notProvisioned**: Cloud PC has not been provisioned yet

## Grace Period Information

When a user license is removed or a Cloud PC is unassigned, it enters a grace period:

- **Duration**: Typically 7 days
- **Purpose**: Allows time to back up data or reassign the Cloud PC
- **During Grace**: The Cloud PC remains accessible but will be automatically deprovisioned when the grace period ends
- **End Grace Early**: Use this tool to deprovision immediately instead of waiting for automatic cleanup

## Security Considerations

⚠️ **Important**: Deprovisioning a Cloud PC is a **permanent action** that cannot be undone. All data on the Cloud PC will be lost.

### Best Practices

1. **Always verify** the Cloud PC and user before deprovisioning
2. **Ensure data backup** has been completed if needed
3. **Confirm with users** that they no longer need access to the Cloud PC
4. **Use with caution** in production environments
5. **Test first** in a non-production tenant if possible

## Troubleshooting

### Authentication Issues

**Problem**: Cannot connect to Microsoft Graph

**Solutions**:
- Ensure you have the Microsoft Graph PowerShell SDK installed
- Check that your account has the necessary permissions
- Try disconnecting and reconnecting: `Disconnect-MgGraph` then run the script again
- Verify you're using the correct tenant

### No Cloud PCs Displayed

**Problem**: The list is empty but you know Cloud PCs exist

**Solutions**:
- Click the Refresh button
- Verify your account has permissions to view Cloud PCs
- Check that you're connected to the correct tenant
- Ensure Cloud PCs are Windows 365 Business (not Enterprise)

### Deprovision Option Grayed Out

**Problem**: Cannot select "Deprovision Now" option

**Expected Behavior**: This is by design. The option is only enabled for Cloud PCs in grace period. If a Cloud PC is not in grace, you cannot use this tool to deprovision it.

### API Errors

**Problem**: Error messages when trying to deprovision

**Solutions**:
- Verify your account has `CloudPC.ReadWrite.All` permissions
- Check your internet connection
- Ensure the Microsoft Graph API is accessible from your network
- Try refreshing your authentication token by restarting the application

## API Endpoints Used

This application uses the following Microsoft Graph API endpoints:

- `GET /deviceManagement/virtualEndpoint/cloudPCs` - List all Cloud PCs
- `POST /deviceManagement/virtualEndpoint/cloudPCs/{id}/endGracePeriod` - End grace period and deprovision

## Technical Details

### Requirements

- **PowerShell Version**: 5.1 or later
- **Modules**: Microsoft.Graph.Authentication, Microsoft.Graph.DeviceManagement.Actions
- **GUI Framework**: WPF (Windows Presentation Foundation)
- **API**: Microsoft Graph Beta endpoint

### Architecture

The application is built using:
- **WPF/XAML** for the user interface
- **Microsoft Graph PowerShell SDK** for API communication
- **ListView control** with custom styling for data presentation
- **Context menus** for action triggers

## License

This tool is provided as-is for managing Windows 365 Cloud PCs. Use at your own risk and always ensure you have proper backups before deprovisioning any Cloud PC.

## Support

For issues related to:
- **Windows 365**: Contact Microsoft Support
- **Microsoft Graph API**: Check the [Microsoft Graph documentation](https://docs.microsoft.com/graph/)
- **This Tool**: Review the troubleshooting section or modify the script as needed

## Version History

### Version 1.0 (October 15, 2025)
- Initial release
- Cloud PC listing with grace period highlighting
- Context menu deprovision functionality
- Confirmation dialogs and error handling
- Automatic list refresh after operations

## Contributing

Feel free to modify and enhance this tool for your specific needs. Some ideas for future enhancements:

- Export Cloud PC list to CSV
- Filter/search functionality
- Additional Cloud PC actions (restart, rename, etc.)
- Scheduled grace period reports
- Email notifications for Cloud PCs entering grace period

## Additional Resources

- [Windows 365 Documentation](https://docs.microsoft.com/windows-365/)
- [Microsoft Graph API - Cloud PC](https://docs.microsoft.com/graph/api/resources/cloudpc)
- [PowerShell Gallery - Microsoft.Graph](https://www.powershellgallery.com/packages/Microsoft.Graph)
