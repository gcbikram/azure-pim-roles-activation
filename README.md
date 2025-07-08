# Azure PIM Roles Activation Tool

A PowerShell script for programmatically activating Privileged Identity Management (PIM) roles in both Entra ID (Azure AD) and Azure Resources. This tool simplifies the process of discovering and activating eligible PIM roles through an interactive command-line interface.

## Features

- **Dual Platform Support**: Works with both Entra ID and Azure Resource PIM roles
- **Interactive Role Selection**: Browse and select roles through a user-friendly menu
- **Batch Activation**: Activate multiple roles simultaneously or all eligible roles at once
- **Status Checking**: Automatically checks if roles are already active to avoid duplicate activations
- **Custom Justification**: Provide custom justification for role activations
- **Error Handling**: Comprehensive error handling with detailed feedback
- **Connection Management**: Automatic connection handling for Microsoft Graph and Azure PowerShell

## Prerequisites

### Required PowerShell Modules

1. **Microsoft.Graph**: For Entra ID PIM operations
   ```powershell
   Install-Module Microsoft.Graph -Scope CurrentUser
   ```

2. **Az**: For Azure Resource PIM operations
   ```powershell
   Install-Module Az -Scope CurrentUser
   ```

### Required Permissions

- **Entra ID**: User must have eligible PIM role assignments
- **Azure Resources**: User must have eligible PIM role assignments on Azure resources
- **Microsoft Graph**: Appropriate permissions to read role eligibility and create role assignment requests

## Usage

### Basic Usage

1. Run the script:
   ```powershell
   .\activate-pim-roles.ps1
   ```

2. The script will:
   - Connect to Microsoft Graph and Azure (if available)
   - Discover all eligible PIM roles
   - Display an interactive menu with available roles
   - Prompt for role selection and justification

### Interactive Options

- **Select specific roles**: Enter role numbers separated by commas (e.g., `1,3,5`)
- **Activate all roles**: Enter `ALL` to activate all eligible roles
- **Quit**: Enter `Q` to exit without making changes

### Example Output

```
Available PIM Roles:
===========================================
1. Global Administrator
   Type: Entra ID
   Scope: / (Directory)

2. Contributor
   Type: Azure Resource
   Scope: Sub: My Subscription (Subscription)

3. Security Administrator
   Type: Entra ID
   Scope: / (Directory)

Options:
  Enter role number(s) (e.g., 1,3,5 for multiple roles)
  Enter 'ALL' to activate all roles
  Enter 'Q' to quit

Please select your option: 1,2
Enter justification for role activation (press Enter for default): Monthly security review
```

## Script Functions

### Core Functions

- **`Initialize-Connections`**: Establishes connections to Microsoft Graph and Azure PowerShell
- **`Get-EntraIDEligibleRoles`**: Retrieves eligible Entra ID PIM roles
- **`Get-AzureResourceEligibleRoles`**: Retrieves eligible Azure Resource PIM roles
- **`Show-AvailableRoles`**: Displays roles in a formatted table
- **`Test-RoleActiveStatus`**: Checks if a role is already active
- **`Invoke-RoleActivation`**: Activates a single role with proper error handling

### Role Activation Details

- **Duration**: Roles are activated for 8 hours by default
- **Justification**: Default justification is "Administrative work requirement" if none provided
- **Scope Support**: Handles Directory, Resource Group, Subscription, and Resource-level scopes

## Technical Details

### Entra ID Role Activation

Uses Microsoft Graph PowerShell SDK:
- `Get-MgRoleManagementDirectoryRoleEligibilitySchedule` - Discover eligible roles
- `New-MgRoleManagementDirectoryRoleAssignmentScheduleRequest` - Activate roles
- `Get-MgRoleManagementDirectoryRoleAssignmentScheduleInstance` - Check active status

### Azure Resource Role Activation

Uses Azure PowerShell and REST API:
- ARM API endpoint: `https://management.azure.com/providers/Microsoft.Authorization/`
- `New-AzRoleAssignmentScheduleRequest` - Activate Azure resource roles
- REST API calls for eligibility discovery and status checking

## Troubleshooting

### Common Issues

1. **Connection Failures**
   - Ensure you have appropriate permissions
   - Try running `Connect-MgGraph` and `Connect-AzAccount` manually first

2. **No Roles Found**
   - Verify you have eligible PIM role assignments
   - Check that PIM is properly configured in your tenant

3. **Activation Failures**
   - Ensure the role isn't already active
   - Verify you have permission to activate the specific role
   - Check that the activation request doesn't violate any approval requirements

### Error Messages

- **"Failed to establish Microsoft Graph connection"**: Authentication issue with Microsoft Graph
- **"Warning: Azure PowerShell not connected"**: Script will continue with Entra ID roles only
- **"Role already active"**: Role is currently activated, no action needed

## Security Considerations

- The script requires privileged access to discover and activate PIM roles
- Always provide meaningful justifications for role activations
- Role activations are logged in Azure AD audit logs
- Activated roles automatically expire after 8 hours

## References

- [Programmatically activate my Entra ID assigned role](https://learn.microsoft.com/en-us/answers/questions/1879083/programmatically-activate-my-entra-id-assigned-rol)
- [Activate PIM role for Azure resources via REST PowerShell](https://stackoverflow.com/questions/77252524/activate-pim-role-for-azure-resources-via-rest-powershell)
- [Get-MgRoleManagementDirectoryRoleEligibilitySchedule Documentation](https://learn.microsoft.com/en-us/powershell/module/microsoft.graph.identity.governance/get-mgrolemanagementdirectoryroleeligibilityschedule?view=graph-powershell-1.0)

## Contributing

Feel free to submit issues and enhancement requests. When contributing:

1. Test changes thoroughly in a development environment
2. Ensure error handling is maintained
3. Update documentation for any new features
4. Follow PowerShell best practices

## License

This project is provided as-is for educational and administrative purposes. Use at your own risk and ensure compliance with your organization's security policies.

# Note
Code was generated by Claude and Gemini with me refactoring here and there and instructing it, enjoy!!