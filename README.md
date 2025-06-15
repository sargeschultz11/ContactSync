# ContactSync - Enterprise Contact Synchronization for Microsoft 365

[![License: MIT](https://img.shields.io/badge/License-MIT-yellow.svg)](https://opensource.org/licenses/MIT)
[![PowerShell](https://img.shields.io/badge/PowerShell-5.1+-blue.svg)](https://github.com/PowerShell/PowerShell)
[![Microsoft 365](https://img.shields.io/badge/Microsoft_365-compatible-brightgreen.svg)](https://www.microsoft.com/microsoft-365)
[![Graph API](https://img.shields.io/badge/Microsoft_Graph-v1.0-blue.svg)](https://developer.microsoft.com/en-us/graph)
[![Azure](https://img.shields.io/badge/Azure_Automation-compatible-0089D6.svg)](https://azure.microsoft.com/en-us/products/automation)
[![GitHub release](https://img.shields.io/github/release/sargeschultz11/ContactSync.svg)](https://GitHub.com/sargeschultz11/ContactSync/releases/)
[![Maintenance](https://img.shields.io/badge/Maintained-yes-green.svg)](https://github.com/swiderski/ContactSync)
[![Made with](https://img.shields.io/badge/Made%20with-PowerShell-1f425f.svg)](https://www.microsoft.com/powershell)

## Overview

ContactSync automates the synchronization of Microsoft 365 users as Exchange contacts for members of a specified security group. This creates a complete, up-to-date company directory for designated users without manual maintenance.

## Key Features

- **Automated contact synchronization** - Creates and maintains contacts for all licensed users
- **Organization-specific directories** - Define which users serve as contacts for different groups  
- **Smart updates** - Updates existing contacts when user information changes
- **Cleanup handling** - Removes contacts for deprovisioned users
- **Secure authentication** - Uses Azure Automation Managed Identity (no stored credentials)
- **Performance optimized** - Batch operations with intelligent fallback and throttling handling

## Prerequisites

- Microsoft 365 tenant with Exchange Online
- Azure Automation account 
- Security group containing users who should receive the contacts

## Setup Instructions

### 1. Enable System-Assigned Managed Identity
1. Navigate to your Azure Automation account
2. Select "Identity" from the sidebar  
3. Under "System assigned" tab, switch the Status to "On" and click "Save"
4. Copy the Object ID - you'll need this for permissions

### 2. Assign API Permissions to Managed Identity
1. Run the `utilities/Add-GraphPermissions.ps1` script from a local machine with global admin permissions:
   ```powershell
   .\Add-GraphPermissions.ps1 -AutomationMSI_ID "<Your-Automation-Account-MSI-Object-ID>"
   ```

### 3. Create a Security Group
1. Create a security group in Microsoft 365 containing users who should receive contacts
2. Note the Object ID of the group

### 4. Import and Configure the Main Script
1. Import `ContactSync.ps1` as a PowerShell runbook in your Automation account
2. Configure the required parameter `TargetGroupId` with your security group Object ID
3. Publish the runbook

### 5. Schedule the Runbook
1. Create a schedule to run ContactSync.ps1 at your desired frequency (daily recommended)
2. Link the schedule to the runbook with the `TargetGroupId` parameter

## Configuration Parameters

| Parameter | Type | Default | Description |
|-----------|------|---------|-------------|
| `TargetGroupId` | string | **Required** | The Microsoft 365 group ID containing users who should receive the contacts |
| `SourceGroupId` | string | "" | The Microsoft 365 group ID containing users who should be synchronized as contacts. If not specified, all licensed users in the tenant will be used. |
| `ExclusionListVariableName` | string | "ExclusionList" | The name of the Automation variable containing users to exclude (line-separated list) |
| `RemoveDeletedContacts` | bool | true | Whether to remove contacts that no longer exist in the source |
| `UpdateExistingContacts` | bool | true | Whether to update existing contacts with current information |
| `IncludeExternalContacts` | bool | true | Whether to include cloud-only users in the contact synchronization |
| `MaxConcurrentUsers` | int | 5 | Maximum number of concurrent users to process |
| `UseBatchOperations` | bool | true | Whether to attempt using batch operations (will fall back if needed) |

## Organization-Specific Contact Management

**Advanced Feature**: Configure separate contact directories for different organizations within the same tenant.

### Example Setup
Create separate schedules for each organization:

**Organization A**:
```powershell
TargetGroupId = "12345678-1234-1234-1234-123456789abc"  # OrgA-Users group ID
SourceGroupId = "12345678-1234-1234-1234-123456789abc"  # Same group - OrgA users get OrgA contacts
```

**Organization B**:
```powershell
TargetGroupId = "87654321-4321-4321-4321-cba987654321"  # OrgB-Users group ID  
SourceGroupId = "87654321-4321-4321-4321-cba987654321"  # Same group - OrgB users get OrgB contacts
```

**Result**: Users in each organization only see contacts from their own organization.

## How It Works

ContactSync uses your Azure Automation account's Managed Identity to securely access the Microsoft Graph API and:

1. **Authenticates** using Managed Identity (no stored credentials required)
2. **Retrieves** source users (from SourceGroupId or all licensed users) and target users (who receive contacts)
3. **Synchronizes** contacts by creating new ones, updating existing ones, and removing obsolete ones
4. **Optimizes** performance with batch operations and intelligent throttling

## Troubleshooting

### Common Issues
- **Authentication failures**: Verify Managed Identity has required Graph API permissions
- **Missing contacts**: Check exclusion list and user license status  
- **Performance issues**: Adjust the `MaxConcurrentUsers` parameter

### View Logs
1. In Azure Automation, go to the Jobs section
2. Select the most recent job run  
3. View the Output tab for detailed logs

## Utility Scripts

Additional maintenance scripts are available in the `utilities/` folder:

- **`ContactCleanup.ps1`** - Removes duplicate contacts and specific categories
- **`DeleteContactFolder.ps1`** - Deletes specific contact folders (e.g., "Administrator")  
- **`ContactDiagnostic.ps1`** - Analyzes contact data for troubleshooting
- **`Add-GraphPermissions.ps1`** - Assigns Graph API permissions to Managed Identity

These are typically used for one-time cleanup operations or diagnostics.

