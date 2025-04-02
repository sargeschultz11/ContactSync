# ContactSync - Enterprise Contact Synchronization Solution for Microsoft 365

[![License: MIT](https://img.shields.io/badge/License-MIT-yellow.svg)](https://opensource.org/licenses/MIT)
[![PowerShell](https://img.shields.io/badge/PowerShell-5.1+-blue.svg)](https://github.com/PowerShell/PowerShell)
[![Microsoft 365](https://img.shields.io/badge/Microsoft_365-compatible-brightgreen.svg)](https://www.microsoft.com/microsoft-365)
[![Graph API](https://img.shields.io/badge/Microsoft_Graph-v1.0-blue.svg)](https://developer.microsoft.com/en-us/graph)
[![Azure](https://img.shields.io/badge/Azure_Automation-compatible-0089D6.svg)](https://azure.microsoft.com/en-us/products/automation)
[![Version](https://img.shields.io/badge/Version-2.1-success.svg)](https://github.com/swiderski/ContactSync)
[![Maintenance](https://img.shields.io/badge/Maintained-yes-green.svg)](https://github.com/swiderski/ContactSync)
[![Made with](https://img.shields.io/badge/Made%20with-PowerShell-1f425f.svg)](https://www.microsoft.com/powershell)

## Overview
This repository contains a suite of PowerShell scripts for managing Microsoft 365 contacts across your organization. The main script, ContactsSync.ps1, automates the synchronization of all licensed Microsoft 365 users to the Exchange contacts of members in a specified security group. Additional utility scripts are provided for cleanup, diagnostics, and maintenance operations.

## Core Features
- Creates contacts for all licensed users in the tenant
- Updates existing contacts when user information changes
- Removes contacts for deprovisioned users
- Supports exclusion lists for specific users
- Configurable to include or exclude cloud-only users
- Optimized performance with batch operations and fallback mechanisms
- Throttling detection and handling

## Included Scripts

| Script | Description |
|--------|-------------|
| **ContactsSync.ps1** | Main script for synchronizing contacts across the organization |
| **ContactCleanup.ps1** | Removes duplicate contacts and contacts with specified categories |
| **DeleteContactFolder.ps1** | Deletes specific contact folders (e.g., "Administrator") |
| **ContactDiagnostic.ps1** | Analyzes contact data for a specific user to help troubleshoot issues |

## Prerequisites
- Microsoft 365 tenant with Exchange Online
- Azure Automation account
- App Registration in Azure AD with appropriate Graph API permissions
- Security group containing users who should receive the contacts

## Required Graph API Permissions
The application requires the following Microsoft Graph API permissions:
- `User.Read.All` - To read all user profiles
- `Group.Read.All` - To read group memberships
- `Contacts.ReadWrite` - To manage user contacts

## Setup Instructions

### 1. Create an App Registration in Azure AD
1. Navigate to Azure Active Directory > App Registrations
2. Create a new registration:
   - Name: ContactsSync
   - Supported account type: Single tenant
   - Redirect URI: (Web) https://portal.azure.com
3. Note the Application (client) ID and Directory (tenant) ID
4. Under Certificates & Secrets, create a new client secret and note the value

### 2. Assign API Permissions
1. Go to API Permissions
2. Add the following Microsoft Graph permissions:
   - User.Read.All
   - Group.Read.All
   - Contacts.ReadWrite
3. Grant admin consent for your organization

### 3. Create a Security Group
1. Create a security group in Microsoft 365 containing users who should receive contacts
2. Note the Object ID of the group

### 4. Set Up Azure Automation
1. Create an Azure Automation account
2. Import the Az modules and any other required modules
3. Create the following Automation variables:
   - `ClientId`: The Application (client) ID from the App Registration
   - `ClientSecret`: The client secret from the App Registration
   - `TenantId`: Your tenant ID
   - `ExclusionList` (optional): Line-separated list of user emails to exclude

### 5. Import the Scripts as Runbooks
1. Import all scripts as PowerShell runbooks
2. Edit the script parameters as needed, especially the `TargetGroupId`
3. Publish the runbooks

### 6. Schedule the Primary Runbook
1. Create a schedule for the ContactsSync.ps1 runbook to run at your desired frequency (daily recommended)
2. Link the schedule to the runbook
3. Configure the parameters for the scheduled run:
   - `TargetGroupId`: The Object ID of your security group
   - Other parameters as needed

## ContactsSync.ps1 Parameters

| Parameter | Type | Default | Description |
|-----------|------|---------|-------------|
| `TargetGroupId` | string | Required | The Microsoft 365 group ID containing users who should receive the contacts |
| `ExclusionListVariableName` | string | "ExclusionList" | The name of the Automation variable containing users to exclude |
| `RemoveDeletedContacts` | bool | true | Whether to remove contacts that no longer exist in the source |
| `UpdateExistingContacts` | bool | true | Whether to update existing contacts with current information |
| `IncludeExternalContacts` | bool | true | Whether to include cloud-only users in the contact synchronization |
| `MaxConcurrentUsers` | int | 5 | Maximum number of concurrent users to process |
| `UseBatchOperations` | bool | true | Whether to attempt using batch operations (will fall back if needed) |

## ContactCleanup.ps1 Parameters

| Parameter | Type | Default | Description |
|-----------|------|---------|-------------|
| `TargetGroupId` | string | Required | The Microsoft 365 group ID containing users whose contacts should be cleaned up |
| `PreserveCategory` | string | "Company Contacts" | The contact category to preserve |
| `RemoveCategory` | string | "Administrator" | The contact category to remove |
| `MaxConcurrentUsers` | int | 5 | Maximum number of concurrent users to process |
| `WhatIf` | switch | true | Run in simulation mode without making any changes |
| `DetailedLogging` | switch | true | Enables more detailed logging of contact processing |

## DeleteContactFolder.ps1 Parameters

| Parameter | Type | Default | Description |
|-----------|------|---------|-------------|
| `TargetGroupId` | string | Required | The Microsoft 365 group ID containing users whose contact folders should be cleaned up |
| `FolderNameToDelete` | string | "Administrator" | The name of the contact folder to delete |
| `WhatIf` | switch | false | Run in simulation mode without making any changes |

## ContactDiagnostic.ps1 Parameters

| Parameter | Type | Default | Description |
|-----------|------|---------|-------------|
| `TargetUserId` | string | Required | The Microsoft 365 user ID or UPN whose contacts should be analyzed |

## How It Works

### ContactsSync.ps1
The primary script that handles the synchronization of contacts.

1. **Authentication**: Uses client credentials flow to obtain an access token for the Microsoft Graph API, managing token expiration and refresh automatically.
2. **Data Retrieval**: Retrieves all licensed users from the Microsoft 365 tenant, filters based on exclusion list and license status, and retrieves all members of the target security group.
3. **Contact Synchronization**: For each user in the target group, retrieves existing contacts, creates new contacts for users that don't exist in the contact list, updates existing contacts if user information has changed, and removes contacts for users who are no longer in the organization.
4. **Performance Optimization**: Attempts to use batch operations for better performance, with fallback to individual operations if needed. Implements throttling detection and exponential backoff.

### ContactCleanup.ps1
A utility script for cleaning up duplicate contacts and removing contacts with specified categories.

1. Identifies all contacts for users in the specified security group
2. Groups contacts by display name and email address to detect duplicates
3. Identifies contacts with specified categories to remove (default is "Administrator")
4. Preserves contacts with a specified category (default is "Company Contacts")
5. Removes duplicate contacts, prioritizing those with the preserved category
6. Provides detailed logging of all operations

### DeleteContactFolder.ps1
A utility script for deleting specific contact folders.

1. Identifies all contact folders for users in the specified security group
2. Locates folders with the specified name (default is "Administrator")
3. Deletes the folder and all its contents
4. Provides detailed reporting of operations

### ContactDiagnostic.ps1
A diagnostic script for analyzing contact data for a specific user.

1. Retrieves all contacts for the specified user
2. Analyzes contacts for duplicates based on display name
3. Reports on contact categories and their distribution
4. Provides detailed output for diagnosing contact-related issues

## Mobile Device Configuration

### Configuring Intune for iOS Devices
To ensure that the synchronized contacts are available on mobile devices, you need to configure an Email profile in Microsoft Intune:

1. In the Microsoft Intune admin center, navigate to **Devices** > **Configuration profiles**
2. Create a new profile using the **Email** template
3. Configure the following settings:
   - **Email server**: outlook.office365.com
   - **Account name**: Corporate Directory
   - **Username attribute from Microsoft Entra ID**: User principal name
   - **Email address attribute from Microsoft Entra ID**: User principal name
   - **Authentication method**: Username and password
   - **SSL**: Enable
   - **OAuth**: Enable
   - **Exchange data to sync**: Contacts only
   - **Allow users to change sync settings**: No
4. Assign the profile to the appropriate groups
5. The contacts will sync to iOS devices without requiring user interaction

## Migration from CiraSync

The ContactCleanup.ps1 and DeleteContactFolder.ps1 scripts are specifically designed to assist with migration from CiraSync to the ContactsSync solution. They help clean up lingering data from the previous system, including:

1. Removing duplicate contacts created during migration
2. Removing contacts or folders with the "Administrator" category (typically used by CiraSync)
3. Preserving contacts with the "Company Contacts" category (used by ContactsSync)

For a smooth migration, use the following workflow:

1. Run ContactDiagnostic.ps1 on a sample user to understand the current state of contacts
2. Run ContactCleanup.ps1 with the WhatIf parameter set to true to simulate cleanup
3. Review the logs and confirm the expected changes
4. Run ContactCleanup.ps1 with WhatIf set to false to perform the actual cleanup
5. If needed, run DeleteContactFolder.ps1 to remove any remaining Administrator folders
6. Start the regular ContactsSync.ps1 process

## Troubleshooting

### Common Issues
- **Authentication failures**: Verify the client ID, client secret, and tenant ID
- **Permission errors**: Ensure the application has the required Graph API permissions
- **Performance issues**: Adjust the `MaxConcurrentUsers` parameter
- **Missing contacts**: Check the exclusion list and verify user license status
- **Duplicate contacts**: Run ContactDiagnostic.ps1 to identify duplicates, then use ContactCleanup.ps1 to resolve them

### Viewing Logs
1. In Azure Automation, go to the Jobs section
2. Select the most recent job run
3. View the Output tab to see detailed logs

## Advanced Configuration

### Excluding Users
To exclude specific users from being created as contacts:
1. In Azure Automation, edit the `ExclusionList` variable
2. Add user email addresses, one per line

### Customizing Contact Properties
To customize the contact properties:
1. Modify the `New-ContactObject` function in the script
2. Add or modify properties as needed

## Maintenance
- Periodically check the Azure Automation job history for errors
- Update the client secret when it expires
- Review and update the exclusion list as needed
- Run ContactDiagnostic.ps1 periodically to check for contact issues
- Run ContactCleanup.ps1 if duplicate contacts are reported