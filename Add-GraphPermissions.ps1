# Requires -Modules "Microsoft.Graph.Applications"
<#
.SYNOPSIS
    Assigns Microsoft Graph API permissions to an Azure Automation Account's System-Assigned Managed Identity.
    
.DESCRIPTION
    This script assigns the necessary Microsoft Graph API permissions to allow the ContactSync solution
    to authenticate using a System-Assigned Managed Identity instead of an App Registration.
    
.NOTES
    Author:         Based on S.C. Swiderski's ContactSync solution
    Version:        1.0
    Creation Date:  April 2025
    
    Required permissions to run this script:
    - AppRoleAssignment.ReadWrite.All
    - Application.Read.All
    
.PARAMETER AutomationMSI_ID
    The Object ID of your Automation Account's System-Assigned Managed Identity.
    This can be found in the Azure Portal under your Automation Account > Identity > System assigned.
#>

[CmdletBinding()]
param (
    [Parameter(Mandatory = $true)]
    [string]$AutomationMSI_ID = "<REPLACE_WITH_YOUR_AUTOMATION_ACCOUNT_MSI_OBJECT_ID>"
)

# Microsoft Graph App ID (constant)
$GRAPH_APP_ID = "00000003-0000-0000-c000-000000000000"

Write-Host "Starting Graph permission assignment process..." -ForegroundColor Cyan

try {
    Write-Host "Connecting to Microsoft Graph API..."
    Connect-MgGraph -Scopes "AppRoleAssignment.ReadWrite.All", "Application.Read.All" -NoWelcome
    Write-Host "Successfully connected to Microsoft Graph API" -ForegroundColor Green
}
catch {
    Write-Host "Failed to connect to Microsoft Graph API: $_" -ForegroundColor Red
    Write-Host "Please ensure you have the required permissions and the Microsoft.Graph.Applications module is installed." -ForegroundColor Yellow
    exit 1
}

try {
    Write-Host "Retrieving Microsoft Graph Service Principal..."
    $GraphServicePrincipal = Get-MgServicePrincipal -Filter "appId eq '$GRAPH_APP_ID'"
    
    if ($null -eq $GraphServicePrincipal) {
        Write-Host "Could not find Microsoft Graph Service Principal. Exiting." -ForegroundColor Red
        exit 1
    }
    
    Write-Host "Found Microsoft Graph Service Principal with ID: $($GraphServicePrincipal.Id)" -ForegroundColor Green
}
catch {
    Write-Host "Error retrieving Microsoft Graph Service Principal: $_" -ForegroundColor Red
    exit 1
}

# Define the Graph permissions required for the ContactSync solution
# These IDs are standard across all tenants for Microsoft Graph
$GraphPermissionsList = @(
    @{Name = "Contacts.ReadWrite"; Id = "6918b873-d17a-4dc1-b314-35f528134491"},
    @{Name = "User.Read.All"; Id = "df021288-bdef-4463-88db-98f22de89214"},
    @{Name = "Group.Read.All"; Id = "5b567255-7703-4780-807c-7be8301ae99b"}
)

Write-Host "Assigning permissions to the Managed Identity ($AutomationMSI_ID)" -ForegroundColor Cyan

foreach ($permission in $GraphPermissionsList) {
    Write-Host "Processing permission: $($permission.Name)"
    
    $existingAssignment = Get-MgServicePrincipalAppRoleAssignment -ServicePrincipalId $AutomationMSI_ID | 
        Where-Object { $_.AppRoleId -eq $permission.Id }
        
    if (-not $existingAssignment) {
        try {
            New-MgServicePrincipalAppRoleAssignment -ServicePrincipalId $AutomationMSI_ID `
                -PrincipalId $AutomationMSI_ID `
                -ResourceId $GraphServicePrincipal.Id `
                -AppRoleId $permission.Id
                
            Write-Host "Permission $($permission.Name) assigned successfully" -ForegroundColor Green
        }
        catch {
            Write-Host "Error assigning permission $($permission.Name): $_" -ForegroundColor Red
        }
    }
    else {
        Write-Host "Permission $($permission.Name) already assigned" -ForegroundColor Yellow
    }
}

Write-Host "Permissions assignment completed" -ForegroundColor Green
Write-Host ""
Write-Host "NEXT STEPS:" -ForegroundColor Cyan
Write-Host "1. Modify your ContactSync scripts to use Managed Identity authentication instead of client credentials" -ForegroundColor White
Write-Host "2. Remove the ClientId, ClientSecret, and TenantId variables from your Automation Account" -ForegroundColor White
Write-Host "3. Update the 'Connect-ToMicrosoftGraph' function in your scripts to use the Managed Identity" -ForegroundColor White
Write-Host ""
Write-Host "Example code for Managed Identity authentication:" -ForegroundColor Cyan
Write-Host "function Connect-ToMicrosoftGraph {" -ForegroundColor White
Write-Host "    try {" -ForegroundColor White
Write-Host "        Write-Log 'Acquiring Microsoft Graph token using Managed Identity...'" -ForegroundColor White
Write-Host "        " -ForegroundColor White
Write-Host "        # Get the access token using the Managed Identity" -ForegroundColor White
Write-Host "        $response = Invoke-RestMethod -Uri 'http://169.254.169.254/metadata/identity/oauth2/token?api-version=2018-02-01&resource=https://graph.microsoft.com' -Headers @{Metadata='true'} -Method GET" -ForegroundColor White
Write-Host "        $script:GraphAccessToken = $response.access_token" -ForegroundColor White
Write-Host "        $script:TokenExpiresIn = $response.expires_in" -ForegroundColor White
Write-Host "        $script:TokenAcquiredTime = Get-Date" -ForegroundColor White
Write-Host "        " -ForegroundColor White
Write-Host "        Write-Log 'Successfully acquired Microsoft Graph API token'" -ForegroundColor White
Write-Host "        return $true" -ForegroundColor White
Write-Host "    }" -ForegroundColor White
Write-Host "    catch {" -ForegroundColor White
Write-Host "        Write-Log \"Failed to connect to Microsoft Graph API: $($_.Exception.Message)\" -Level \"ERROR\"" -ForegroundColor White
Write-Host "        throw $_" -ForegroundColor White
Write-Host "    }" -ForegroundColor White
Write-Host "}" -ForegroundColor White