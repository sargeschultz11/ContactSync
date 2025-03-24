<#
.SYNOPSIS
    Delete Contact Folder Script for Microsoft 365
    
.DESCRIPTION
    DeleteContactFolder.ps1 identifies and deletes a specific contact folder
    (such as "Administrator") for members of a specified security group.
    
.FUNCTIONALITY
    - Identifies contact folders by name
    - Deletes specified contact folder and all its contacts
    - Supports simulation mode for testing
    
.NOTES
    Author:         Based on Ryan Schultz's ContactsSync.ps1
    Version:        1.0
    Creation Date:  March 2025
    
.PARAMETER TargetGroupId
    The Microsoft 365 group ID containing users whose contact folders should be cleaned up
    
.PARAMETER FolderNameToDelete
    The name of the contact folder to delete (default is "Administrator")
    
.PARAMETER WhatIf
    Run in simulation mode without making any changes, only reporting what would be done
#>

param(
    [Parameter(Mandatory = $true)]
    [string]$TargetGroupId,
    [string]$FolderNameToDelete = "Administrator",
    [switch]$WhatIf = $false
)

# Global variables
$ErrorActionPreference = "Stop"
$script:GraphAccessToken = $null
$script:TokenAcquiredTime = $null
$script:TokenExpiresIn = 3600

# Logging function
function Write-Log {
    param(
        [Parameter(Mandatory = $true)]
        [string]$Message,
        [ValidateSet("INFO", "WARNING", "ERROR", "VERBOSE")]
        [string]$Level = "INFO"
    )
    
    $timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    $logEntry = "$timestamp [$Level] $Message"
    Write-Output $logEntry
}

# Token management
function Connect-ToMicrosoftGraph {
    try {
        $clientId = Get-AutomationVariable -Name "ClientId"
        $clientSecret = Get-AutomationVariable -Name "ClientSecret" 
        $tenantId = Get-AutomationVariable -Name "TenantId"
        
        Write-Log "Acquiring Microsoft Graph token..."
        
        $tokenUrl = "https://login.microsoftonline.com/$tenantId/oauth2/v2.0/token"
        $body = @{
            grant_type    = "client_credentials"
            client_id     = $clientId
            client_secret = $clientSecret
            scope         = "https://graph.microsoft.com/.default"
        }
        
        $tokenResponse = Invoke-RestMethod -Uri $tokenUrl -Method Post -Body $body -ContentType "application/x-www-form-urlencoded"
        $script:GraphAccessToken = $tokenResponse.access_token
        $script:TokenExpiresIn = $tokenResponse.expires_in
        $script:TokenAcquiredTime = Get-Date
        
        Write-Log "Successfully acquired Microsoft Graph API token"
        return $true
    }
    catch {
        Write-Log "Failed to connect to Microsoft Graph API: $($_.Exception.Message)" -Level "ERROR"
        throw $_
    }
}

function Ensure-ValidGraphToken {
    $currentTime = Get-Date
    $tokenAge = 0
    
    if ($script:TokenAcquiredTime) {
        $tokenAge = ($currentTime - $script:TokenAcquiredTime).TotalSeconds
    }
    
    if ($null -eq $script:TokenAcquiredTime -or $tokenAge -gt 3000) {
        Write-Log "Graph token is expired or about to expire. Refreshing token..." -Level "INFO"
        Connect-ToMicrosoftGraph
        Write-Log "Graph token refreshed successfully" -Level "INFO"
    }
}

# Graph API Request Functions
function Invoke-GraphRequest {
    param (
        [Parameter(Mandatory = $true)]
        [string]$Method,
        
        [Parameter(Mandatory = $true)]
        [string]$Uri,
        
        [Parameter(Mandatory = $false)]
        [object]$Body = $null,
        
        [Parameter(Mandatory = $false)]
        [string]$ContentType = "application/json",
        
        [Parameter(Mandatory = $false)]
        [int]$MaxRetries = 5
    )
    
    try {
        Ensure-ValidGraphToken
        
        $headers = @{
            "Authorization" = "Bearer $script:GraphAccessToken"
            "Content-Type" = $ContentType
            "ConsistencyLevel" = "eventual"
        }
        
        $params = @{
            Method = $Method
            Uri = "https://graph.microsoft.com/v1.0$Uri"
            Headers = $headers
        }
        
        if ($null -ne $Body -and $Method -ne "GET") {
            if ($ContentType -eq "application/json") {
                $params.Body = ($Body | ConvertTo-Json -Depth 10)
            }
            else {
                $params.Body = $Body
            }
        }
        
        $retryCount = 0
        $retryDelay = 2
        $success = $false
        $response = $null
        
        while (-not $success -and $retryCount -le $MaxRetries) {
            try {
                $response = Invoke-RestMethod @params -ErrorAction Stop
                $success = $true
            } 
            catch {
                $statusCode = 0
                
                if ($_.Exception.Response) {
                    $statusCode = [int]$_.Exception.Response.StatusCode
                }
                
                if ($statusCode -eq 429 -or $statusCode -eq 503 -or $statusCode -eq 504) {
                    $retryCount++
                    
                    if ($retryCount -le $MaxRetries) {
                        Write-Log "Request throttled (status $statusCode). Waiting $retryDelay seconds before retry $retryCount of $MaxRetries..." -Level "WARNING"
                        Start-Sleep -Seconds $retryDelay
                        $retryDelay = [Math]::Min($retryDelay * 2, 60)
                    }
                    else {
                        Write-Log "Request failed after $MaxRetries retries: $($_.Exception.Message)" -Level "ERROR"
                        throw $_
                    }
                }
                else {
                    Write-Log "Graph API error (status $statusCode): $($_.Exception.Message)" -Level "ERROR"
                    throw $_
                }
            }
        }
        
        return $response
    }
    catch {
        Write-Log "Graph API request failed: $($_.Exception.Message)" -Level "ERROR"
        throw $_
    }
}

function Get-GroupMembers {
    param(
        [Parameter(Mandatory = $true)]
        [string]$GroupId
    )
    
    try {
        Write-Log "Retrieving members of target group $GroupId..."
        
        $uri = "/groups/$GroupId/members?`$select=id,userPrincipalName,displayName,mail,accountEnabled"
        
        $allMembers = @()
        
        do {
            $response = Invoke-GraphRequest -Method "GET" -Uri $uri
            
            if ($response -and $response.value) {
                $allMembers += $response.value
            }
            
            $uri = $null
            if ($response.'@odata.nextLink') {
                $uri = $response.'@odata.nextLink' -replace "https://graph.microsoft.com/v1.0", ""
            }
        } while ($uri)
        
        $enabledMembers = $allMembers | Where-Object { 
            $_.'@odata.type' -like "*#microsoft.graph.user*" -and 
            $_.accountEnabled -eq $true 
        }
        
        Write-Log "Found $($enabledMembers.Count) enabled members in the target group"
        return $enabledMembers
    }
    catch {
        Write-Log "Error retrieving group members: $($_.Exception.Message)" -Level "ERROR"
        throw $_
    }
}

function Delete-ContactFolder {
    param(
        [Parameter(Mandatory = $true)]
        [object]$User,
        
        [Parameter(Mandatory = $true)]
        [string]$FolderName,
        
        [Parameter(Mandatory = $false)]
        [bool]$WhatIf = $false
    )
    
    try {
        Write-Log "Starting contact folder cleanup for user $($User.userPrincipalName)..."
        
        $uri = "/users/$($User.id)/contactFolders?`$top=999"
        $allFolders = @()
        
        do {
            $response = Invoke-GraphRequest -Method "GET" -Uri $uri
            
            if ($response -and $response.value) {
                $allFolders += $response.value
            }
            
            $uri = $null
            if ($response.'@odata.nextLink') {
                $uri = $response.'@odata.nextLink' -replace "https://graph.microsoft.com/v1.0", ""
            }
        } while ($uri)
        
        Write-Log "Retrieved $($allFolders.Count) contact folders for $($User.userPrincipalName)"
        
        $targetFolder = $allFolders | Where-Object { $_.displayName -eq $FolderName }
        
        if ($null -eq $targetFolder) {
            Write-Log "No contact folder named '$FolderName' found for user $($User.userPrincipalName)" -Level "INFO"
            return @{
                User = $User.userPrincipalName
                FolderFound = $false
                ContactsDeleted = 0
                FolderDeleted = $false
            }
        }
        
        Write-Log "Found '$FolderName' folder for $($User.userPrincipalName) with ID: $($targetFolder.id)" -Level "INFO"
        
        $contactsUri = "/users/$($User.id)/contactFolders/$($targetFolder.id)/contacts?`$top=999&`$select=id,displayName"
        $folderContacts = @()
        
        do {
            $response = Invoke-GraphRequest -Method "GET" -Uri $contactsUri
            
            if ($response -and $response.value) {
                $folderContacts += $response.value
            }
            
            $contactsUri = $null
            if ($response.'@odata.nextLink') {
                $contactsUri = $response.'@odata.nextLink' -replace "https://graph.microsoft.com/v1.0", ""
            }
        } while ($contactsUri)
        
        Write-Log "Folder contains $($folderContacts.Count) contacts" -Level "INFO"
        
        if (-not $WhatIf) {
            try {
                $deleteUri = "/users/$($User.id)/contactFolders/$($targetFolder.id)"
                Write-Log "Deleting folder '$FolderName' for $($User.userPrincipalName)..." -Level "INFO"
                Invoke-GraphRequest -Method "DELETE" -Uri $deleteUri
                Write-Log "Successfully deleted folder '$FolderName' and all its contents" -Level "INFO"
                
                return @{
                    User = $User.userPrincipalName
                    FolderFound = $true
                    ContactsDeleted = $folderContacts.Count
                    FolderDeleted = $true
                }
            }
            catch {
                Write-Log "Error deleting folder: $($_.Exception.Message)" -Level "ERROR"
                return @{
                    User = $User.userPrincipalName
                    FolderFound = $true
                    ContactsDeleted = 0
                    FolderDeleted = $false
                    Error = $_.Exception.Message
                }
            }
        }
        else {
            Write-Log "WhatIf mode: Would delete folder '$FolderName' containing $($folderContacts.Count) contacts" -Level "INFO"
            return @{
                User = $User.userPrincipalName
                FolderFound = $true
                ContactsDeleted = $folderContacts.Count
                FolderDeleted = $false
                WhatIf = $true
            }
        }
    }
    catch {
        Write-Log "Error processing contact folder for $($User.userPrincipalName): $($_.Exception.Message)" -Level "ERROR"
        throw $_
    }
}

# Main execution
try {
    $startTime = Get-Date
    Write-Log "Starting Delete Contact Folder process"
    
    if ($WhatIf) {
        Write-Log "Running in SIMULATION MODE (WhatIf). No changes will be made." -Level "WARNING"
    }
    
    Write-Log "Parameters: TargetGroupId='$TargetGroupId', FolderNameToDelete='$FolderNameToDelete'"
    
    Connect-ToMicrosoftGraph
    
    $targetUsers = Get-GroupMembers -GroupId $TargetGroupId
    
    if ($targetUsers.Count -eq 0) {
        Write-Log "No users found in the target group. Exiting." -Level "WARNING"
        exit
    }
    
    $totalUsers = $targetUsers.Count
    $processedCount = 0
    $completedCount = 0
    $errorCount = 0
    $foldersFound = 0
    $foldersDeleted = 0
    $totalContactsDeleted = 0
    
    Write-Log "Processing $totalUsers users"
    
    foreach ($currentUser in $targetUsers) {
        $processedCount++
        
        if ([string]::IsNullOrEmpty($currentUser.userPrincipalName)) {
            Write-Log "Skipping user $processedCount of $totalUsers - Invalid user information" -Level "WARNING"
            $errorCount++
            continue
        }
        
        Write-Log "Processing user $processedCount of $totalUsers : $($currentUser.userPrincipalName)"
        
        try {
            $result = Delete-ContactFolder -User $currentUser -FolderName $FolderNameToDelete -WhatIf $WhatIf
            $completedCount++
            
            if ($result.FolderFound) {
                $foldersFound++
                if ($result.FolderDeleted) {
                    $foldersDeleted++
                    $totalContactsDeleted += $result.ContactsDeleted
                }
            }
        }
        catch {
            $errorCount++
            Write-Log "Error processing user $($currentUser.userPrincipalName): $($_.Exception.Message)" -Level "ERROR"
        }
    }
    
    $endTime = Get-Date
    $duration = $endTime - $startTime
    
    $modeStr = if ($WhatIf) { "would be" } else { "were" }
    Write-Log "DeleteContactFolder completed in $($duration.TotalMinutes.ToString("F2")) minutes"
    Write-Log "Final stats: $completedCount users processed successfully, $errorCount errors"
    Write-Log "Found $foldersFound users with '$FolderNameToDelete' folders"
    Write-Log "$foldersDeleted folders $modeStr deleted, containing $totalContactsDeleted contacts"
}
catch {
    $errorMessage = $_.Exception.Message
    Write-Log "DeleteContactFolder failed with error: $errorMessage" -Level "ERROR"
    throw $_
}