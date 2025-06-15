<#
.SYNOPSIS
    Contact Cleanup Script for Microsoft 365 - Removes Duplicates and Administrator Category Contacts
    
.DESCRIPTION
    ContactCleanup.ps1 is a PowerShell script that removes duplicate contacts and
    contacts tagged with the "Administrator" category for members of a specified security group.
    This script is designed to be run as a one-time cleanup after migrating from CiraSync to
    the ContactsSync solution.
    
.FUNCTIONALITY
    - Identifies and removes duplicate contacts based on email address
    - Removes contacts with the "Administrator" category tag
    - Preserves contacts with the "Company Contacts" category
    - Provides detailed logging of all operations
    
.NOTES
    Author:         Based on Ryan Schultz's ContactsSync.ps1
    Version:        1.1
    Creation Date:  March 2025
    Modified Date:  April 2025
    
.PARAMETER TargetGroupId
    The Microsoft 365 group ID containing users whose contacts should be cleaned up
    
.PARAMETER PreserveCategory
    The contact category to preserve (default is "Company Contacts")
    
.PARAMETER RemoveCategory
    The contact category to remove (default is "Administrator")
    
.PARAMETER MaxConcurrentUsers
    Maximum number of concurrent users to process (default is 5)
    
.PARAMETER WhatIf
    Run in simulation mode without making any changes, only reporting what would be done
    
.PARAMETER DetailedLogging
    Enables more detailed logging of contact processing
#>

param(
    [Parameter(Mandatory = $true)]
    [string]$TargetGroupId, 
    [string]$PreserveCategory = "Company Contacts",
    [string]$RemoveCategory = "Administrator",
    [int]$MaxConcurrentUsers = 5,
    [switch]$WhatIf = $true,
    [switch]$DetailedLogging = $true
)

# Global variables
$ErrorActionPreference = "Stop"
$script:GraphAccessToken = $null
$script:TokenAcquiredTime = $null
$script:TokenExpiresIn = 3600 # Default 1 hour (in seconds)
$script:Throttled = $false

# Logging function
function Write-Log {
    param(
        [Parameter(Mandatory = $true)]
        [string]$Message,
        [ValidateSet("INFO", "WARNING", "ERROR", "VERBOSE")]
        [string]$Level = "INFO"
    )
    
    if ($Level -eq "VERBOSE" -and -not $DetailedLogging) {
        return
    }
    
    $timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    $logEntry = "$timestamp [$Level] $Message"
    Write-Output $logEntry
}

# Token management with Managed Identity
function Connect-ToMicrosoftGraph {
    try {
        Write-Log "Acquiring Microsoft Graph token using Managed Identity via Az modules..."
        
        Connect-AzAccount -Identity | Out-Null
        
        $cmdInfo = Get-Command Get-AzAccessToken -ErrorAction SilentlyContinue
        $hasAsSecureStringParam = $false
        
        if ($cmdInfo -and $cmdInfo.Parameters.ContainsKey("AsSecureString")) {
            $hasAsSecureStringParam = $true
        }
        
        if ($hasAsSecureStringParam) {
            $secureToken = (Get-AzAccessToken -ResourceUrl "https://graph.microsoft.com" -AsSecureString).Token
            $bstr = [System.Runtime.InteropServices.Marshal]::SecureStringToBSTR($secureToken)
            $token = [System.Runtime.InteropServices.Marshal]::PtrToStringAuto($bstr)
            [System.Runtime.InteropServices.Marshal]::ZeroFreeBSTR($bstr)
        } else {
            $token = (Get-AzAccessToken -ResourceUrl "https://graph.microsoft.com").Token
        }
        
        if ([string]::IsNullOrEmpty($token)) {
            throw "Failed to acquire token from managed identity"
        }
        
        $script:GraphAccessToken = $token
        $script:TokenAcquiredTime = Get-Date
        $script:TokenExpiresIn = 3600
        
        Write-Log "Successfully acquired Microsoft Graph API token via Managed Identity"
        return $true
    }
    catch {
        Write-Log "Failed to connect to Microsoft Graph API using Managed Identity: $($_.Exception.Message)" -Level "ERROR"
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
        $retryDelay = 2 # Initial retry delay in seconds
        $success = $false
        $response = $null
        
        while (-not $success -and $retryCount -le $MaxRetries) {
            try {
                $response = Invoke-RestMethod @params -ErrorAction Stop
                $success = $true
                $script:Throttled = $false
            } 
            catch {
                $statusCode = 0
                
                if ($_.Exception.Response) {
                    $statusCode = [int]$_.Exception.Response.StatusCode
                }
                
                if ($statusCode -eq 429 -or $statusCode -eq 503 -or $statusCode -eq 504) {
                    $retryCount++
                    $script:Throttled = $true
                    
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

function Clean-UserContacts {
    param(
        [Parameter(Mandatory = $true)]
        [object]$User
    )
    
    try {
        Write-Log "Starting contact cleanup for user $($User.userPrincipalName)..."
        
        $existingContacts = @()
        $uri = "/users/$($User.id)/contacts?`$top=999&`$select=id,displayName,givenName,surname,emailAddresses,businessPhones,mobilePhone,companyName,categories"
        
        do {
            $response = Invoke-GraphRequest -Method "GET" -Uri $uri
            
            if ($response -and $response.value) {
                $existingContacts += $response.value
            }
            
            $uri = $null
            if ($response.'@odata.nextLink') {
                $uri = $response.'@odata.nextLink' -replace "https://graph.microsoft.com/v1.0", ""
            }
        } while ($uri)
        
        Write-Log "Retrieved $($existingContacts.Count) existing contacts for $($User.userPrincipalName)"
        
        $contactsByDisplayName = @{}
        $contactsByEmail = @{}
        $contactsWithoutEmail = @()
        $administratorContacts = @()
        
        foreach ($contact in $existingContacts) {
            if ($contact.categories -and $contact.categories -contains $RemoveCategory) {
                Write-Log "Found contact with '$RemoveCategory' category: $($contact.displayName) ($($contact.id))" -Level "VERBOSE"
                $administratorContacts += $contact
            }
            
            $displayName = $contact.displayName.ToLower()
            if (-not [string]::IsNullOrEmpty($displayName)) {
                if (-not $contactsByDisplayName.ContainsKey($displayName)) {
                    $contactsByDisplayName[$displayName] = @()
                }
                $contactsByDisplayName[$displayName] += $contact
            }
            
            if ($contact.emailAddresses.Count -gt 0) {
                $email = $contact.emailAddresses[0].address.ToLower()
                
                if (-not [string]::IsNullOrEmpty($email)) {
                    if (-not $contactsByEmail.ContainsKey($email)) {
                        $contactsByEmail[$email] = @()
                    }
                    
                    $contactsByEmail[$email] += $contact
                }
            }
            else {
                $contactsWithoutEmail += $contact
            }
        }
        
        $contactsToDelete = @()
        $duplicateCount = 0
        $administratorCount = 0
        $processedIds = @{}
        
        foreach ($contact in $administratorContacts) {
            if (-not $processedIds.ContainsKey($contact.id)) {
                $contactsToDelete += $contact
                $processedIds[$contact.id] = $true
                $administratorCount++
                Write-Log "Marking contact for deletion (has '$RemoveCategory' category): $($contact.displayName)" -Level "VERBOSE"
            }
        }
        
        foreach ($displayName in $contactsByDisplayName.Keys) {
            $contacts = $contactsByDisplayName[$displayName]
            
            if ($contacts.Count -gt 1) {
                Write-Log "Found $($contacts.Count) duplicate contacts for name: $displayName" -Level "VERBOSE"
                $duplicateCount += ($contacts.Count - 1)
                
                $contactToKeep = $null
                
                foreach ($contact in $contacts) {
                    if ($contact.categories -and $contact.categories -contains $PreserveCategory -and 
                        (-not $contact.categories -contains $RemoveCategory)) {
                        $contactToKeep = $contact
                        break
                    }
                }
                
                if ($null -eq $contactToKeep) {
                    $contactToKeep = $contacts[0]
                }
                
                foreach ($contact in $contacts) {
                    if ($contact.id -ne $contactToKeep.id -and -not $processedIds.ContainsKey($contact.id)) {
                        $contactsToDelete += $contact
                        $processedIds[$contact.id] = $true
                        Write-Log "Marking duplicate contact for deletion: $($contact.displayName) ($($contact.id))" -Level "VERBOSE"
                    }
                }
            }
        }
        
        foreach ($email in $contactsByEmail.Keys) {
            $contacts = $contactsByEmail[$email]
            
            if ($contacts.Count -gt 1) {
                Write-Log "Found $($contacts.Count) duplicate contacts for email: $email" -Level "VERBOSE"
                
                $contactToKeep = $null
                
                foreach ($contact in $contacts) {
                    if ($contact.categories -and $contact.categories -contains $PreserveCategory -and 
                        (-not $contact.categories -contains $RemoveCategory) -and 
                        -not $processedIds.ContainsKey($contact.id)) {
                        $contactToKeep = $contact
                        break
                    }
                }
                
                if ($null -eq $contactToKeep) {
                    foreach ($contact in $contacts) {
                        if (-not $processedIds.ContainsKey($contact.id)) {
                            $contactToKeep = $contact
                            break
                        }
                    }
                }
                
                if ($null -ne $contactToKeep) {
                    foreach ($contact in $contacts) {
                        if ($contact.id -ne $contactToKeep.id -and -not $processedIds.ContainsKey($contact.id)) {
                            $contactsToDelete += $contact
                            $processedIds[$contact.id] = $true
                            $duplicateCount++
                            Write-Log "Marking email duplicate for deletion: $($contact.displayName) ($($contact.id))" -Level "VERBOSE"
                        }
                    }
                }
            }
        }
        
        foreach ($contact in $administratorContacts) {
            if (-not $processedIds.ContainsKey($contact.id)) {
                $contactsToDelete += $contact
                $processedIds[$contact.id] = $true
                $administratorCount++
                Write-Log "Marking additional contact with '$RemoveCategory' for deletion: $($contact.displayName)" -Level "VERBOSE"
            }
        }
        
        Write-Log "Found $duplicateCount duplicates and $administratorCount contacts with '$RemoveCategory' category" -Level "INFO"
        
        if ($contactsToDelete.Count -gt 0) {
            Write-Log "Contacts marked for deletion:" -Level "INFO"
            foreach ($contact in $contactsToDelete) {
                $categoryInfo = if ($contact.categories) { "Categories: $($contact.categories -join ', ')" } else { "No categories" }
                $emailInfo = if ($contact.emailAddresses.Count -gt 0) { "Email: $($contact.emailAddresses[0].address)" } else { "No email" }
                Write-Log "- $($contact.displayName) [$categoryInfo] [$emailInfo]" -Level "VERBOSE"
            }
        }
        
        if (-not $WhatIf) {
            $deletedCount = 0
            $errorCount = 0
            
            foreach ($contact in $contactsToDelete) {
                try {
                    $uri = "/users/$($User.id)/contacts/$($contact.id)"
                    Invoke-GraphRequest -Method "DELETE" -Uri $uri
                    $deletedCount++
                    
                    Start-Sleep -Milliseconds 50
                }
                catch {
                    Write-Log "Error deleting contact $($contact.id): $($_.Exception.Message)" -Level "ERROR"
                    $errorCount++
                }
            }
            
            Write-Log "Successfully deleted $deletedCount contacts with $errorCount errors"
        }
        else {
            Write-Log "WhatIf mode: Would delete $($contactsToDelete.Count) contacts"
        }
        
        return @{
            User = $User.userPrincipalName
            TotalContacts = $existingContacts.Count
            DuplicatesFound = $duplicateCount
            AdministratorContacts = $administratorCount
            ContactsToDelete = $contactsToDelete.Count
        }
    }
    catch {
        Write-Log "Error cleaning contacts for $($User.userPrincipalName): $($_.Exception.Message)" -Level "ERROR"
        throw $_
    }
}

# Main execution
try {
    $startTime = Get-Date
    Write-Log "Starting ContactCleanup process"
    
    if ($WhatIf) {
        Write-Log "Running in SIMULATION MODE (WhatIf). No changes will be made." -Level "WARNING"
    }
    
    Write-Log "Parameters: TargetGroupId='$TargetGroupId', PreserveCategory='$PreserveCategory', RemoveCategory='$RemoveCategory', MaxConcurrentUsers=$MaxConcurrentUsers"
    
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
    $totalDuplicates = 0
    $totalAdministrator = 0
    $totalDeleted = 0
    
    Write-Log "Processing $totalUsers users"
    
    foreach ($currentUser in $targetUsers) {
        $processedCount++
        
        if ([string]::IsNullOrEmpty($currentUser.userPrincipalName)) {
            Write-Log "Skipping user $processedCount of $totalUsers - Invalid user information" -Level "WARNING"
            $errorCount++
            continue
        }
        
        Write-Log "Starting cleanup for user $processedCount of $totalUsers : $($currentUser.userPrincipalName)"
        
        try {
            $result = Clean-UserContacts -User $currentUser
            $completedCount++
            
            $totalDuplicates += $result.DuplicatesFound
            $totalAdministrator += $result.AdministratorContacts
            $totalDeleted += $result.ContactsToDelete
            
            Write-Log "Completed cleanup for $($result.User): Found $($result.DuplicatesFound) duplicates and $($result.AdministratorContacts) '$RemoveCategory' contacts"
        }
        catch {
            $errorCount++
            Write-Log "Error processing user $($currentUser.userPrincipalName): $($_.Exception.Message)" -Level "ERROR"
        }
        
        if ($script:Throttled) {
            Write-Log "Detected API throttling, pausing for 5 seconds..." -Level "WARNING"
            Start-Sleep -Seconds 5
        }
    }
    
    $endTime = Get-Date
    $duration = $endTime - $startTime
    
    Write-Log "ContactCleanup completed in $($duration.TotalMinutes.ToString("F2")) minutes"
    
    $modeStr = if ($WhatIf) { "would be" } else { "were" }
    Write-Log "Final stats: $completedCount users processed successfully, $errorCount errors"
    Write-Log "Found $totalDuplicates duplicate contacts and $totalAdministrator '$RemoveCategory' contacts"
    Write-Log "Total $totalDeleted contacts $modeStr deleted"
}
catch {
    $errorMessage = $_.Exception.Message
    Write-Log "ContactCleanup failed with error: $errorMessage" -Level "ERROR"
    throw $_
}