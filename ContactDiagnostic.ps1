<#
.SYNOPSIS
    Contact Diagnostic Script for Microsoft 365
    
.DESCRIPTION
    ContactDiagnostic.ps1 dumps detailed information about contacts for a specific user
    to help diagnose issues with duplicate contacts and categories.
    
.NOTES
    Author:         Based on Ryan Schultz's ContactsSync.ps1
    Version:        1.0
    Creation Date:  March 2025
    
.PARAMETER TargetUserId
    The Microsoft 365 user ID or UPN whose contacts should be analyzed
#>

param(
    [Parameter(Mandatory = $true)]
    [string]$TargetUserId
)

# Global variables
$ErrorActionPreference = "Stop"
$script:GraphAccessToken = $null
$script:TokenAcquiredTime = $null
$script:TokenExpiresIn = 3600 # Default 1 hour (in seconds)

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

function Analyze-UserContacts {
    param(
        [Parameter(Mandatory = $true)]
        [string]$UserId
    )
    
    try {
        $userId = $UserId
        if ($UserId -like "*@*") {
            Write-Log "Resolving user ID for UPN: $UserId"
            $userResponse = Invoke-GraphRequest -Method "GET" -Uri "/users/$UserId"
            $userId = $userResponse.id
            Write-Log "Resolved user ID: $userId"
        }
        
        Write-Log "Retrieving all contacts for user $UserId..."
        
        $existingContacts = @()
        $uri = "/users/$userId/contacts?`$top=999&`$select=id,displayName,givenName,surname,emailAddresses,businessPhones,mobilePhone,companyName,categories,createdDateTime"
        
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
        
        Write-Log "Retrieved $($existingContacts.Count) contacts for user"
        
        $contactsByDisplayName = @{}
        foreach ($contact in $existingContacts) {
            $displayName = $contact.displayName
            
            if (-not $contactsByDisplayName.ContainsKey($displayName)) {
                $contactsByDisplayName[$displayName] = @()
            }
            
            $contactsByDisplayName[$displayName] += $contact
        }
        
        Write-Log "=== DUPLICATE CONTACTS ANALYSIS ==="
        $duplicateCount = 0
        
        foreach ($displayName in $contactsByDisplayName.Keys) {
            $contacts = $contactsByDisplayName[$displayName]
            
            if ($contacts.Count -gt 1) {
                $duplicateCount++
                Write-Log "Found $($contacts.Count) contacts with display name: $displayName" -Level "INFO"
                
                foreach ($contact in $contacts) {
                    $email = if ($contact.emailAddresses.Count -gt 0) { $contact.emailAddresses[0].address } else { "No email" }
                    $categories = if ($contact.categories -and $contact.categories.Count -gt 0) { $contact.categories -join ", " } else { "No categories" }
                    $created = if ($contact.createdDateTime) { $contact.createdDateTime } else { "Unknown" }
                    
                    Write-Log "  - ID: $($contact.id)" -Level "INFO"
                    Write-Log "    Email: $email" -Level "INFO"
                    Write-Log "    Categories: $categories" -Level "INFO"
                    Write-Log "    Created: $created" -Level "INFO"
                    Write-Log "    Full object: $($contact | ConvertTo-Json -Depth 3)" -Level "INFO"
                }
            }
        }
        
        Write-Log "=== CATEGORIES ANALYSIS ==="
        $allCategories = @{}
        
        foreach ($contact in $existingContacts) {
            if ($contact.categories -and $contact.categories.Count -gt 0) {
                foreach ($category in $contact.categories) {
                    if (-not $allCategories.ContainsKey($category)) {
                        $allCategories[$category] = 0
                    }
                    
                    $allCategories[$category]++
                }
            }
        }
        
        Write-Log "Found the following categories:" -Level "INFO"
        foreach ($category in $allCategories.Keys) {
            Write-Log "  - $category`: $($allCategories[$category]) contacts" -Level "INFO"
        }
        
        $contactsWithoutCategories = $existingContacts | Where-Object { -not $_.categories -or $_.categories.Count -eq 0 }
        Write-Log "Contacts without categories: $($contactsWithoutCategories.Count)" -Level "INFO"
        
        return @{
            TotalContacts = $existingContacts.Count
            DuplicateContactSets = $duplicateCount
            Categories = $allCategories
            ContactsWithoutCategories = $contactsWithoutCategories.Count
        }
    }
    catch {
        Write-Log "Error analyzing contacts: $($_.Exception.Message)" -Level "ERROR"
        throw $_
    }
}

try {
    $startTime = Get-Date
    Write-Log "Starting Contact Diagnostic process"
    
    Connect-ToMicrosoftGraph
    
    $results = Analyze-UserContacts -UserId $TargetUserId
    
    Write-Log "=== SUMMARY ===" -Level "INFO"
    Write-Log "Total contacts: $($results.TotalContacts)" -Level "INFO"
    Write-Log "Duplicate contact sets: $($results.DuplicateContactSets)" -Level "INFO"
    Write-Log "Contacts without categories: $($results.ContactsWithoutCategories)" -Level "INFO"
    
    $endTime = Get-Date
    $duration = $endTime - $startTime
    
    Write-Log "Diagnostic completed in $($duration.TotalSeconds.ToString("F2")) seconds"
}
catch {
    $errorMessage = $_.Exception.Message
    Write-Log "ContactDiagnostic failed with error: $errorMessage" -Level "ERROR"
    throw $_
}