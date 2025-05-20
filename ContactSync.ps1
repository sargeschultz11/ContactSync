<#
.SYNOPSIS
    Enterprise Contact Synchronization Solution for Microsoft 365 - Working Version
    
.DESCRIPTION
    ContactsSync.ps1 automates the synchronization of all licensed Microsoft 365 users 
    to the Exchange contacts of members in a specified security group. The script 
    maintains a complete and current company directory for all designated users.
    
.FUNCTIONALITY
    - Creates contacts for all licensed users in the tenant
    - Updates existing contacts when user information changes (optional)
    - Removes contacts for deprovisioned users (optional)
    - Supports exclusion lists for specific users
    - Configurable to include or exclude cloud only users (optional)
    - Optimized performance with fallback for compatibility
    
.NOTES
    Author:         Ryan Schultz
    Version:        2.3.2
    Creation Date:  March 2025
    
.PARAMETER TargetGroupId
    The Microsoft 365 group ID containing users who should receive the contacts
    
.PARAMETER ExclusionListVariableName
    The name of the Automation variable containing users to exclude from synchronization (line separated list)
    
.PARAMETER RemoveDeletedContacts
    Removes contacts that no longer exist in the source when set to $true (default is $true)
    
.PARAMETER UpdateExistingContacts
    Updates existing contacts with current information when set to $true (default is $true)

.PARAMETER IncludeExternalContacts
    Includes cloud only users in the contact synchronization when set to $true (default is $true)
    
.PARAMETER MaxConcurrentUsers
    Maximum number of concurrent users to process (default is 5)
    
.PARAMETER UseBatchOperations
    Whether to attempt using batch operations (will fall back to individual operations if needed)
#>

param(
    [Parameter(Mandatory = $true)]
    [string]$TargetGroupId, 
    [string]$ExclusionListVariableName = "ExclusionList",
    [bool]$RemoveDeletedContacts = $true,
    [bool]$UpdateExistingContacts = $true,
    [bool]$IncludeExternalContacts = $true,
    [int]$MaxConcurrentUsers = 5,
    [bool]$UseBatchOperations = $true,
    [Parameter(Mandatory = $false)]
    [string]$CharacterEncoding = "UTF-8"
)

# Global variables
$ErrorActionPreference = "Stop"
$script:GraphAccessToken = $null
$script:TokenAcquiredTime = $null
$script:TokenExpiresIn = 3600
$script:Throttled = $false
$script:BatchOperationsSupported = $UseBatchOperations

# Logging function
function Write-Log {
    param(
        [Parameter(Mandatory = $true)]
        [string]$Message,
        [ValidateSet("INFO", "WARNING", "ERROR")]
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
            "Content-Type" = "$ContentType; charset=utf-8"
            "ConsistencyLevel" = "eventual"
            "Accept-Charset" = "utf-8"
        }
        
        $params = @{
            Method = $Method
            Uri = "https://graph.microsoft.com/v1.0$Uri"
            Headers = $headers
        }
        
        if ($null -ne $Body -and $Method -ne "GET") {
            if ($ContentType -eq "application/json") {
                $params.Body = ConvertTo-Json -InputObject $Body -Depth 10 -Encoding UTF8
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
                elseif ($statusCode -eq 405 -and $Uri -eq "/$batch") {
                    $script:BatchOperationsSupported = $false
                    Write-Log "Batch operations not supported for this tenant. Switching to individual operations." -Level "WARNING"
                    throw $_
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
        if ($Uri -eq "/$batch" -and -not $script:BatchOperationsSupported) {
            throw [System.InvalidOperationException]::new("Batch operations not supported")
        }
        
        Write-Log "Graph API request failed: $($_.Exception.Message)" -Level "ERROR"
        throw $_
    }
}

function Execute-ContactOperations {
    param (
        [Parameter(Mandatory = $true)]
        [array]$Operations,
        
        [Parameter(Mandatory = $true)]
        [string]$UserId
    )
    
    try {
        if ($Operations.Count -eq 0) {
            return @()
        }
        
        if ($script:BatchOperationsSupported) {
            try {
                $batchRequests = @()
                foreach ($op in $Operations) {
                    $batchRequests += @{
                        id = [Guid]::NewGuid().ToString()
                        method = $op.Method
                        url = $op.Url
                        body = $op.Body
                        headers = @{ "Content-Type" = "application/json" }
                    }
                }
                
                $batchRequestBody = @{
                    requests = $batchRequests
                }
                
                $batchResponse = Invoke-GraphRequest -Method "POST" -Uri "/$batch" -Body $batchRequestBody
                
                return $batchResponse.responses
            }
            catch [System.InvalidOperationException] {
                Write-Log "Falling back to individual operations for user $UserId" -Level "INFO"
            }
            catch {
                Write-Log "Error with batch operations: $($_.Exception.Message). Falling back to individual operations." -Level "WARNING"
                $script:BatchOperationsSupported = $false
            }
        }
        
        $responses = @()
        
        foreach ($op in $Operations) {
            try {
                $uri = $op.Url
                $method = $op.Method
                $body = $op.Body
                
                $response = Invoke-GraphRequest -Method $method -Uri $uri -Body $body
                
                $responses += @{
                    id = [Guid]::NewGuid().ToString()
                    status = 200
                    body = $response
                }
            }
            catch {
                $responses += @{
                    id = [Guid]::NewGuid().ToString()
                    status = 500
                    body = @{ error = @{ message = $_.Exception.Message } }
                }
            }
            
            Start-Sleep -Milliseconds 50
        }
        
        return $responses
    }
    catch {
        Write-Log "Error executing contact operations: $($_.Exception.Message)" -Level "ERROR"
        throw $_
    }
}

# Data Retrieval Functions
function Get-ExclusionList {
    try {
        $exclusionListStr = Get-AutomationVariable -Name $ExclusionListVariableName
        
        if ($exclusionListStr) {
            $exclusionList = $exclusionListStr -split "`r`n|`r|`n" | Where-Object { $_ -notmatch "^\s*$" }
            Write-Log "Loaded exclusion list with $($exclusionList.Count) entries"
            return $exclusionList
        }
        else {
            Write-Log "Exclusion list variable not found or empty. No exclusions will be applied." -Level "INFO"
            return @()
        }
    }
    catch {
        $errorMessage = $_.Exception.Message
        Write-Log "Error loading exclusion list: $errorMessage" -Level "WARNING"
        return @()
    }
}

function Get-AllLicensedUsers {
    param(
        [string[]]$ExclusionList
    )
    
    try {
        Write-Log "Retrieving all licensed users from Microsoft Graph..."
        
        $normalizedExclusionList = @()
        if ($ExclusionList -and $ExclusionList.Count -gt 0) {
            $normalizedExclusionList = $ExclusionList | ForEach-Object { $_.ToLower().Trim() }
            Write-Log "Normalized $($normalizedExclusionList.Count) entries in exclusion list"
        }
        
        $filter = "userType eq 'Member'"
        if (-not $IncludeExternalContacts) {
            $filter += " and onPremisesSyncEnabled eq true"
        }
        
        $select = "id,userPrincipalName,displayName,givenName,surname,mail,jobTitle,department,businessPhones,mobilePhone,companyName,accountEnabled,assignedLicenses"
        $uri = "/users?`$filter=$([System.Web.HttpUtility]::UrlEncode($filter))&`$select=$select&`$top=999"
        
        $allUsers = @()
        
        do {
            $response = Invoke-GraphRequest -Method "GET" -Uri $uri
            
            if ($response -and $response.value) {
                $allUsers += $response.value
            }
            
            $uri = $null
            if ($response.'@odata.nextLink') {
                $uri = $response.'@odata.nextLink' -replace "https://graph.microsoft.com/v1.0", ""
            }
        } while ($uri)
        
        $licensedUsers = $allUsers | Where-Object { 
            $_.assignedLicenses.Count -gt 0 -and
            $_.accountEnabled -eq $true -and
            (
                ($normalizedExclusionList.Count -eq 0) -or 
                (
                    ($_.userPrincipalName -and ($_.userPrincipalName.ToLower().Trim() -notin $normalizedExclusionList)) -and
                    ((-not $_.mail) -or ($_.mail -and ($_.mail.ToLower().Trim() -notin $normalizedExclusionList)))
                )
            )
        }
        
        Write-Log "Found $($licensedUsers.Count) licensed users to be used as contacts"
        
        $contactsToSync = $licensedUsers | ForEach-Object {
            [PSCustomObject]@{
                Id = $_.id
                DisplayName = $_.displayName
                GivenName = $_.givenName
                Surname = $_.surname
                EmailAddress = $_.mail
                JobTitle = $_.jobTitle
                Department = $_.department
                BusinessPhone = if ($_.businessPhones.Count -gt 0) { $_.businessPhones[0] } else { "" }
                MobilePhone = $_.mobilePhone
                CompanyName = if ([string]::IsNullOrEmpty($_.companyName)) { "-" } else { $_.companyName }
            }
        }
        
        return $contactsToSync
    }
    catch {
        Write-Log "Error retrieving licensed users: $($_.Exception.Message)" -Level "ERROR"
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

# Contact Management Functions
function New-ContactObject {
    param(
        [Parameter(Mandatory = $true)]
        [PSCustomObject]$ContactData
    )

    # Ensure all text fields are properly encoded
    # $displayName = [System.Web.HttpUtility]::HtmlEncode($ContactData.DisplayName)
    # $givenName = [System.Web.HttpUtility]::HtmlEncode($ContactData.GivenName)
    # $surname = [System.Web.HttpUtility]::HtmlEncode($ContactData.Surname)
    
    $contactObject = @{
        givenName = $ContactData.GivenName
        surname = $ContactData.Surname
        displayName = $ContactData.DisplayName
        companyName = $ContactData.CompanyName
        jobTitle = $ContactData.JobTitle
        mobilePhone = $ContactData.MobilePhone
        businessPhones = @()
        emailAddresses = @()
        department = $ContactData.Department
        categories = @("Company Contacts")
    }
    
    if (-not [string]::IsNullOrEmpty($ContactData.BusinessPhone)) {
        $contactObject.businessPhones += $ContactData.BusinessPhone
    }
    
    if (-not [string]::IsNullOrEmpty($ContactData.EmailAddress)) {
        $contactObject.emailAddresses += @{
            address = $ContactData.EmailAddress
            name = $ContactData.DisplayName
        }
    }
    
    return $contactObject
}

function Get-ContactHash {
    param(
        [Parameter(Mandatory = $true)]
        [object]$Contact
    )
    
    $hashStr = "$($Contact.givenName)#$($Contact.surname)#$($Contact.jobTitle)#$($Contact.department)#$($Contact.companyName)"
    
    if ($Contact.businessPhones -and $Contact.businessPhones.Count -gt 0) {
        $hashStr += "#$($Contact.businessPhones[0])"
    }
    
    if (-not [string]::IsNullOrEmpty($Contact.mobilePhone)) {
        $hashStr += "#$($Contact.mobilePhone)"
    }
    
    return $hashStr
}

function Sync-UserContacts {
    param(
        [Parameter(Mandatory = $true)]
        [object]$User,
        
        [Parameter(Mandatory = $true)]
        [PSCustomObject[]]$ContactsToSync
    )
    
    try {
        Write-Log "Starting contact sync for user $($User.userPrincipalName)..."
        
        $existingContacts = @()
        $uri = "/users/$($User.id)/contacts?`$top=999"
        
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
        
        $existingContactMap = @{}
        foreach ($contact in $existingContacts) {
            $key = ""
            
            if ($contact.emailAddresses.Count -gt 0) {
                $key = $contact.emailAddresses[0].address.ToLower()
                if (-not [string]::IsNullOrEmpty($key)) {
                    $existingContactMap[$key] = $contact
                }
            }
        }
        
        $userContacts = $ContactsToSync | Where-Object { $_.Id -ne $User.id }
        
        $operations = @()
        $createdCount = 0
        $updatedCount = 0
        $deletedCount = 0
        $unchangedCount = 0
        
        # Process contacts to create or update
        foreach ($contactData in $userContacts) {
            # Skip contacts without email
            if ([string]::IsNullOrEmpty($contactData.EmailAddress)) {
                continue
            }
            
            $emailKey = $contactData.EmailAddress.ToLower()
            $contactObject = New-ContactObject -ContactData $contactData
            
            if ($existingContactMap.ContainsKey($emailKey)) {
                $existingContact = $existingContactMap[$emailKey]
                
                $existingContactMap.Remove($emailKey)
                
                if ($UpdateExistingContacts) {
                    $existingHash = Get-ContactHash -Contact $existingContact
                    $newHash = Get-ContactHash -Contact $contactObject
                    
                    if ($existingHash -ne $newHash) {
                        $operations += @{
                            Method = "PATCH"
                            Url = "/users/$($User.id)/contacts/$($existingContact.id)"
                            Body = $contactObject
                        }
                        $updatedCount++
                    }
                    else {
                        $unchangedCount++
                    }
                }
                else {
                    $unchangedCount++
                }
            }
            else {
                $operations += @{
                    Method = "POST"
                    Url = "/users/$($User.id)/contacts"
                    Body = $contactObject
                }
                $createdCount++
            }
        }
        
        if ($RemoveDeletedContacts -and $existingContactMap.Count -gt 0) {
            foreach ($key in $existingContactMap.Keys) {
                $contactToRemove = $existingContactMap[$key]
                
                if ($contactToRemove.categories -contains "Company Contacts") {
                    $operations += @{
                        Method = "DELETE"
                        Url = "/users/$($User.id)/contacts/$($contactToRemove.id)"
                        Body = $null
                    }
                    $deletedCount++
                }
            }
        }
        
        if ($operations.Count -gt 0) {
            $responses = Execute-ContactOperations -Operations $operations -UserId $User.id
            $errors = $responses | Where-Object { $_.status -ge 400 }
            if ($errors -and $errors.Count -gt 0) {
                Write-Log "Some contact operations failed for $($User.userPrincipalName): $($errors.Count) errors" -Level "WARNING"
            }
        }
        
        Write-Log "Completed sync for $($User.userPrincipalName): Created=$createdCount, Updated=$updatedCount, Deleted=$deletedCount, Unchanged=$unchangedCount"
        
        return @{
            Created = $createdCount
            Updated = $updatedCount
            Deleted = $deletedCount
            Unchanged = $unchangedCount
        }
    }
    catch {
        Write-Log "Error syncing contacts for $($User.userPrincipalName): $($_.Exception.Message)" -Level "ERROR"
        throw $_
    }
}

# Main execution
try {
    $startTime = Get-Date
    Write-Log "Starting optimized ContactSync process"
    Write-Log "Parameters: RemoveDeletedContacts=$RemoveDeletedContacts, UpdateExistingContacts=$UpdateExistingContacts, IncludeExternalContacts=$IncludeExternalContacts, MaxConcurrentUsers=$MaxConcurrentUsers, UseBatchOperations=$UseBatchOperations"
    
    Connect-ToMicrosoftGraph
    
    $exclusionList = Get-ExclusionList
    
    $allUsersAsContacts = Get-AllLicensedUsers -ExclusionList $exclusionList
    
    if ($allUsersAsContacts.Count -eq 0) {
        Write-Log "No licensed users found to use as contacts. Exiting." -Level "WARNING"
        exit
    }
    
    $targetUsers = Get-GroupMembers -GroupId $TargetGroupId
    
    if ($targetUsers.Count -eq 0) {
        Write-Log "No users found in the target group. Exiting." -Level "WARNING"
        exit
    }
    
    $totalUsers = $targetUsers.Count
    $processedCount = 0
    $waitingCount = $totalUsers
    $concurrentCount = 0
    $completedCount = 0
    $errorCount = 0
    
    Write-Log "Processing $totalUsers users with max $MaxConcurrentUsers concurrent jobs"
    
    foreach ($currentUser in $targetUsers) {
        $processedCount++
        $waitingCount--
        
        if ([string]::IsNullOrEmpty($currentUser.userPrincipalName)) {
            Write-Log "Skipping user $processedCount of $totalUsers - Invalid user information" -Level "WARNING"
            $errorCount++
            continue
        }
        
        Write-Log "Starting sync for user $processedCount of $totalUsers : $($currentUser.userPrincipalName)"
        
        try {
            $result = Sync-UserContacts -User $currentUser -ContactsToSync $allUsersAsContacts
            $completedCount++
            
            Write-Log "Completed user $($currentUser.userPrincipalName): Created=$($result.Created), Updated=$($result.Updated), Deleted=$($result.Deleted), Unchanged=$($result.Unchanged)"
        }
        catch {
            $errorCount++
            Write-Log "Error processing user $($currentUser.userPrincipalName): $($_.Exception.Message)" -Level "ERROR"
        }
        
        Write-Log "Progress: $completedCount completed, $errorCount errors, $waitingCount waiting"
        
        if ($script:Throttled) {
            Write-Log "Detected API throttling, pausing for 5 seconds..." -Level "WARNING"
            Start-Sleep -Seconds 5
        }
    }
    
    $endTime = Get-Date
    $duration = $endTime - $startTime
    
    Write-Log "ContactSync completed successfully in $($duration.TotalMinutes) minutes"
    Write-Log "Final stats: $completedCount users processed successfully, $errorCount users with errors"
}
catch {
    $errorMessage = $_.Exception.Message
    Write-Log "ContactSync failed with error: $errorMessage" -Level "ERROR"
    throw $_
}