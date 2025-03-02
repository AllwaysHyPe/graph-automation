# Ensures you do not inherit an AzContext in your runbook
Disable-AzContextAutosave -Scope Process

# Function to authenticate and retrieve the access token
function Get-AccessToken {
    $invokeRestMethodSplat = @{
        Uri     = "https://login.microsoftonline.com/{0}/oauth2/v2.0/token" -f $env:TENANTID
        Method  = "POST"
        Headers = @{ "Content-Type" = "application/x-www-form-urlencoded" }
        Body    = @{
            client_id     = $env:CLIENTID
            client_secret = $env:APP_SECRET
            scope         = "https://graph.microsoft.com/.default"
            grant_type    = "client_credentials"
        }
    }

    # Retrieve access token
    $access_token = Invoke-RestMethod @invokeRestMethodSplat
    return $access_token.access_token
}

function Invoke-MgGraphCall {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory=$True)]
        [string]$URI,
        [Parameter(Mandatory=$True)]
        [string]$Method,
        [Parameter(Mandatory=$False)]
        [string]$Body
        
    )

    $AccessToken = Get-AccessToken

    # Create Splat hash table for API call
    $graphSplatParams = @{
        Headers      = @{
            "Content-Type" = "application/json"
            "Authorization" = "Bearer $($AccessToken)"
        }
        Method = $Method
        URI = $URI
        ErrorAction = "SilentlyContinue"
        StatusCodeVariable = "scv"
    }

    # If method requires body, add body to splat
    if($Method -in('PUT', 'PATCH', 'POST')){
        $graphSplatParams["Body"] = $Body

    }

    # Make API call using Invoke-RestMethod
    $MgGraphResult = Invoke-RestMethod @graphSplatParams

    #Return status code and result for the API call
    return $SCV, $MgGraphResult
}

# Function to get users with pagination handling
function Get-UsersData {
    param (
        [string]$Uri
    )

    $AllUserData = [System.Collections.ArrayList]::new()
    $UserData = Invoke-MgGraphCall -URI $Uri -Method "GET"

    do {
        $UserBatch = $UserData.value
        $AllUserData.AddRange($UserBatch)
        $NextLink = $UserData.'@odata.nextLink'

        if ($NextLink -ne $null) {
            $UserData = Invoke-MgGraphCall -URI $NextLink -Method "GET"
        }
    } while ($NextLink -ne $null)

    return $AllUserData
}

function Send-BatchRequest {
    param (
        [array]$Requests
    )

    $BatchBody = @{ "requests" = $Requests } | ConvertTo-Json -Depth 3
    $BatchUri = "https://graph.microsoft.com/v1.0/`$batch"

    $MaxRetries = 3
    $RetryDelay = 5  # Seconds

    for ($Attempt = 1; $Attempt -le $MaxRetries; $Attempt++) {
        try {
            return Invoke-MgGraphCall -URI $BatchUri -Method "POST" -Body $BatchBody
        } catch {
            Write-Host "Batch request failed (Attempt $Attempt). Retrying in $RetryDelay seconds..."
            Start-Sleep -Seconds $RetryDelay
        }
    }

    Write-Host "Batch request failed after $MaxRetries attempts."
    return $null
}

function Get-UserPhotos {
    param (
        [System.Collections.ArrayList]$Users
    )

    $BatchSize = 20
    $BatchRequests = [System.Collections.ArrayList]::new()
    $UserPhotos = [System.Collections.ArrayList]::new()
    $PhotoDictionary = @{}
    $ProcessedUsers = @{}  # Hash table to prevent duplicate API calls

    $Counter = 1

    foreach ($User in $Users) {
        if ($User.id -ne $null -and -not $ProcessedUsers.ContainsKey($User.id)) {
            $ProcessedUsers[$User.id] = $true  # Mark user as processed

            $BatchRequest = @{
                "id"     = "$Counter"
                "method" = "GET"
                "url"    = "https://graph.microsoft.com/v1.0/users/$($User.id)/photo/`$value"
            }
            $BatchRequests.Add($BatchRequest)
        }

        if ($BatchRequests.Count -ge $BatchSize -or $User -eq $Users[$Users.Count - 1]) {
            Write-Host "Sending batch request for $($BatchRequests.Count) user photos..."
            $BatchResponses = Send-BatchRequest -Requests $BatchRequests

            foreach ($Response in $BatchResponses.responses) {
                if ($Response.PSObject.Properties['id'] -and $Response.PSObject.Properties['status']) {
                    $UserId = $Response.id
                    $StatusCode = $Response.status

                    if (![string]::IsNullOrEmpty($UserId) -and -not $PhotoDictionary.ContainsKey($UserId)) {
                        $UserDetails = $Users | Where-Object { $_.id -eq $UserId }
                        $HasPhoto = if ($StatusCode -eq 200) { $true } else { $false }

                        if ($null -ne $UserDetails) {
                            $PhotoData = [PSCustomObject]@{
                                Id          = $UserId
                                DisplayName = $UserDetails.DisplayName
                                Email       = $UserDetails.UserPrincipalName
                                HasPhoto    = $HasPhoto
                            }

                            Write-Host "Adding User to PhotoDictionary: $UserId - $($UserDetails.DisplayName) - HasPhoto: $HasPhoto"
                            $UserPhotos.Add($PhotoData)
                            $PhotoDictionary[$UserId] = $PhotoData
                        } else {
                            Write-Host "Skipping: User ID $UserId not found in Users list"
                        }
                    } else {
                        Write-Host "Skipping invalid or duplicate user: $UserId"
                    }
                } else {
                    Write-Host "Warning: Unexpected response structure from Graph API."
                    Write-Host ($Response | ConvertTo-Json -Depth 3)
                }
            }

            $BatchRequests.Clear()
        }

        $Counter++
    }

    # Debugging: Ensure photos were actually stored
    Write-Host "Users in PhotoDictionary: $($PhotoDictionary.Count)"
    return $PhotoDictionary
}

function Update-UserPhotos {
    param (
        [System.Collections.ArrayList]$Users,
        [System.Collections.ArrayList]$UserPhotos,
        [string]$PhotoDirectory
    )

    $MatchedUsers = [System.Collections.ArrayList]::new()
    $BatchSize = 20
    $BatchRequests = [System.Collections.ArrayList]::new()
    $Counter = 1
    $GraphBaseUrl = "https://graph.microsoft.com/v1.0"

    # Debugging: Confirm no disabled accounts are in UserDictionary
    Write-Host "Total Users in UserDictionary: $($UserDictionary.Count)"
    Write-Host "Disabled Users in UserDictionary: $(($UserDictionary.Values | Where-Object { $_.AccountEnabled -eq $false }).Count)"

    # Dictionary for photo metadata (ensures only enabled users are considered)
    $PhotoDictionary = @{}
    foreach ($UserPhoto in $UserPhotos) {
        if ($UserPhoto.Id -ne $null -and $UserDictionary.ContainsKey($UserPhoto.Id)) {
            $PhotoDictionary[$UserPhoto.Id] = $UserPhoto
        }
    }

    # Debugging: Ensure that only users from Get-UserPhotos exist in the dictionary
    Write-Host "Users in PhotoDictionary: $($PhotoDictionary.Count)"

    # Process photo files
    $PhotoFiles = Get-ChildItem -Path "$PhotoDirectory\*.jpg"

    foreach ($File in $PhotoFiles) {
        # Normalize filename to match DisplayName-based lookups
        $NormalizedFileName = $File.BaseName.Replace("'", "''") # Escape single quotes for OData

        # Find user by ID instead of DisplayName (ensuring correct lookup)
        $User = $null
        foreach ($Key in $UserDictionary.Keys) {
            if ($UserDictionary[$Key].DisplayName.ToLower() -eq $NormalizedFileName) {
                $User = $UserDictionary[$Key]
                break
            }
        }

        if ($User -ne $null -and $User.id -ne $null) {
            if ($PhotoDictionary.ContainsKey($User.id)) {
                $UserPhoto = $PhotoDictionary[$User.id]

                if ($UserPhoto -ne $null -and -not $UserPhoto.HasPhoto) {
                    $FileBytes = [System.IO.File]::ReadAllBytes($File.FullName)
                    $Base64String = [Convert]::ToBase64String($FileBytes)

                    $BatchRequest = @{
                        "id"     = "$Counter"
                        "method" = "PUT"
                        "url"    = "$GraphBaseUrl/users/$($User.id)/photo/`$value"
                        "headers"= @{ "Content-Type" = "image/jpeg" }
                        "body"   = $Base64String
                    }
                    $BatchRequests.Add($BatchRequest)

                    $MatchedUser = [PSCustomObject]@{
                        Id              = $User.id
                        DisplayName     = $User.DisplayName
                        Email           = $User.UserPrincipalName
                        CreatedDateTime = $User.createdDateTime
                        AccountEnabled  = $User.AccountEnabled
                        PhotoPath       = $File.FullName
                    }

                    $MatchedUsers.Add($MatchedUser)
                } else {
                    Write-Host "Skipping user '$NormalizedFileName' - already has a photo."
                }
            } else {
                Write-Host "Error: User '$NormalizedFileName' (ID: $($User.id)) not found in photo dictionary."
            }
        } else {
            Write-Host "Warning: No matching user found for file '$($File.Name)'."
        }

        if ($BatchRequests.Count -ge $BatchSize -or $File -eq $PhotoFiles[$PhotoFiles.Count - 1]) {
            if ($BatchRequests.Count -gt 0) {
                Write-Host "Sending batch update for $($BatchRequests.Count) users..."
                Send-BatchRequest -Requests $BatchRequests
                $BatchRequests.Clear()
            } else {
                Write-Host "Skipping empty batch update."
            }
        }

        $Counter++
    }

    return $MatchedUsers
}


# Set Paths
$PhotoDirectory = "/photos/in_progress"
$Destination = "/photos/completed"
$ReportPath = "/photos/completed/Report.csv"

#Get All Users (excluding guests)
$Uri = "https://graph.microsoft.com/v1.0/users?`$filter=accountEnabled eq true and usertype eq 'Member'&`$select=id,displayName,userPrincipalName,userType,accountEnabled,createdDateTime"
$Users = Get-UsersData -Uri $Uri

# Get all user photos 
$UserPhotos = Get-UserPhotos -Users $Users

# Match users with local photos and update Graph
$UpdatedUsers = Update-UserPhotos -Users $Users -UserPhotos $UserPhotos -PhotoDirectory $PhotoDirectory

