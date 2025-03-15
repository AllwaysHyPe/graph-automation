# Import required modules
Import-Module Az.Accounts

# Connect using Managed Identity
try {
    Connect-AzAccount -Identity -ErrorAction Stop | Out-Null
    $AccessToken = Get-AzAccessToken -ResourceUrl "https://graph.microsoft.com" -ErrorAction Stop
    Write-Output "Successfully authenticated using Managed Identity!"
} catch {
    Write-Error "Failed to authenticate using Managed Identity: $_" -ErrorAction Stop
    Exit
}

# Define Headers for Graph API requests
$Headers = @{
    "Authorization" = "Bearer $($AccessToken.Token)"
    "Content-Type" = "application/json"
}

Write-Output $Headers

# Get All Users (excluding guests)
$Uri = "https://graph.microsoft.com/beta/users?`$filter=accountEnabled eq true and usertype eq 'Member'&`$select=id,displayName,userPrincipalName,userType,accountEnabled,createdDateTime"
$UsersTable = [ordered]@{}
$RequestUri = $Uri

Write-Output "Fetching users..."

while ($RequestUri) {
    # Make Graph API call
    $Response = Invoke-RestMethod -Uri $RequestUri -Headers $Headers -Method Get

    # Store users in ordered hash table
    foreach ($User in $Response.value) {
        $UsersTable[$User.id] = $User
    }

    Write-Output "Retrieved $($Response.value.Count) users. Total so far: $($UsersTable.Count)"

    # Check for additional page of results
    $RequestUri = $Response.'@odata.nextLink'
}

Write-Output "Total Users Retrieved: $($UsersTable.Count)"



$BasePhotoPath = "C:\Photos\InProgress\"
$ExistingPhotos = Get-ChildItem -Path $BasePhotoPath -Filter "*.jpg" | Select-Object -ExpandProperty Name

Write-Output "Total photos available in directory: $($ExistingPhotos.Count)"

# Find users with missing photos, but have a matching photo in folder

$UsersToPatch = $UsersTable.Keys | Where-Object {
    -not $UsersWithPhotos.Contains($_) -and ("$($UsersTable[$_].displayName).jpg" -in $ExistingPhotos)
}

Write-Output "Users to be patched: $($UsersToPatch.Count)"

# Patch users that need photos
$UsersPatched = [ordered]@{}

foreach ($UserID in $UsersToPatch) {
    $User = $UsersTable[$UserID]
    $UserPhotoPath = "$BasePhotoPath$($User.displayName).jpg"
    
    if (Test-Path $UserPhotoPath) {
        # Convert photo to Base64
        $PhotoBytes = [System.IO.File]::ReadAllBytes($UserPhotoPath)
        Write-Output "Binary Data Debugging for $($User.displayName) ($UserID)"
        Write-Output "File Size: $($PhotoBytes.Length) bytes"
        Write-Output "First 10 Bytest (Hex): $([BitConverter]::ToString($PhotoBytes[0..9]))"
        $EncodedPhoto = [Convert]::ToBase64String($PhotoBytes)

        # Create PATCH request
        $PatchBody = $EncodedPhoto | ConvertTo-Json -Depth 10

        $PatchRequestProperties = @{
            Uri = "https://graph.microsoft.com/beta/users/$UserID/photo/`$value"
            Method = "PUT"
            Headers = @{
                "Authorization" = "Bearer $($AccessToken.Token)"
                "Content-Type" = "image/jpeg"
            }
            Body = $PhotoBytes
        }
        
        try {
            Invoke-RestMethod @PatchRequestProperties
            Write-Host "Successfully updated photo for $($User.displayName) ($UserID)"
            
            # Store successful patch attempt
            $UsersPatched[$UserID] = [PSCustomObject]@{
                UserID = $UserID
                DisplayName = $User.displayName
                Email = $User.userPrincipalName
                PhotoPath = $UserPhotoPath
            }
        } catch {
            Write-Output "Failed to update photo for $($User.displayName) ($UserID): $_.Exception.Message"
            Write-Output "Full API Response: $($_ | ConvertTo-JSON -Depth 10)"
        }
    } else {
        Write-Output "Photo not found for $($User.displayName) ($UserID)"
    }
}

Write-Host "Total Users Successfully Patched: $($UsersPatched.Count)"
