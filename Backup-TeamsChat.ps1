<#

.SYNOPSIS
 
    Backup-TeamsChat.ps1 - Backup Teams chat messages
    https://www.lee-ford.co.uk/backup-teamschat
 
.DESCRIPTION
    Author: Lee Ford

    This tool allows you to backup Teams chat messages (not channel messages) for safe keeping. See https://www.github.com/leeford/Backup-TeamsChat for latest details.

    This tool has been written for PowerShell "Core" on Windows, Mac and Linux - it will not work with "Windows PowerShell".

    Note: A predefined Azure AD application has been used, but you can use your own with the same scope.

.LINK
    Blog: https://www.lee-ford.co.uk
    Twitter: http://www.twitter.com/lee_ford
    LinkedIn: https://www.linkedin.com/in/lee-ford/
 
.EXAMPLE 
    
    To backup your chat messages:
    Backup-TeamsChat.ps1 -Path <folder to store backup>

#>

Param (

    [Parameter(mandatory = $true)][string]$Path

)

# Application (client) ID, resource and scope
$script:clientId = "56a68496-f8f3-41c6-abac-eaa2b276e736"
$script:scope = "Chat.Read"
$script:tenantId = "common"

function Invoke-GraphAPICall {

    param (

        [Parameter(mandatory = $true)][uri]$URI,
        [Parameter(mandatory = $false)][switch]$WriteStatus,
        [Parameter(mandatory = $false)][string]$Method,
        [Parameter(mandatory = $false)][string]$Body

    )

    # Is method speficied (if not assume GET)
    if ([string]::IsNullOrEmpty($method)) { $method = 'GET' }

    # Access token still valid?
    $currentEpoch = [int][double]::Parse((Get-Date (get-date).ToUniversalTime() -UFormat %s))

    if ($currentEpoch -gt [int]$script:token.expires_on) {

        Refresh-UserToken

    }

    $Headers = @{"Authorization" = "Bearer $($script:token.access_token)" }

    $currentUri = $URI

    $content = while (-not [string]::IsNullOrEmpty($currentUri)) {

        # API Call
        $apiCall = try {
            
            Invoke-RestMethod -Method $method -Uri $currentUri -ContentType "application/json" -Headers $Headers -Body $body -ResponseHeadersVariable script:responseHeaders

        }
        catch {
            
            $errorMessage = $_.ErrorDetails.Message | ConvertFrom-Json

        }
        
        $currentUri = $null
    
        if ($apiCall) {
    
            # Check if any data is left
            if ($apiCall.'@odata.count' -gt 0) {

                $currentUri = $apiCall.'@odata.nextLink'

            }

            $apiCall

        }

    }

    if ($WriteStatus) {

        # If error returned
        if ($errorMessage) {

            Write-Host "FAILED $($errormessage.error.message)" -ForegroundColor Red

        }
        else {

            Write-Host "SUCCESS" -ForegroundColor Green

        }
        
    }

    return $content
    
}

function Get-UserToken {

    $script:token = $null

    $resource = "https://graph.microsoft.com/"

    $codeBody = @{ 

        resource  = $resource
        client_id = $script:clientId
        scope     = $script:scope
        

    }

    # Get OAuth Code
    $codeRequest = Invoke-RestMethod -Method POST -Uri "https://login.microsoftonline.com/$script:tenantId/oauth2/devicecode" -Body $codeBody

    # Print Code to host
    Write-Host "`n$($codeRequest.message)"

    $tokenBody = @{

        grant_type = "urn:ietf:params:oauth:grant-type:device_code"
        code       = $codeRequest.device_code
        client_id  = $clientId

    }

    # Get OAuth Token
    while ([string]::IsNullOrEmpty($tokenRequest.access_token)) {

        $tokenRequest = try {

            Invoke-RestMethod -Method POST -Uri "https://login.microsoftonline.com/$script:tenantId/oauth2/token" -Body $tokenBody

        }
        catch {

            $errorMessage = $_.ErrorDetails.Message | ConvertFrom-Json

            # If not waiting for auth, throw error
            if ($errorMessage.error -ne "authorization_pending") {

                Throw

            }

        }

    }

    $script:token = $tokenRequest

}

function Refresh-UserToken {
    param (
        
    )

    $refreshBody = @{

        client_id     = $script:clientId
        scope         = "$script:scope offline_access" # Add offline_access to scope to ensure refresh_token is issued
        grant_type    = "refresh_token"
        refresh_token = $script:token.refresh_token

    }

    $tokenRequest = try {

        Invoke-RestMethod -Method POST -Uri "https://login.microsoftonline.com/$script:tenantId/oauth2/token" -Body $refreshBody

    }
    catch {

        Throw
    
    }

    $script:token = $tokenRequest

}

function Get-Chats {
    param (

        [Parameter(mandatory = $true)][object]$User

    )

    Write-Host "Processing User: $($user.displayName) - $($user.mail)" -ForegroundColor Yellow

    # Create Chat folder
    $date = Get-Date -UFormat "%Y-%m-%d %H%M"
    $folder = "$($user.mail)_$date" -replace '([\\/:*?"<>|\s])+', "_"
    $fullPath = "$path/$folder"

    Write-Host " - Creating folder $fullPath..." 
    try {
        New-Item -ItemType Directory -Force -Path $fullPath | Out-Null
        Write-Host "SUCCESS" -ForegroundColor Green

    }
    catch {

        Write-Host "FAILED" -ForegroundColor Red

    }

    Write-Host " - Getting chats..." -NoNewline
    $chats = Invoke-GraphAPICall -URI "https://graph.microsoft.com/beta/$($user.Id)/chats" -Method "GET" -WriteStatus

    # Loop through each chat thread and get messages, members etc.
    $chats.value | ForEach-Object {

        Write-Host " - Getting chat - $($_.id) - $($_.topic)..."

        # Get Members
        Write-Host "    - Getting chat members - $($_.id)..." -NoNewline
        $members = Invoke-GraphAPICall -URI "https://graph.microsoft.com/beta/$($user.Id)/chats/$($_.id)/members" -WriteStatus
        $chatMembers = $members.value.displayName -join ", "

        # Get Messages
        Write-Host "    - Getting chat messages - $($_.id)..." -NoNewline
        $messages = Invoke-GraphAPICall -URI "https://graph.microsoft.com/beta/$($user.Id)/chats/$($_.id)/messages" -WriteStatus

        $htmlMessages = "<br /><h3>Chat Transcript:</h3>"

        $messages.value | Sort-Object -Property createdDateTime | ForEach-Object {

            $important = Check-MessageImportance $_

            $htmlMessages += "
                            <div class='card'>
                                <div class='card-header bg-light'><b>$($_.from.user.displayName)</b> $($_.createdDateTime)</div>
                                $important
                                <div class='card-body'>
                                    <h4 class='card-title'>$($_.subject)</h4>
                                    <p>$($_.body.content)</p>
                                </div>
                                "
            $htmlMessages += "</div><br />"

        }

        $html = "
        <br /><div class='card'>
            <h5 class='card-header bg-light'>Chat Overview</h5>
            <div class='card-body'>
            <table class='table table-borderless'>
            <tbody>
                <tr>
                    <th scope='row'>Backup Date:</th>
                    <td>$date</td>
                </tr>
                <tr>
                    <th scope='row'>Topic:</th>
                    <td>$($_.topic)</td>
                </tr>
                <tr>
                    <th scope='row'>Members:</th>
                    <td>$chatMembers</td>
                </tr>
            </tbody>
          </table>
        </div>
        </div>
        $htmlMessages
"
        # Save file with topic name
        if ($_.topic) {

            $file = "Chat_$($_.topic)" -replace '([\\/:*?"<>|\s])+', "_"

        # Or members if no topic
        }
        else {

            $file = "Chat_$($members.value.displayName -join "_")" -replace '([\\/:*?"<>|\s])+', "_"

        }

        Write-Host "    - Saving chat to $fullPath/$file.htm... " -NoNewline
        Create-HTMLPage -Content $html -PageTitle "$($user.displayName) - Chat with $chatMembers" -Path "$fullPath/$file.htm"

    }

}

function Create-HTMLPage {
    param (

        [Parameter(mandatory = $true)][string]$Content,
        [Parameter(mandatory = $true)][string]$PageTitle,
        [Parameter(mandatory = $true)][string]$Path

    )

    $html = "
    <div class='p-0 m-0'>
        <div class='container m-3'>
            <div class='page-header'>
                <h1>$pageTitle</h1>
                <h5>Created with <a href='https://www.lee-ford.co.uk/Backup-TeamsChat'>Backup-TeamsChat</a></h5>
            </div>

            $Content

            </div>
    </div>"

    try {
            
        ConvertTo-Html -CssUri "https://stackpath.bootstrapcdn.com/bootstrap/4.3.1/css/bootstrap.min.css" -Body $html -Title $PageTitle | Out-File $Path
        Write-Host "SUCCESS" -ForegroundColor Green

    }
    catch {

        Write-Host "FAILED" -ForegroundColor Red

    }

}

function Check-MessageImportance {

    param (

        [Parameter(mandatory = $true)][System.Object]$message

    )

    if ($message.importance -eq "high") {

        return "<div class='alert alert-danger m-2' role='alert'>IMPORTANT!</div>"
        
    }
    else {
        
        return $null

    }

}

Write-Host "`n----------------------------------------------------------------------------------------------
            `n Backup-TeamsChat.ps1 - Lee Ford - https://www.lee-ford.co.uk/
            `n----------------------------------------------------------------------------------------------" -ForegroundColor Yellow

# Get Azure AD User Token
Get-UserToken | Out-Null

# Get logged in User
$script:me = Invoke-GraphAPICall "https://graph.microsoft.com/v1.0/me"

if ($script:me.id) {

    Write-Host "`nSIGNED-IN as $($me.mail)" -ForegroundColor Green

    $myChats = Get-Chats -User $me

}