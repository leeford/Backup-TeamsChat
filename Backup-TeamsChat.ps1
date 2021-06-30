<#

.SYNOPSIS
 
    Backup-TeamsChat.ps1 - Backup Teams chat messages
    https://github.com/leeford/Backup-TeamsChat
 
.DESCRIPTION
    Author: Lee Ford

    This tool allows you to backup Teams chat messages (not channel messages) for safe keeping. See https://www.github.com/leeford/Backup-TeamsChat for latest details.

    This tool has been written for PowerShell "Core" on Windows, Mac and Linux - it will not work with "Windows PowerShell".

.LINK
    Blog: https://www.lee-ford.co.uk
    Twitter: http://www.twitter.com/lee_ford
    LinkedIn: https://www.linkedin.com/in/lee-ford/
 
.EXAMPLE 
    
    To backup all chat messages for tenant:
    Backup-TeamsChat.ps1 -Path <directory to store backup>

#>

Param (
    [Parameter(mandatory = $true)][string]$Path,
    [Parameter(mandatory = $false)][int]$Days,
    [Parameter(mandatory = $false)][string]$User
)

function Check-ModuleInstalled {
    param (
        [Parameter (mandatory = $true)][String]$module,
        [Parameter (mandatory = $true)][String]$moduleName
    )

    # Do you have module installed?
    Write-Host "`n`rChecking $moduleName installed..." -NoNewline

    if (Get-Module -ListAvailable -Name $module) {
        Write-Host " INSTALLED" -ForegroundColor Green
    }
    else {
        Write-Host " NOT INSTALLED" -ForegroundColor Red
        break
    }
    
}

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
    if ($currentEpoch -gt $script:tokenExpiresOn) {
        Get-ApplicationToken
    }

    $maxRetries = 15
    $retryIntervalSec = 3
    $retryCount = 0

    $headers = @{"Authorization" = "Bearer $($script:token.access_token)" }
    $currentUri = $URI
    $content = while (-not [string]::IsNullOrEmpty($currentUri)) {
        # API Call
        try {
            $response = Invoke-RestMethod -Method $method -Uri $currentUri -ContentType "application/json" -Headers $headers -Body $body -ResponseHeadersVariable script:responseHeaders

            if ($response) {
                # Check if any data is left
                if ($response.'@odata.count' -gt 0) {
                    # Set URI to nextLink
                    $currentUri = $response.'@odata.nextLink'
                    # Reset retry counter
                    $retryCount = 0
                }
                else {
                    $currentUri = $null
                }
                $response
            }
        }
        catch {
            if (($_.Exception.Response.StatusCode -eq 403) -or ($retryCount -ge $maxRetries)) {
                # Max retries reached or forbidden
                $errorMessage = $_.ErrorDetails.Message | ConvertFrom-Json
                $currentUri = $null
            }
            else {
                $retryCount += 1
                Start-Sleep -Seconds $retryIntervalSec
            }
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

function Get-ApplicationToken {

    $script:token = $null

    # Construct URI
    $uri = "https://login.microsoftonline.com/$tenantId/oauth2/v2.0/token"

    # Construct Body
    $body = @{
        client_id     = $clientId
        scope         = "https://graph.microsoft.com/.default"
        client_secret = $clientSecret
        grant_type    = "client_credentials"
    }

    # Get OAuth 2.0 Token
    $tokenRequest = Invoke-WebRequest -Method Post -Uri $uri -ContentType "application/x-www-form-urlencoded" -Body $body -UseBasicParsing

    # Access Token
    $script:token = ($tokenRequest.Content | ConvertFrom-Json)

    # Note token expiry (minus 3 minutes)
    $script:tokenExpiresOn = $script:token.expires_in + [int][double]::Parse((Get-Date (get-date).ToUniversalTime() -UFormat %s)) - 180

}

function Get-Chats {
    param (
        [Parameter(mandatory = $true)][object]$userObject,
        [Parameter(mandatory = $true)][string]$rootDirectory
    )

    Write-Host "Processing User: $($userObject.displayName) - $($userObject.userPrincipalName)" -ForegroundColor Yellow

    # Create Chat directory for user
    $directory = "$($userObject.userPrincipalName)" -replace '([\\/:*?"<>|\s])+', "_"
    $userPath = "$rootDirectory/$directory"

    $chatThreads = @()
    if ($Days) {
        $fromDateTime = (Get-Date).AddDays(-$Days) | Get-Date -AsUTC -Format o
        $toDateTime = (Get-Date).AddDays(+1) | Get-Date -AsUTC -Format o
        Write-Host " - Getting chat messages for last $Days days..." -NoNewline
        $chatMessages = Invoke-GraphAPICall -URI "https://graph.microsoft.com/beta/users/$($userObject.id)/chats/getAllMessages?`$filter=lastModifiedDateTime gt $fromDateTime and lastModifiedDateTime lt $toDateTime" -Method "GET" -WriteStatus
    }
    else {
        Write-Host " - Getting chat messages..." -NoNewline
        $chatMessages = Invoke-GraphAPICall -URI "https://graph.microsoft.com/beta/users/$($userObject.id)/chats/getAllMessages" -Method "GET" -WriteStatus
    }
    # Loop through each chat thread and get messages, members etc
    Write-Host " - Parsing $($chatMessages.value.count) chat messages..."
    foreach ($chatMessage in $chatMessages.value) {
        # Part of chat (oneOnOne, group, meeting etc.)
        if ($chatMessage.chatId) {
            # Do we have a chat object already for this chat ID?
            # If not, create it
            if ($chatThreads.id -notcontains $chatMessage.chatId) {
                Write-Host "     - Getting chat details: $($chatMessage.chatId)..." -NoNewline
                $chatResponse = Invoke-GraphAPICall -URI "https://graph.microsoft.com/v1.0/chats/$($chatMessage.chatId)" -Method "GET" -WriteStatus
                $chatObject = [PSCustomObject]@{
                    id       = $chatMessage.chatId
                    chat     = $chatResponse
                    messages = @()
                    members  = @()
                    chatType = $chatResponse.chatType
                }
                $chatThreads += $chatObject
            }

            # Add message, member (if not already) to chat object
            $chatThreads | Where-Object { $_.id -eq $chatMessage.chatId } | ForEach-Object {
                $_.messages += $chatMessage
                if ($chatMessage.from.user.displayName -and $_.members -notcontains $chatMessage.from.user.displayName) {
                    $_.members += $chatMessage.from.user.displayName
                }
            }

        }
        # Part of channel
        elseif ($chatMessage.channelIdentity.channelId -and $chatMessage.channelIdentity.teamId) {
            # Create 'chat ID' based on combined team and channel IDs
            $chatId = "$($chatMessage.channelIdentity.channelId)_$($chatMessage.channelIdentity.teamId)"
            if ($chatThreads.id -notcontains $chatId) {
                Write-Host "     - Getting details for team: $($chatMessage.channelIdentity.teamId)..." -NoNewline
                $teamResponse = Invoke-GraphAPICall -URI "https://graph.microsoft.com/v1.0/teams/$($chatMessage.channelIdentity.teamId)?`$select=displayName,description,id" -Method "GET" -WriteStatus
                Write-Host "     - Getting details for channel: $($chatMessage.channelIdentity.channelId)..." -NoNewline
                $channelResponse = Invoke-GraphAPICall -URI "https://graph.microsoft.com/v1.0/teams/$($chatMessage.channelIdentity.teamId)/channels/$($chatMessage.channelIdentity.channelId)?`$select=displayName,description,id" -Method "GET" -WriteStatus

                $chatObject = [PSCustomObject]@{
                    id       = $chatId
                    team     = $teamResponse
                    channel  = $channelResponse
                    members  = @()
                    messages = @()
                    chatType = "channel"
                }
                $chatThreads += $chatObject
            }

            # Add message, members to chat object
            $chatThreads | Where-Object { $_.id -eq $chatId } | ForEach-Object {
                $_.messages += $chatMessage
                if ($chatMessage.from.user.displayName -and $_.members -notcontains $chatMessage.from.user.displayName) {
                    $_.members += $chatMessage.from.user.displayName
                }
            }

        }

    }

    # If there are chat threads, parse them
    if ($chatThreads.count -gt 0) {
        # Create user directory
        Write-Host " - Creating directory $userPath..." -NoNewline
        try {
            New-Item -ItemType Directory -Force -Path "$userPath/chats" | Out-Null
            Write-Host "SUCCESS" -ForegroundColor Green
        }
        catch {
            Write-Host "FAILED" -ForegroundColor Red
        }

        $chatIndex = @()
        # Loop through each chat thread now messages have been parsed
        foreach ($chat in $chatThreads) {
            $dateTaken = Get-Date -UFormat "%Y-%m-%d %H:%M"
            $chatMembers = $chat.members[0..2] -join ", "
            if ($chat.members.count -gt 3) {
                $chatMembers = @"
                $($chatMembers) and <a href="#" title="$($chat.members -join ", ")">$($chat.members.count - 3) others</a>
"@
            }

            # Title of chat
            if ($chat.chat.topic) {
                $chatTitle = $chat.chat.topic
            } 
            elseif ($chat.team.displayName -and $chat.channel.displayName) {
                $chatTitle = "$($chat.team.displayName)/$($chat.channel.displayName) channel conversation"
            }
            elseif ($chat.chatType -eq "meeting") {
                $chatTitle = "Meeting chat with $chatMembers"
            }
            else {
                $chatTitle = "Chat with $chatMembers"
            }

            # Create HTML output of messages
            $htmlMessages = Get-HTMLChatMessages -message $chat.messages -user $userObject

            $html = @"
                <br />
                <div class="card">
                <h5 class="card-header bg-light">Chat Overview</h5>
                <div class="card-body">
                <table class="table">
                <tbody>
                    <tr>
                        <th scope="row">Backup Taken:</th>
                        <td>$dateTaken</td>
                    </tr>
                    <tr>
                        <th scope="row">Type:</th>
                        <td>$( 
                            switch ($chat.chatType) {
                            oneOnOne { "One on one" }
                            group { "Group" }
                            meeting { "Meeting"}
                            channel { "Channel" }
                            Default { $chat.chatType }
                            })</td>
                    </tr>
                    $(
                        if ($chat.chat.topic) {
                            @"
                            <tr>
                                <th scope="row">Topic:</th>
                                <td>$($chat.chat.topic)</td>
                            </tr>
"@
                        }
                        if ($chat.team.displayName) {
                            @"
                            <tr>
                                <th scope="row">Team:</th>
                                <td>$($chat.team.displayName)</td>
                            </tr>
"@
                        }
                        if ($chat.channel.displayName) {
                            @"
                            <tr>
                                <th scope="row">Channel:</th>
                                <td>$($chat.channel.displayName)</td>
                            </tr>
"@
                        }
                    )
                    <tr>
                        <th scope="row">Messages:</th>
                        <td>$($chat.messages.count)</td>
                    </tr>
                    <tr>
                        <th scope="row">Members:</th>
                        <td>$($chat.members.count) - $($chatMembers)</td>
                    </tr>
                </tbody>
            </table>
            </div>
            </div>
            <br />
            <div class="card">
                <h5 class="card-header bg-light">Transcript</h5>
                <div class="card-body">
                $htmlMessages
                </div>
            </div>
"@

            # Save file with topic name
            $file = "$($chat.id).htm" -replace '([\\/:*?"<>|\s])+', "_"
            $filePath = "$userPath/chats/$file"
            Write-Host "    - Saving chat to $filePath... " -NoNewline
            New-HTMLPage -Content $html -PageTitle "$($userObject.displayName) - $chatTitle" -Path $filePath

            # Add to index
            $chatIndexObject = [PSCustomObject]@{
                chatTitle     = $chatTitle
                link          = "./chats/$file"
                totalMembers  = $chat.members.count
                totalMessages = $chat.messages.count
                lastMessage   = $chat.messages | Sort-Object -Property createdDateTime | Select-Object -Last 1
            }
            $chatIndex += $chatIndexObject
        }
        # Create chat index html file
        $html = @"
                <br />
                <div class="card">
                <h5 class="card-header bg-light">$($chatIndex.count) Chats</h5>
                <div class="card-body">
                <table class="table">
                    <thead>
                        <tr>
                            <th scope="col">Title</th>
                            <th scope="col">Members</th>
                            <th scope="col">Messages</th>
                            <th scope="col">Lastest Message</th>
                        </tr>
                    </thead>
                    <tbody>
                        $(
                            foreach($chatIndexObject in $chatIndex | Sort-Object -Property lastMessage.createdDateTime -Descending) {
                                @"
                                <tr>
                                    <td><a href="$($chatIndexObject.link)">$($chatIndexObject.chatTitle)</a></td>
                                    <td>$($chatIndexObject.totalMembers)</td>
                                    <td>$($chatIndexObject.totalMessages)</td>
                                    <td>$($chatIndexObject.lastMessage.createdDateTime)</td>
                                </tr>
"@
                            }
                        )
                    </tbody>
                </table>
                </div>
                </div>
"@

        $filePath = "$userPath/index.htm"
        Write-Host "    - Saving chat index to $filePath... " -NoNewline
        New-HTMLPage -Content $html -PageTitle "Chats: $($userObject.displayName)" -Path $filePath

        # Return user object
        return [PSCustomObject]@{
            displayName = $userObject.displayName
            totalChats  = $chatIndex.count
            link        = "$directory/index.htm"
        }
    }
}

function Get-MessageReactions {
    param (
        [Parameter(mandatory = $true)][System.Object]$reactions
    )

    if ($reactions.count -gt 0) {
        $totalReactions = @()
        $reactions | Group-Object -Property reactionType -NoElement | Sort-Object -Property Count -Descending | ForEach-Object {
            switch ($_.Name) {
                like { $emoji = "128077" }
                angry { $emoji = "128545" }
                sad { $emoji = "128577" }
                laugh { $emoji = "128518" }
                heart { $emoji = "128156" }
                surprised { $emoji = "128562" }
            }
            $totalReactions += "$($_.count) &#$emoji; "
        }
        return @"
        <div align="right">
            $totalReactions
        </div>
"@
    }

}

function Get-HTMLChatMessage {
    param (
        [Parameter(mandatory = $true)][System.Object]$message
    )

    return @"
                <div class="card-header bg-light">
                    <div class="row">
                        <div class="col-8">
                            <b>$($message.from.user.displayName)</b> at $($message.createdDateTime | Get-Date -UFormat "%Y-%m-%d %H:%M:%S") 
                            $(if ($message.lastEditedDateTime) { "(Edited)" })
                        </div>
                        <div class="col">
                            $(Get-MessageReactions -reactions $message.reactions)
                        </div>
                    </div>
                </div>
                $(if ($message.importance -eq "high") { '<div class="alert alert-danger m-2" role="alert">IMPORTANT!</div>'})
                $(if ($message.importance -eq "urgent") { '<div class="alert alert-danger m-2" role="alert">URGENT!</div>'})
                <div class="card-body">
                    <h4 class="card-title">$($_.subject)</h4>
                    $(if ($message.deletedDateTime) { '<i>This message has been deleted.</i>' })
                    <p>$($message.body.content)</p>
                </div>
"@

}

function Get-HTMLChatMessages {
    param (
        [Parameter(mandatory = $true)][System.Object]$messages,
        [Parameter(mandatory = $true)][System.Object]$user
    )

    # Loop through each message
    $htmlMessages = @()
    foreach ($message in $messages | Sort-Object -Property createdDateTime) {
        # Check it is not a reply to another message in thread
        if (!$message.replyToId) {
            # Check for replies to message
            $replies = @()
            $messages | Where-Object { $_.replyToId -eq $message.id } | ForEach-Object {
                $replies += $_
            }

            $messageStyle = $message.from.user.id -eq $user.id ? 'style="margin: 0 0 0 6rem; background-color: #e9eaf6"' : 'style="margin: 0 6rem 0 0"'
            $htmlMessages += @"
            <div class="card" $messageStyle>
                $(Get-HTMLChatMessage -message $message)
                $(if ($replies.count -gt 0) {
                    @"
                    <ul class="list-group list-group-flush">
                        <li class="list-group-item">$($replies.count) replies:</li>
                    </ul>
"@
                    foreach ($reply in $replies | Sort-Object -Property createdDateTime) {
                        $(Get-HTMLChatMessage -message $reply)
                    }
                })
            </div><br />
"@
        }
    }
    return $htmlMessages
}

function New-HTMLPage {
    param (
        [Parameter(mandatory = $true)][string]$Content,
        [Parameter(mandatory = $true)][string]$pageTitle,
        [Parameter(mandatory = $true)][string]$Path
    )

    $html = @"
        <div class="container-md">
            <div class="page-header">
                <h1>$pageTitle</h1>
                <h5>Created with <a href="https://github.com/leeford/Backup-TeamsChat">Backup-TeamsChat</a></h5>
            </div>
            $Content
            </div>
"@

    try {
        ConvertTo-Html -CssUri "https://cdn.jsdelivr.net/npm/bootstrap@5.0.1/dist/css/bootstrap.min.css" -Body $html -Title $pageTitle | Out-File $Path
        Write-Host "SUCCESS" -ForegroundColor Green
    }
    catch {
        Write-Host "FAILED" -ForegroundColor Red
    }

}

Write-Host "`n----------------------------------------------------------------------------------------------
            `n Backup-TeamsChat.ps1 - Lee Ford - https://github.com/leeford/Backup-TeamsChat
            `n----------------------------------------------------------------------------------------------" -ForegroundColor Yellow

# Check secret modules are installed to store connection details
Check-ModuleInstalled -module Microsoft.PowerShell.SecretManagement -moduleName "Microsoft PowerShell SecretManagement"
Check-ModuleInstalled -module Microsoft.PowerShell.SecretStore -moduleName "Microsoft PowerShell SecretStore"

# Create vault script (if required)
Write-Host "Creating secret vault..." -NoNewline
try {
    Register-SecretVault -Name Backup-TeamsChat -ModuleName Microsoft.PowerShell.SecretStore
    Write-Host "SUCCESS" -ForegroundColor Green
}
catch {
    Write-Host "SUCCESS (Vault already exists)" -ForegroundColor Green
}

# Check Days is a postive number
if ($Days -and $Days -lt 0) {
    Write-Host "Please specify a valid day range (greater than 0 days)" -ForegroundColor Red
    break
}

# Get secrets and ask if not present
# Client ID
try {
    $clientId = Get-Secret -Name ClientId -AsPlainText -Vault Backup-TeamsChat -ErrorAction Stop
}
catch {
    $clientId = Read-Host "Azure AD Application (client) ID not found. Please enter this"
    Set-Secret -Name ClientId -Secret $clientId -Vault Backup-TeamsChat
}
# Tenant ID
try {
    $tenantId = Get-Secret -Name TenantId -AsPlainText -Vault Backup-TeamsChat -ErrorAction Stop
}
catch {
    $tenantId = Read-Host "Azure AD Directory (tenant) ID not found. Please enter this"
    Set-Secret -Name TenantId -Secret $tenantId -Vault Backup-TeamsChat
}
# Client Secret
try {
    $clientSecret = Get-Secret -Name ClientSecret -AsPlainText -Vault Backup-TeamsChat -ErrorAction Stop
}
catch {
    $clientSecret = Read-Host "Azure AD Client secret not found. Please enter this"
    Set-Secret -Name ClientSecret -Secret $clientSecret -Vault Backup-TeamsChat
}

# Get OAuth token for Graph
Get-ApplicationToken

# Create root directory
$date = Get-Date -UFormat "%Y_%m_%d_%H%M"
$directoryName = "/TeamsChatBackup_$date"
$rootDirectory = "$Path/$directoryName"
New-Item -Path $rootDirectory -ItemType Directory -ErrorAction SilentlyContinue | Out-Null

# Get Chats for each user
$userOutput = @()
if ($user) {
    $userResponse = Invoke-GraphAPICall "https://graph.microsoft.com/v1.0/users/$($user)?`$select=displayName,userPrincipalName,id"
    $userOutput += Get-Chats -userObject $userResponse -rootDirectory $rootDirectory
}
else {
    # All users
    $usersResponse = Invoke-GraphAPICall "https://graph.microsoft.com/v1.0/users?`$select=displayName,userPrincipalName,id"
    foreach ($userObject in $usersResponse.value) {
        $userOutput += Get-Chats -UserObject $userObject -rootDirectory $rootDirectory
    }
}

# Create user index
if ($userOutput.count -gt 0) {
    $html = @"
    <br />
    <div class="card">
    <h5 class="card-header bg-light">$($userOutput.count) Users</h5>
    <div class="card-body">
    <table class="table">
        <thead>
            <tr>
                <th scope="col">User</th>
                <th scope="col">Chats</th>
            </tr>
        </thead>
        <tbody>
            $(
                foreach($userObject in $userOutput | Sort-Object -Property displayName) {
                    @"
                    <tr>
                        <td><a href="$($userObject.link)">$($userObject.displayName)</a></td>
                        <td>$($userObject.totalChats)</td>
                    </tr>
"@
                }
            )
        </tbody>
    </table>
    </div>
    </div>
"@
    
    $filePath = "$rootDirectory/index.htm"
    Write-Host "Saving chat index to $filePath... " -NoNewline -ForegroundColor Yellow
    New-HTMLPage -Content $html -PageTitle "Backup-TeamsChat" -Path $filePath
}