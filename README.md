# Backup-TeamsChat

> **Disclaimer:** This tool is provided ‘as-is’ without any warranty or support. Use of this tool is at your own risk and I accept no responsibility for any damage caused.

Backup Teams chat messages (not standard Team channel messages) for safe keeping - messages are saved in to a HTML report for easy viewing. Written in PowerShell Core and using Graph API, it can be used on Windows, Mac and Linux.

It can export chat messages for all users within a Teams tenant - including guests. This allows you to backup the following chat types:

* One on one (2 participants)
* Group (more than 2 participants but not in a Team)
* Meeting
* Private Team channel conversation

All of this is possible by using [Microsoft Teams Protected APIs](https://docs.microsoft.com/en-us/graph/teams-protected-apis) in Microsoft Graph API. By gaining access to Teams Protected APIs in Microsoft Graph you will have unrestricted access to all Team resources which may contain sesintive information in your organisation. **Please ensure you have the relevant permission within your organisation to access this information before requesting access to the APIs**

## Pre-requisites

### PowerShell
You need to ensure you have PowerShell _Core_ (6+) installed. **This tool will not work with Windows PowerShell (5.1)**.

### Azure AD App registration
To connect to Microsoft Graph API, you will need to use an Azure AD App registration. Follow the below instructions to create one:

1. Login to https://portal.azure.com (under the tenant you wish to use the tool) and under **Azure Active Directory > App registrtions** create a **New registration**. Set the `Name` to something descriptive and `Supported acccount types` to **Single tenant** (unless you wish to use the tool on multiple tenants). Click **Register**
![image](https://user-images.githubusercontent.com/472320/123973930-1aafc680-d9b4-11eb-9560-63af528f5bcf.png)
2. Take a note of the `Application (client) ID` and `Directory (tenant) ID` for later on
![image](https://user-images.githubusercontent.com/472320/123974188-5a76ae00-d9b4-11eb-914b-4f7046a5c225.png)
3. Under **Certificates & secrets** create a new **Client secret**. Set the `Description` and when it `Expires` (remember to renew) and click **Add**. Take a copy of the `Value` for later.
![image](https://user-images.githubusercontent.com/472320/123975945-a83fe600-d9b5-11eb-9761-bc00a2ffe15d.png)
![image](https://user-images.githubusercontent.com/472320/123976051-bee63d00-d9b5-11eb-9c99-6653e7c34df3.png)
4. Under **API Permissions**, add the following **Microsoft Graph** permissions:

| Permission | Type | Description |
| ---- | ---- | ---- |
| Chat.Read.All | Application | Read all chat messages |
| Channel.ReadBasic.All | Application | Read the names and descriptions of Teams channels|
| Team.ReadBasic.All | Application | Read the names and descriptions of Teams |
| User.Read.All | Application | Read all users profiles |
5. With the permissions added, you will need to **Grant admin consent**. It should look like the following:
![image](https://user-images.githubusercontent.com/472320/123975328-264fbd00-d9b5-11eb-9c05-f1e4de29884a.png)
6. That is the App registation configured

### Teams Protected APIs
As mentioned above, you will need to request access to the Teams Protected APIs. This is acheived by filling in this [form](https://aka.ms/teamsgraph/requestaccess) with details of your Azure AD App registration and why you require access. It can take around a week to hear back.

### PowerShell Modules
Two modules are required to store secrets about your Azure AD App safely. Run the following:

```pwsh
Install-Module -Name Microsoft.PowerShell.SecretManagement  
Install-Module -Name Microsoft.PowerShell.SecretStore
```

## Usage
With the Azure AD App registration created and the Teams Protected APIs granted to it, it is now possible to use the tool.

Firstly, download the latest release (2.0+) at https://github.com/leeford/Backup-TeamsChat/releases

You can then run the .ps1 file from a PowerShell _Core_ prompt. If it is the first time running it, it will create a secret vault called **Backup-TeamsChat**. Within this secret vault it will securely store your Azure AD App registration details (client ID, tenant ID and client secret):

![image](https://user-images.githubusercontent.com/472320/123989672-0625fb00-d9c1-11eb-8bca-5658608f7819.png)

> Each time you run the tool, you will need to enter the password you used when creating the secret vault.

* Backup all chat messages for all users in tenant:
```pwsh
Backup-TeamsChat.ps1 -Path <directory to save backup>
```

* Backup chat messages for a specific user:
```pwsh
Backup-TeamsChat.ps1 -Path <directory to save backup> -User <UPN or ID>
```

* Backup chat messages for all users in tenant for the last _X_ days:
```pwsh
Backup-TeamsChat.ps1 -Path <directory to save backup> -Days 30
```

Within the specified path there will be a new folder created and inside `index.htm`. Open this to navigate through users and their chats.

* List of users:
![image](https://user-images.githubusercontent.com/472320/123978679-f950d980-d9b7-11eb-9dba-fdc8cd75e9cf.png)
* List of chats:
![image](https://user-images.githubusercontent.com/472320/123979000-3ae18480-d9b8-11eb-86a3-23208b7c6b84.png)
* Chat thread:
![image](https://user-images.githubusercontent.com/472320/123980122-35d10500-d9b9-11eb-9fdd-7f5edbb2a4ac.png)


