# Backup-TeamsChat

> **Disclaimer:** This tool is provided ‘as-is’ without any warranty or support. Use of this tool is at your own risk and I accept no responsibility for any damage caused.

Backup Teams chat messages (not channel messages) for safe keeping - messages are saved in to a HTML report for easy viewing. Written in PowerShell Core and using Graph API (no modules required), it can be used on Windows, Mac and Linux.

> I recently (Oct 2019) attended a Microsoft 365 Developer Bootcamp in Nottingham around Microsoft Graph. Part of the day was around creating a tool/script to demonstrate how Graph could be used. This was my entry.

## Pre-requisites ##

You need to ensure you have PowerShell _Core_ (6+) installed. **This tool will not work with Windows PowerShell**.

In addition, to connect to Graph API, you will need to use an Azure AD v2.0 Application. The application requires that it is granted the following (delegated) Graph API permissions:

* **Chat.Read** - Allows read of logged in user's chat messages

The tool is pre-configured with a Azure AD application that has these permissions configured.

If you would prefer to use your own, you can create an application that supports device login and populate the _$script:clientId_ and _$script:tenantId_ with the client and tenant ID of your application. For instructions on how to create an device-code application, see https://www.lee-ford.co.uk/graph-api-device-code/

## Usage ##

1. Download latest release at https://github.com/leeford/Backup-TeamsChat/releases

2. Run the .ps1 file from a PowerShell _Core_ prompt:
   
    ```Backup-TeamsChat.ps1 -Path <directory to save backup>```

3. Copy the code from the console and enter it at https://microsoft.com/devicelogin and sign in to your Office 365 tenant. You may be asked to grant consent to the application

4. Locate backup file and unzip. Each chat will have its own HTML file:

![](https://www.lee-ford.co.uk/images/backup-teamschat/sample-html.png)

