# DRADelegation PowerShell Module

This PowerShell module is to allow MicroFocus/NetIQ DRA administrators to automate DRA delegation operations.

Included commands:

**Get-DRAActiveView  
Get-DRAActiveViewRule  
Get-DRAAssistantAdmin  
Get-DRAAssistantAdminRule  
Get-DRARole  
Grant-DRADelegation  
New-DRAActiveView  
New-DRAActiveViewActiveViewRule  
New-DRAActiveViewDomainRule  
New-DRAActiveViewGroupRule  
New-DRAActiveViewOURule  
New-DRAActiveViewUserRule  
New-DRAAssistantAdmin  
New-DRAAssistantAdminGroupRule  
New-DRAAssistantAdminUserRule  
Remove-DRAActiveView  
Remove-DRAActiveViewRule  
Remove-DRAAssistantAdmin  
Remove-DRAAssistantAdminRule  
Rename-DRAActiveView  
Rename-DRAActiveViewRule  
Rename-DRAAssistantAdmin  
Rename-DRAAssistantAdminRule  
Revoke-DRADelegation  
Set-DRAActiveViewComment  
Set-DRAActiveViewDescription  
Set-DRAActiveViewRuleComment  
Set-DRAAssistantAdminComment  
Set-DRAAssistantAdminDescription  
Set-DRAAssistantAdminRuleComment**

I wrote this module mainly because the EA.exe command-line interface is awkward and difficult to use.

## System Requirements

* Windows 7/Windows Server 2008 R2 or later
* Windows PowerShell 3.0 or later
* DRA client installation including the **Command-line interface** feature

On Windows 7/Windows Server 2008 R2, Windows Management Framework (WMF) 3.0 or later is required to meet the Windows PowerShell prerequisite. The DRA client installation with the **Command-line interface** feature is required because the module is a "wrapper" that automates the EA.exe command.

## Installation

Choose **Releases** and download and run the installer.

## Limitations

* Module is limited to what's available in EA.exe
* Module speed is limited by performance of EA.exe
* All actions connect to the primary server, so recommendation is only to use this module on primary server or server in same site as primary with a fast network connection
* ActiveView resource rules are not supported

## Contributions

If you would like to contribute to this project, use this link:

https://paypal.me/wastewart

Thank you!