# DRADelegation PowerShell Module

This PowerShell module is to allow MicroFocus/NetIQ DRA administrators to automate DRA delegation operations.

Included commands:

**Get-DRAActiveView  
Get-DRAActiveViewRule  
Get-DRAAssistantAdmin  
Get-DRAAssistantAdminMember  
Get-DRAAssistantAdminRule  
Get-DRADelegation  
Get-DRAPower  
Get-DRARole  
Get-DRARoleMember  
Get-DRAServer  
Grant-DRADelegation  
New-DRAActiveView  
New-DRAActiveViewActiveViewRule  
New-DRAActiveViewDomainRule  
New-DRAActiveViewGroupRule  
New-DRAActiveViewOURule  
New-DRAActiveViewUserRule  
New-DRAAssistantAdmin  
New-DRAAssistantAdminAARule  
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
Rename-DRARole  
Revoke-DRADelegation  
Set-DRAActiveViewComment  
Set-DRAActiveViewDescription  
Set-DRAActiveViewRuleComment  
Set-DRAAssistantAdminComment  
Set-DRAAssistantAdminDescription  
Set-DRAAssistantAdminRuleComment**

I wrote this module mainly because the EA.exe command-line interface is awkward and difficult to use. This module also adds some features missing from EA.exe.

## System Requirements

* Windows 7/Windows Server 2008 R2 or later
* Windows PowerShell 3.0 or later
* DRA server installation including the **Command-line interface** feature (only tested on DRA 9.2)

On Windows 7/Windows Server 2008 R2, Windows Management Framework (WMF) 3.0 or later is required to meet the Windows PowerShell prerequisite. The DRA server must be installed with the **Command-line interface** feature.

## Installation

Choose **Releases** and download and run the installer.

## Contributions

If you would like to contribute to this project, use this link:

https://paypal.me/wastewart

Thank you!