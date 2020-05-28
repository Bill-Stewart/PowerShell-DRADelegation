# DRADelegation.psm1 Written by Bill Stewart
#
# Provides a PowerShell interface for getting and setting MicroFocus/NetIQ DRA
# delegation details (ActiveViews, AssistantAdmins, and Roles). For the most
# part the module is a wrapper for the EA.exe command-line interface. Some of
# the functions use the DRA COM objects. (The COM interfaces are not documented
# for customers, but I managed to cobble some functionality together based on
# some info from the DRA developers and from viewing the DRA admin service
# logs.)
#
# For output operations that use EA.exe, the module captures the output of the
# command, parses it, and outputs as PowerShell objects that can be filtered,
# sorted, etc. Pipeline operations are supported and should work as expected.
#
# The module uses Write-Progress to show the EA.exe command line in cases of
# slow operations. The EA.exe command line is also shown when debug is enabled.
#
# Prerequisites:
# * Current computer must be a DRA server
# * The "command-line interface" feature must be installed
# * DRA 9.2 (or later?)
#
# Current limitations:
#
# * Requires DRA 9.2 (or later?).
#
# * The module uses EA.exe for many actions, so it can be somewhat slow if
#   EA.exe gets called repeatedly (still faster than point-and-click, though).
#
# * Errors from EA.exe actions are parsed from its output, so if the EA.exe
#   output format changes, the module may behave incorrectly in various ways.
#
# * The module only discovers DRA servers in the current domain.
#
# * All actions connect to the primary server, so it's recommended only to use
#   this module on the primary server or server in same site as primary with a
#   fast network connection.
#
# * The EA.exe error messages don't align well with the PowerShell objects, but
#   they should give enough information to tell what's wrong.
#
# * The Get-DRADelegation "Role" property can contain Power assignments (Power
#   assignments are read-only for purposes of the module). For this reason,
#   it's recommended to only create delegations using Role objects. (If you
#   need a delegation for a single Power, create a Role for it.)
#
# * DRA Power objects have limited visibility and are read-only (we have
#   Get-DRAPower and enumeration of assigned Power objects in
#   Get-DRADelegation, but that's all for now).
#
# * ActiveView resource rules are not supported.
#
# * Messages are not localized (English only).
#
# Version history:
#
# 1.0 (2019-09-16)
# * Initial version.
#
# 1.5 (2019-11-20)
# * Replaced Out-Object with [PSCustomObject] (PSv3 is prerequisite)
# * Added support for COM object interface
# * Changed GetDRAObject and GetDRAObjectRule functions to use COM object
#   interface (this is an improvement because objects are output as they get
#   enumerated; previously when parsing EA.exe output, we would have to collect
#   ALL output up front before parsing, which makes the module feel less
#   responsive to the user)
# * "Type" property of Get-* renamed to "Builtin" and changed from string
#   ("Built-in" or "Custom") to boolean
# * "Assigned" property of Get-* changed from string ("Yes" or "No") to boolean
# * Added Get-DRADelegation (uses COM as EA.exe does not support)
# * Added Get-DRAPower (uses COM as EA.exe does not support)
# * Added Rename-DRARole (uses COM as EA.exe does not support)
# * Get-DRAActiveViewRule -ActiveView parameter has "*" as default
# * Get-DRAAssistantAdminRule -AssistantAdmin parameter has "*" as default
# * Added "AV" as parameter alias for "ActiveView"
# * Added "AA" as parameter alias for "AssistantAdmin"
# * Added pipeline inputs to Get-*
# * Cleaned up error handling (correct object names, etc.)
#
# 1.6 (2020-05-28)
# * Added New-DRAAssistantAdminAARule

#requires -version 3

#------------------------------------------------------------------------------
# CATEGORY: Initialization
#------------------------------------------------------------------------------
$DRAInstallDir = Get-ItemProperty "HKLM:\SOFTWARE\WOW6432Node\Mission Critical Software\OnePoint\Administration" -ErrorAction SilentlyContinue
if ( -not $DRAInstallDir ) {
  $DRAInstallDir = Get-ItemProperty "HKLM:\SOFTWARE\Mission Critical Software\OnePoint\Administration" -ErrorAction SilentlyContinue
}
$DRAInstallDir = $DRAInstallDir.InstallDir
if ( -not $DRAInstallDir ) {
  throw "This module requires that the current computer be a DRA server."
}

$EA = Join-Path $DRAInstallDir "EA.exe"
if ( -not (Test-Path $EA) ) {
  throw "Unable to find '$EA'. This module requires the DRA command-line interface to be installed."
}

try {
  [Void] (New-Object -ComObject "EAServer.EAServe")
}
catch {
  throw "This module requires that the current computer be a DRA server."
}

# Always use primary server?
$FORCE_PRIMARY = $true
#------------------------------------------------------------------------------

#------------------------------------------------------------------------------
# CATEGORY: Pathname object support
#------------------------------------------------------------------------------
$ADS_ESCAPEDMODE_ON = 2
$ADS_SETTYPE_DN = 4
$ADS_FORMAT_X500_DN = 7

$Pathname = New-Object -ComObject "Pathname"
[Void] $Pathname.GetType().InvokeMember("EscapedMode","SetProperty",$null,$Pathname,$ADS_ESCAPEDMODE_ON)

function Get-EscapedName {
  param(
    [String] $distinguishedName
  )
  [Void] $Pathname.GetType().InvokeMember("Set","InvokeMethod",$null,$Pathname,($distinguishedName,$ADS_SETTYPE_DN))
  $Pathname.GetType().InvokeMember("Retrieve","InvokeMethod",$null,$Pathname,$ADS_FORMAT_X500_DN)
}
#------------------------------------------------------------------------------

#------------------------------------------------------------------------------
# CATEGORY: Parameter validation support
#------------------------------------------------------------------------------
function Get-EAParamName {
  param(
    [String] $objectType
  )
  switch ( $objectType ) {
    "ActiveView"     { return "AV" }
    "AssistantAdmin" { return "AA" }
    "Role"           { return "ROLE" }
  }
}

# Used in ValidateScript parameter validations for DRA object names; If
# -wildcard is used, the wildcard characters (? and *) are permitted
function Test-DRAValidObjectNameParameter {
  param(
    [String] $argument,
    [Switch] $wildcard
  )
  $invalidChars = '$','#','%','\'
  if ( -not $wildcard ) { $invalidChars += '?','*' }
  if ( $argument ) {
    foreach ( $invalidChar in $invalidChars ) {
      if ( $argument.Contains($invalidChar) ) {
        throw [Management.Automation.ValidationMetadataException] "The argument cannot contain any of the following characters: $invalidChars"
      }
    }
  }
  else {
    throw [Management.Automation.ValidationMetadataException] "The argument is null or empty."
  }
  $true
}

# Implements custom errors for ValidateScript for list-like parameters
function Test-ValidListParameter {
  param(
    [String[]] $argument,
    [String] $matchOnly,
    [String[]] $matchOneOrMore
  )
  if ( -not $argument ) {
    if ( $matchOnly ) {
      throw [Management.Automation.ValidationMetadataException] "The argument must be $matchOnly or a list of one or more of: $matchOneOrMore"
    }
    else {
      throw [Management.Automation.ValidationMetadataException] "The argument must be a list of one or more of: $matchOneOrMore"
    }
  }
  if ( $matchOnly ) {
    $set = $matchOneOrMore + $matchOnly
  }
  else {
    $set = $matchOneOrMore
  }
  if ( $set -notcontains $argument ) {
    $OFS = " "
    if ( $matchOnly ) {
      throw [Management.Automation.ValidationMetadataException] "The argument must be $matchOnly or a list of one or more of: $matchOneOrMore"
    }
    else {
      throw [Management.Automation.ValidationMetadataException] "The argument must be a list of one or more of: $matchOneOrMore"
    }
  }
  return $true
}

# Returns a parameter's argument ($argument) that can contain only a single
# value ($matchOnly) or one or more values ($matchOneOrMore); returns the value
# or set of values if valid, or throws an exception otherwise
function Get-ValidListParameter {
  param(
    [String] $parameterName,
    [String[]] $argument,
    [String] $matchOnly,
    [String[]] $matchOneOrMore
  )
  if ( $argument ) {
    if ( $argument.Count -eq 1 ) {
      if ( $argument -eq $matchOnly ) {
        return $argument
      }
    }
    $argument = $argument | Select-Object -Unique
    foreach ( $item in $argument ) {
      if ( $matchOneOrMore -notcontains $item ) {
        $OFS = " "
        if ( $matchOnly ) {
          throw [ArgumentException] "Cannot validate argument on parameter '$parameterName'. The argument must be $matchOnly or a list of one or more of: $matchOneOrMore"
        }
        else {
          throw [ArgumentException] "Cannot validate argument on parameter '$parameterName'. The argument must be a list of one or more of: $matchOneOrMore"
        }
      }
    }
    return $argument
  }
}
#------------------------------------------------------------------------------

#------------------------------------------------------------------------------
# CATEGORY: DRA server discovery
#------------------------------------------------------------------------------
# EXPORT
function Get-DRAServer {
  <#
  .SYNOPSIS
  Gets information about DRA servers in the current domain.

  .DESCRIPTION
  Gets information about DRA servers in the current domain.

  .PARAMETER Site
  Specifies to get information about DRA servers in the current site.

  .PARAMETER All
  Specifies to get information about DRA servers in the current domain.

  .PARAMETER Primary
  Specifies to get information about the primary DRA server.

  .INPUTS
  None

  .OUTPUTS
  None
  #>
  [CmdletBinding(DefaultParameterSetName = "Site")]
  param(
    [Parameter(ParameterSetName = "Site")]
    [Switch] $Site,

    [Parameter(ParameterSetName = "All",Mandatory = $true)]
    [Switch] $All,

    [Parameter(ParameterSetName = "Primary",Mandatory = $true)]
    [Switch] $Primary
  )
  try {
    $defaultNC = ([ADSI] "LDAP://rootDSE").Get("defaultNamingContext")
  }
  catch [Runtime.InteropServices.COMException] {
    throw $_
  }
  # Get container containing DRA serviceConnectionPoint objects
  $scpContainer = [ADSI] "LDAP://CN=DraServer,CN=System,$defaultNC"
  if ( -not $scpContainer.Children ) {
    throw [Management.Automation.ItemNotFoundException] "Unable to discover DRA servers."
  }
  $draServers = New-Object Collections.Generic.List[PSObject]
  foreach ( $dirEntry in $scpContainer.Children ) {
    $keywords = ($dirEntry.Properties["keywords"] -join [System.Environment]::NewLine) | ConvertFrom-StringData
    [Void] $draServers.Add([PSCustomObject] @{
      "Name"    = $dirEntry.Properties["name"][0]
      "Domain"  = $keywords["Domain"]
      "Forest"  = $keywords["Forest"]
      "Site"    = $keywords["Site"]
      "Type"    = $keywords["Type"]
      "Version" = [Version] $keywords["Version"]
    })
  }
  switch ( $PSCmdlet.ParameterSetName ) {
    "Site" {
      # Get list of DRA servers in current site
      $draServersInSite = $draServers | Where-Object { $_.Site -eq [DirectoryServices.ActiveDirectory.ActiveDirectorySite]::GetComputerSite().Name }
      # If no DRA servers in current site, get list of all DRA servers
      if ( -not $draServersInSite ) {
        $draServersInSite = $draServers
      }
      $draServersInSite
    }
    "All" {
      $draServers
    }
    "Primary" {
      # Output only the primary DRA server
      $draServers | Where-Object { ($_.Type -eq "Primary") }
    }
  }
}

# Abort module load if we can't discover any DRA servers
[Void] (Get-DRAServer -All)
#------------------------------------------------------------------------------

#------------------------------------------------------------------------------
# CATEGORY: Executable processing
#------------------------------------------------------------------------------

# Invoke-Executable uses this function to convert stdout and stderr to
# arrays; trailing newlines are ignored
function ConvertTo-Array {
  param(
    [String] $stringToConvert
  )
  $stringWithoutTrailingNewLine = $stringToConvert -replace '(\r\n)*$',''
  if ( $stringWithoutTrailingNewLine.Length -gt 0 ) {
    return ,($stringWithoutTrailingNewLine -split [Environment]::NewLine)
  }
}

# Invokes an executable ($filePath) with the specified command line
# ($arguments); stdout and stderr are captured to $output; function output is
# the executable's exit code, if available
function Invoke-Executable {
  param(
    [String] $filePath,
    [String] $arguments,
    [Ref] $output
  )
  $result = 0
  $process = New-Object Diagnostics.Process
  $startInfo = $process.StartInfo
  $startInfo.FileName = $filePath
  $startInfo.Arguments = $arguments
  $startInfo.CreateNoWindow = $true
  $startInfo.UseShellExecute = $false
  $startInfo.RedirectStandardError = $true
  $startInfo.RedirectStandardOutput = $true
  try {
    Write-Progress ('EA {0}' -f $arguments) "Processing"
    Write-Debug ('EA {0}' -f $arguments)
    if ( $process.Start() ) {
      $output.Value = [String[]] @()
      $standardOutput = ConvertTo-Array $process.StandardOutput.ReadToEnd()
      $standardError = ConvertTo-Array $process.StandardError.ReadToEnd()
      if ( $standardOutput.Count -gt 0 ) {
        $output.Value += $standardOutput
      }
      if ( $standardError.Count -gt 0 ) {
        $output.Value += $standardError
      }
      $process.WaitForExit()
      $result = $process.ExitCode
    }
  }
  catch {
    $result = $_.Exception.InnerException.ErrorCode
    if ( -not $result ) {
      $result = 13  # ERROR_INVALID_DATA
    }
  }
  Write-Progress " " " " -Completed
  return $result
}

function Invoke-EA {
  [CmdletBinding(DefaultParameterSetName = "Auto")]
  param(
    [Parameter(Position = 0,ParameterSetName = "Auto",Mandatory = $true)]
    [Parameter(Position = 0,ParameterSetName = "Primary",Mandatory = $true)]
    [Parameter(Position = 0,ParameterSetName = "Server",Mandatory = $true)]
    [ValidateNotNullOrEmpty()]
    [String[]] $argumentList,

    [Parameter(Position = 1,ParameterSetName = "Auto",Mandatory = $true)]
    [Parameter(Position = 1,ParameterSetName = "Primary",Mandatory = $true)]
    [Parameter(Position = 1,ParameterSetName = "Server",Mandatory = $true)]
    [Ref] $output,

    [Parameter(ParameterSetName = "Primary",Mandatory = $true)]
    [Switch] $primary,

    [Parameter(ParameterSetName = "Server",Mandatory = $true)]
    [ValidateNotNullOrEmpty()]
    [String] $draServer
  )
  $arguments = "/NOCR /NOLOGO"
  if ( -not $FORCE_PRIMARY ) {
    switch ( $PSCmdlet.ParameterSetName ) {
      "Primary" { $arguments += " /MASTER" }
      "Server"  { $arguments += " /SERVER:$draServer" }
    }
  }
  else {
    $arguments += " /MASTER"
  }
  $argumentList | ForEach-Object {
    if ( $_ -match '\s' ) {
      $arguments += ' "{0}"' -f $_
    }
    else {
      $arguments += ' {0}' -f $_
    }
  }
  Invoke-Executable $EA $arguments $output
}
#------------------------------------------------------------------------------

#------------------------------------------------------------------------------
# CATEGORY: Get objects
#------------------------------------------------------------------------------
function GetDRAObject {
  [CmdletBinding()]
  param(
    [Parameter(Mandatory = $true)]
    [ValidateSet("ActiveView","AssistantAdmin","Power","Role")]
    [String] $objectType,

    [Parameter(Mandatory = $true)]
    [ValidateNotNullOrEmpty()]
    [String] $objectName,

    [String] $draServerName,

    [Switch] $emptyIsError
  )
  if ( -not $FORCE_PRIMARY ) {
    if ( -not $draServerName ) {
      $draServerName = (Get-DRAServer | Get-Random).Name
    }
  }
  else {
    $draServerName = (Get-DRAServer -Primary).Name
  }
  $eaType = [Type]::GetTypeFromProgID("EAServer.EAServe",$draServerName)
  if ( -not $eaType ) {
    throw [Management.Automation.ItemNotFoundException] "Computer '$draServerName' is not a DRA server, or is not reachable."
  }
  try {
    $eaServer = [Activator]::CreateInstance($eaType)
    $varSetIn = New-Object -ComObject "NetIQDraVarSet.VarSet"
  }
  catch [Management.Automation.MethodInvocationException],[Runtime.InteropServices.COMException] {
    throw $_
  }
  Write-Debug "GetDRAObject: Connect to server '$draServerName'"
  switch ( $objectType ) {
    "ActiveView" {
      $container = "OnePoint://module=security"
      $filter = @("ActiveView(`$McsNameValue='$objectName')")
      $andFilter = $null
    }
    "AssistantAdmin" {
      $container = "OnePoint://module=security"
      $filter = @("AssistantAdmin(`$McsNameValue='$objectName')")
      $andFilter = @("AssistantAdmin(`$McsHidden='0'),AssistantAdmin(type='assistantadmin')")
    }
    "Power" {
      $container = "OnePoint://module=operations"
      $filter = @("PowerTemplate(`$McsNameValue='$objectName')")
      $andFilter = @("PowerTemplate(`$McsIsHidden='false')")
    }
    "Role" {
      $container = "OnePoint://module=security"
      $Filter = @("Role(`$McsNameValue='$objectName')")
      $andFilter = $null
    }
  }
  $varSetIn.put("Hints",@('$McsNameValue','Description','Comment','$McsSysFlag','$McsIsAssigned'))
  $varSetIn.put("OperationName","ContainerEnum")
  $varSetIn.put("Scope",0)
  $varSetIn.put("ManagedObjectsOnly",$true)
  $varSetIn.put("Container",$container)
  $varSetIn.put("Filter",$filter)
  if ( $andFilter ) { $varSetIn.put("AndFilter",$andFilter) }
  $varSetOut = $eaServer.ScriptSubmit($varSetIn)
  $lastError = $varSetOut.get("Errors.LastError")
  if ( $lastError -eq 0 ) {
    $objectCount = $varSetOut.get("TotalNumberObjects")
    if ( $objectCount -gt 0 ) {
      $tableBuffer = $varSetOut.get("IEaEnumerateBuf")
      if ( $tableBuffer ) {
        for ( $i = 0; $i -lt $tableBuffer.NumberOfRows; $i++ ) {
          $outObj = [PSCustomObject] @{
            $objectType   = $null
            "Description" = $null
            "Comment"     = $null
            "Builtin"     = $null
          }
          if ( ($objectType -eq "AssistantAdmin") -or ($objectType -eq "Role") ) {
            $outObj | Add-Member NoteProperty "Assigned" $null
          }
          for ( $j = 0; $j -lt $tableBuffer.NumberOfColumns; $j++ ) {
            $fieldValue = $null
            $tableBuffer.GetField([Ref] $fieldValue)
            switch ( $j ) {
              0 { $outObj.$objectType = $fieldValue }
              1 { $outObj.Description = $fieldValue }
              2 { $outObj.Comment     = $fieldValue }
              3 { $outObj.Builtin     = $fieldValue -as [Boolean] }
              4 {
                  if ( ($outObj.PSObject.Properties | Select-Object -ExpandProperty Name) -contains "Assigned" ) {
                    $outObj.Assigned = $fieldValue
                  }
              }
            }
          }
          $outObj
          $tableBuffer.NextRow()
        }
      }
    }
    else {
      if ( $emptyIsError ) {
        (Get-Variable PSCmdlet -Scope 1).Value.WriteError((New-Object Management.Automation.ErrorRecord "DRA $objectType '$objectName' not found.",
        (Get-Variable MyInvocation -Scope 1).Value.MyCommand.Name,
        ([Management.Automation.ErrorCategory]::ObjectNotFound),
        $objectName))
      }
    }
  }
  else {
    $PSCmdlet.ThrowTerminatingError((New-Object Management.Automation.ErrorRecord ("GetDRAObject returned error 0x{0:X8}." -f $lastError),
      (Get-Variable MyInvocation -Scope 1).Value.MyCommand.Name,
      ([Management.Automation.ErrorCategory]::NotSpecified),
      $objectName))
  }
}

# EXPORT
function Get-DRAActiveView {
  <#
  .SYNOPSIS
  Gets information about DRA ActiveView objects.

  .DESCRIPTION
  Gets information about DRA ActiveView objects.

  .PARAMETER ActiveView
  Specifies the name of one or more ActiveView objects. This parameter supports wildcards ('?' and '*').

  .INPUTS
  * System.String
  * Objects with property 'ActiveView'

  .OUTPUTS
  Objects with the following properties:
    ActiveView   The name of the ActiveView object
    Description  The object's description
    Comment      The object's comment
    Builtin      True if a built-in DRA object, or False otherwise
  #>
  param(
    [Parameter(ValueFromPipeline = $true,ValueFromPipelineByPropertyName = $true)]
    [ValidateScript({Test-DRAValidObjectNameParameter $_ -wildcard})]
    [SupportsWildcards()]
    [Alias("AV")]
    [String[]] $ActiveView = "*"
  )
  process {
    foreach ( $activeViewItem in $ActiveView ) {
      try {
        GetDRAObject "ActiveView" $activeViewItem -emptyIsError
      }
      catch {
        $PSCmdlet.WriteError($_)
      }
    }
  }
}

# EXPORT
function Get-DRAAssistantAdmin {
  <#
  .SYNOPSIS
  Gets information about DRA AssistantAdmin objects.

  .DESCRIPTION
  Gets information about DRA AssistantAdmin objects.

  .PARAMETER AssistantAdmin
  Specifies the name of one or more AssistantAdmin objects. This parameter supports wildcards ('?' and '*').

  .INPUTS
  * System.String
  * Objects with property 'AssistantAdmin'

  .OUTPUTS
  Objects with the following properties:
    AssistantAdmin  The name of the AssistantAdmin object
    Description     The object's description
    Comment         The object's comment
    Builtin         True if a built-in DRA object, or False otherwise
    Assigned        True if assigned in a delegation, or False otherwise
  #>
  param(
    [Parameter(ValueFromPipeline = $true,ValueFromPipelineByPropertyName = $true)]
    [ValidateScript({Test-DRAValidObjectNameParameter $_ -wildcard})]
    [SupportsWildcards()]
    [Alias("AA")]
    [String[]] $AssistantAdmin = "*"
  )
  process {
    foreach ( $assistantAdminItem in $AssistantAdmin ) {
      try {
        GetDRAObject "AssistantAdmin" $assistantAdminItem -emptyIsError
      }
      catch {
        $PSCmdlet.WriteError($_)
      }
    }
  }
}

# EXPORT
function Get-DRARole {
  <#
  .SYNOPSIS
  Gets information about DRA Role objects.

  .DESCRIPTION
  Gets information about DRA Role objects.

  .PARAMETER Role
  Specifies the name of one or more DRA Role objects. This parameter supports wildcards ('?' and '*').

  .INPUTS
  * System.String
  * Objects with property 'Role'

  .OUTPUTS
  Objects with the following properties:
    Role         The name of the Role object
    Description  The object's description
    Comment      The object's comment
    Builtin      True if a built-in DRA object, or False otherwise
    Assigned     True if assigned in a delegation, or False otherwise
  #>
  param(
    [Parameter(ValueFromPipeline = $true,ValueFromPipelineByPropertyName = $true)]
    [ValidateScript({Test-DRAValidObjectNameParameter $_ -wildcard})]
    [SupportsWildcards()]
    [String[]] $Role = "*"
  )
  process {
    foreach ( $roleItem in $Role ) {
      try {
        GetDRAObject "Role" $roleItem -emptyIsError
      }
      catch {
        $PSCmdlet.WriteError($_)
      }
    }
  }
}

# EXPORT
function Get-DRAPower {
  <#
  .SYNOPSIS
  Gets information about DRA Power objects.

  .DESCRIPTION
  Gets information about DRA Power objects.

  .PARAMETER Role
  Specifies the name of one or more DRA Power objects. This parameter supports wildcards ('?' and '*').

  .INPUTS
  System.String

  .OUTPUTS
  Objects with the following properties:
    Power        The name of the Power object
    Description  The object's description
    Comment      The object's comment
    Builtin      True if a built-in DRA object, or False otherwise
  #>
  param(
    [Parameter(ValueFromPipeline = $true)]
    [ValidateScript({Test-DRAValidObjectNameParameter $_ -wildcard})]
    [SupportsWildcards()]
    [String[]] $Power = "*"
  )
  process {
    foreach ( $powerItem in $Power ) {
      try {
        GetDRAObject "Power" $powerItem -emptyIsError
      }
      catch {
        $PSCmdlet.WriteError($_)
      }
    }
  }
}
#------------------------------------------------------------------------------

#------------------------------------------------------------------------------
# CATEGORY: Get delegation info
#------------------------------------------------------------------------------
# EXPORT
function Get-DRADelegation {
  <#
  .SYNOPSIS
  Gets information about DRA delegations.

  .DESCRIPTION
  Gets information about DRA delegations.

  .PARAMETER AssistantAdmin
  Specifies the name of one or more DRA AssistantAdmin objects. This parameter supports wildcards ('?' and '*').

  .INPUTS
  * System.String
  * Objects with property 'AssistantAdmin'

  .OUTPUTS
  Objects with the following properties:
    AssistantAdmin  The AssistantAdmin to which the delegation applies
    Role            The delegated Role (or Power)
    ActiveView      The ActiveView over which the delegation applies
  #>
  [CmdletBinding()]
  param(
    [Parameter(ValueFromPipeline = $true,ValueFromPipelineByPropertyName = $true)]
    [ValidateScript({Test-DRAValidObjectNameParameter $_ -wildcard})]
    [SupportsWildcards()]
    [Alias("AA")]
    [String[]] $AssistantAdmin = "*"
  )
  begin {
    if ( -not $FORCE_PRIMARY ) {
      $draServerName = (Get-DRAServer | Get-Random).Name
    }
    else {
      $draServerName = (Get-DRAServer -Primary).Name
    }
    $eaType = [Type]::GetTypeFromProgID("EAServer.EAServe",$draServerName)
    if ( -not $eaType ) {
      throw [Management.Automation.ItemNotFoundException] "Computer '$draServerName' is not a DRA server, or is not reachable."
    }
    try {
      $eaServer = [Activator]::CreateInstance($eaType)
      $varSetIn = New-Object -ComObject "NetIQDraVarSet.VarSet"
    }
    catch [Management.Automation.MethodInvocationException] {
      throw $_
    }
    Write-Debug "Get-DRADelegation: Connect to server '$draServerName'"
    $varSetIn.put("Hints",@('$McsNameValue'))
    $varSetIn.put("OperationName","SecurityAssignmentEnum")
    $varSetIn.put("Scope",0)
    $varSetIn.put("Type","AdminAssignment")
    $varSetIn.put("ManagedObjectsOnly",$true)
    $varSetIn.put("ResumeStr","")
    $varSetIn.put("nextrows",-1)
  }
  process {
    foreach ( $assistantAdminItem in $AssistantAdmin ) {
      $assistantAdminNames = GetDRAObject "AssistantAdmin" $assistantAdminItem -emptyIsError | Select-Object -ExpandProperty AssistantAdmin
      foreach ( $assistantAdminName in $assistantAdminNames ) {
        $varSetIn.put("AA","OnePoint://aa=$(Get-EscapedName $assistantAdminName),module=security")
        $varSetOut = $eaServer.ScriptSubmit($varSetIn)
        $lastError = $varSetOut.get("Errors.LastError")
        if ( $lastError -eq 0 ) {
          $objectCount = $varSetOut.get("TotalNumberObjects")
          if ( $objectCount -gt 0 ) {
            $tableBuffer = $varSetOut.get("IEaEnumerateBuf")
            if ( $tableBuffer ) {
              for ( $i = 0; $i -lt $objectCount; $i++ ) {
                $role = $null
                $tableBuffer.GetField([Ref] $role)
                $tableBuffer.NextRow()
                $activeView = $null
                $tableBuffer.GetField([Ref] $activeView)
                $tableBuffer.NextRow()
                [PSCustomObject] @{
                  "AssistantAdmin" = $assistantAdminName
                  "Role"           = $role
                  "ActiveView"     = $activeView
                }
              }
            }
          }
          else {
            $PSCmdlet.WriteError((New-Object Management.Automation.ErrorRecord "DRA AssistantAdmin '$assistantAdminName' has no delegations.",
            $MyInvocation.MyCommand.Name,
            ([Management.Automation.ErrorCategory]::ObjectNotFound),
            $assistantAdminName))
          }
        }
        else {
          $PSCmdlet.ThrowTerminatingError((New-Object Management.Automation.ErrorRecord ("Get-DRADelegation returned error 0x{0:X8}." -f $lastError),
            $MyInvocation.MyCommand.Name,
            ([Management.Automation.ErrorCategory]::NotSpecified),
            $assistantAdminName))
        }
      }
    }
  }
}
#------------------------------------------------------------------------------

#------------------------------------------------------------------------------
# CATEGORY: Get object rules
#------------------------------------------------------------------------------
function GetDRAObjectRule {
  [CmdletBinding()]
  param(
    [Parameter(Mandatory = $true)]
    [ValidateSet("ActiveView","AssistantAdmin")]
    [String] $objectType,

    [Parameter(Mandatory = $true)]
    [ValidateNotNullOrEmpty()]
    [String] $objectName,

    [String] $ruleName,

    [String] $draServerName,

    [Switch] $emptyIsError
  )
  if ( -not $FORCE_PRIMARY ) {
    if ( -not $draServerName ) {
      $draServerName = (Get-DRAServer | Get-Random).Name
    }
  }
  else {
    $draServerName = (Get-DRAServer -Primary).Name
  }
  $eaType = [Type]::GetTypeFromProgID("EAServer.EAServe",$draServerName)
  if ( -not $eaType ) {
    throw [Management.Automation.ItemNotFoundException] "Computer '$draServerName' is not a DRA server, or is not reachable."
  }
  try {
    $eaServer = [Activator]::CreateInstance($eaType)
    $varSetIn = New-Object -ComObject "NetIQDraVarSet.VarSet"
  }
  catch [Management.Automation.MethodInvocationException],[Runtime.InteropServices.COMException] {
    throw $_
  }
  Write-Debug "GetDRAObject: Connect to server '$draServerName'"
  $containers = GetDRAObject $objectType $objectName -emptyIsError
  foreach ( $container in $containers ) {
    switch ( $objectType ) {
      "ActiveView" {
        $containerObjectName = $container | Select-Object -ExpandProperty ActiveView
        $containerName = "OnePoint://av=$(Get-EscapedName $containerObjectName),module=security"
      }
      "AssistantAdmin" {
        $containerObjectName = $container | Select-Object -ExpandProperty AssistantAdmin
        $containerName = "OnePoint://aa=$(Get-EscapedName $containerObjectName),module=security"
      }
    }
    $varSetIn.put("Hints",@('$McsNameValue','Description','Comment'))
    $varSetIn.put("OperationName","ContainerEnum")
    $varSetIn.put("Scope",0)
    $varSetIn.put("Container",$containerName)
    $varSetIn.put("Filter",@("Rule(`$McsNameValue='$ruleName')"))
    $varSetIn.put("ManagedObjectsOnly",$true)
    $varSetOut = $eaServer.ScriptSubmit($varSetIn)
    $lastError = $varSetOut.get("Errors.LastError")
    if ( $lastError -eq 0 ) {
      $objectCount = $varSetOut.get("TotalNumberObjects")
      if ( $objectCount -gt 0 ) {
        $tableBuffer = $varSetOut.get("IEaEnumerateBuf")
        if ( $tableBuffer ) {
          for ( $i = 0; $i -lt $tableBuffer.NumberOfRows; $i++ ) {
            $outObj = [PSCustomObject] @{
              $objectType    = $containerObjectName
              "Rule"         = $null
              "Description"  = $null
              "Comment"      = $null
            }
            for ( $j = 0; $j -lt $tableBuffer.NumberOfColumns; $j++ ) {
              $fieldValue = $null
              $tableBuffer.GetField([Ref] $fieldValue)
              switch ( $j ) {
                0 { $outObj.Rule        = $fieldValue }
                1 { $outObj.Description = $fieldValue }
                2 { $outObj.Comment     = $fieldValue }
              }
            }
            $outObj
            $tableBuffer.NextRow()
          }
        }
      }
      else {
        if ( $emptyIsError ) {
          (Get-Variable PSCmdlet -Scope 1).Value.WriteError((New-Object Management.Automation.ErrorRecord "Rule(s) matching '$ruleName' not found in DRA $objectType '$containerObjectName'.",
          (Get-Variable MyInvocation -Scope 1).Value.MyCommand.Name,
          ([Management.Automation.ErrorCategory]::ObjectNotFound),
          $containerObjectName))
        }
      }
    }
    else {
      $PSCmdlet.ThrowTerminatingError((New-Object Management.Automation.ErrorRecord ("GetDRAObjectRule returned error 0x{0:X8}." -f $lastError),
        (Get-Variable MyInvocation -Scope 1).Value.MyCommand.Name,
        ([Management.Automation.ErrorCategory]::NotSpecified),
        $containerObjectName))
    }
  }
}

# EXPORT
function Get-DRAActiveViewRule {
  <#
  .SYNOPSIS
  Gets information about a DRA ActiveView object's rules.

  .DESCRIPTION
  Gets information about a DRA ActiveView object's rules.

  .PARAMETER ActiveView
  Specifies the name of one or more DRA ActiveView objects. This parameter supports wildcards ('?' and '*').

  .PARAMETER Rule
  Specifies the rule name. This parameter supports wildcards ('?' and '*').

  .INPUTS
  * System.String
  * Objects with property 'ActiveView'

  .OUTPUTS
  Objects with the following properties:
    ActiveView   The name of the ActiveView object containing the rule
    Rule         The name of the rule
    Description  The rule's description
    Comment      The rule's comment
  #>
  [CmdletBinding()]
  param(
    [Parameter(Position = 0,ValueFromPipeline = $true,ValueFromPipelineByPropertyName = $true)]
    [ValidateScript({Test-DRAValidObjectNameParameter $_ -wildcard})]
    [SupportsWildcards()]
    [Alias("AV")]
    [String[]] $ActiveView = "*",

    [Parameter(Position = 1)]
    [ValidateScript({Test-DRAValidObjectNameParameter $_ -wildcard})]
    [SupportsWildcards()]
    [String] $Rule = "*"
  )
  process {
    foreach ( $activeViewItem in $ActiveView ) {
      try {
        GetDRAObjectRule "ActiveView" $activeViewItem $Rule -emptyIsError
      }
      catch {
        $PSCmdlet.WriteError($_)
      }
    }
  }
}

# EXPORT
function Get-DRAAssistantAdminRule {
  <#
  .SYNOPSIS
  Gets information about a DRA AssistantAdmin object's rules.

  .DESCRIPTION
  Gets information about a DRA AssistantAdmin object's rules.

  .PARAMETER AssistantAdmin
  Specifies the name of one or more DRA AssistantAdmin objects. This parameter supports wildcards ('?' and '*').

  .PARAMETER Rule
  Specifies the rule name. This parameter supports wildcards ('?' and '*').

  .INPUTS
  * System.String
  * Objects with property 'AssistantAdmin'

  .OUTPUTS
  Objects with the following properties:
    AssistantAdmin  The name of the AssistantAdmin object containing the rule
    Rule            The name of the rule
    Description     The rule's description
    Comment         The rule's comment
  #>
  [CmdletBinding()]
  param(
    [Parameter(Position = 0,ValueFromPipeline = $true,ValueFromPipelineByPropertyName = $true)]
    [ValidateScript({Test-DRAValidObjectNameParameter $_ -wildcard})]
    [SupportsWildcards()]
    [Alias("AA")]
    [String[]] $AssistantAdmin = "*",

    [Parameter(Position = 1)]
    [ValidateScript({Test-DRAValidObjectNameParameter $_ -wildcard})]
    [SupportsWildcards()]
    [String] $Rule = "*"
  )
  process {
    foreach ( $assistantAdminItem in $AssistantAdmin ) {
      try {
        GetDRAObjectRule "AssistantAdmin" $assistantAdminItem $Rule -emptyIsError
      }
      catch {
        $PSCmdlet.WriteError($_)
      }
    }
  }
}
#------------------------------------------------------------------------------

#------------------------------------------------------------------------------
# CATEGORY: New objects
#------------------------------------------------------------------------------
function NewDRAObject {
  [CmdletBinding()]
  param(
    [Parameter(Mandatory = $true)]
    [ValidateSet("ActiveView","AssistantAdmin")]
    [String] $objectType,

    [Parameter(Mandatory = $true)]
    [ValidateNotNullOrEmpty()]
    [String] $objectName,

    [String] $description,

    [String] $comment
  )
  $argList = (Get-EAParamName $objectType),$objectName,"CREATE"
  if ( $description ) { $argList += "DESCRIPTION:$Description" }
  if ( $comment )     { $argList += "COMMENT:$Comment" }
  $output = $null
  $result = Invoke-EA $argList ([Ref] $output) -primary
  if ( $result -ne 0 ) {
    $PSCmdlet.ThrowTerminatingError((New-Object Management.Automation.ErrorRecord $output[$output.Count - 1],
      (Get-Variable MyInvocation -Scope 1).Value.MyCommand.Name,
      ([Management.Automation.ErrorCategory]::NotSpecified),
      $objectName))
  }
}

# EXPORT
function New-DRAActiveView {
  <#
  .SYNOPSIS
  Creates one or more DRA ActiveView objects.

  .DESCRIPTION
  Creates one or more DRA ActiveView objects.

  .PARAMETER ActiveView
  Specifies the name of one or more DRA ActiveView objects to create.

  .PARAMETER Description
  Specifies a description for the new object(s).

  .PARAMETER Comment
  Specifies a comment for the new object(s).

  .INPUTS
  None

  .OUTPUTS
  None
  #>
  [CmdletBinding(SupportsShouldProcess = $true)]
  param(
    [Parameter(Position = 0,Mandatory = $true)]
    [ValidateScript({Test-DRAValidObjectNameParameter $_})]
    [Alias("AV")]
    [String[]] $ActiveView,

    [String] $Description,

    [String] $Comment
  )
  process {
    foreach ( $activeViewItem in $ActiveView ) {
      if ( $PSCmdlet.ShouldProcess("DRA ActiveView '$activeViewItem'","Create") ) {
        try {
          NewDRAObject "ActiveView" $activeViewItem $Description $Comment
        }
        catch {
          $PSCmdlet.WriteError($_)
        }
      }
    }
  }
}

# EXPORT
function New-DRAAssistantAdmin {
  <#
  .SYNOPSIS
  Creates one or more DRA AssistantAdmin objects.

  .DESCRIPTION
  Creates one or more DRA AssistantAdmin objects.

  .PARAMETER AssistantAdmin
  Specifies the name of one or more DRA AssistantAdmin objects to create.

  .PARAMETER Description
  Specifies a description for the new object(s).

  .PARAMETER Comment
  Specifies a comment for the new object(s).

  .INPUTS
  None

  .OUTPUTS
  None
  #>
  [CmdletBinding(SupportsShouldProcess = $true)]
  param(
    [Parameter(Position = 0,Mandatory = $true)]
    [ValidateScript({Test-DRAValidObjectNameParameter $_})]
    [Alias("AA")]
    [String[]] $AssistantAdmin,

    [String] $Description,

    [String] $Comment
  )
  process {
    foreach ( $assistantAdminItem in $AssistantAdmin ) {
      if ( $PSCmdlet.ShouldProcess("DRA AssistantAdmin '$assistantAdminItem'","Create") ) {
        try {
          NewDRAObject "AssistantAdmin" $assistantAdminItem $Description $Comment
        }
        catch {
          $PSCmdlet.WriteError($_)
        }
      }
    }
  }
}
#------------------------------------------------------------------------------

#------------------------------------------------------------------------------
# CATEGORY: Remove objects
#------------------------------------------------------------------------------
function RemoveDRAObject {
  [CmdletBinding()]
  param(
    [Parameter(Mandatory = $true)]
    [ValidateSet("ActiveView","AssistantAdmin")]
    [String] $objectType,

    [Parameter(Mandatory = $true)]
    [ValidateNotNullOrEmpty()]
    [String] $objectName
  )
  $output = $null
  $result = Invoke-EA (Get-EAParamName $objectType),$objectName,"DELETE","MODE:B" ([Ref] $output) -primary
  if ( $result -ne 0 ) {
    $PSCmdlet.ThrowTerminatingError((New-Object Management.Automation.ErrorRecord $output[0],
      (Get-Variable MyInvocation -Scope 1).Value.MyCommand.Name,
      ([Management.Automation.ErrorCategory]::NotSpecified),
      $objectName))
  }
}

# EXPORT
function Remove-DRAActiveView {
  <#
  .SYNOPSIS
  Removes one or more DRA ActiveView objects.

  .DESCRIPTION
  Removes one or more DRA ActiveView objects.

  .PARAMETER ActiveView
  Specifies the name of one or more DRA ActiveView objects to remove. This parameter supports wildcards ('?' and '*').

  .INPUTS
  * System.String
  * Objects with property 'ActiveView'

  .OUTPUTS
  None

  .NOTES
  Use wildcards ('?' and '*') with caution as it is possible to delete multiple objects with a single command.

  To prevent the confirmation prompt, use '-Confirm:0' or set the $ConfirmPreference variable to 'Low'.
  #>
  [CmdletBinding(SupportsShouldProcess = $true,ConfirmImpact = "High")]
  param(
    [Parameter(Position = 0,Mandatory = $true,ValueFromPipeline = $true,ValueFromPipelineByPropertyName = $true)]
    [ValidateScript({Test-DRAValidObjectNameParameter $_ -wildcard})]
    [SupportsWildcards()]
    [Alias("AV")]
    [String[]] $ActiveView
  )
  process {
    foreach ( $activeViewItem in $ActiveView ) {
      if ( $PSCmdlet.ShouldProcess("DRA ActiveView '$ActiveViewItem'","Remove") ) {
        try {
          RemoveDRAObject "ActiveView" $activeViewItem
        }
        catch {
          $PSCmdlet.WriteError($_)
        }
      }
    }
  }
}

# EXPORT
function Remove-DRAAssistantAdmin {
  <#
  .SYNOPSIS
  Removes or more DRA AssistantAdmin objects.

  .DESCRIPTION
  Removes or more DRA AssistantAdmin objects.

  .PARAMETER AssistantAdmin
  Specifies the name of one or more DRA AssistantAdmin objects to remove. This parameter supports wildcards ('?' and '*').

  .INPUTS
  * System.String
  * Objects with property 'AssistantAdmin'

  .OUTPUTS
  None

  .NOTES
  Use wildcards ('?' and '*') with caution as it is possible to delete multiple objects with a single command.

  To prevent the confirmation prompt, use '-Confirm:0' or set the $ConfirmPreference variable to 'Low'.
  #>
  [CmdletBinding(SupportsShouldProcess = $true,ConfirmImpact = "High")]
  param(
    [Parameter(Position = 0,Mandatory = $true,ValueFromPipeline = $true,ValueFromPipelineByPropertyName = $true)]
    [ValidateScript({Test-DRAValidObjectNameParameter $_ -wildcard})]
    [SupportsWildcards()]
    [Alias("AA")]
    [String[]] $AssistantAdmin
  )
  process {
    foreach ( $assistantAdminItem in $AssistantAdmin ) {
      if ( $PSCmdlet.ShouldProcess("DRA AssistantAdmin '$assistantAdminItem'","Remove") ) {
        try {
          RemoveDRAObject "AssistantAdmin" $assistantAdminItem
        }
        catch {
          $PSCmdlet.WriteError($_)
        }
      }
    }
  }
}
#------------------------------------------------------------------------------

#------------------------------------------------------------------------------
# CATEGORY: New object rules
#------------------------------------------------------------------------------
# EXPORT
function New-DRAActiveViewActiveViewRule {
  <#
  .SYNOPSIS
  Creates a new ActiveView rule in a DRA ActiveView object.

  .DESCRIPTION
  Creates a new ActiveView rule in a DRA ActiveView object.

  .PARAMETER ActiveView
  Specifies the name of the ActiveView to which the rule should be added.

  .PARAMETER Rule
  Specifies the name of the rule.

  .PARAMETER Name
  Specifies the name of the ActiveView to add to the rule.

  .PARAMETER Comment
  Specifies a comment for the rule.

  .PARAMETER Exclude
  Specifies that the rule is an exclusion rule (i.e., the rule excludes rather than includes matching objects). The default is an include rule.

  .PARAMETER MatchWildcard
  Specifies that the -Name parameter can contain wildcard characters ('?' and '*').

  .PARAMETER RestrictSource
  Specifies source restriction for the rule (only allow included objects to be cloned, moved, or added to groups).

  .PARAMETER RestrictTarget
  Specifies target restriction for the rule (do not allow included objects to be cloned, moved, or added to groups).

  .INPUTS
  None

  .OUTPUTS
  None
  #>
  [CmdletBinding(SupportsShouldProcess = $true,DefaultParameterSetName = "NoRestrict")]
  param(
    [Parameter(Position = 0,ParameterSetName = "NoRestrict",Mandatory = $true)]
    [Parameter(Position = 0,ParameterSetName = "RestrictSource",Mandatory = $true)]
    [Parameter(Position = 0,ParameterSetName = "RestrictTarget",Mandatory = $true)]
    [ValidateScript({Test-DRAValidObjectNameParameter $_})]
    [Alias("AV")]
    [String] $ActiveView,

    [Parameter(Position = 1,ParameterSetName = "NoRestrict",Mandatory = $true)]
    [Parameter(Position = 1,ParameterSetName = "RestrictSource",Mandatory = $true)]
    [Parameter(Position = 1,ParameterSetName = "RestrictTarget",Mandatory = $true)]
    [ValidateScript({Test-DRAValidObjectNameParameter $_})]
    [String] $Rule,

    [Parameter(Position = 2,ParameterSetName = "NoRestrict",Mandatory = $true)]
    [Parameter(Position = 2,ParameterSetName = "RestrictSource",Mandatory = $true)]
    [Parameter(Position = 2,ParameterSetName = "RestrictTarget",Mandatory = $true)]
    # WAS 2019-07-17 DRA seems to support wildcard match, but GUI won't enumerate
    # [ValidateScript({Test-DRAValidObjectNameParameter $_ -wildcard})]
    [ValidateScript({Test-DRAValidObjectNameParameter $_})]
    [String] $Name,

    # WAS 2019-07-17 EA ignores
    # [Parameter(ParameterSetName = "NoRestrict")]
    # [Parameter(ParameterSetName = "RestrictSource")]
    # [Parameter(ParameterSetName = "RestrictTarget")]
    # [String] $Comment,

    [Parameter(ParameterSetName = "NoRestrict")]
    [Parameter(ParameterSetName = "RestrictSource")]
    [Parameter(ParameterSetName = "RestrictTarget")]
    [Switch] $Exclude,

    # WAS 2019-07-17 DRA seems to support wildcard match, but GUI won't enumerate
    # [Parameter(ParameterSetName = "NoRestrict")]
    # [Parameter(ParameterSetName = "RestrictSource")]
    # [Parameter(ParameterSetName = "RestrictTarget")]
    # [Switch] $MatchWildcard,

    [Parameter(ParameterSetName = "RestrictSource")]
    [Switch] $RestrictSource,

    [Parameter(ParameterSetName = "RestrictTarget")]
    [Switch] $RestrictTarget
  )
  # WAS 2019-07-17 DRA seems to support wildcard match, but GUI won't enumerate
  # if ( (-not $MatchWildcard) -and ($Name -match '\?|\*') ) {
  #   throw [ArgumentException] "Cannot validate argument on parameter 'Name'. Wildcards are only permitted with the -MatchWildcard parameter."
  # }
  if ( $PSCmdlet.ShouldProcess("ActiveView rule '$Rule' in DRA ActiveView '$ActiveView'","Create") ) {
    # Build argument list
    $argList = "AV",$ActiveView,"ADD",$Rule,"TYPE:AV","MATCH:$Name"
    # WAS 2019-07-17 EA ignores
    # if ( $Comment )        { $argList += "COMMENT:$Comment" }
    if ( $Exclude )          { $argList += "ACTION:EXCLUDE" }
    # WAS 2019-07-17 DRA seems to support wildcard match, but GUI won't enumerate
    # if ( $MatchWildcard )  { $argList += "MATCHWILDCARD" }
    if ( $RestrictSource )   { $argList += "RESTRICTION:S" }
    if ( $RestrictTarget )   { $argList += "RESTRICTION:T" }
    $argList += "MODE:B"
    # Invoke EA command
    $output = $null
    $result = Invoke-EA $argList ([Ref] $output) -primary
    if ( ($result -ne 0) -or ($output[$output.Count - 1] -match 'Failed$') ) {
      $PSCmdlet.WriteError((New-Object Management.Automation.ErrorRecord $output[$output.Count - 2],
        $MyInvocation.MyCommand.Name,
        ([Management.Automation.ErrorCategory]::NotSpecified),
        $objectName))
    }
  }
}

# EXPORT
function New-DRAActiveViewDomainRule {
  <#
  .SYNOPSIS
  Creates a new domain rule in a DRA ActiveView object.

  .DESCRIPTION
  Creates a new domain rule in a DRA ActiveView object.

  .PARAMETER ActiveView
  Specifies the name of the ActiveView to which the rule should be added.

  .PARAMETER Rule
  Specifies the name of the rule.

  .PARAMETER Name
  Specifies the name of the domain to add to the rule.

  .PARAMETER Comment
  Specifies a comment for the rule.

  .PARAMETER Exclude
  Specifies that the rule is an exclusion rule (i.e., the rule excludes rather than includes matching objects). The default is an include rule.

  .PARAMETER MemberType
  Specifies the types of member objects matched in the rule. This parameter can be 'NONE' or one or more of the following object types: 'COMPUTER', 'CONTACT', 'GROUP', 'OU', and 'USER'. The default is all object types.

  .PARAMETER OmitBase
  Specifies to exclude the base object from the rule. The default is to include the base object.

  .PARAMETER OneLevel
  Specifies not to match objects in child containers. The default is to match objects in child containers.

  .PARAMETER RestrictSource
  Specifies source restriction for the rule (only allow included objects to be cloned, moved, or added to groups).

  .PARAMETER RestrictTarget
  Specifies target restriction for the rule (do not allow included objects to be cloned, moved, or added to groups).

  .INPUTS
  None

  .OUTPUTS
  None
  #>
  [CmdletBinding(SupportsShouldProcess = $true,DefaultParameterSetName = "NoRestrict")]
  param(
    [Parameter(Position = 0,ParameterSetName = "NoRestrict",Mandatory = $true)]
    [Parameter(Position = 0,ParameterSetName = "RestrictSource",Mandatory = $true)]
    [Parameter(Position = 0,ParameterSetName = "RestrictTarget",Mandatory = $true)]
    [ValidateScript({Test-DRAValidObjectNameParameter $_})]
    [Alias("AV")]
    [String] $ActiveView,

    [Parameter(Position = 1,ParameterSetName = "NoRestrict",Mandatory = $true)]
    [Parameter(Position = 1,ParameterSetName = "RestrictSource",Mandatory = $true)]
    [Parameter(Position = 1,ParameterSetName = "RestrictTarget",Mandatory = $true)]
    [ValidateScript({Test-DRAValidObjectNameParameter $_})]
    [String] $Rule,

    [Parameter(Position = 2,ParameterSetName = "NoRestrict",Mandatory = $true)]
    [Parameter(Position = 2,ParameterSetName = "RestrictSource",Mandatory = $true)]
    [Parameter(Position = 2,ParameterSetName = "RestrictTarget",Mandatory = $true)]
    # WAS 2019-07-17 DRA supports wildcard matching, but EA creates broken rule
    # [ValidateScript({Test-DRAValidObjectNameParameter $_ -wildcard})]
    [ValidateScript({Test-DRAValidObjectNameParameter $_})]
    [String] $Name,

    # WAS 2019-07-17 EA ignores
    # [Parameter(ParameterSetName = "NoRestrict")]
    # [Parameter(ParameterSetName = "RestrictSource")]
    # [Parameter(ParameterSetName = "RestrictTarget")]
    # [String] $Comment,

    [Parameter(ParameterSetName = "NoRestrict")]
    [Parameter(ParameterSetName = "RestrictSource")]
    [Parameter(ParameterSetName = "RestrictTarget")]
    [Switch] $Exclude,

    # WAS 2019-07-17 DRA supports wildcard matching, but EA creates broken rule
    # [Parameter(ParameterSetName = "NoRestrict")]
    # [Parameter(ParameterSetName = "RestrictSource")]
    # [Parameter(ParameterSetName = "RestrictTarget")]
    # [Switch] $MatchWildcard,

    [Parameter(ParameterSetName = "NoRestrict")]
    [Parameter(ParameterSetName = "RestrictSource")]
    [Parameter(ParameterSetName = "RestrictTarget")]
    [ValidateScript({Test-ValidListParameter $_ "NONE" "COMPUTER","CONTACT","GROUP","OU","USER"})]
    [String[]] $MemberType,

    [Parameter(ParameterSetName = "NoRestrict")]
    [Parameter(ParameterSetName = "RestrictSource")]
    [Parameter(ParameterSetName = "RestrictTarget")]
    [Switch] $OmitBase,

    [Parameter(ParameterSetName = "NoRestrict")]
    [Parameter(ParameterSetName = "RestrictSource")]
    [Parameter(ParameterSetName = "RestrictTarget")]
    [Switch] $OneLevel,

    [Parameter(ParameterSetName = "RestrictSource")]
    [Switch] $RestrictSource,

    [Parameter(ParameterSetName = "RestrictTarget")]
    [Switch] $RestrictTarget
  )
  # WAS 2019-07-17 DRA supports wildcard matching, but EA creates broken rule
  # if ( (-not $MatchWildcard) -and ($Name -match '\?|\*') ) {
  #   throw [ArgumentException] "Cannot validate argument on parameter 'Name'. Wildcards are only permitted with the -MatchWildcard parameter."
  # }
  # Do custom runtime parameter validation for -MemberType
  if ( $MemberType ) { $MemberType = Get-ValidListParameter "MemberType" $MemberType "NONE" "COMPUTER","CONTACT","GROUP","OU","USER" }
  if ( $PSCmdlet.ShouldProcess("Domain rule '$Rule' in DRA ActiveView '$ActiveView'","Create") ) {
    # Build argument list
    $argList = "AV",$ActiveView,"ADD",$Rule,"TYPE:DOMAIN","MATCH:$Name"
    # WAS 2019-07-17 EA ignores
    # if ( $Comment )      { $argList += "COMMENT:$Comment" }
    if ( $Exclude )        { $argList += "ACTION:EXCLUDE" }
    if ( $MemberType )     { $argList += "MEMBERS:{0}" -f ($MemberType -join ',') }
    if ( $OmitBase )       { $argList += "SELECTBASE:N" }
    if ( $OneLevel )       { $argList += "RECURSIVE:N" }
    if ( $RestrictSource ) { $argList += "RESTRICTION:S" }
    if ( $RestrictTarget ) { $argList += "RESTRICTION:T" }
    $argList += "MODE:B"
    # Invoke EA command
    $output = $null
    $result = Invoke-EA $argList ([Ref] $output) -primary
    if ( ($result -ne 0) -or ($output[$output.Count - 1] -match 'Failed$') ) {
      $PSCmdlet.WriteError((New-Object Management.Automation.ErrorRecord $output[$output.Count - 2],
        $MyInvocation.MyCommand.Name,
        ([Management.Automation.ErrorCategory]::NotSpecified),
        $objectName))
    }
  }
}

# EXPORT
function New-DRAActiveViewGroupRule {
  <#
  .SYNOPSIS
  Creates a new group rule in a DRA ActiveView object.

  .DESCRIPTION
  Creates a new group rule in a DRA ActiveView object.

  .PARAMETER ActiveView
  Specifies the name of the ActiveView to which the rule should be added.

  .PARAMETER Rule
  Specifies the name of the rule.

  .PARAMETER Name
  Specifies the name of the group to add to the new rule. If -MatchWildcard is specified, this parameter supports wildcards ('?' and '*') and will cause the rule to match all groups matching the wildcard.

  .PARAMETER Comment
  Specifies a comment for the rule.

  .PARAMETER Exclude
  Specifies that the rule is an exclusion rule (i.e., the rule excludes rather than includes matching objects). The default is an include rule.

  .PARAMETER GroupScope
  Specifies the group scope for the rule. This parameter can be one or more of the following group scopes: 'LOCAL', 'GLOBAL', and 'UNIVERSAL'. The default is all group scopes.

  .PARAMETER GroupType
  Specifies the group type for the rule. This parameter can be one or more of the following group types: 'SECURITY' and 'DISTRIBUTION'. The default is all group types.

  .PARAMETER MatchWildcard
  Specifies that the -Name parameter can contain wildcard characters ('?' and '*').

  .PARAMETER MemberType
  Specifies the types of member objects matched in the rule. This parameter can be 'NONE' or one or more of the following object types: 'COMPUTER', 'CONTACT', 'GROUP', and 'USER'. The default is all object types.

  .PARAMETER NoNestedMembers
  Specifies not to manage nested members. The default is to manage nested members.

  .PARAMETER OmitBase
  Specifies to exclude the base object from the rule. The default is to include the base object.

  .PARAMETER RestrictSource
  Specifies source restriction for the rule (only allow included objects to be cloned, moved, or added to groups).

  .PARAMETER RestrictTarget
  Specifies target restriction for the rule (do not allow included objects to be cloned, moved, or added to groups).

  .PARAMETER SearchBase
  Specifies the distinguished name of an OU where objects should match. Matching objects will be found only in this OU and child OUs.

  .INPUTS
  None

  .OUTPUTS
  None
  #>
  [CmdletBinding(SupportsShouldProcess = $true,DefaultParameterSetName = "NoRestrict")]
  param(
    [Parameter(Position = 0,ParameterSetName = "NoRestrict",Mandatory = $true)]
    [Parameter(Position = 0,ParameterSetName = "RestrictSource",Mandatory = $true)]
    [Parameter(Position = 0,ParameterSetName = "RestrictTarget",Mandatory = $true)]
    [ValidateScript({Test-DRAValidObjectNameParameter $_})]
    [Alias("AV")]
    [String] $ActiveView,

    [Parameter(Position = 1,ParameterSetName = "NoRestrict",Mandatory = $true)]
    [Parameter(Position = 1,ParameterSetName = "RestrictSource",Mandatory = $true)]
    [Parameter(Position = 1,ParameterSetName = "RestrictTarget",Mandatory = $true)]
    [ValidateScript({Test-DRAValidObjectNameParameter $_})]
    [String] $Rule,

    [Parameter(Position = 2,ParameterSetName = "NoRestrict",Mandatory = $true)]
    [Parameter(Position = 2,ParameterSetName = "RestrictSource",Mandatory = $true)]
    [Parameter(Position = 2,ParameterSetName = "RestrictTarget",Mandatory = $true)]
    [ValidateScript({Test-DRAValidObjectNameParameter $_ -wildcard})]
    [String] $Name,

    # WAS 2019-07-17 EA ignores
    # [Parameter(ParameterSetName = "NoRestrict")]
    # [Parameter(ParameterSetName = "RestrictSource")]
    # [Parameter(ParameterSetName = "RestrictTarget")]
    # [String] $Comment,

    [Parameter(ParameterSetName = "NoRestrict")]
    [Parameter(ParameterSetName = "RestrictSource")]
    [Parameter(ParameterSetName = "RestrictTarget")]
    [Switch] $Exclude,

    [Parameter(ParameterSetName = "NoRestrict")]
    [Parameter(ParameterSetName = "RestrictSource")]
    [Parameter(ParameterSetName = "RestrictTarget")]
    [ValidateScript({Test-ValidListParameter $_ "" "LOCAL","GLOBAL","UNIVERSAL"})]
    [String[]] $GroupScope,

    [Parameter(ParameterSetName = "NoRestrict")]
    [Parameter(ParameterSetName = "RestrictSource")]
    [Parameter(ParameterSetName = "RestrictTarget")]
    [ValidateScript({Test-ValidListParameter $_ "" "SECURITY","DISTRIBUTION"})]
    [String[]] $GroupType,

    [Parameter(ParameterSetName = "NoRestrict")]
    [Parameter(ParameterSetName = "RestrictSource")]
    [Parameter(ParameterSetName = "RestrictTarget")]
    [Switch] $MatchWildcard,

    [Parameter(ParameterSetName = "NoRestrict")]
    [Parameter(ParameterSetName = "RestrictSource")]
    [Parameter(ParameterSetName = "RestrictTarget")]
    [ValidateScript({Test-ValidListParameter $_ "NONE" "COMPUTER","CONTACT","GROUP","USER"})]
    [String[]] $MemberType,

    [Parameter(ParameterSetName = "NoRestrict")]
    [Parameter(ParameterSetName = "RestrictSource")]
    [Parameter(ParameterSetName = "RestrictTarget")]
    [Switch] $NoNestedMembers,

    [Parameter(ParameterSetName = "NoRestrict")]
    [Parameter(ParameterSetName = "RestrictSource")]
    [Parameter(ParameterSetName = "RestrictTarget")]
    [Switch] $OmitBase,

    [Parameter(ParameterSetName = "RestrictSource")]
    [Switch] $RestrictSource,

    [Parameter(ParameterSetName = "RestrictTarget")]
    [Switch] $RestrictTarget,

    [Parameter(ParameterSetName = "NoRestrict")]
    [Parameter(ParameterSetName = "RestrictSource")]
    [Parameter(ParameterSetName = "RestrictTarget")]
    [String] $SearchBase
  )
  if ( (-not $MatchWildcard) -and ($Name -match '\?|\*') ) {
    throw [ArgumentException] "Cannot validate argument on parameter 'Name'. Wildcards are only permitted with the -MatchWildcard parameter."
  }
  # Do custom runtime parameter validation for -GroupScope, -GroupType, and -MemberType
  if ( $GroupScope ) { $GroupScope = Get-ValidListParameter "GroupScope" $GroupScope "" "LOCAL","GLOBAL","UNIVERSAL" }
  if ( $GroupType  ) { $GroupType  = Get-ValidListParameter "GroupType"  $GroupType  "" "SECURITY","DISTRIBUTION" }
  if ( $MemberType ) { $MemberType = Get-ValidListParameter "MemberType" $MemberType "NONE" "COMPUTER","CONTACT","GROUP","USER" }
  if ( $PSCmdlet.ShouldProcess("Group rule '$Rule' in DRA ActiveView '$ActiveView'","Create") ) {
    # Build argument list
    $argList = "AV",$ActiveView,"ADD",$Rule,"TYPE:GROUP","MATCH:$Name"
    # WAS 2019-07-17 EA ignores
    # if ( $Comment )       { $argList += "COMMENT:$Comment" }
    if ( $Exclude )         { $argList += "ACTION:EXCLUDE" }
    if ( $GroupScope )      { $argList += "GROUPSCOPE:{0}" -f ($GroupScope -join ',') }
    if ( $GroupType )       { $argList += "GROUPTYPE:{0}" -f ($GroupType -join ',') }
    if ( $MatchWildcard )   { $argList += "MATCHWILDCARD" }
    if ( $MemberType )      { $argList += "MEMBERS:{0}" -f ($MemberType -join ',') }
    if ( $NoNestedMembers ) { $argList += "MATCHNESTED:N" }
    if ( $OmitBase )        { $argList += "SELECTBASE:N" }
    if ( $RestrictSource )  { $argList += "RESTRICTION:S" }
    if ( $RestrictTarget )  { $argList += "RESTRICTION:T" }
    if ( $SearchBase )      { $argList += "OU:$SearchBase" }
    $argList += "MODE:B"
    # Invoke EA command
    $output = $null
    $result = Invoke-EA $argList ([Ref] $output) -primary
    if ( ($result -ne 0) -or ($output[$output.Count - 1] -match 'Failed$') ) {
      $PSCmdlet.WriteError((New-Object Management.Automation.ErrorRecord $output[$output.Count - 2],
        $MyInvocation.MyCommand.Name,
        ([Management.Automation.ErrorCategory]::NotSpecified),
        $objectName))
    }
  }
}

# EXPORT
function New-DRAActiveViewOURule {
  <#
  .SYNOPSIS
  Creates a new OU rule in a DRA ActiveView object.

  .DESCRIPTION
  Creates a new OU rule in a DRA ActiveView object.

  .PARAMETER ActiveView
  Specifies the name of the ActiveView to which the rule should be added.

  .PARAMETER Rule
  Specifies the name of the rule.

  .PARAMETER Name
  Specifies the name of the OU to add to the new rule. If -MatchWildcard is specified, this parameter supports wildcards ('?' and '*') and will cause the rule to match all OUs matching the wildcard.

  .PARAMETER Comment
  Specifies a comment for the rule.

  .PARAMETER Exclude
  Specifies that the rule is an exclusion rule (i.e., the rule excludes rather than includes matching objects). The default is an include rule.

  .PARAMETER MatchWildcard
  Specifies that the -Name parameter can contain wildcard characters ('?' and '*').

  .PARAMETER MemberType
  Specifies the types of member objects matched in the rule. This parameter can be 'NONE' or one or more of the following object types: 'COMPUTER', 'CONTACT', 'GROUP', 'OU', and 'USER'. The default is all object types.

  .PARAMETER OmitBase
  Specifies to exclude the base object from the rule. The default is to include the base object.

  .PARAMETER OneLevel
  Specifies not to match objects in child containers. The default is to match objects in child containers.

  .PARAMETER RestrictSource
  Specifies source restriction for the rule (only allow included objects to be cloned, moved, or added to groups).

  .PARAMETER RestrictTarget
  Specifies target restriction for the rule (do not allow included objects to be cloned, moved, or added to groups).

  .INPUTS
  None

  .OUTPUTS
  None
  #>
  [CmdletBinding(SupportsShouldProcess = $true,DefaultParameterSetName = "NoRestrict")]
  param(
    [Parameter(Position = 0,ParameterSetName = "NoRestrict",Mandatory = $true)]
    [Parameter(Position = 0,ParameterSetName = "RestrictSource",Mandatory = $true)]
    [Parameter(Position = 0,ParameterSetName = "RestrictTarget",Mandatory = $true)]
    [ValidateScript({Test-DRAValidObjectNameParameter $_})]
    [Alias("AV")]
    [String] $ActiveView,

    [Parameter(Position = 1,ParameterSetName = "NoRestrict",Mandatory = $true)]
    [Parameter(Position = 1,ParameterSetName = "RestrictSource",Mandatory = $true)]
    [Parameter(Position = 1,ParameterSetName = "RestrictTarget",Mandatory = $true)]
    [ValidateScript({Test-DRAValidObjectNameParameter $_})]
    [String] $Rule,

    [Parameter(Position = 2,ParameterSetName = "NoRestrict",Mandatory = $true)]
    [Parameter(Position = 2,ParameterSetName = "RestrictSource",Mandatory = $true)]
    [Parameter(Position = 2,ParameterSetName = "RestrictTarget",Mandatory = $true)]
    [ValidateScript({Test-DRAValidObjectNameParameter $_ -wildcard})]
    [String] $Name,

    # WAS 2019-07-17 EA ignores
    # [Parameter(ParameterSetName = "NoRestrict")]
    # [Parameter(ParameterSetName = "RestrictSource")]
    # [Parameter(ParameterSetName = "RestrictTarget")]
    # [String] $Comment,

    [Parameter(ParameterSetName = "NoRestrict")]
    [Parameter(ParameterSetName = "RestrictSource")]
    [Parameter(ParameterSetName = "RestrictTarget")]
    [Switch] $Exclude,

    [Parameter(ParameterSetName = "NoRestrict")]
    [Parameter(ParameterSetName = "RestrictSource")]
    [Parameter(ParameterSetName = "RestrictTarget")]
    [Switch] $MatchWildcard,

    [Parameter(ParameterSetName = "NoRestrict")]
    [Parameter(ParameterSetName = "RestrictSource")]
    [Parameter(ParameterSetName = "RestrictTarget")]
    [ValidateScript({Test-ValidListParameter $_ "NONE" "COMPUTER","CONTACT","GROUP","OU","USER"})]
    [String[]] $MemberType,

    [Parameter(ParameterSetName = "NoRestrict")]
    [Parameter(ParameterSetName = "RestrictSource")]
    [Parameter(ParameterSetName = "RestrictTarget")]
    [Switch] $OmitBase,

    [Parameter(ParameterSetName = "NoRestrict")]
    [Parameter(ParameterSetName = "RestrictSource")]
    [Parameter(ParameterSetName = "RestrictTarget")]
    [Switch] $OneLevel,

    [Parameter(ParameterSetName = "RestrictSource")]
    [Switch] $RestrictSource,

    [Parameter(ParameterSetName = "RestrictTarget")]
    [Switch] $RestrictTarget
  )
  if ( (-not $MatchWildcard) -and ($Name -match '\?|\*') ) {
    throw [ArgumentException] "Cannot validate argument on parameter 'Name'. Wildcards are only permitted with the -MatchWildcard parameter."
  }
  # Do custom runtime parameter validation for -MemberType
  if ( $MemberType ) { $MemberType = Get-ValidListParameter "MemberType" $MemberType "NONE" "COMPUTER","CONTACT","GROUP","OU","USER" }
  if ( $PSCmdlet.ShouldProcess("OU rule '$Rule' in DRA ActiveView '$ActiveView'","Create") ) {
    # Build argument list
    $argList = "AV",$ActiveView,"ADD",$Rule,"TYPE:OU","MATCH:$Name"
    # WAS 2019-07-17 EA ignores
    # if ( $Comment )       { $argList += "COMMENT:$Comment" }
    if ( $Exclude )         { $argList += "ACTION:EXCLUDE" }
    if ( $MatchWildcard )   { $argList += "MATCHWILDCARD" }
    if ( $MemberType )      { $argList += "MEMBERS:{0}" -f ($MemberType -join ',') }
    if ( $OmitBase )        { $argList += "SELECTBASE:N" }
    if ( $OneLevel )        { $argList += "RECURSIVE:N" }
    if ( $RestrictSource )  { $argList += "RESTRICTION:S" }
    if ( $RestrictTarget )  { $argList += "RESTRICTION:T" }
    $argList += "MODE:B"
    # Invoke EA command
    $output = $null
    $result = Invoke-EA $argList ([Ref] $output) -primary
    if ( ($result -ne 0) -or ($output[$output.Count - 1] -match 'Failed$') ) {
      $PSCmdlet.WriteError((New-Object Management.Automation.ErrorRecord $output[$output.Count - 2],
        $MyInvocation.MyCommand.Name,
        ([Management.Automation.ErrorCategory]::NotSpecified),
        $objectName))
    }
  }
}

# EXPORT
function New-DRAActiveViewUserRule {
  <#
  .SYNOPSIS
  Creates a new user rule in a DRA ActiveView object.

  .DESCRIPTION
  Creates a new user rule in a DRA ActiveView object.

  .PARAMETER ActiveView
  Specifies the name of the ActiveView to which the rule should be added.

  .PARAMETER Rule
  Specifies the name of the rule.

  .PARAMETER Name
  Specifies the name of the user to add to the new rule. If -MatchWildcard is specified, this parameter supports wildcards ('?' and '*') and will cause the rule to match all users matching the wildcard.

  .PARAMETER Comment
  Specifies a comment for the rule.

  .PARAMETER Exclude
  Specifies that the rule is an exclusion rule (i.e., the rule excludes rather than includes matching objects). The default is an include rule.

  .PARAMETER MatchWildcard
  Specifies that the -Name parameter can contain wildcard characters ('?' and '*').

  .PARAMETER RestrictSource
  Specifies source restriction for the rule (only allow included objects to be cloned, moved, or added to groups).

  .PARAMETER RestrictTarget
  Specifies target restriction for the rule (do not allow included objects to be cloned, moved, or added to groups).

  .PARAMETER SearchBase
  Specifies the distinguished name of an OU where objects should match (i.e., matching objects will only be found within this OU or a child OU).

  .INPUTS
  None

  .OUTPUTS
  None
  #>
  [CmdletBinding(SupportsShouldProcess = $true,DefaultParameterSetName = "NoRestrict")]
  param(
    [Parameter(Position = 0,ParameterSetName = "NoRestrict",Mandatory = $true)]
    [Parameter(Position = 0,ParameterSetName = "RestrictSource",Mandatory = $true)]
    [Parameter(Position = 0,ParameterSetName = "RestrictTarget",Mandatory = $true)]
    [ValidateScript({Test-DRAValidObjectNameParameter $_})]
    [Alias("AV")]
    [String] $ActiveView,

    [Parameter(Position = 1,ParameterSetName = "NoRestrict",Mandatory = $true)]
    [Parameter(Position = 1,ParameterSetName = "RestrictSource",Mandatory = $true)]
    [Parameter(Position = 1,ParameterSetName = "RestrictTarget",Mandatory = $true)]
    [ValidateScript({Test-DRAValidObjectNameParameter $_})]
    [String] $Rule,

    [Parameter(Position = 2,ParameterSetName = "NoRestrict",Mandatory = $true)]
    [Parameter(Position = 2,ParameterSetName = "RestrictSource",Mandatory = $true)]
    [Parameter(Position = 2,ParameterSetName = "RestrictTarget",Mandatory = $true)]
    [ValidateScript({Test-DRAValidObjectNameParameter $_ -wildcard})]
    [String] $Name,

    # WAS 2019-07-17 EA ignores
    # [Parameter(ParameterSetName = "NoRestrict")]
    # [Parameter(ParameterSetName = "RestrictSource")]
    # [Parameter(ParameterSetName = "RestrictTarget")]
    # [String] $Comment,

    [Parameter(ParameterSetName = "NoRestrict")]
    [Parameter(ParameterSetName = "RestrictSource")]
    [Parameter(ParameterSetName = "RestrictTarget")]
    [Switch] $Exclude,

    [Parameter(ParameterSetName = "NoRestrict")]
    [Parameter(ParameterSetName = "RestrictSource")]
    [Parameter(ParameterSetName = "RestrictTarget")]
    [Switch] $MatchWildcard,

    [Parameter(ParameterSetName = "RestrictSource")]
    [Switch] $RestrictSource,

    [Parameter(ParameterSetName = "RestrictTarget")]
    [Switch] $RestrictTarget,

    [Parameter(ParameterSetName = "NoRestrict")]
    [Parameter(ParameterSetName = "RestrictSource")]
    [Parameter(ParameterSetName = "RestrictTarget")]
    [String] $SearchBase
  )
  if ( (-not $MatchWildcard) -and ($Name -match '\?|\*') ) {
    throw [ArgumentException] "Cannot validate argument on parameter 'Name'. Wildcards are only permitted with the -MatchWildcard parameter."
  }
  if ( $PSCmdlet.ShouldProcess("User rule '$Rule' in DRA ActiveView '$ActiveView'","Create") ) {
    # Build argument list
    $argList = "AV",$ActiveView,"ADD",$Rule,"TYPE:USER","MATCH:$Name"
    # WAS 2019-07-17 EA ignores
    # if ( $Comment )      { $argList += "COMMENT:$Comment" }
    if ( $Exclude )        { $argList += "ACTION:EXCLUDE" }
    if ( $MatchWildcard )  { $argList += "MATCHWILDCARD" }
    if ( $RestrictSource ) { $argList += "RESTRICTION:S" }
    if ( $RestrictTarget ) { $argList += "RESTRICTION:T" }
    if ( $SearchBase )     { $argList += "OU:$SearchBase" }
    $argList += "MODE:B"
    # Invoke EA command
    $output = $null
    $result = Invoke-EA $argList ([Ref] $output) -primary
    if ( ($result -ne 0) -or ($output[$output.Count - 1] -match 'Failed$') ) {
      $PSCmdlet.WriteError((New-Object Management.Automation.ErrorRecord $output[$output.Count - 2],
        $MyInvocation.MyCommand.Name,
        ([Management.Automation.ErrorCategory]::NotSpecified),
        $objectName))
    }
  }
}

# EXPORT
function New-DRAAssistantAdminAARule {
  <#
  .SYNOPSIS
  Creates a new DRA AssitantAdmin rule in a DRA AssistantAdmin object.

  .DESCRIPTION
  Creates a new DRA AssitantAdmin rule in a DRA AssistantAdmin object.

  .PARAMETER AssistantAdmin
  Specifies the name of the AssistantAdmin to which the rule should be added.

  .PARAMETER Rule
  Specifies the name of the rule.

  .PARAMETER Name
  Specifies the name of the DRA AssistantAdmin to add to the new rule.

  .PARAMETER MatchWildcard
  Specifies that the -Name parameter can contain wildcard characters ('?' and '*').

  .INPUTS
  None

  .OUTPUTS
  None
  #>
  [CmdletBinding(SupportsShouldProcess = $true)]
  param(
    [Parameter(Position = 0,Mandatory = $true)]
    [ValidateScript({Test-DRAValidObjectNameParameter $_})]
    [Alias("AA")]
    [String] $AssistantAdmin,

    [Parameter(Position = 1,Mandatory = $true)]
    [ValidateScript({Test-DRAValidObjectNameParameter $_})]
    [String] $Rule,

    [Parameter(Position = 2,Mandatory = $true)]
    [ValidateScript({Test-DRAValidObjectNameParameter $_ -wildcard})]
    [String] $Name,

    [Switch] $MatchWildcard
  )
  if ( $AssistantAdmin -eq $Name ) {
    $PSCmdlet.ThrowTerminatingError((New-Object Management.Automation.ErrorRecord "The Rule cannot be created because you cannot add a DRA AssistantAdmin as a rule for itself.",
      $MyInvocation.MyCommand.Name,
      ([Management.Automation.ErrorCategory]::InvalidArgument),
      $AssistantAdmin))
    return
  }
  if ( (-not $MatchWildcard) -and ($Name -match '\?|\*') ) {
    throw [ArgumentException] "Cannot validate argument on parameter 'Name'. Wildcards are only permitted with the -MatchWildcard parameter."
  }
  if ( -not (GetDRAObject "AssistantAdmin" $AssistantAdmin) ) {
    $PSCmdlet.ThrowTerminatingError((New-Object Management.Automation.ErrorRecord "DRA AssistantAdmin '$AssistantAdmin' not found.",
      $MyInvocation.MyCommand.Name,
      ([Management.Automation.ErrorCategory]::ObjectNotFound),
      $AssistantAdmin))
    return
  }
  if ( -not (GetDRAObject "AssistantAdmin" $Name) ) {
    $PSCmdlet.ThrowTerminatingError((New-Object Management.Automation.ErrorRecord "The Rule cannot be created because no DRA AssistantAdmin matching the name '$Name' was found.",
      $MyInvocation.MyCommand.Name,
      ([Management.Automation.ErrorCategory]::ObjectNotFound),
      $AssistantAdmin))
    return
  }
  if ( GetDRAObjectRule "AssistantAdmin" $AssistantAdmin $Rule ) {
    $PSCmdlet.ThrowTerminatingError((New-Object Management.Automation.ErrorRecord "The Rule cannot be created because an object with the name '$Rule' already exists.",
      $MyInvocation.MyCommand.Name,
      ([Management.Automation.ErrorCategory]::ResourceExists),
      $AssistantAdmin))
    return
  }
  if ( -not $FORCE_PRIMARY ) {
    if ( -not $draServerName ) {
      $draServerName = (Get-DRAServer | Get-Random).Name
    }
  }
  else {
    $draServerName = (Get-DRAServer -Primary).Name
  }
  $eaType = [Type]::GetTypeFromProgID("EAServer.EAServe",$draServerName)
  if ( -not $eaType ) {
    throw [Management.Automation.ItemNotFoundException] "Computer '$draServerName' is not a DRA server, or is not reachable."
  }
  try {
    $eaServer = [Activator]::CreateInstance($eaType)
    $varSetIn = New-Object -ComObject "NetIQDraVarSet.VarSet"
  }
  catch [Management.Automation.MethodInvocationException],[Runtime.InteropServices.COMException] {
    throw $_
  }
  Write-Debug "New-DRAAssistantAdminAARule: Connect to server '$draServerName'"
  $varSetIn.put("Container","OnePoint://aa=$(Get-EscapedName $AssistantAdmin),module=Security")
  $varSetIn.put("OperationName","RuleCreate")
  $varSetIn.put("Rule","rule=$Rule")
  $varSetIn.put("Properties.ClientFlags",0x12)
  $varSetIn.put("Properties.Description","Include Assistant Admin Groups matching $Name")
  $varSetIn.put("Properties.IncludeFlag",$true)
  $varSetIn.put("Properties.NestedFlag",$true)
  $varSetIn.put("Properties.SourceFlag",$true)
  $varSetIn.put("Properties.TargetFlag",$true)
  $varSetIn.put("RuleSpecification","AssistantAdminByNameRule")
  $varSetIn.put("RuleSpecification.MatchString",$Name)
  $varSetIn.put("RuleSpecification.SelectBase",$false)
  $varSetOut = $eaServer.ScriptSubmit($varSetIn)
  $lastError = $varSetOut.get("Errors.LastError")
  if ( $lastError -ne 0 ) {
    $PSCmdlet.ThrowTerminatingError((New-Object Management.Automation.ErrorRecord ("New-DRAAssistantAdminAARule returned error 0x{0:X8}." -f $lastError),
      (Get-Variable MyInvocation -Scope 1).Value.MyCommand.Name,
      ([Management.Automation.ErrorCategory]::NotSpecified),
      $objectName))
  }
}

# EXPORT
function New-DRAAssistantAdminGroupRule {
  <#
  .SYNOPSIS
  Creates a new group rule in a DRA AssistantAdmin object.

  .DESCRIPTION
  Creates a new group rule in a DRA AssistantAdmin object.

  .PARAMETER AssistantAdmin
  Specifies the name of the AssistantAdmin to which the rule should be added.

  .PARAMETER Rule
  Specifies the name of the rule.

  .PARAMETER Name
  Specifies the name of the group to add to the new rule.

  .PARAMETER Comment
  Specifies a comment for the rule.

  .PARAMETER GroupScope
  Specifies the group scope for the rule. This parameter can be one or more of the following group scopes: 'LOCAL', 'GLOBAL', and 'UNIVERSAL'. The default is all group scopes.

  .PARAMETER GroupType
  Specifies the group type for the rule. This parameter can be one or more of the following group types: 'SECURITY' and 'DISTRIBUTION'. The default is all group types.

  .PARAMETER MatchWildcard
  Specifies that the -Name parameter can contain wildcard characters ('?' and '*').

  .PARAMETER MemberType
  Specifies the types of member objects matched in the rule. This parameter can be one or more of the following object types: 'COMPUTER', 'CONTACT', 'GROUP', and 'USER'. The default is all object types.

  .PARAMETER NoNestedMembers
  Specifies not to manage nested members. The default is to manage nested members.

  .PARAMETER OmitBase
  Specifies to exclude the base object from the rule. The default is to include the base object.

  .PARAMETER SearchBase
  Specifies the distinguished name of an OU where objects should match (i.e., matching objects will only be found within this OU or a child OU).

  .INPUTS
  None

  .OUTPUTS
  None
  #>
  [CmdletBinding(SupportsShouldProcess = $true)]
  param(
    [Parameter(Position = 0,Mandatory = $true)]
    [ValidateScript({Test-DRAValidObjectNameParameter $_})]
    [Alias("AA")]
    [String] $AssistantAdmin,

    [Parameter(Position = 1,Mandatory = $true)]
    [ValidateScript({Test-DRAValidObjectNameParameter $_})]
    [String] $Rule,

    [Parameter(Position = 2,Mandatory = $true)]
    # WAS 2019-07-17 DRA supports wildcard matching, but EA creates broken rule
    # [ValidateScript({Test-DRAValidObjectNameParameter $_ -wildcard})]
    [ValidateScript({Test-DRAValidObjectNameParameter $_})]
    [String] $Name,

    # WAS 2019-07-17 EA ignores
    # [String] $Comment,

    [ValidateScript({Test-ValidListParameter $_ "" "LOCAL","GLOBAL","UNIVERSAL"})]
    [String[]] $GroupScope,

    [ValidateScript({Test-ValidListParameter $_ "" "SECURITY","DISTRIBUTION"})]
    [String[]] $GroupType,

    [ValidateScript({Test-ValidListParameter $_ "NONE" "COMPUTER","CONTACT","GROUP","USER"})]
    [String[]] $MemberType,

    # WAS 2019-07-17 DRA supports wildcard matching, but EA creates broken rule
    # [Switch] $MatchWildcard,

    [Switch] $NoNestedMembers,

    [Switch] $OmitBase,

    [String] $SearchBase
  )
  # WAS 2019-07-17 DRA supports wildcard matching, but EA creates broken rule
  # if ( (-not $MatchWildcard) -and ($Name -match '\?|\*') ) {
  #   throw [ArgumentException] "Cannot validate argument on parameter 'Name'. Wildcards are only permitted with the -MatchWildcard parameter."
  # }
  # Do custom runtime parameter validation for -GroupScope, -GroupType, and -MemberType
  if ( $GroupScope ) { $GroupScope = Get-ValidListParameter "GroupScope" $GroupScope "" "LOCAL","GLOBAL","UNIVERSAL" }
  if ( $GroupType  ) { $GroupType  = Get-ValidListParameter "GroupType"  $GroupType  "" "SECURITY","DISTRIBUTION" }
  if ( $MemberType ) { $MemberType = Get-ValidListParameter "MemberType" $MemberType "NONE" "COMPUTER","CONTACT","GROUP","USER" }
  if ( $PSCmdlet.ShouldProcess("Group rule '$Rule' in DRA AssistantAdmin '$AssistantAdmin'","Create") ) {
    # Build argument list
    # WAS 2019-07-17 EA ignores
    # if ( $Comment )       { $argList += "COMMENT:$Comment" }
    $argList = "AA",$AssistantAdmin,"ADD",$Rule,"TYPE:GROUP","MATCH:$Name"
    if ( $GroupScope )      { $argList += "GROUPSCOPE:{0}" -f ($GroupScope -join ',') }
    if ( $GroupType )       { $argList += "GROUPTYPE:{0}" -f ($GroupType -join ',') }
    # WAS 2019-07-17 DRA supports wildcard matching, but EA creates broken rule
    # if ( $MatchWildcard ) { $argList += "MATCHWILDCARD" }
    if ( $MemberType )      { $argList += "MEMBERS:{0}" -f ($MemberType -join ',') }
    if ( $NoNestedMembers ) { $argList += "MATCHNESTED:N" }
    if ( $OmitBase )        { $argList += "SELECTBASE:N" }
    if ( $SearchBase )      { $argList += "OU:$SearchBase" }
    $argList += "MODE:B"
    # Invoke EA command
    $output = $null
    $result = Invoke-EA $argList ([Ref] $output) -primary
    if ( ($result -ne 0) -or ($output[$output.Count - 1] -match 'Failed$') ) {
      $PSCmdlet.WriteError((New-Object Management.Automation.ErrorRecord $output[$output.Count - 2],
        $MyInvocation.MyCommand.Name,
        ([Management.Automation.ErrorCategory]::NotSpecified),
        $objectName))
    }
  }
}

# EXPORT
function New-DRAAssistantAdminUserRule {
  <#
  .SYNOPSIS
  Creates a new user rule in a DRA AssistantAdmin object.

  .DESCRIPTION
  Creates a new user rule in a DRA AssistantAdmin object.

  .PARAMETER AssistantAdmin
  Specifies the name of the AssistantAdmin to which the rule should be added.

  .PARAMETER Rule
  Specifies the name of the rule.

  .PARAMETER Name
  Specifies the name of the user to add to the new rule.

  .PARAMETER Comment
  Specifies a comment for the rule.

  .PARAMETER MatchWildcard
  Specifies that the -Name parameter can contain wildcard characters ('?' and '*').

  .PARAMETER SearchBase
  Specifies the distinguished name of an OU where objects should match (i.e., matching objects will only be found within this OU or a child OU).

  .INPUTS
  None

  .OUTPUTS
  None
  #>
  [CmdletBinding(SupportsShouldProcess = $true)]
  param(
    [Parameter(Position = 0,Mandatory = $true)]
    [ValidateScript({Test-DRAValidObjectNameParameter $_})]
    [Alias("AA")]
    [String] $AssistantAdmin,

    [Parameter(Position = 1,Mandatory = $true)]
    [ValidateScript({Test-DRAValidObjectNameParameter $_})]
    [String] $Rule,

    [Parameter(Position = 2,Mandatory = $true)]
    # WAS 2019-07-17 DRA supports wildcard matching, but EA creates broken rule
    # [ValidateScript({Test-DRAValidObjectNameParameter $_ -wildcard})]
    [ValidateScript({Test-DRAValidObjectNameParameter $_})]
    [String] $Name,

    # WAS 2019-07-17 EA ignores
    # [String] $Comment,

    # WAS 2019-07-17 DRA supports, but EA creates broken rule
    # [Switch] $MatchWildcard,

    [String] $SearchBase
  )
  # WAS 2019-07-17 DRA supports wildcard matching, but EA creates broken rule
  # if ( (-not $MatchWildcard) -and ($Name -match '\?|\*') ) {
  #   throw [ArgumentException] "Cannot validate argument on parameter 'Name'. Wildcards are only permitted with the -MatchWildcard parameter."
  # }
  if ( $PSCmdlet.ShouldProcess("User rule '$Rule' in DRA AssistantAdmin '$AssistantAdmin'","Create") ) {
    # Build argument list
    $argList = "AA",$AssistantAdmin,"ADD",$Rule,"TYPE:USER","MATCH:$Name"
    # WAS 2019-07-17 EA ignores
    # if ( $Comment )       { $argList += "COMMENT:$Comment" }
    # WAS 2019-07-17 DRA supports wildcard matching, but EA creates broken rule
    # if ( $MatchWildcard ) { $argList += "MATCHWILDCARD" }
    if ( $SearchBase )      { $argList += "OU:$SearchBase" }
    $argList += "MODE:B"
    # Invoke EA command
    $output = $null
    $result = Invoke-EA $argList ([Ref] $output) -primary
    if ( ($result -ne 0) -or ($output[$output.Count - 1] -match 'Failed$') ) {
      $PSCmdlet.WriteError((New-Object Management.Automation.ErrorRecord $output[$output.Count - 2],
        $MyInvocation.MyCommand.Name,
        ([Management.Automation.ErrorCategory]::NotSpecified),
        $objectName))
    }
  }
}
#------------------------------------------------------------------------------

#------------------------------------------------------------------------------
# CATEGORY: Remove object rules
#------------------------------------------------------------------------------
function RemoveDRAObjectRule {
  param(
    [Parameter(Mandatory = $true)]
    [ValidateSet("ActiveView","AssistantAdmin")]
    [String] $objectType,

    [Parameter(Mandatory = $true)]
    [ValidateNotNullOrEmpty()]
    [String] $objectName,

    [Parameter(Mandatory = $true)]
    [ValidateNotNullOrEmpty()]
    [String] $ruleName
  )
  $output = $null
  $result = Invoke-EA (Get-EAParamName $objectType),$objectName,"REMOVE",$ruleName,"MODE:B" ([Ref] $output) -primary
  if ( $result -ne 0 ) {
    $errorMsg = $output[0]
  }
  else {
    if ( $output[1] -match 'not found\.+$' ) {
      $result = 2
      $name = [Regex]::Match($output[0],'''(.+)''').Groups[1].Value
      $errorMsg = "{0} in DRA $objectType '$name'." -f ($output[1] -replace '^\s+','' -replace '\.+$','')
    }
  }
  if ( $result -ne 0 ) {
    (Get-Variable PSCmdlet -Scope 1).Value.WriteError((New-Object Management.Automation.ErrorRecord $errorMsg,
      (Get-Variable MyInvocation -Scope 1).Value.MyCommand.Name,
      ([Management.Automation.ErrorCategory]::NotSpecified),
      $objectName))
  }
}

# EXPORT
function Remove-DRAActiveViewRule {
  <#
  .SYNOPSIS
  Removes one or more rules from a DRA ActiveView object.

  .DESCRIPTION
  Removes one or more rules from a DRA ActiveView object.

  .PARAMETER ActiveView
  Specifies the name of the DRA ActiveView from which rules will be removed. This parameter supports wildcards ('?' and '*').

  .PARAMETER Rule
  Specifies the name of one or more rules to be removed. This parameter supports wildcards ('?' and '*').

  .INPUTS
  * System.String
  * Objects with property 'Rule'
  * Objects with properties 'ActiveView' and 'Rule'

  .OUTPUTS
  None

  .NOTES
  Use wildcards ('?' and '*') with caution as it is possible to delete multiple object rules with a single command.

  To prevent the confirmation prompt, use '-Confirm:0' or set the $ConfirmPreference variable to 'Low'.
  #>
  [CmdletBinding(SupportsShouldProcess = $true,ConfirmImpact = "High")]
  param(
    [Parameter(Position = 0,Mandatory = $true,ValueFromPipelineByPropertyName = $true)]
    [ValidateScript({Test-DRAValidObjectNameParameter $_ -wildcard})]
    [SupportsWildcards()]
    [Alias("AV")]
    [String] $ActiveView,

    [Parameter(Position = 1,Mandatory = $true,ValueFromPipeline = $true,ValueFromPipelineByPropertyName = $true)]
    [ValidateScript({Test-DRAValidObjectNameParameter $_ -wildcard})]
    [SupportsWildcards()]
    [String[]] $Rule
  )
  process {
    foreach ( $ruleItem in $Rule ) {
      if ( $PSCmdlet.ShouldProcess("Rule '$ruleItem' in DRA ActiveView '$ActiveView'","Remove") ) {
        try {
          RemoveDRAObjectRule "ActiveView" $ActiveView $ruleItem
        }
        catch {
          $PSCmdlet.WriteError($_)
        }
      }
    }
  }
}

# EXPORT
function Remove-DRAAssistantAdminRule {
  <#
  .SYNOPSIS
  Removes one or more rules from a DRA AssistantAdmin object.

  .DESCRIPTION
  Removes one or more rules from a DRA AssistantAdmin object.

  .PARAMETER AssistantAdmin
  Specifies the name of the DRA AssistantAdmin from which rules will be removed. This parameter supports wildcards ('?' and '*').

  .PARAMETER Rule
  Specifies the name of one or more rules to be removed. This parameter supports wildcards ('?' and '*').

  .INPUTS
  * System.String
  * Objects with property 'Rule'
  * Objects with properties 'AssistantAdmin' and 'Rule'

  .OUTPUTS
  None

  .NOTES
  Use wildcards ('?' and '*') with caution as it is possible to delete multiple object rules with a single command.

  To prevent the confirmation prompt, use '-Confirm:0' or set the $ConfirmPreference variable to 'Low'.
  #>
  [CmdletBinding(SupportsShouldProcess = $true,ConfirmImpact = "High")]
  param(
    [Parameter(Position = 0,Mandatory = $true,ValueFromPipelineByPropertyName = $true)]
    [ValidateScript({Test-DRAValidObjectNameParameter $_ -wildcard})]
    [SupportsWildcards()]
    [Alias("AA")]
    [String] $AssistantAdmin,

    [Parameter(Position = 1,Mandatory = $true,ValueFromPipeline = $true,ValueFromPipelineByPropertyName = $true)]
    [ValidateScript({Test-DRAValidObjectNameParameter $_ -wildcard})]
    [SupportsWildcards()]
    [String[]] $Rule
  )
  process {
    foreach ( $ruleItem in $Rule ) {
      if ( $PSCmdlet.ShouldProcess("Rule '$ruleItem' in DRA AssistantAdmin '$AssistantAdmin'","Remove") ) {
        try {
          RemoveDRAObjectRule "AssistantAdmin" $AssistantAdmin $ruleItem
        }
        catch {
          $PSCmdlet.WriteError($_)
        }
      }
    }
  }
}
#------------------------------------------------------------------------------

#------------------------------------------------------------------------------
# CATEGORY: Grant/revoke delegation
#------------------------------------------------------------------------------
# EXPORT
function Grant-DRADelegation {
  <#
  .SYNOPSIS
  Grants an AssistantAdmin a delegated Role over an ActiveView.

  .DESCRIPTION
  Grants an AssistantAdmin a delegated Role over an ActiveView.

  .PARAMETER AssistantAdmin
  Specifies the name of the AssistantAdmin to which the Role will be delegated.

  .PARAMETER Role
  Specifies the name of the Role to delegate to the AssistantAdmin. (DRA Powers are not supported.)

  .PARAMETER ActiveView
  Specifies the name of the ActiveView to which the delegation applies.

  .INPUTS
  None

  .OUTPUTS
  None
  #>
  [CmdletBinding(SupportsShouldProcess = $true)]
  param(
    [Parameter(Position = 0,Mandatory = $true)]
    [ValidateScript({Test-DRAValidObjectNameParameter $_})]
    [Alias("AA")]
    [String] $AssistantAdmin,

    [Parameter(Position = 1,Mandatory = $true)]
    [ValidateScript({Test-DRAValidObjectNameParameter $_})]
    [String] $Role,

    [Parameter(Position = 2,Mandatory = $true)]
    [ValidateScript({Test-DRAValidObjectNameParameter $_})]
    [Alias("AV")]
    [String] $ActiveView
  )
  if ( $PSCmdlet.ShouldProcess("DRA AssistantAdmin '$AssistantAdmin'","Grant Role '$Role' over ActiveView '$ActiveView'") ) {
    $output = $null
    $result = Invoke-EA "AV",$ActiveView,"DELEGATE","ADMIN:$AssistantAdmin","ROLE:$Role","MODE:B" ([Ref] $output) -primary
    if ( $result -ne 0 ) {
      $PSCmdlet.WriteError((New-Object Management.Automation.ErrorRecord $output[$output.Count - 1],
        $MyInvocation.MyCommand.Name,
        ([Management.Automation.ErrorCategory]::NotSpecified),
        $AssistantAdmin))
    }
  }
}

# EXPORT
function Revoke-DRADelegation {
  <#
  .SYNOPSIS
  Revokes an AssistantAdmin's delegated Role over an ActiveView.

  .DESCRIPTION
  Revokes an AssistantAdmin's delegated Role over an ActiveView.

  .PARAMETER AssistantAdmin
  Specifies the name of the AssistantAdmin from which to revoke the delegation.

  .PARAMETER Role
  Specifies the name of the Role to be revoked. (DRA Powers are not supported.)

  .PARAMETER ActiveView
  Specifies the name of the ActiveView to which the delegation applies.

  .INPUTS
  * System.String
  * Objects with property 'AssistantAdmin'
  * Objects with properties 'AssistantAdmin', 'Role', and 'ActiveView'

  .OUTPUTS
  None

  .NOTES
  To prevent the confirmation prompt, use '-Confirm:0' or set the $ConfirmPreference variable to 'Low'.
  #>
  [CmdletBinding(SupportsShouldProcess = $true,ConfirmImpact = "High")]
  param(
    [Parameter(Position = 0,Mandatory = $true,ValueFromPipeline = $true,ValueFromPipelineByPropertyName = $true)]
    [ValidateScript({Test-DRAValidObjectNameParameter $_})]
    [Alias("AA")]
    [String] $AssistantAdmin,

    [Parameter(Position = 1,Mandatory = $true,ValueFromPipelineByPropertyName = $true)]
    [ValidateScript({Test-DRAValidObjectNameParameter $_})]
    [String] $Role,

    [Parameter(Position = 2,Mandatory = $true,ValueFromPipelineByPropertyName = $true)]
    [ValidateScript({Test-DRAValidObjectNameParameter $_})]
    [Alias("AV")]
    [String] $ActiveView
  )
  process {
    foreach ( $assistantAdminItem in $AssistantAdmin ) {
      if ( $PSCmdlet.ShouldProcess("DRA AssistantAdmin '$assistantAdminItem'","Revoke Role '$Role' over ActiveView '$ActiveView'") ) {
        $output = $null
        $result = Invoke-EA "AV",$ActiveView,"REVOKE","ADMIN:$assistantAdminItem","ROLE:$Role","MODE:B" ([Ref] $output) -primary
        if ( $result -ne 0 ) {
          $errorMsg = $output[$output.Count - 1]
        }
        else {
          if ( $output[$output.Count - 2] -match 'not associated' ) {
            $result = 2
            $errorMsg = "The specified delegation for DRA AssistantAdmin '$assistantAdminItem' was not found."
          }
        }
        if ( $result -ne 0 ) {
          $PSCmdlet.WriteError((New-Object Management.Automation.ErrorRecord $errorMsg,
            $MyInvocation.MyCommand.Name,
            ([Management.Automation.ErrorCategory]::NotSpecified),
            $AssistantAdmin))
        }
      }
    }
  }
}
#------------------------------------------------------------------------------

#------------------------------------------------------------------------------
# CATEGORY: Rename objects
#------------------------------------------------------------------------------
function RenameDRAObject {
  [CmdletBinding()]
  param(
    [Parameter(Mandatory = $true)]
    [ValidateSet("ActiveView","AssistantAdmin")]
    [String] $objectType,

    [Parameter(Mandatory = $true)]
    [ValidateNotNullOrEmpty()]
    [String] $objectName,

    [Parameter(Mandatory = $true)]
    [ValidateNotNullOrEmpty()]
    [String] $newName
  )
  $output = $null
  $result = Invoke-EA (Get-EAParamName $objectType),$objectName,"UPDATE","NAME:$newName","MODE:B" ([Ref] $output) -primary
  if ( $result -ne 0 ) {
    if ( $output[$output.Count - 1] -match 'Failed$' ) {
      $errorMsg = $output[1]
    }
    else {
      $errorMsg = $output[0]
    }
    $PSCmdlet.ThrowTerminatingError((New-Object Management.Automation.ErrorRecord $errorMsg,
      (Get-Variable MyInvocation -Scope 1).Value.MyCommand.Name,
      ([Management.Automation.ErrorCategory]::NotSpecified),
      $objectName))
  }
}

# EXPORT
function Rename-DRAActiveView {
  <#
  .SYNOPSIS
  Renames a DRA ActiveView object.

  .DESCRIPTION
  Renames a DRA ActiveView object.

  .PARAMETER ActiveView
  Specifies the name of the ActiveView to be renamed.

  .PARAMETER NewName
  Specifies the new name of the ActiveView.

  .INPUTS
  None

  .OUTPUTS
  None
  #>
  [CmdletBinding(SupportsShouldProcess = $true)]
  param(
    [Parameter(Position = 0,Mandatory = $true)]
    [ValidateScript({Test-DRAValidObjectNameParameter $_})]
    [Alias("AV")]
    [String] $ActiveView,

    [Parameter(Position = 1,Mandatory = $true)]
    [ValidateScript({Test-DRAValidObjectNameParameter $_})]
    [String] $NewName
  )
  if ( $PSCmdlet.ShouldProcess("DRA ActiveView '$ActiveView'","Rename") ) {
    try {
      RenameDRAObject "ActiveView" $ActiveView $NewName
    }
    catch {
      $PSCmdlet.WriteError($_)
    }
  }
}

# EXPORT
function Rename-DRAAssistantAdmin {
  <#
  .SYNOPSIS
  Renames a DRA AssistantAdmin object.

  .DESCRIPTION
  Renames a DRA AssistantAdmin object.

  .PARAMETER AssistantAdmin
  Specifies the name of the AssistantAdmin to be renamed.

  .PARAMETER NewName
  Specifies the new name of the AssistantAdmin.

  .INPUTS
  None

  .OUTPUTS
  None
  #>
  [CmdletBinding(SupportsShouldProcess = $true)]
  param(
    [Parameter(Position = 0,Mandatory = $true)]
    [ValidateScript({Test-DRAValidObjectNameParameter $_})]
    [Alias("AA")]
    [String] $AssistantAdmin,

    [Parameter(Position = 1,Mandatory = $true)]
    [ValidateScript({Test-DRAValidObjectNameParameter $_})]
    [String] $NewName
  )
  if ( $PSCmdlet.ShouldProcess("DRA AssistantAdmin '$AssistantAdmin'","Rename") ) {
    try {
      RenameDRAObject "AssistantAdmin" $AssistantAdmin $NewName
    }
    catch {
      $PSCmdlet.WriteError($_)
    }
  }
}

# EXPORT
function Rename-DRARole {
  <#
  .SYNOPSIS
  Renames a DRA Role object.

  .DESCRIPTION
  Renames a DRA Role object.

  .PARAMETER Role
  Specifies the name of the Role to be renamed.

  .PARAMETER NewName
  Specifies the new name of the Role.

  .INPUTS
  None

  .OUTPUTS
  None
  #>
  [CmdletBinding(SupportsShouldProcess = $true)]
  param(
    [Parameter(Position = 0,Mandatory = $true)]
    [ValidateScript({Test-DRAValidObjectNameParameter $_})]
    [String] $Role,

    [Parameter(Position = 1,Mandatory = $true)]
    [ValidateScript({Test-DRAValidObjectNameParameter $_})]
    [String] $NewName
  )
  if ( $PSCmdlet.ShouldProcess("DRA Role '$Role'","Rename") ) {
    $draServerName = (Get-DRAServer -Primary).Name
    $eaType = [Type]::GetTypeFromProgID("EAServer.EAServe",$draServerName)
    if ( -not $eaType ) {
      throw [Management.Automation.ItemNotFoundException] "Computer '$draServerName' is not a DRA server, or is not reachable."
    }
    try {
      $eaServer = [Activator]::CreateInstance($eaType)
      $varSetIn = New-Object -ComObject "NetIQDraVarSet.VarSet"
    }
    catch [Management.Automation.MethodInvocationException] {
      throw $_
    }
    Write-Debug "Rename-DRARole: Connect to server '$draServerName'"
    $roleName = GetDRAObject "Role" $Role -emptyIsError | Select-Object -ExpandProperty Role
    if ( $roleName ) {
      if ( -not (GetDRAObject "Role" $NewName) ) {
        $varSetIn.put("Role","OnePoint://role=$(Get-EscapedName $roleName),module=Security")
        $varSetIn.put("OperationName","RoleMoveHere")
        $varSetIn.put("NewName","role=$NewName")
        $varSetOut = $eaServer.ScriptSubmit($varSetIn)
        $lastError = $varSetOut.get("Errors.LastError")
        if ( $lastError -ne 0 ) {
          $PSCmdlet.ThrowTerminatingError((New-Object Management.Automation.ErrorRecord ("Rename-DRARole returned 0x{0:X8}." -f $lastError),
            (Get-Variable MyInvocation -Scope 1).Value.MyCommand.Name,
            ([Management.Automation.ErrorCategory]::NotSpecified),
            $Role))
        }
      }
      else {
        $PSCmdlet.WriteError((New-Object Management.Automation.ErrorRecord "The Role cannot be renamed because an object with the name '$newName' already exists.",
        $MyInvocation.MyCommand.Name,
        ([Management.Automation.ErrorCategory]::ObjectNotFound),
        $Role))
      }
    }
  }
}
#------------------------------------------------------------------------------

#------------------------------------------------------------------------------
# CATEGORY: Rename object rules
#------------------------------------------------------------------------------
function RenameDRAObjectRule {
  [CmdletBinding()]
  param(
    [Parameter(Mandatory = $true)]
    [ValidateSet("ActiveView","AssistantAdmin")]
    [String] $objectType,

    [Parameter(Mandatory = $true)]
    [ValidateNotNullOrEmpty()]
    [String] $objectName,

    [Parameter(Mandatory = $true)]
    [ValidateNotNullOrEmpty()]
    [String] $ruleName,

    [Parameter(Mandatory = $true)]
    [ValidateNotNullOrEmpty()]
    [String] $newName
  )
  $output = $null
  $result = Invoke-EA (Get-EAParamName $objectType),$objectName,"UPDATERULES",$ruleName,"NAME:$newName","MODE:B" ([Ref] $output) -primary
  if ( $result -ne 0 ) {
    if ( $output[$output.Count - 1] -match 'Failed$' ) {
      $errorMsg = $output[$output.Count - 2]
    }
    else {
      $errorMsg = $output[0]
    }
  }
  else {
    if ( $output[$output.Count - 1] -match 'not found\.$' ) {
      $result = 2
      $errorMsg = "{0} in DRA $objectType '$objectName'." -f ($output[$output.Count - 1] -replace '^\s+','' -replace '\.+$','')
    }
  }
  if ( $result -ne 0 ) {
    $PSCmdlet.ThrowTerminatingError((New-Object Management.Automation.ErrorRecord $errorMsg,
      (Get-Variable MyInvocation -Scope 1).Value.MyCommand.Name,
      ([Management.Automation.ErrorCategory]::NotSpecified),
      $objectName))

  }
}

# EXPORT
function Rename-DRAActiveViewRule {
  <#
  .SYNOPSIS
  Renames a DRA ActiveView object rule.

  .DESCRIPTION
  Renames a DRA ActiveView object rule.

  .PARAMETER ActiveView
  Specifies the name of the ActiveView.

  .PARAMETER Rule
  Specifies the name of the rule to be renamed.

  .PARAMETER NewName
  Specifies the new rule name.

  .INPUTS
  None

  .OUTPUTS
  None
  #>
  [CmdletBinding(SupportsShouldProcess = $true)]
  param(
    [Parameter(Position = 0,Mandatory = $true)]
    [ValidateScript({Test-DRAValidObjectNameParameter $_})]
    [Alias("AV")]
    [String] $ActiveView,

    [Parameter(Position = 1,Mandatory = $true)]
    [ValidateScript({Test-DRAValidObjectNameParameter $_})]
    [String] $Rule,

    [Parameter(Position = 2,Mandatory = $true)]
    [ValidateScript({Test-DRAValidObjectNameParameter $_})]
    [String] $NewName
  )
  if ( $PSCmdlet.ShouldProcess("Rule '$Rule' in DRA ActiveView '$ActiveView'","Rename") ) {
    try {
      RenameDRAObjectRule "ActiveView" $ActiveView $Rule $NewName
    }
    catch {
      $PSCmdlet.WriteError($_)
    }
  }
}

# EXPORT
function Rename-DRAAssistantAdminRule {
  <#
  .SYNOPSIS
  Renames a DRA AssistantAdmin object rule.

  .DESCRIPTION
  Renames a DRA AssistantAdmin object rule.

  .PARAMETER AssistantAdmin
  Specifies the name of the AssistantAdmin.

  .PARAMETER Rule
  Specifies the name of the rule to be renamed.

  .PARAMETER NewName
  Specifies the new rule name.

  .INPUTS
  None

  .OUTPUTS
  None
  #>
  [CmdletBinding(SupportsShouldProcess = $true)]
  param(
    [Parameter(Position = 0,Mandatory = $true)]
    [ValidateScript({Test-DRAValidObjectNameParameter $_})]
    [Alias("AA")]
    [String] $AssistantAdmin,

    [Parameter(Position = 1,Mandatory = $true)]
    [ValidateScript({Test-DRAValidObjectNameParameter $_})]
    [String] $Rule,

    [Parameter(Position = 2,Mandatory = $true)]
    [ValidateScript({Test-DRAValidObjectNameParameter $_})]
    [String] $NewName
  )
  if ( $PSCmdlet.ShouldProcess("Rule '$Rule' in DRA AssistantAdmin '$AssistantAdmin'","Rename") ) {
    try {
      RenameDRAObjectRule "AssistantAdmin" $AssistantAdmin $Rule $NewName
    }
    catch {
      $PSCmdlet.WriteError($_)
    }
  }
}

#------------------------------------------------------------------------------
# CATEGORY: Set object description/comment
#------------------------------------------------------------------------------
function SetDRAObjectDetail {
  [CmdletBinding()]
  param(
    [Parameter(Mandatory = $true)]
    [ValidateSet("ActiveView","AssistantAdmin")]
    [String] $objectType,

    [Parameter(Mandatory = $true)]
    [ValidateNotNullOrEmpty()]
    [String] $objectName,

    [String] $comment,

    [String] $description
  )
  $argList = (Get-EAParamName $objectType),$objectName,"UPDATE"
  if ( $PSBoundParameters.ContainsKey("Comment") )     { $argList += "COMMENT:$comment" }
  if ( $PSBoundParameters.ContainsKey("Description") ) { $argList += "DESCRIPTION:$description" }
  $argList += "MODE:B"
  $output = $null
  $result = Invoke-EA $argList ([Ref] $output) -primary
  if ( $result -ne 0 ) {
    $PSCmdlet.ThrowTerminatingError((New-Object Management.Automation.ErrorRecord $output[0],
      (Get-Variable MyInvocation -Scope 1).Value.MyCommand.Name,
      ([Management.Automation.ErrorCategory]::NotSpecified),
      $objectName))
  }
}

# EXPORT
function Set-DRAActiveViewComment {
  <#
  .SYNOPSIS
  Sets the comment for one or more DRA ActiveView objects.

  .DESCRIPTION
  Sets the comment for one or more DRA ActiveView objects.

  .PARAMETER ActiveView
  Specifies the name of one or more DRA ActiveView objects. This parameter supports wildcards ('?' and '*').

  .PARAMETER Comment
  Specifies the comment. To remove the comment, specify an empty string as the argument to this parameter.

  .INPUTS
  * System.String
  * Objects with property 'ActiveView'

  .OUTPUTS
  None
  #>
  [CmdletBinding(SupportsShouldProcess = $true)]
  param(
    [Parameter(Position = 0,Mandatory = $true,ValueFromPipeline = $true,ValueFromPipelineByPropertyName = $true)]
    [ValidateScript({Test-DRAValidObjectNameParameter $_ -wildcard})]
    [SupportsWildcards()]
    [Alias("AV")]
    [String[]] $ActiveView,

    [Parameter(Position = 1,Mandatory = $true)]
    [AllowEmptyString()]
    [String] $Comment
  )
  process {
    foreach ( $activeViewItem in $ActiveView ) {
      if ( $PSCmdlet.ShouldProcess("DRA ActiveView '$activeViewItem'","Set comment") ) {
        try {
          SetDRAObjectDetail "ActiveView" $activeViewItem -comment $Comment
        }
        catch {
          $PSCmdlet.WriteError($_)
        }
      }
    }
  }
}

# EXPORT
function Set-DRAActiveViewDescription {
  <#
  .SYNOPSIS
  Sets the description for one or more DRA ActiveView objects.

  .DESCRIPTION
  Sets the description for one or more DRA ActiveView objects.

  .PARAMETER ActiveView
  Specifies the name of one or more DRA ActiveView objects. This parameter supports wildcards ('?' and '*').

  .PARAMETER Description
  Specifies the description. To remove the description, specify an empty string as the argument to this parameter.

  .INPUTS
  * System.String
  * Objects with property 'ActiveView'

  .OUTPUTS
  None
  #>
  [CmdletBinding(SupportsShouldProcess = $true)]
  param(
    [Parameter(Position = 0,Mandatory = $true,ValueFromPipeline = $true,ValueFromPipelineByPropertyName = $true)]
    [ValidateScript({Test-DRAValidObjectNameParameter $_ -wildcard})]
    [SupportsWildcards()]
    [Alias("AV")]
    [String[]] $ActiveView,

    [Parameter(Position = 1,Mandatory = $true)]
    [AllowEmptyString()]
    [String] $Description
  )
  process {
    foreach ( $activeViewItem in $ActiveView ) {
      if ( $PSCmdlet.ShouldProcess("DRA ActiveView '$activeViewItem'","Set description") ) {
        try {
          SetDRAObjectDetail "ActiveView" $activeViewItem -description $Description
        }
        catch {
          $PSCmdlet.WriteError($_)
        }
      }
    }
  }
}

# EXPORT
function Set-DRAAssistantAdminComment {
  <#
  .SYNOPSIS
  Sets the comment for one or more DRA AssistantAdmin objects.

  .DESCRIPTION
  Sets the comment for one or more DRA AssistantAdmin objects.

  .PARAMETER AssistantAdmin
  Specifies the name of one or more DRA AssistantAdmin objects. This parameter supports wildcards ('?' and '*').

  .PARAMETER Comment
  Specifies the comment. To remove the comment, specify an empty string as the argument to this parameter.

  .INPUTS
  * System.String
  * Objects with property 'AssistantAdmin'

  .OUTPUTS
  None
  #>
  [CmdletBinding(SupportsShouldProcess = $true)]
  param(
    [Parameter(Position = 0,Mandatory = $true,ValueFromPipeline = $true,ValueFromPipelineByPropertyName = $true)]
    [ValidateScript({Test-DRAValidObjectNameParameter $_ -wildcard})]
    [SupportsWildcards()]
    [Alias("AA")]
    [String[]] $AssistantAdmin,

    [Parameter(Position = 1,Mandatory = $true)]
    [AllowEmptyString()]
    [String] $Comment
  )
  process {
    foreach ( $assistantAdminItem in $AssistantAdmin ) {
      if ( $PSCmdlet.ShouldProcess("DRA AssistantAdmin '$assistantAdminItem'","Set comment") ) {
        try {
          SetDRAObjectDetail "AssistantAdmin" $assistantAdminItem -comment $Comment
        }
        catch {
          $PSCmdlet.WriteError($_)
        }
      }
    }
  }
}

# EXPORT
function Set-DRAAssistantAdminDescription {
  <#
  .SYNOPSIS
  Sets the description for one or more DRA AssistantAdmin objects.

  .DESCRIPTION
  Sets the description for one or more DRA AssistantAdmin objects.

  .PARAMETER AssistantAdmin
  Specifies the name of one or more DRA AssistantAdmin objects. This parameter supports wildcards ('?' and '*').

  .PARAMETER Description
  Specifies the description. To remove the description, specify an empty string as the argument to this parameter.

  .INPUTS
  * System.String
  * Objects with property 'AssistantAdmin'

  .OUTPUTS
  None
  #>
  [CmdletBinding(SupportsShouldProcess = $true)]
  param(
    [Parameter(Position = 0,Mandatory = $true,ValueFromPipeline = $true,ValueFromPipelineByPropertyName = $true)]
    [ValidateScript({Test-DRAValidObjectNameParameter $_ -wildcard})]
    [SupportsWildcards()]
    [Alias("AA")]
    [String[]] $AssistantAdmin,

    [Parameter(Position = 1,Mandatory = $true)]
    [AllowEmptyString()]
    [String] $Description
  )
  process {
    foreach ( $assistantAdminItem in $AssistantAdmin ) {
      if ( $PSCmdlet.ShouldProcess("DRA AssistantAdmin '$assistantAdminItem'","Set description") ) {
        try {
          SetDRAObjectDetail "AssistantAdmin" $assistantAdminItem -description $Description
        }
        catch {
          $PSCmdlet.WriteError($_)
        }
      }
    }
  }
}
#------------------------------------------------------------------------------

#------------------------------------------------------------------------------
# CATEGORY: Set object rule comment
#------------------------------------------------------------------------------
function SetDRAObjectRuleComment {
  [CmdletBinding()]
  param(
    [Parameter(Mandatory = $true)]
    [ValidateSet("ActiveView","AssistantAdmin")]
    [String] $objectType,

    [Parameter(Mandatory = $true)]
    [ValidateNotNullOrEmpty()]
    [String] $objectName,

    [Parameter(Mandatory = $true)]
    [ValidateNotNullOrEmpty()]
    [String] $ruleName,

    [String] $comment
  )
  $output = $null
  $result = Invoke-EA (Get-EAParamName $objectType),$objectName,"UPDATERULES",$ruleName,"COMMENT:$comment","MODE:B" ([Ref] $output) -primary
  if ( $result -ne 0 ) {
    $errorMsg = $output[0]
  }
  else {
    if ( $output[1] -match 'not found\.+$' ) {
      $result = 2
      $name = [Regex]::Match($output[0],'''(.+)''').Groups[1].Value
      $errorMsg = "{0} in DRA $objectType '$name'." -f ($output[1] -replace '^\s+','' -replace '\.+$','')
    }
  }
  if ( $result -ne 0 ) {
    (Get-Variable PSCmdlet -Scope 1).Value.ThrowTerminatingError((New-Object Management.Automation.ErrorRecord $errorMsg,
      (Get-Variable MyInvocation -Scope 1).Value.MyCommand.Name,
      ([Management.Automation.ErrorCategory]::NotSpecified),
      $objectName))
  }
}

# EXPORT
function Set-DRAActiveViewRuleComment {
  <#
  .SYNOPSIS
  Sets the comment for one or more DRA ActiveView object rules.

  .DESCRIPTION
  Sets the comment for one or more DRA ActiveView object rules.

  .PARAMETER ActiveView
  Specifies the name of one or more DRA ActiveView objects. This parameter supports wildcards ('?' and '*').

  .PARAMETER Rule
  Specifies the name of the rule. This parameter supports wildcards ('?' and '*').

  .PARAMETER Comment
  Specifies the comment for the rule. To remove the comment, specify an empty string as the argument to this parameter.

  .INPUTS
  * System.String
  * Objects with property 'Rule'
  * Objects with properties 'ActiveView' and 'Rule'

  .OUTPUTS
  None
  #>
  [CmdletBinding(SupportsShouldProcess = $true)]
  param(
    [Parameter(Position = 0,Mandatory = $true,ValueFromPipelineByPropertyName = $true)]
    [ValidateScript({Test-DRAValidObjectNameParameter $_ -wildcard})]
    [SupportsWildcards()]
    [Alias("AV")]
    [String] $ActiveView,

    [Parameter(Position = 1,Mandatory = $true,ValueFromPipeline = $true,ValueFromPipelineByPropertyName = $true)]
    [ValidateScript({Test-DRAValidObjectNameParameter $_ -wildcard})]
    [SupportsWildcards()]
    [String[]] $Rule,

    [Parameter(Position = 2,Mandatory = $true)]
    [AllowEmptyString()]
    [String] $Comment
  )
  process {
    foreach ( $ruleItem in $Rule ) {
      if ( $PSCmdlet.ShouldProcess("Rule '$ruleItem' in DRA ActiveView '$ActiveView'","Set comment") ) {
        try {
          SetDRAObjectRuleComment "ActiveView" $ActiveView $ruleItem -comment $Comment
        }
        catch {
          $PSCmdlet.WriteError($_)
        }
      }
    }
  }
}

# EXPORT
function Set-DRAAssistantAdminRuleComment {
  <#
  .SYNOPSIS
  Sets the comment for one or more DRA AssistantAdmin object rules.

  .DESCRIPTION
  Sets the comment for one or more DRA AssistantAdmin object rules.

  .PARAMETER AssistantAdmin
  Specifies the name of one or more DRA AssistantAdmin objects. This parameter supports wildcards ('?' and '*').

  .PARAMETER Rule
  Specifies the name of the rule. This parameter supports wildcards ('?' and '*').

  .PARAMETER Comment
  Specifies the comment for the rule. To remove the comment, specify an empty string as the argument to this parameter.

  .INPUTS
  * System.String
  * Objects with property 'Rule'
  * Objects with properties 'AssistantAdmin' and 'Rule'

  .OUTPUTS
  None
  #>
  [CmdletBinding(SupportsShouldProcess = $true)]
  param(
    [Parameter(Position = 0,Mandatory = $true,ValueFromPipelineByPropertyName = $true)]
    [ValidateScript({Test-DRAValidObjectNameParameter $_ -wildcard})]
    [SupportsWildcards()]
    [Alias("AA")]
    [String] $AssistantAdmin,

    [Parameter(Position = 1,Mandatory = $true,ValueFromPipeline = $true,ValueFromPipelineByPropertyName = $true)]
    [ValidateScript({Test-DRAValidObjectNameParameter $_ -wildcard})]
    [SupportsWildcards()]
    [String[]] $Rule,

    [Parameter(Position = 2,Mandatory = $true)]
    [AllowEmptyString()]
    [String] $Comment
  )
  process {
    foreach ( $ruleItem in $Rule ) {
      if ( $PSCmdlet.ShouldProcess("Rule '$ruleItem' in DRA AssistantAdmin '$AssistantAdmin'","Set comment") ) {
        try {
          SetDRAObjectRuleComment "AssistantAdmin" $AssistantAdmin $ruleItem -comment $Comment
        }
        catch {
          $PSCmdlet.WriteError($_)
        }
      }
    }
  }
}
