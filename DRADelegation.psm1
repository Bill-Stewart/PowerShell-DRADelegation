# DRADelegation.psm1
# Written by Bill Stewart for IHS
#
# Provides a PowerShell interface for getting and setting DRA delegation
# details (ActiveViews, AssistantAdmins, and Roles). The module is a wrapper for
# the EA.exe command-line interface.
#
# For output operations, the module captures the output of the EA.exe command,
# parses it, and outputs as PowerShell objects that can be filtered, sorted,
# etc. Pipeline operations are supported and should work as expected.
#
# The module uses Write-Progress to show the EA.exe command line in cases of
# slow operations. The EA.exe command line is also shown when debug is enabled.
#
# Prerequisite:
# * DRA client installation including the "Command-line interface" feature
#
# Current limitations:
#
# * Module is limited to what's available in EA.exe
# * Module speed is limited by performance of EA.exe
# * All actions connect to the primary server, so recommendation is only to use
#   this module on primary server or server in same site as primary with a fast
#   network connection
# * ActiveView resource rules are not supported
#
# Version history:
#
# 1.0 (2019-09-17)
# * Initial version.

#requires -version 3

$DRAInstallDir = Get-ItemProperty "HKLM:\SOFTWARE\WOW6432Node\Mission Critical Software\OnePoint\Administration" -ErrorAction SilentlyContinue
if ( -not $DRAInstallDir ) {
  $DRAInstallDir = Get-ItemProperty "HKLM:\SOFTWARE\Mission Critical Software\OnePoint\Administration" -ErrorAction SilentlyContinue
}
$DRAInstallDir = $DRAInstallDir.InstallDir
if ( -not $DRAInstallDir ) {
  throw "Unable to find DRA client installation on this computer."
}

$EA = Join-Path $DRAInstallDir "EA.exe"
if ( -not (Test-Path $EA) ) {
  throw "Unable to find '$EA'. This module requires the DRA command-line interface to be installed."
}

# Always use primary server?
$FORCE_PRIMARY = $true

#------------------------------------------------------------------------------
# CATEGORY: Miscellaneous
#------------------------------------------------------------------------------
function Out-Object {
  param(
    [Collections.Hashtable[]] $hashData
  )
  $order = @()
  $result = @{}
  $hashData | ForEach-Object {
    $order += ($_.Keys -as [Array])[0]
    $result += $_
  }
  New-Object PSObject -Property $result | Select-Object $order
}

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
# CATEGORY: Executable processing
#------------------------------------------------------------------------------

# Invoke-Executable uses this function to convert stdout and stderr to
# arrays; trailing newlines are ignored
function ConvertTo-Array {
  param(
    [String] $stringToConvert
  )
  $stringWithoutTrailingNewLine = $stringToConvert -replace '(\r\n)*$', ''
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
    [ValidateSet("ActiveView","AssistantAdmin","Role")]
    [String] $objectType,

    [Parameter(Mandatory = $true)]
    [ValidateNotNullOrEmpty()]
    [String] $objectName
  )
  $output = $null
  $result = Invoke-EA "/DELI:TAB",(Get-EAParamName $objectType),$objectName,"DISPLAY","ALL" ([Ref] $output)
  if ( $result -eq 0 ) {
    switch ( $objectType ) {
      "ActiveView" {
        $output | Select-String '^(.+)\tComment:"(.*)"\tDescription:"(.*)"\tType:"(.*)"' | ForEach-Object {
          Out-Object `
            @{"ActiveView"  = $_.Matches[0].Groups[1].Value},
            @{"Comment"     = $_.Matches[0].Groups[2].Value},
            @{"Description" = $_.Matches[0].Groups[3].Value},
            @{"Type"        = $_.Matches[0].Groups[4].Value}
        }
      }
      "AssistantAdmin" {
        $output | Select-String '^(.+)\tComment:"(.*)"\tDescription:"(.*)"\tType:"(.*)"\tAssigned:"(.*)"' | ForEach-Object {
          Out-Object `
            @{"AssistantAdmin" = $_.Matches[0].Groups[1].Value},
            @{"Comment"        = $_.Matches[0].Groups[2].Value},
            @{"Description"    = $_.Matches[0].Groups[3].Value},
            @{"Type"           = $_.Matches[0].Groups[4].Value},
            @{"Assigned"       = $_.Matches[0].Groups[5].Value}
        }
      }
      "Role" {
        $output | Select-String '^(.+)\tComment:"(.*)"\tDescription:"(.*)"\tType:"(.*)"\tAssigned:"(.*)"' | ForEach-Object {
          Out-Object `
            @{"Role"        = $_.Matches[0].Groups[1].Value},
            @{"Comment"     = $_.Matches[0].Groups[2].Value},
            @{"Description" = $_.Matches[0].Groups[3].Value},
            @{"Type"        = $_.Matches[0].Groups[4].Value},
            @{"Assigned"    = $_.Matches[0].Groups[5].Value}
        }
      }
    }
  }
  else {
    $OFS = [Environment]::NewLine
    $PSCmdlet.ThrowTerminatingError((New-Object Management.Automation.ErrorRecord "$output",
      (Get-Variable MyInvocation -Scope 1).Value.MyCommand.Name,
      ([Management.Automation.ErrorCategory]::ObjectNotFound),
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
  System.String

  .OUTPUTS
  Objects with the following properties:
    ActiveView   The name of the ActiveView object
    Comment      The object's comment
    Description  The object's description
    Type         The object's type (Built-in or Custom)
  #>
  param(
    [Parameter(ValueFromPipeline = $true)]
    [ValidateScript({Test-DRAValidObjectNameParameter $_ -wildcard})]
    [SupportsWildcards()]
    [String[]] $ActiveView = "*"
  )
  process {
    foreach ( $activeViewItem in $ActiveView ) {
      try {
        GetDRAObject "ActiveView" $activeViewItem
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
  System.String

  .OUTPUTS
  Objects with the following properties:
    AssistantAdmin  The name of the AssistantAdmin object
    Comment         The object's comment
    Description     The object's description
    Type            The object's type (Built-in or Custom)
  #>
  param(
    [Parameter(ValueFromPipeline = $true)]
    [ValidateScript({Test-DRAValidObjectNameParameter $_ -wildcard})]
    [SupportsWildcards()]
    [String[]] $AssistantAdmin = "*"
  )
  process {
    foreach ( $assistantAdminItem in $AssistantAdmin ) {
      try {
        GetDRAObject "AssistantAdmin" $assistantAdminItem
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
  System.String

  .OUTPUTS
  Objects with the following properties:
    Role         The name of the Role object
    Comment      The object's comment
    Description  The object's description
    Type         The object's type (Built-in or Custom)
  #>
  param(
    [Parameter(ValueFromPipeline = $true)]
    [ValidateScript({Test-DRAValidObjectNameParameter $_ -wildcard})]
    [SupportsWildcards()]
    [String[]] $Role = "*"
  )
  process {
    foreach ( $roleItem in $Role ) {
      try {
        GetDRAObject "Role" $roleItem
      }
      catch {
        $PSCmdlet.WriteError($_)
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

    [String] $ruleName
  )
  $output = $null
  $result = Invoke-EA "/DELI:TAB",(Get-EAParamName $objectType),$objectName,"DISPLAYRULES",$ruleName,"ALL" ([Ref] $output)
  if ( $result -eq 0 ) {
    for ( $i = 0; $i -lt $output.Count; $i++ ) {
      # Lines without leading whitespace give object type and name
      if ( $output[$i] -match '^[^\s]' ) {
        $output[$i] | Select-String '^(.+) ''(.+)''\.\.\.' | ForEach-Object {
          # Construct output object
          $outObj = Out-Object `
            @{$_.Matches[0].Groups[1].Value = $_.Matches[0].Groups[2].Value},
            @{"Rule"                        = $null},
            @{"Description"                 = $null},
            @{"Comment"                     = $null}
        }
      }
      else {
        if ( $output[$i] -notmatch 'not found\.$' ) {
          # Update output object properties
          $output[$i] | Select-String '^  (.+)\tdescription:"(.*)"\tcomment:"(.*)"' | ForEach-Object {
            $outObj.Rule        = $_.Matches[0].Groups[1].Value
            $outObj.Description = $_.Matches[0].Groups[2].Value
            $outObj.Comment     = $_.Matches[0].Groups[3].Value
          }
          $outObj
        }
        else {
          $name = [Regex]::Match($output[$i - 1],'''(.+)''').Groups[1].Value
          (Get-Variable PSCmdlet -Scope 1).Value.WriteError((New-Object Management.Automation.ErrorRecord "Rule(s) matching '$ruleName' not found in DRA $objectType '$name'.",
            (Get-Variable MyInvocation -Scope 1).Value.MyCommand.Name,
            ([Management.Automation.ErrorCategory]::ObjectNotFound),
            $objectName))
        }
      }
    }
  }
  else {
    $OFS = [Environment]::NewLine
    $PSCmdlet.ThrowTerminatingError((New-Object Management.Automation.ErrorRecord "$output",
      (Get-Variable MyInvocation -Scope 1).Value.MyCommand.Name,
      ([Management.Automation.ErrorCategory]::ObjectNotFound),
      $objectName))
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
    [Parameter(Position = 0,Mandatory = $true,ValueFromPipeline = $true,ValueFromPipelineByPropertyName = $true)]
    [ValidateScript({Test-DRAValidObjectNameParameter $_ -wildcard})]
    [SupportsWildcards()]
    [String[]] $ActiveView,

    [Parameter(Position = 1)]
    [ValidateScript({Test-DRAValidObjectNameParameter $_ -wildcard})]
    [SupportsWildcards()]
    [String] $Rule = "*"
  )
  process {
    foreach ( $activeViewItem in $ActiveView ) {
      try {
        GetDRAObjectRule "ActiveView" $activeViewItem $Rule
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
    [Parameter(Position = 0,Mandatory = $true,ValueFromPipeline = $true,ValueFromPipelineByPropertyName = $true)]
    [ValidateScript({Test-DRAValidObjectNameParameter $_ -wildcard})]
    [SupportsWildcards()]
    [String[]] $AssistantAdmin,

    [Parameter(Position = 1)]
    [ValidateScript({Test-DRAValidObjectNameParameter $_ -wildcard})]
    [SupportsWildcards()]
    [String] $Rule = "*"
  )
  process {
    foreach ( $assistantAdminItem in $AssistantAdmin ) {
      try {
        GetDRAObjectRule "AssistantAdmin" $assistantAdminItem $Rule
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
    [String] $AssistantAdmin,

    [Parameter(Position = 1,Mandatory = $true)]
    [ValidateScript({Test-DRAValidObjectNameParameter $_})]
    [String] $Role,

    [Parameter(Position = 2,Mandatory = $true)]
    [ValidateScript({Test-DRAValidObjectNameParameter $_})]
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
    [String] $AssistantAdmin,

    [Parameter(Position = 1,Mandatory = $true,ValueFromPipelineByPropertyName = $true)]
    [ValidateScript({Test-DRAValidObjectNameParameter $_})]
    [String] $Role,

    [Parameter(Position = 2,Mandatory = $true,ValueFromPipelineByPropertyName = $true)]
    [ValidateScript({Test-DRAValidObjectNameParameter $_})]
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
    $PSCmdlet.ThrowTerminatingError((New-Object Management.Automation.ErrorRecord $errorMsg,
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
#------------------------------------------------------------------------------
