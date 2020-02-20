<#
.Synopsis
  Returns an PSObject containing SCCM patching data

.Description
  Returns an PSObject containing SCCM patching data.  To view
  said object you will need to add the following to the command
  Command: | select-object *
  Example:  Get-SCCMPatchStatusPublic | select-object *

.Parameter SCCMComputerObject
  A PSObject created via a previous command that called this command

.Parameter ComputerName
  The name of the computer(s) you want to get patch data from.
  The local computer is the default.

.Parameter Credential
  Specifies a user account that has permission to perform this action.
  The current users credentials is the default.

.Example
  Get-SCCMPatchStatusPublic

.Example
  Get-SCCMPatchStatusPublic -SCCMComputerObject Object

.Example
  Get-SCCMPatchStatusPublic -ComputerName 'Name of a computer'

.Example
  Get-SCCMPatchStatusPublic -Credential 'A credential object'

.Example
  Get-SCCMPatchStatusPublic -ComputerName 'Name of a computer' -Credential 'A credential object'

#>

#Requires -RunAsAdministrator

Function Get-SCCMPatchStatusPublic {
  [CmdletBinding(DefaultParameterSetName = 'User Input')]
  param(
    [Parameter(
      ParameterSetName = 'Custom PSObject',
      Mandatory = $false,
      ValueFromPipelineByPropertyName = $true,
      HelpMessage = 'Custom PSObject created by a previous SCCMTools function')]
    [ValidateNotNullOrEmpty()]
    [PSObject]
    $SCCMComputerObject = $null,

    [Parameter(
      ParameterSetName = 'User Input',
      Position = 0,
      Mandatory = $false,
      ValueFromPipeline = $true,
      ValueFromPipelineByPropertyName = $true,
      HelpMessage = "Computer(s) you want to create a CimSession to.  Default is local computer")]
    [ValidateNotNullOrEmpty()]
    [string[]]
    $ComputerName = $env:COMPUTERNAME,

    [Parameter(
      ParameterSetName = 'User Input',
      Position = 1,
      Mandatory = $false,
      ValueFromPipelineByPropertyName = $true,
      HelpMessage = "The credentials needed to connect to remote computer(s)")]
    [ValidateNotNullOrEmpty()]
    [System.Management.Automation.Credential()]
    [System.Management.Automation.PSCredential]
    $Credential = [System.Management.Automation.PSCredential]::Empty
  )

  begin {
    Write-Verbose "Get-SCCMPatchStatusPublic: Started"
    Set-StrictMode -Version 1.0 #Option Explicit
    $ErrorActionPreference = 'Stop'

    # The default MaxEnvelopeSizeKb is 500, this is not big enough for the data that is pulled from using CIM
    Set-Item -Path WSMan:\localhost\MaxEnvelopeSizeKb -Value 2048 -Force

    # If a SCCMComputerObject was not given, create it
    if ($null -eq $SCCMComputerObject) {
      if ($PSBoundParameters.ContainsKey('Credential')) {
        Write-Verbose -Message 'Creating SCCMComputerObject with a specific credential'
        [PSObject]$SCCMComputerObject = New-SCCMComputerObjectPublic -ComputerName $ComputerName -Credential $Credential
      }
      else {
        Write-Verbose -Message 'Creating SCCMComputerObject with current user credentials'
        [PSObject]$SCCMComputerObject = New-SCCMComputerObjectPublic -ComputerName $ComputerName
      }
    }
  }

  process {
    foreach ($Computer in $SCCMComputerObject) {
      if ($Computer.CimSessionConnected -eq $true -and $Computer.SCCMInstalled -eq $true) {
        # Creates the key 'AllPatchDeploymentData' which is an empty array
        $Computer | Add-Member -MemberType NoteProperty -Name 'AllPatchDeploymentData' -Value @()

        # Create the Patches Missing CIM parameters
        $CimInstanceParameters = @{
          Query       = "SELECT * FROM CCM_SoftwareUpdate WHERE ComplianceState = '0'"
          NameSpace   = 'ROOT\ccm\ClientSDK'
          CimSession  = $Computer.CimSession
          ErrorAction = 'SilentlyContinue'
        }

        # Get all patches currently pending install
        $PatchesMissing = @(Get-CimInstance @CimInstanceParameters) | Sort-Object Name

        if ($null -eq $PatchesMissing) {
          $PatchesMissing = 0
        }
        $Computer | Add-Member -MemberType NoteProperty -Name 'PatchesMissing' -Value $PatchesMissing

        # Create the Deployments CIM parameters
        $CimInstanceParameters = @{
          Query       = 'Select * FROM CCM_UpdateCIAssignment'
          NameSpace   = 'ROOT\ccm\Policy\Machine\RequestedConfig'
          CimSession  = $Computer.CimSession
          ErrorAction = 'SilentlyContinue'
        }

        # Gets any patch deployments currently deployed to this server
        $AllDeploymentsData = Get-CimInstance @CimInstanceParameters | Sort-Object AssignmentName

        # Builds the missing patch object if it exists so long as $AllDeploymentsData does not equal $null
        foreach ($Deployment in $AllDeploymentsData) {
          $Computer.AllPatchDeploymentData += $Deployment.AssignmentName
          $PatchesToConvert = $Deployment.AssignedCIs

          [System.Collections.ArrayList] $AllDeployedSCCMPatches = @()
          foreach ($Patch in $PatchesToConvert) {
            # Create custom PS object with empty variables
            $PatchObj = [PSCustomObject]@{
              "Article"                = "";
              "Id"                     = "";
              "ModelName"              = "";
              "Version"                = "";
              "CIVersion"              = "";
              "ApplicabilityCondition" = "";
              "EnforcementEnabled"     = "";
              "DisplayName"            = "";
              "UpdateClassification"   = ""
            }

            # Assign Patch object variables
            [xml]$Ux = $Patch
            $PatchObj.Article = $Ux.ci.DisplayName | Select-String -Pattern 'KB\d*' -AllMatches | ForEach-Object { $_.Matches } | ForEach-Object { $_.value }
            $PatchObj.id = $ux.ci.id
            $PatchObj.ModelName = $Ux.ci.ModelName
            $PatchObj.Version = $ux.ci.version
            $PatchObj.CIVersion = $ux.ci.CIVersion
            $PatchObj.ApplicabilityCondition = $ux.ci.ApplicabilityCondition
            $PatchObj.EnforcementEnabled = $ux.ci.EnforcementEnabled
            $PatchObj.DisplayName = $Ux.ci.DisplayName
            $PatchObj.UpdateClassification = $ux.ci.UpdateClassification

            # Some objects in $AssignedCIs are null.  If it isn't, add the new patch object to the array.
            if ($null -ne $PatchObj) {
              $AllDeployedSCCMPatches += $PatchObj
            }
          }

          foreach ($Patch in $PatchesMissing) {
            # Determine and list which patch deployment the patch is being deployed from
            if ($AllDeployedSCCMPatches.ModelName -contains $Patch.UpdateID) {
              $NoteExists = $false
              $NoteExists = [bool]($Patch.PSObject.Properties | Where-Object { $_.Name -eq "PatchBeingDeployedFrom" } -ErrorAction SilentlyContinue)
              # Is this the first time the patch is found in a deployment to this system
              if ($NoteExists -eq $false) {
                $Patch | Add-Member -MemberType NoteProperty -Name PatchBeingDeployedFrom -Value @()
              }
              # Add one or more deployments that this patch is being deployed from
              $Patch.PatchBeingDeployedFrom += $Deployment.AssignmentName
            }
          }
        }

        if ($Computer.PatchesMissing -eq 0) {
          $Computer | Add-Member -MemberType NoteProperty -Name 'PatchesAvailableCount' -Value 0
        }
        else {
          $Computer | Add-Member -MemberType NoteProperty -Name 'PatchesAvailableCount' -Value @($Computer.PatchesMissing).count
        }
      }
    }
  }

  end {
    Write-Verbose "Get-SCCMPatchStatusPublic: Completed"
    return $SCCMComputerObject
  }
}
