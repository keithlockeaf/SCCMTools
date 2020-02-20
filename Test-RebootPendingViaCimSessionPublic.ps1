<#
.Synopsis
  What your functions end goal is

.Description
  Everything your function does

.Parameter Parameter1
  The reason for this parameter

.Parameter Parameter2
  The reason for this parameter

.Example
  Test-RebootPendingViaCimSessionPublic

.Example
  Test-RebootPendingViaCimSessionPublic -Parameter1 ParameterValue

.Example
  Test-RebootPendingViaCimSessionPublic -Parameter1 ParameterValue -Parameter2 Parameter2Value

#>

#Requires -RunAsAdministrator

Function Test-RebootPendingViaCimSessionPublic {
  [CmdletBinding()]
  param(
    [Parameter(
      Mandatory = $false,
      ValueFromPipelineByPropertyName = $true,
      HelpMessage = "Cim Session of the remote or local server you want to check the reboot status of.  Default:  Local Host")]
    [ValidateNotNullOrEmpty()]
    [CimSession]
    $CimSession
  )

  begin {
    Write-Verbose "Test-RebootPendingViaCimSessionPublic: Started"
    Set-StrictMode -Version 1.0 #Option Explicit
    $ErrorActionPreference = 'Stop'

    # The default MaxEnvelopeSizeKb is 500, this is not big enough for the data that is pulled in from using CIM
    Set-Item -Path WSMan:\localhost\MaxEnvelopeSizeKb -Value 2048
  }

  process {
    # Assumption is that no reboot is pending for the server
    $RebootPending = $false

    # Arguments designed to allow CimMethod to search the registry
    $HKLM = [uINT32]2147483650
    $RegistryArguments = @{
      hDefKey     = $HKLM
      sSubKeyName = "SYSTEM\CurrentControlSet\Control\Session Manager"
      sValueName  = "PendingFileRenameOperations"
    }

    # CimMethod parameters for easy reading
    $CimMethodParameters = @{
      Namespace   = 'ROOT\CIMv2'
      ClassName   = 'StdRegProv'
      MethodName  = 'GetMultiStringValue'
      Arguments   = $RegistryArguments
      ErrorAction = 'SilentlyContinue'
      Verbose     = $false
    }

    # Used if there is a specific session that should be used (Remote systems)
    if ($PSBoundParameters.ContainsKey('CimSession')) {
      $CimMethodParameters.Add('CimSession', $CimSession)
    }

    # Check for file name change reboot pending flag
    $RebootPending = $RebootPending -or ((Invoke-CimMethod @CimMethodParameters).sValue.length -ne 0)

    # CimMethod parameters for easy reading
    $CimMethodParameters = @{
      Namespace   = 'ROOT\ccm\ClientSDK'
      ClassName   = 'CCM_ClientUtilities'
      MethodName  = 'DetermineIfRebootPending'
      Verbose     = $false
      ErrorAction = 'SilentlyContinue'
    }

    # Used if there is a specific session that should be used (Remote systems)
    if ($PSBoundParameters.ContainsKey('CimSession')) {
      $CimMethodParameters.Add('CimSession', $CimSession)
    }

    # Check the SCCM Utilities for pending reboot (Soft and Hard)
    $RebootPending = $RebootPending -or (Invoke-CimMethod @CimMethodParameters).RebootPending
    $RebootPending = $RebootPending -or (Invoke-CimMethod @CimMethodParameters).IsHardRebootPending

    # CimInstance parameters for easy reading
    $CimInstanceParameters = @{
      Namespace   = 'ROOT\ccm\ClientSDK'
      Query       = 'SELECT * FROM CCM_SoftwareUpdate'
      Verbose     = $false
      ErrorAction = 'SilentlyContinue'
    }

    # Used if there is a specific session that should be used (Remote systems)
    if ($PSBoundParameters.ContainsKey('CimSession')) {
      $CimInstanceParameters.Add('CimSession', $CimSession)
    }

    # Check the SCCM installed patches for pending reboot or reboot required before patch install
    $RebootPending = $RebootPending -or ((@(Get-CimInstance @CimInstanceParameters) | Where-Object {
          $_.EvaluationState -eq 8 -or # patch pending soft reboot
          $_.EvaluationState -eq 9 -or # patch pending hard reboot
          $_.EvaluationState -eq 10 } # reboot needed before installing patch (the prior prerequisite patch not yet installed)
      ).length -ne 0
    )
  }

  end {
    Write-Verbose "Test-RebootPendingViaCimSessionPublic: Completed"
    return $RebootPending
  }
}
