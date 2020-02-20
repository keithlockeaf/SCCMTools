<#
.Synopsis
  Creates and updates a SCCM specific computer PSObject for use with the SCCMTools module.

.Description
  Creates and updates a SCCM specific computer PSObject for use with the SCCMTools module.
    Computer object is created via "New-CustomCimSessionPublic"
    Determines if the computer has the SCCM client installed
    Test for any pending reboot flags via "Test-RebootPendingViaCimSessionPublic"

.Parameter ComputerName
    The name of the computer(s). The local computer is the default.

.Parameter Credential
    Specifies a user account that has permission to perform this action.
    The default is the current user.

.Example
  New-SCCMComputerObjectPublic

.EXAMPLE
  New-SCCMComputerObjectPublic -ComputerName Server01, Server02

.EXAMPLE
  New-SCCMComputerObjectPublic -ComputerName Server01, Server02 -Credential (Get-Credential)

#>

#Requires -RunAsAdministrator

Function New-SCCMComputerObjectPublic {
  [CmdletBinding(
    SupportsShouldProcess = $true
  )]
  param(
    [Parameter(
      Mandatory = $false,
      ValueFromPipelineByPropertyName = $true,
      HelpMessage = 'Name of the computer(s) you want generate the SCCM computer object for.  Default: Local Host')]
    [ValidateNotNullOrEmpty()]
    [string[]]
    $ComputerName = $env:COMPUTERNAME,

    [Parameter(
      Mandatory = $false,
      ValueFromPipelineByPropertyName = $true,
      HelpMessage = "The credentials needed to connect to the computer(s)")]
    [ValidateNotNullOrEmpty()]
    [System.Management.Automation.Credential()]
    [System.Management.Automation.PSCredential]
    $Credential = [System.Management.Automation.PSCredential]::Empty
  )

  begin {
    Write-Verbose "New-SCCMComputerObjectPublic: Started"
    Set-StrictMode -Version 1.0 #Option Explicit
    $ErrorActionPreference = 'Stop'
  }

  process {
    # The default MaxEnvelopeSizeKb is 500, this is not big enough for the data that is pulled from using CIM
    Set-Item -Path WSMan:\localhost\MaxEnvelopeSizeKb -Value 2048

    # Creates the Custom CIM Session object
    if ($PSBoundParameters.ContainsKey('Credential')) {
      $SCCMComputerObject = New-CustomCimSessionPublic -ComputerName $ComputerName -Credential $Credential
    }
    else {
      $SCCMComputerObject = New-CustomCimSessionPublic -ComputerName $ComputerName
    }

    # For each computer that had a cim session successfully created
    foreach ($Computer in $SCCMComputerObject) {
      if ($Computer.CimSessionConnected -eq $true) {
        # Add the following keys to the computer object
        $Computer | Add-Member -MemberType NoteProperty -Name 'SCCMInstalled' -Value $false
        $Computer | Add-Member -MemberType NoteProperty -Name 'RebootPending' -Value $null

        # Determine if the computer has a SCCM Client installed
        if ((Get-CimInstance -CimSession $($Computer.CimSession) -ClassName Win32_Product -Verbose:$false).Name -Contains "Configuration Manager Client") {
          $Computer.SCCMInstalled = $true
        }

        # Check for reboot pending flags
        $Computer.RebootPending = (Test-RebootPendingViaCimSessionPublic -CimSession $Computer.CimSession)
      }
    }
  }

  end {
    Write-Verbose "New-SCCMComputerObjectPublic: Completed`n"
    return $SCCMComputerObject
  }
}
