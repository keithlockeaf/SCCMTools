<#
.SYNOPSIS
    Creates CimSessions to computer(s) using the Dcom protocol.

.DESCRIPTION
    New-CustomCimSessionPublic is a function that is designed to create CimSessions to one or more
    computers, using the Dcom protocol. PowerShell version 3 is required on the
    computer that this function is being run on, but PowerShell does not need to be
    installed at all on the computer.

.Parameter ComputerName
    The name of the computer(s). The local computer is the default.

.Parameter Credential
    Specifies a user account that has permission to perform this action.
    The default is the current user.

.EXAMPLE
     New-CustomCimSessionPublic -ComputerName Server01, Server02

.EXAMPLE
     New-CustomCimSessionPublic -ComputerName Server01, Server02 -Credential (Get-Credential)

.EXAMPLE
     Get-Content -Path C:\Servers.txt | New-CustomCimSessionPublic

#>

#Requires -RunAsAdministrator

Function New-CustomCimSessionPublic {
  [CmdletBinding(
    SupportsShouldProcess = $true
  )]
  param(
    [Parameter(
      Position = 0,
      Mandatory = $false,
      ValueFromPipeline = $true,
      ValueFromPipelineByPropertyName = $true,
      HelpMessage = "Computer(s) you want to create a CimSession to")]
    [ValidateNotNullOrEmpty()]
    [string[]]
    $ComputerName = $env:COMPUTERNAME,

    [Parameter(
      Position = 1,
      Mandatory = $false,
      ValueFromPipeline = $true,
      ValueFromPipelineByPropertyName = $true,
      HelpMessage = "The credential (Get-Credential) you want to use to connect to your remote servers")]
    [ValidateNotNullOrEmpty()]
    [System.Management.Automation.Credential()]
    [System.Management.Automation.PSCredential]
    $Credential = [System.Management.Automation.PSCredential]::Empty
  )

  begin {
    Write-Verbose "New-CustomCimSessionPublic: Started"
    Set-StrictMode -Version 1.0 #Option Explicit
    $ErrorActionPreference = 'Stop'

    # The default MaxEnvelopeSizeKb is 500, this is not big enough for the data that is pulled in from using CIM
    Set-Item -Path WSMan:\localhost\MaxEnvelopeSizeKb -Value 2048
  }

  process {
    # Create an array to store all the computer objects created
    [System.Collections.ArrayList] $AllComputerObjects = @()

    # Create the DCOM protocol for use with older systems
    $Opt = New-CimSessionOption -Protocol DCOM
    # Parameters to be used when connecting to the
    $SessionParameters = @{
      ErrorAction   = 'Stop'
      SessionOption = $Opt
    }

    # Write out all the CimSession Parameters using Verbose
    Write-Verbose -Message "CimSession Parameters:"
    foreach ($Parameter in $SessionParameters.Keys) {
      Write-Verbose -Message "  $Parameter $($SessionParameters[$Parameter])"
    }

    # If the computer is not the local computer, use the supplied credentials or get credentials
    if ($ComputerName -ne $env:COMPUTERNAME) {
      if ($PSBoundParameters['Credential']) {
        $SessionParameters.Credential = $Credential
      }
      else {
        $SessionParameters.Credential = Get-Credential
      }
    }

    # Create a new computer object for each computer
    foreach ($Computer in $ComputerName) {
      # If the computer name contains any domain information, strip off everything but the computer name
      if ($Computer.contains('.')) {
        $Computer = $Computer.Split('.')[0]
      }
      $Computer = ($Computer).ToLower()

      # Create a new computer object
      $NewComputerObject = [PSObject] @{ }
      $NewComputerObject | Add-Member -MemberType NoteProperty -Name 'ComputerName' -Value $Computer
      $NewComputerObject | Add-Member -MemberType NoteProperty -Name 'FQDN' -Value $null
      $NewComputerObject | Add-Member -MemberType NoteProperty -Name 'Credential' -Value $SessionParameters.Credential
      $NewComputerObject | Add-Member -MemberType NoteProperty -Name 'CimSessionConnected' -Value $false

      try {
        $NewComputerObject.FQDN = (([System.Net.Dns]::GetHostByName("$Computer")).Hostname).ToLower()
      }
      catch { }

      if ($null -eq $NewComputerObject.FQDN) {
        try {
          $NewComputerObject.FQDN = (([System.Net.Dns]::GetHostByName($Computer)).Hostname).ToLower()
        }
        catch {
          $NewComputerObject.FQDN = 'Could not determine FQDN. Please verify system name'
        }
      }

      # Try to initiate a CimSession to the computer
      if ($NewComputerObject.FQDN -ne 'Could not determine FQDN. Please verify system name') {
        try {
          $SessionParameters.ComputerName = $NewComputerObject.FQDN
          $NewComputerObject | Add-Member -MemberType NoteProperty -Name 'CimSession' -Value (New-CimSession @SessionParameters -Verbose:$false)
        }
        catch { }
      }

      # If the CimSession was able to connect successfully
      if ($null -ne $NewComputerObject.CimSession) {
        $NewComputerObject.CimSessionConnected = $true
      }

      # Add the new computer object to the computer objects array
      $AllComputerObjects += $NewComputerObject
      Write-Verbose -Message "New object created:  "
      Write-Verbose -Message "  ComputerName:         $($NewComputerObject.ComputerName)"
      Write-Verbose -Message "  FQDN:                 $($NewComputerObject.FQDN)"
      Write-Verbose -Message "  CimSession:           $($NewComputerObject.CimSession)"
      Write-Verbose -Message "  CimSessionConnected:  $($NewComputerObject.CimSessionConnected)"
    }

    Write-Verbose -Message "Total objects created:  $($AllComputerObjects.Count)"
  }

  end {
    Write-Verbose "New-CustomCimSessionPublic: Completed"
    return , $AllComputerObjects
  }
}
