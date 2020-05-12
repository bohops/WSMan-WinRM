function Invoke-WSManWinRM
{
<#
    .SYNOPSIS
        Purpose: A simple POC for leveraging a WMI class to execute a remote command over the WinRM (WSMan) protocol using the WSMan.Automation COM object (wsmauto.dll)
        Inspiration: Invoke-WSManAction
        Author: @bohops
        License: BSD 3-Clause
    .PARAMETER hostname
        The hostname (or FQDN) of the remote WinRM host. Required.
    .PARAMETER command
        The command to execute remotely. Required.
    .PARAMETER user
        Domain\Username (credential). Optional.
    .PARAMETER password
        Password (credential). Optional. 
    .EXAMPLE
        PS C:\> Invoke-WSManWinRM -hostname MyServer.domain.local -command calc.exe
        Returns XML Blog with PID if successful
    .EXAMPLE
        PS C:\> Invoke-WSManWinRM -hostname MyServer.domain.local -command calc.exe -user domain\joe.user -password P@ssw0rd
        Returns XML Blog with PID if successful
    .LINK
        https://docs.microsoft.com/en-us/windows/win32/winrm/wsman
#>
    Param
    (
         [Parameter(Mandatory=$true, Position=0)]
         [string] $hostname,
         [Parameter(Mandatory=$true, Position=1)]
         [string] $command,
         [Parameter(Mandatory=$false, Position=2)]
         [string] $user,
         [Parameter(Mandatory=$false, Position=3)]
         [string] $password
    )

    $protocol = "http"
    $port = "5985"

    $wsman = new-object -com WSMan.Automation
    $options = $wsman.CreateConnectionOptions()
    $sessionUrl = $protocol + "://" + $hostname + ":" + $port + "/wsman"

    $session = $wsman.CreateSession($sessionUrl, 0, $options)
    if (($user.Length -gt 0) -and ($password.Length -gt 0))
    {
        $options.Username = $user
        $options.Password = $password
        $session = $wsman.CreateSession($sessionUrl, $wsman.SessionFlagCredUsernamePassword(), $options)
    }

    
    $resource = "http://schemas.microsoft.com/wbem/wsman/1/wmi/root/cimv2/Win32_Process"
    $parameters = "<p:Create_INPUT xmlns:p=`"http://schemas.microsoft.com/wbem/wsman/1/wmi/root/cimv2/Win32_Process`"><p:CommandLine>" + $command + "</p:CommandLine></p:Create_INPUT>"
    $response = $session.Invoke("Create", $resource, $parameters)
    write-host $response
}
