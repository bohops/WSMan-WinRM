'WSManWinRM.vbs
'       Purpose: A simple POC for leveraging a WMI class to execute a remote command over the WinRM(WSMan) protocol using the WSMan.Automation COM object (wsmauto.dll)
'       Inspiration: Invoke-WSManAction (PowerShell)
'       Author: @bohops
'       License: BSD 3-Clause
'
'       Usage: cscript.exe WSManWinRM.vbs [hostname] [command]
'       Usage: cscript.exe WSManWinRM.vbs [hostname] [command] [domain\user] [password]
'
'       Example: cscript.exe WSManWinRM.vbs host.domain.local notepad.exe
'       Example: cscript.exe WSManWinRM.vbs host.domain.local "cmd /c notepad.exe" domain\joe.user P@ssw0rd
'
'       Link: https://docs.microsoft.com/en-us/windows/win32/winrm/wsman

dim arg, protocol, port, wsman, session, sessionUrl, resource, options, parameters, response

set args = wscript.arguments
protocol = "http"
port = "5985"

set wsman = CreateObject("Wsman.Automation")
set options = wsman.CreateConnectionOptions
sessionUrl = protocol + "://" + args(0) + ":" + port + "/wsman"

set session = wsman.CreateSession(sessionUrl, 0, options)
if args.count = 4 then
    options.UserName = args(2)
    options.Password = args(3)
    set session = wsman.CreateSession(sessionUrl, wsman.SessionFlagCredUsernamePassword, options)
end if


resource = "http://schemas.microsoft.com/wbem/wsman/1/wmi/root/cimv2/Win32_Process"
parameters = "<p:Create_INPUT xmlns:p=""http://schemas.microsoft.com/wbem/wsman/1/wmi/root/cimv2/Win32_Process""><p:CommandLine>" + args(1) + "</p:CommandLine></p:Create_INPUT>"
response = session.invoke("Create", resource, parameters)
wscript.echo response