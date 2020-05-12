/*
WSManWinRM.js
       Purpose: A simple POC for leveraging a WMI class to execute a remote command over the WinRM(WSMan) protocol using the WSMan.Automation COM object (wsmauto.dll)
       Author: @bohops
       License: BSD 3-Clause

       Usage: cscript.exe WSManWinRM.js [hostname] [command]
       Usage: cscript.exe WSManWinRM.js [hostname] [command] [domain\user] [password]

       Example: cscript.exe WSManWinRM.js host.domain.local notepad.exe
       Example: cscript.exe WSManWinRM.js host.domain.local "cmd /c notepad.exe" domain\joe.user P@ssw0rd

       Link: https://docs.microsoft.com/en-us/windows/win32/winrm/wsman
*/

var args = WScript.Arguments;
var protocol = "http";
var port = "5985";

var wsman = new ActiveXObject("Wsman.Automation");
var options = wsman.CreateConnectionOptions();
var sessionUrl = protocol + "://" + args(0) + ":" + port + "/wsman";

var session = wsman.CreateSession(sessionUrl, 0, options);
if (args.Count() == 4) 
{
    options.UserName = args(2);
    options.Password = args(3);
    session = wsman.CreateSession(sessionUrl, wsman.SessionFlagCredUsernamePassword(), options);
}

var resource = "http://schemas.microsoft.com/wbem/wsman/1/wmi/root/cimv2/Win32_Process";
var parameters = "<p:Create_INPUT xmlns:p=\"http://schemas.microsoft.com/wbem/wsman/1/wmi/root/cimv2/Win32_Process\"><p:CommandLine>" + args(1) + "</p:CommandLine></p:Create_INPUT>";
var response = session.invoke("Create", resource, parameters);
WScript.Echo(response);