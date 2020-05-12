# WSMan-WinRM
A collection of proof-of-concept source code and scripts for executing remote commands over WinRM using the WSMan.Automation COM object. 

## Background
For background information, please refer to the following blog post: [WS-Management COM: Another Approach for WinRM Lateral Movement](https://bohops.com/2020/05/12/ws-management-com-another-approach-for-winrm-lateral-movement/)

## Notes
- SharpWSManWinRM.cs and CppWsManWinRM.cpp compile in Visual Studio 2019.  Refer to the code comments for required imports/references/etc.
- All examples leverage the WMI Win32_Process class and WMI Create method for invocation. 

## Usage

### SharpWSManWinRM.cs
```
 Usage: SharpWSManWinRM.exe <hostname> <command>
 Usage: SharpWSManWinRM.exe <hostname> <command> <domain\user> <password>

 Example: SharpWSManWinRM.exe host.domain.local notepad.exe
 Example: SharpWSManWinRM.exe host.domain.local "cmd /c notepad.exe" domain\joe.user P@ssw0rd
```
### WSManWinRM.ps1
```
 Usage: Invoke-WSManWinRM -hostname <hostname> -command <command>
 Usage: Invoke-WSManWinRM -hostname <hostname> -command <command> -user <domain\user> -password <password>

 Example: import-module .\WSManWinRM.ps1
          Invoke-WSManWinRM -hostname MyServer.domain.local -command calc.exe
 Example: import-module .\WSManWinRM.ps1
          Invoke-WSManWinRM -hostname MyServer.domain.local -command calc.exe -user domain\joe.user -password P@ssw0rd
```		  
		  
### WSManWinRM.vbs
```
 Usage: cscript.exe SharpWSManWinRM.vbs <hostname> <command>
 Usage: cscript.exe SharpWSManWinRM.vbs <hostname> <command> <domain\user> <password>

 Example: cscript.exe SharpWSManWinRM.vbs host.domain.local notepad.exe
 Example: cscript.exe SharpWSManWinRM.vbs host.domain.local "cmd /c notepad.exe" domain\joe.user P@ssw0rd	
```
### WSManWinRM.js
```
 Usage: cscript.exe SharpWSManWinRM.js <hostname> <command>
 Usage: cscript.exe SharpWSManWinRM.js <hostname> <command> <domain\user> <password>

 Example: cscript.exe SharpWSManWinRM.js host.domain.local notepad.exe
 Example: cscript.exe SharpWSManWinRM.js host.domain.local "cmd /c notepad.exe" domain\joe.user P@ssw0rd	 
```
### CppWSManWinRM.cpp
```
 Usage: CppWSManWinRM.exe <hostname> <command>

 Example: CppWSManWinRM.exe host.domain.local notepad.exe
 
 Note: Username/password option does not work yet
 ```
 
## Ethics

WSMan-WinRM is designed to help security professionals perform ethical and legal security assessments and penetration tests. Do not use for nefarious purposes.

