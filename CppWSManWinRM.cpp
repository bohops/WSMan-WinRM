/*
	CppWSManWinRM.cpp

	   Purpose: A simple POC for leveraging a WMI class to execute a remote command over the WinRM (WSMan) protocol using the WSMan.Automation COM object (wsmauto.dll)
	   Author: @bohops
	   License: BSD 3-Clause

	   Usage: CppWSManWinRM.exe [hostname] [command]
	   //Usage: CppWSManWinRM.exe [hostname] [command] [domain\user] [password] <-- Does not work...yet

	   Example: CppWSManWinRM.exe host.domain.local notepad.exe
	   //Example: CppWSManWinRM.exe host.domain.local "cmd /c notepad.exe" domain\joe.user P@ssw0rd <-- Does not work...yet

	   Link: http://forums.codeguru.com/showthread.php?499875-WSManAutomation-lib-is-missing (Primary Code Source -> credit: ssoma from CodeGuru forums)
	   Link: https://docs.microsoft.com/en-us/windows/win32/winrm/wsman
*/

#include <comdef.h>
#include <windows.h>
#include <wsmandisp.h>
#include "wsmandisp_i.c"
#include <iostream>

using namespace std;

int ExecRemoteCmd(string host, string cmd, string user, string passwd);

int main(int argc, char** argv)
{
	if (argc == 3) //Args: 2 (+ 1)
		int res = ExecRemoteCmd(argv[1], argv[2], "", "");
	else if (argc == 5) //Args: 4 (+1)
		cout << ":-|"; //int res = ExecRemoteCmd(argv[1], argv[2], argv[3], argv[4]); <--- not implemented....yet
	else
	{
		cout << "Usage:   CppWSManWinRM.exe [hostname] [command]" << endl;
		cout << "Example: CppWSManWinRM.exe host.domain.local \"cmd /c notepad.exe\"" << endl;
		//cout << "  -or-" << endl;
		//cout << "Usage:   CppWSManWinRM.exe [hostname] [command] [domain\\user] [password]" << endl;
		//cout << "Example: CppWSManWinRM.exe host.domain.local \"cmd /c notepad.exe\" domain\\joe.user P@ssw0rd" << endl;
	}
}
int ExecRemoteCmd(string host, string cmd, string user, string passwd)
{
	_bstr_t protocol = "http";
	_bstr_t port = "5985";

	_bstr_t hostname = (_bstr_t)host.c_str();
	_bstr_t command = (_bstr_t)cmd.c_str();
	//_bstr_t username = (_bstr_t)user.c_str();
	//_bstr_t password = (_bstr_t)passwd.c_str();

	HRESULT hres = NULL;
	HRESULT hr = S_OK;
	IDispatch* pSvc = NULL;
	IDispatch* pOptions = NULL;
	IWSManSession* pSvc1 = NULL;

	// Initialize COM
	hres = CoInitializeEx(0, COINIT_MULTITHREADED);
	if (FAILED(hres))
	{
		cout << "Failed to initialize COM library. Error code = 0x" << hex << hres << endl;
		return 1; // Program has failed.
	}

	//Instatliate WSMAN COM Object
	IWSMan* pLoc = NULL;
	hres = CoCreateInstance(CLSID_WSMan, 0, CLSCTX_INPROC_SERVER, IID_IWSMan, (LPVOID*)&pLoc);
	if (FAILED(hres))
	{
		cout << "Failed to create IWSMan object." << " Err code = 0x" << hex << hres << endl;
		CoUninitialize();
		return 1; // Program has failed.
	}

	//Create Remote Session
	cout << "[*] Creating session with the remote system..." << endl;
	pSvc = reinterpret_cast<IDispatch*>(pSvc1);
	long f = 0;

	// This needs to be fixed for supplied credentials.  Pull requests welcome...
	//if ("creds")
	//{
	//	//HRESULT SessionFlagUseNoAuthentication(f);
	//	HRESULT SessionFlagCredUsernamePassword(f);
	//	IWSManConnectionOptions* options = NULL;
	//	options->put_UserName(bstr_t("corp\\corpmin"));
	//	options->put_Password(bstr_t("CorpM@ster"));
	//	pOptions = reinterpret_cast<IDispatch*>(options);
	//	hres = pLoc->CreateConnectionOptions(&pOptions);
	//	hres = pLoc->CreateSession(_bstr_t("http://corp-dc:5985/wsman"), f, pOptions, &pSvc);
	//}
	//else
	//{
	//	hres = pLoc->CreateSession(_bstr_t("http://corp-dc:5985/wsman"), 0, NULL, &pSvc);
	//}

	hres = pLoc->CreateSession(protocol + _bstr_t("://") + hostname + _bstr_t(":") + port + _bstr_t("/wsman"), 0, NULL, &pSvc);
	if (FAILED(hres))
	{
		cout << "Could not connect. Error code = 0x" << hex << hres << endl;
		pLoc->Release();
		CoUninitialize();
		return 1; // Program has failed.
	}

	//Invoke Command
	cout << "[*] Connected to the remote WinRM system" << endl;
	pSvc1 = reinterpret_cast<IWSManSession*>(pSvc);
	_variant_t resource = "http://schemas.microsoft.com/wbem/wsman/1/wmi/root/cimv2/Win32_Process";
	_bstr_t parameters = "<p:Create_INPUT xmlns:p=\"http://schemas.microsoft.com/wbem/wsman/1/wmi/root/cimv2/Win32_Process\"><p:CommandLine>" + command + "</p:CommandLine></p:Create_INPUT>";
	wchar_t* response = new wchar_t(1000);
	hres = pSvc1->Invoke(_bstr_t("Create"), resource, parameters, 0, &response);

	//Print Response 
	cout << "[*] Result Code: " << response;

	//Cleanup
	pSvc->Release();
	pSvc1->Release();
	pOptions->Release();
	pLoc->Release();
	CoUninitialize();

	return 0;
}