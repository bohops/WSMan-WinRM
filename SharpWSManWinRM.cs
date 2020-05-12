/*
    SharpWSManWinRM.cs

       Purpose: A simple POC for leveraging a WMI class to execute a remote command over the WinRM (WSMan) protocol using the WSMan.Automation COM object (wsmauto.dll)
       Author: @bohops
       License: BSD 3-Clause

       Usage: SharpWSManWinRM.exe [hostname] [command]
       Usage: SharpWSManWinRM.exe [hostname] [command] [domain\user] [password]

       Example: SharpWSManWinRM.exe host.domain.local notepad.exe
       Example: SharpWSManWinRM.exe host.domain.local "cmd /c notepad.exe" domain\joe.user P@ssw0rd

       Link: https://docs.microsoft.com/en-us/windows/win32/winrm/wsman
*/

using System;
using System.Runtime.InteropServices;
using System.Xml;
using WSManAutomation;  //Add Reference -> windows\system32\wsmauto.dll (or COM: Microsoft WSMan Automation V 1.0 Library)

namespace SharpWinRM
{
    //Globals vars to store protocol/port info for now before testing out TLS/SSL
    public class Globals
    {
        public const string protocol = "http";
        public const string port = "5985";
    }
    class Program
    {
        static void Main(string[] args)
        {
            if (args.Length == 2)
                ExecRemoteCmd(args[0], args[1], "", "");
            else if (args.Length == 4)
                ExecRemoteCmd(args[0], args[1], args[2], args[3]);
            else
            {
                Console.WriteLine("Usage:   SharpWSManWinRM.exe [hostname] [command]\n" +
                                  "Example: SharpWSManWinRM.exe host.domain.local \"cmd /c notepad.exe\"");
                Console.WriteLine("\n    -or-\n");
                Console.WriteLine("Usage:   SharpWSManWinRM.exe [hostname] [command] [domain\\user] [password]\n" +
                                  "Example: SharpWSManWinRM.exe host.domain.local \"cmd /c notepad.exe\" domain\\joe.user P@ssw0rd");
            }
        }

        static void ExecRemoteCmd(string host, string cmd, string user, string password)
        {
            try
            {
                //Setup session & execute remote command
                IWSManEx wsman = new WSMan();
                IWSManConnectionOptions options = (IWSManConnectionOptions)wsman.CreateConnectionOptions();
                string sessionUrl = Globals.protocol + "://" + host + ":" + Globals.port + "/wsman";

                IWSManSession session = (IWSManSession)wsman.CreateSession(sessionUrl, 0, options);
                if ((user.Length > 0) && (password.Length > 0))
                {
                    options.UserName = user;
                    options.Password = password;
                    session = (IWSManSession)wsman.CreateSession(sessionUrl, wsman.SessionFlagCredUsernamePassword(), options);
                }

                string resource = "http://schemas.microsoft.com/wbem/wsman/1/wmi/root/cimv2/Win32_Process";
                string parameters = "<p:Create_INPUT xmlns:p=\"http://schemas.microsoft.com/wbem/wsman/1/wmi/root/cimv2/Win32_Process\"><p:CommandLine>" + cmd + "</p:CommandLine></p:Create_INPUT>";
                string response = session.Invoke("Create", resource, parameters);

                //(Poorly) Parse XML response and print values
                XmlDocument xml = new XmlDocument();
                xml.LoadXml(response);
                Console.WriteLine("xsi         : " + xml.DocumentElement.GetAttribute("xmlns:xsi"));
                Console.WriteLine("p           : " + xml.DocumentElement.GetAttribute("xmlns:p"));
                Console.WriteLine("cim         : " + xml.DocumentElement.GetAttribute("xmlns:cim"));
                Console.WriteLine("lang        : " + xml.DocumentElement.GetAttribute("xml:lang"));
                Console.WriteLine("ProcessId   : " + xml.DocumentElement.ChildNodes[0].InnerText);
                Console.WriteLine("ReturnValue : " + xml.DocumentElement.ChildNodes[1].InnerText);

                //Cleanup COM Object
                Marshal.ReleaseComObject(session);
                Marshal.ReleaseComObject(options);
            }
            catch (Exception e)
            {
                Console.WriteLine(e.Message.ToString());
            }
        }
    }
}