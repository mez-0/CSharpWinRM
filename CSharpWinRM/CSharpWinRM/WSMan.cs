using System;
using System.ComponentModel;
using System.Linq.Expressions;
using System.Runtime.InteropServices;
using System.Text;
using System.Xml;
using WSManAutomation;

namespace CSharpWinRM
{
    class WSMan
    {
        public static Boolean Execute(string[] args)
        {
            Boolean Encryption = false;

            String Target = "";
            String Domain = "";
            String Username = "";
            String Password = "";
            String Command = "";

            if (args.Length == 2)
            {
                Target = args[0];
                Command = args[1];
            }

            else if (args.Length == 5)
            {
                try
                {
                    Target = args[0];
                    Domain = args[1];
                    Username = args[2];
                    Password = args[3];
                    Command = args[4];
                }
                catch (Exception e)
                {
                    Logger.Print(Logger.STATUS.ERROR, e.Message);
                    return false;
                }
            }
            else
            {
                Help();
                return false;
            }

            String hostUri = String.Format("http://{0}:5985", Target);
            String sessionURI = String.Format("http://{0}/wsman", Target);

            WSManClass wsmanClass = new WSManClass();

            IWSManConnectionOptions connectionOptions = (IWSManConnectionOptions)wsmanClass.CreateConnectionOptions();
            if(Domain != "" && Username != "" && Password != "")
            {
                connectionOptions.UserName = Domain + "\\" + Username;
                connectionOptions.Password = Password;
                Logger.Print(Logger.STATUS.INFO, "Authenticating with: " + Domain + "\\"+ Username + ":" + Password);

            }
            else if (Domain == "" && Username == "" && Password == "")
            {
                Logger.Print(Logger.STATUS.INFO, "Authenticating as: " + Environment.UserDomainName + "\\" + Environment.UserName);
            }
            else
            {
                Help();
                return false;
            }

            Logger.Print(Logger.STATUS.INFO, "Command: " + Command);

            int wsmanFlags = wsmanClass.SessionFlagUTF8() | wsmanClass.SessionFlagCredUsernamePassword();

            if (!Encryption)
            {
                // https://stackoverflow.com/questions/1469791/powershell-v2-remoting-how-do-you-enable-unencrypted-traffic
                wsmanClass.SessionFlagNoEncryption();
            }

            try
            {
                // https://docs.microsoft.com/en-us/windows/win32/api/wsmandisp/nn-wsmandisp-iwsmansession
                IWSManSession wsmanSession = (IWSManSession)wsmanClass.CreateSession(hostUri, wsmanFlags, connectionOptions);

                if (wsmanSession != null)
                {
                    // https://docs.microsoft.com/en-us/windows/win32/winrm/windows-remote-management-and-wmi
                    // https://csharp.hotexamples.com/examples/-/IWSManSession/-/php-iwsmansession-class-examples.html
                    
                    XmlDocument xmlIdentifyResponse = new XmlDocument();

                    // https://docs.microsoft.com/en-us/windows/win32/api/wsmandisp/nf-wsmandisp-iwsmansession-identify
                    // Queries a remote computer to determine if it supports the WS-Management protocol.
                    try
                    {
                         xmlIdentifyResponse.LoadXml(wsmanSession.Identify());
                        if (!xmlIdentifyResponse.HasChildNodes)
                        {
                            Logger.Print(Logger.STATUS.INFO, "Failed to Identify() host");
                            Marshal.ReleaseComObject(wsmanSession);
                            return false;
                        }
                        Logger.Print(Logger.STATUS.GOOD, "Successfully identified host: ");
                        foreach(XmlNode node in xmlIdentifyResponse.ChildNodes)
                        {
                            foreach(XmlNode innernodes in node.ChildNodes)
                            {
                                Logger.Print(Logger.STATUS.GOOD, innernodes.InnerText);
                            }
                        }
                    }
                    catch (Exception e)
                    {
                        Logger.Print(Logger.STATUS.ERROR, "Failed to Identify() host");
                        Logger.Print(Logger.STATUS.ERROR, e.Message);
                        Marshal.ReleaseComObject(wsmanSession);
                        return false;
                    }

                    string resourceURI = "http://schemas.microsoft.com/wbem/wsman/1/wmi/root/cimv2/Win32_Process";

                    // https://social.microsoft.com/Forums/en-US/d6ec5087-33dc-4967-8183-f8524683a3ea/using-remote-powershellwinrm-within-caspnet
                    StringBuilder parameters = new StringBuilder();
                    parameters.Append("<p:Create_INPUT ");
                    parameters.Append("xmlns:p=\"http://schemas.microsoft.com/wbem/wsman/1/wmi/root/cimv2/Win32_Process\">");
                    parameters.Append("<p:CommandLine>" + Command + "</p:CommandLine>");
                    parameters.Append("</p:Create_INPUT>");

                    Logger.Print(Logger.STATUS.INFO, "Sending the following XML: ");
                    Console.WriteLine("\n" + parameters + "\n");

                    String responseFromInvoke = wsmanSession.Invoke("Create", resourceURI, parameters.ToString(), 0);

                    if (responseFromInvoke != null)
                    {
                        Logger.Print(Logger.STATUS.GOOD, "Got a response from invoke:");
                        Console.WriteLine("\n" + responseFromInvoke + "\n");
                        XmlDocument xmlInvokeResponse = new XmlDocument();
                        try
                        {
                            xmlInvokeResponse.LoadXml(responseFromInvoke);
                            foreach (XmlNode node in xmlInvokeResponse.ChildNodes)
                            {
                                foreach (XmlNode innernodes in node.ChildNodes)
                                {
                                    String NodeName = innernodes.Name.Replace("p:", "");
                                    String NodeValue = innernodes.InnerText;
                                    Logger.Print(Logger.STATUS.GOOD, NodeName + ": " + NodeValue);
                                    if(NodeName == "ReturnValue")
                                    {
                                        if(NodeValue == "0")
                                        {
                                            Logger.Print(Logger.STATUS.GOOD, "Process Sucessfully Started!");
                                        }
                                        else
                                        {
                                            Logger.Print(Logger.STATUS.ERROR, "Process failed to start, got error: " + NodeValue);
                                        }
                                    }
                                }
                            }
                        }
                        catch(Exception e)
                        {
                            Logger.Print(Logger.STATUS.ERROR, "Got an erorr whilst parsing response: " + e.Message);
                            Marshal.ReleaseComObject(wsmanSession);
                        }
                        
                    }
                    else
                    {
                        Logger.Print(Logger.STATUS.ERROR, "Got no response, not too sure what this means...");
                        Marshal.ReleaseComObject(wsmanSession);
                        return false;
                    }

                    Marshal.ReleaseComObject(wsmanSession);
                }
                else
                {
                    Logger.Print(Logger.STATUS.ERROR, "Failed to create session with IWSManSession");
                    Marshal.ReleaseComObject(wsmanSession);
                    ErrorMsg();
                }
            }
            finally
            {
                Marshal.ReleaseComObject(connectionOptions);
            }
            return true;
        }
        private static Int32 ErrorMsg()
        {
            Win32Exception errorMessage = new Win32Exception(Marshal.GetLastWin32Error());
            Logger.Print(Logger.STATUS.ERROR, String.Format("{0} (Error Code: {1})", errorMessage.Message, errorMessage.NativeErrorCode.ToString()));
            return (Int32)errorMessage.NativeErrorCode;
        }
        private static void Help()
        {
            Console.WriteLine("[*] Usage: .\\CSharpWinRM.exe <Target> [Domain] [Username] [Password] <Command>");
            Console.WriteLine("[*] Example 1: .\\CSharpWinRM.exe 192.168.0.1 DomainName Administrator Password123! \"powershell.exe -e blah\"");
            Console.WriteLine("[*] Example 2: .\\CSharpWinRM.exe 192.168.0.1 \"powershell.exe -e blah\"");
            return
        }
    }
}
