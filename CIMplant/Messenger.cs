﻿using System;
using System.Collections.Generic;

namespace CIMplant
{
    public class Messenger
    {
        public static void ErrorMessage(string error)
        {
            // Make errors pop
            if (Console.BackgroundColor == ConsoleColor.Black)
            {
                Console.ForegroundColor = ConsoleColor.Red;
                Console.WriteLine(error);
                Console.ResetColor();
            }
            else
                Console.WriteLine(error);
        }

        public static void WarningMessage(string output)
        {
            // Make cool messages pop
            if (Console.BackgroundColor == ConsoleColor.Black)
            {
                Console.ForegroundColor = ConsoleColor.Yellow;
                Console.WriteLine(output);
                Console.ResetColor();
            }
            else
                Console.WriteLine(output);
        }

        public static void BlueMessage(string output)
        {
            // Make cool messages pop
            if (Console.BackgroundColor == ConsoleColor.Black)
            {
                Console.ForegroundColor = ConsoleColor.Cyan;
                Console.WriteLine(output);
                Console.ResetColor();
            }
            else
                Console.WriteLine(output);
        }

        public static void GoodMessage(string output)
        {
            // Make cool messages pop
            if (Console.BackgroundColor == ConsoleColor.Black)
            {
                Console.ForegroundColor = ConsoleColor.Green;
                Console.WriteLine(output);
                Console.ResetColor();
            }
            else
                Console.WriteLine(output);
        }

        public static void ImportantAsciiArt()
        {
            const string art = @"
           _____ _____ __  __       _             _   
          / ____|_   _|  \/  |     | |           | |  
         | |      | | | \  / |_ __ | | __ _ _ __ | |_ 
         | |      | | | |\/| | '_ \| |/ _` | '_ \| __|
         | |____ _| |_| |  | | |_) | | (_| | | | | |_ 
          \_____|_____|_|  |_| .__/|_|\__,_|_| |_|\__|
                             | |                      
          by @Matt_Grandy_   |_|  (@FortyNorthSec)                    
            ";
            Console.WriteLine(art);
        }

        public static void GetCommands()
        {
            //Creates the command dictionary - {header:commands}
            Dictionary<string, string[]> commandDict = new Dictionary<string, string[]>
            {
                {"File Operations", new string[] { "cat - Reads the contents of a file",
                        "copy - Copies a file from one location to another",
                        "download** - Download a file from the targeted machine",
                        "ls - File/Directory listing of a specific directory",
                        "search - Search for a file on a user-specified drive",
                        "upload** - Upload a file to the targeted machine"
                    }
                },
                {"Lateral Movement Facilitation", new string[] {"command_exec** - Run a command line command and receive the output. Run with nops flag to disable PowerShell",
                        "disable_wdigest - Sets the registry value for UseLogonCredential to zero",
                        "enable_wdigest - Adds registry value UseLogonCredential",
                        "disable_winrm** - Disables WinRM on the targeted system",
                        "enable_winrm** - Enables WinRM on the targeted system",
                        "reg_mod - Modify the registry on the targeted machine",
                        "reg_create - Create the registry value on the targeted machine",
                        "reg_delete - Delete the registry on the targeted machine",
                        "remote_posh** - Run a PowerShell script on a remote machine and receive the output",
                        "sched_job - Not implimented due to the Win32_ScheduledJobs accessing an outdated API",
                        "service_mod - Create, delete, or modify system services"
                    }
                },
                {"Process Operations", new string[] { "process_kill - Kill a process via name or process id on the targeted machine",
                        "process_start - Start a process on the targeted machine",
                        "ps - Process listing"
                    }
                },
                {"System Operations", new string[]
                    {
                        "active_users - List domain users with active processes on the targeted system",
                        "basic_info - Used to enumerate basic metadata about the targeted system",
                        "drive_list  - List local and network drives",
                        "ifconfig - Receive IP info from NICs with active network connections",
                        "installed_programs - Receive a list of the installed programs on the targeted machine",
                        "logoff - Log users off the targeted machine",
                        "reboot (or restart) - Reboot the targeted machine",
                        "power_off (or shutdown) - Power off the targeted machine",
                        "vacant_system - Determine if a user is away from the system",
                         "edr_query - Query the local or remote system for EDR vendors"

                    }
                },
                {"Log Operations", new string[] {"logon_events - Identify users that have logged onto a system"}}
            };

            //Print out all commands
            if (Console.BackgroundColor == ConsoleColor.Black)
            {
                bool printDescription = true;
                // Iterate through headers and print out all the available commands
                foreach (KeyValuePair<string, string[]> entry in commandDict)
                {
                    Console.ForegroundColor = ConsoleColor.Red;
                    if (printDescription)
                    {
                        Console.WriteLine("\n{0,-30}{1,-50}", entry.Key, "Description");
                        Console.WriteLine("{0,-30}{1,-50}", "---------------------", "----------------------");

                        printDescription = false;
                    }
                    else
                    {
                        Console.WriteLine("\n{0,-30}", entry.Key);
                        Console.WriteLine("{0,-30}", "-----------------------------");
                    }

                    Console.ResetColor();
                    foreach (string command in entry.Value)
                    {
                        string[] commandSplit = command.Split('-');
                        Console.WriteLine("{0,-30}{1,-50}", commandSplit[0], commandSplit[1].TrimStart());
                    }
                }

                Console.WriteLine("\n* PowerShell can be disabled by using the --nops flag");
                Console.WriteLine("** Denotes PowerShell usage (either using a PowerShell Runspace or through Win32_Process::Create method)");
            }

            else
            {
                bool printDescription = true;
                foreach (KeyValuePair<string, string[]> entry in commandDict)
                {
                    if (printDescription)
                    {
                        Console.WriteLine("\n{0,-30}{1,-50}", entry.Key, "Description");
                        Console.WriteLine("{0,-30}{1,-50}", "---------------------", "----------------------");

                        printDescription = false;
                    }
                    else
                    {
                        Console.WriteLine("\n{0,-30}", entry.Key);
                        Console.WriteLine("{0,-30}", "-----------------------------");
                    }

                    Console.ResetColor();
                    foreach (string command in entry.Value)
                    {
                        string[] commandSplit = command.Split('-');
                        Console.WriteLine("{0,-30}{1,-50}", commandSplit[0], commandSplit[1].TrimStart());
                    }
                }

                Console.WriteLine("\n* PowerShell can be disabled by using the --nops flag");
                Console.WriteLine("** Denotes PowerShell usage (either using a PowerShell Runspace or Win32_Process::Create method)");
            }

            Environment.Exit(0);
        }

        public static void GetExamples()
        {
            //Creates the command dictionary - {header:commands}
            Dictionary<string, string[]> commandDict = new Dictionary<string, string[]>
            {
                {"File Operations", new string[] { @"cat , -c cat -f [file] -s [ip] -d [domain] -u [username] -p [password] , -c cat -f c:\users\test\desktop\test.txt -s 192.168.64.4 -u test -p 1",
                        @"copy , -c copy -f [file] --fileto [new file] , -c copy -f c:\users\test\desktop\test.txt --fileto c:\users\test\desktop\testnew.txt -s 192.168.64.4 -u test -p 1",
                        @"download , -c download -f [file] (optional: --fileto [file]) , -c download -f c:\users\test\desktop\test.exe -s 192.168.64.4 -u test -p 1",
                        @"ls , -c ls --directory [directory] , -c ls --directory c:\users\test\desktop\ -s 192.168.64.4 -u test -p 1",
                        @"search , -c search --directory [directory] -f [search term] (search term can contain wildcards ex: *pass*) , -c search --directory c:\users\test\desktop\ -f *test* -s 192.168.64.4 -u test -p 1",
                        @"upload , -c upload -f [file to upload] --fileto [upload location and filename] (ex: --fileto c:\test.exe) , -c upload -f file.exe --fileto c:\users\test\desktop\file.exe -s 192.168.64.4 -u test -p 1"
                    }
                },
                {"Lateral Movement Facilitation", new string[] { @"command_exec , -c command_exec --execute [sub command ex: ps; tasklist; whoami] , -c command_exec --execute ps -s 192.168.64.4 -u test -p 1",
                        @"disable_wdigest , -c disable_wdigest , -c disable_wdigest -s 192.168.64.4 -u test -p 1",
                        @"enable_wdigest , -c enable_wdigest , -c enable_wdigest -s 192.168.64.4 -u test -p 1",
                        @"disable_winrm , -c disable_winrm , -c disable_winrm -s 192.168.64.4 -u test -p 1",
                        @"enable_winrm , -c enable_winrm , -c enable_winrm -s 192.168.64.4 -u test -p 1",
                        @"reg_mod , -c reg_mod --regkey [full path to regkey] --regsubkey [regsubkey val] --regval [data value] , -c reg_mod -s 192.168.64.4 -u test -p 1 --regkey HKLM\SYSTEM\CurrentControlSet\Control\SecurityProviders\WDigest --regsubkey test --regval test",
                        @"reg_create , -c reg_create --regkey [full path to regkey] --regsubkey [regsubkey val] --regval [data val] --regvaltype [data type ex: reg_sz] , -c reg_create -s 192.168.64.4 -u test -p 1 --regkey HKLM\SYSTEM\CurrentControlSet\Control\SecurityProviders\WDigest --regsubkey test --regval test --regvaltype reg_sz",
                        @"reg_delete , -c reg_delete --regkey [full path to regkey] --regsubkey [regsubkey val] , -c reg_delete -s 192.168.64.4 -u test -p 1 --regkey HKLM\SYSTEM\CurrentControlSet\Control\SecurityProviders\WDigest --regsubkey test --regval test",
                        @"remote_posh , -c remote_posh -f [local PS file] --cmdlet [cmdlet] , -c remote_posh -s 192.168.64.4 -u test -p 1 1 -f Invoke-Mimi.ps1 --cmdlet invoke-mimikittenz",
                        @"sched_job , Not implimented due to the Win32_ScheduledJobs accessing an outdated API , ",
                        @"service_mod , -c service_mod --execute list , -c service_mod --execute list -s 192.168.64.4 -u test -p 1",
                        @"service_mod , -c service_mod --execute create --servicebin [path to service exe] --service [service] , -c service_mod --execute create -s 192.168.64.4 -u test -p 1 --servicebin c:\users\test\desktop\temp.exe --service fortynorth",
                        @"service_mod , -c service_mod --execute start --service [service] , -c service_mod --execute create -s 192.168.64.4 -u test -p 1 --service fortynorth",
                        @"service_mod , -c service_mod --execute delete --service [service] , -c service_mod --execute delete -s 192.168.64.4 -u test -p 1 --service fortynorth",
                    }
                },
                {"Process Operations", new string[] { @"process_kill , -c process_kill --process [process (accepts wildcards)] , -c process_kill --process power* -s 192.168.64.4 -u test -p 1",
                        @"process_start , -c process_start --process [process (use quotes if spaces)] , -c process_start --process C:\users\test\desktop\test.exe -s 192.168.64.4 -u test -p 1",
                        @"ps , -c ps , -c ps -s 192.168.64.4 -u test -p 1"
                    }
                },
                {"System Operations", new string[]
                    {
                        @"active_users , -c active_users , -c active_users -s 192.168.64.4 -u test -p 1",
                        @"basic_info , -c basic_info , -c basic_info -s 192.168.64.4 -u test -p 1",
                        @"drive_list  , -c drive_list , -c drive_list -s 192.168.64.4 -u test -p 1",
                        @"ifconfig , -c ifconfig , -c ifconfig -s 192.168.64.4 -u test -p 1",
                        @"installed_programs , -c installed_programs , -c installed_programs -s 192.168.64.4 -u test -p 1",
                        @"logoff (or logout) , -c logoff , -c logoff -s 192.168.64.4 -u test -p 1",
                        @"reboot (or restart) , -c reboot , -c reboot -s 192.168.64.4 -u test -p 1",
                        @"power_off (or shutdown) , -c power_off , -c shutdown -s 192.168.64.4 -u test -p 1",
                        @"vacant_system , -c vacant_system , -c vacant_system -s 192.168.64.4 -u test -p 1",
                        @"edr_query , -c edr_query, -c edr_query -s 192.168.64.4 -u test -p 1"
                    }
                },
                {"Log Operations", new string[] { @"logon_events , -c logon_events , -c logon_events -s 192.168.64.4 -u test -p 1" }}
            };

            //Print out all commands
            if (Console.BackgroundColor == ConsoleColor.Black)
            {
                // Iterate through headers and print out all the available commands
                bool printDescription = true;
                // Iterate through headers and print out all the available commands
                foreach (KeyValuePair<string, string[]> entry in commandDict)
                {
                    Console.ForegroundColor = ConsoleColor.Red;
                    if (printDescription)
                    {
                        Console.WriteLine("\n{0,-30}{1,-50}", entry.Key, "Example Command");
                        Console.WriteLine("{0,-30}{1,-50}", "---------------------", "----------------------");

                        printDescription = false;
                    }
                    else
                    {
                        Console.WriteLine("\n{0,-30}", entry.Key);
                        Console.WriteLine("{0,-30}", "-----------------------------");
                    }

                    Console.ResetColor();
                    foreach (string command in entry.Value)
                    {
                        string[] commandSplit = command.Split(new[] { ',' }, 3);
                        Console.WriteLine("{0,-30}{1,-50}", commandSplit[0], commandSplit[1].TrimStart());
                        Console.WriteLine("{0,-30}{1,-50}\n", "", commandSplit[2].TrimStart());
                    }
                }
            }

            else
            {
                bool printDescription = true;
                foreach (KeyValuePair<string, string[]> entry in commandDict)
                {
                    if (printDescription)
                    {
                        Console.WriteLine("\n{0,-30}{1,-50}", entry.Key, "Example Command");
                        Console.WriteLine("{0,-30}{1,-50}", "---------------------", "----------------------");

                        printDescription = false;
                    }
                    else
                    {
                        Console.WriteLine("\n{0,-30}", entry.Key);
                        Console.WriteLine("{0,-30}", "-----------------------------");
                    }

                    Console.ResetColor();
                    foreach (string command in entry.Value)
                    {
                        string[] commandSplit = command.Split(new[] { ',' }, 3);
                        Console.WriteLine("{0,-30}{1,-50}", commandSplit[0], commandSplit[1].TrimStart());
                        Console.WriteLine("{0,-30}{1,-50}\n", "", commandSplit[2].TrimStart());
                    }
                }
            }

            Environment.Exit(0);
        }

        // Thanks Harley!
        // https://github.com/harleyQu1nn/AggressorScripts/blob/master/ProcessColor.cna
        public static readonly string[] AVs = {"Tanium.exe", "360RP.exe", "360SD.exe", "360Safe.exe", "360leakfixer.exe", "360rp.exe", "360safe.exe", "360sd.exe",
            "360tray.exe", "AAWTray.exe", "ACAAS.exe", "ACAEGMgr.exe", "ACAIS.exe", "AClntUsr.EXE", "ALERT.EXE", "ALERTSVC.EXE", "ALMon.exe", "ALUNotify.exe",
            "ALUpdate.exe", "ALsvc.exe", "AVENGINE.exe", "AVGCHSVX.EXE", "AVGCSRVX.EXE", "AVGIDSAgent.exe", "AVGIDSMonitor.exe", "AVGIDSUI.exe", "AVGIDSWatcher.exe",
            "AVGNSX.EXE", "AVKProxy.exe", "AVKService.exe", "AVKTray.exe", "AVKWCtl.exe", "AVP.EXE", "AVP.exe", "AVPDTAgt.exe", "AcctMgr.exe", "Ad-Aware.exe",
            "Ad-Aware2007.exe", "AddressExport.exe", "AdminServer.exe", "Administrator.exe", "AeXAgentUIHost.exe", "AeXNSAgent.exe", "AeXNSRcvSvc.exe", "AlertSvc.exe",
            "AlogServ.exe", "AluSchedulerSvc.exe", "AnVir.exe", "AppSvc32.exe", "AtrsHost.exe", "Auth8021x.exe", "AvastSvc.exe", "AvastUI.exe", "Avconsol.exe", "AvpM.exe",
            "Avsynmgr.exe", "Avtask.exe", "BLACKD.exe", "BWMeterConSvc.exe", "CAAntiSpyware.exe", "CALogDump.exe", "CAPPActiveProtection.exe", "CAPPActiveProtection.exe",
            "CB.exe", "CCAP.EXE", "CCenter.exe", "CClaw.exe", "CLPS.exe", "CLPSLA.exe", "CLPSLS.exe", "CNTAoSMgr.exe", "CPntSrv.exe", "CTDataLoad.exe",
            "CertificationManagerServiceNT.exe", "ClShield.exe", "ClamTray.exe", "ClamWin.exe", "Console.exe", "CylanceUI.exe", "DAO_Log.exe", "DLService.exe",
            "DLTray.EXE", "DLTray.exe", "DRWAGNTD.EXE", "DRWAGNUI.EXE", "DRWEB32W.EXE", "DRWEBSCD.EXE", "DRWEBUPW.EXE", "DRWINST.EXE", "DSMain.exe", "DWHWizrd.exe",
            "DefWatch.exe", "DolphinCharge.exe", "EHttpSrv.exe", "EMET_Agent.exe", "EMET_Service.exe", "EMLPROUI.exe", "EMLPROXY.exe", "EMLibUpdateAgentNT.exe",
            "ETConsole3.exe", "ETCorrel.exe", "ETLogAnalyzer.exe", "ETReporter.exe", "ETRssFeeds.exe", "EUQMonitor.exe", "EndPointSecurity.exe", "EngineServer.exe",
            "EntityMain.exe", "EtScheduler.exe", "EtwControlPanel.exe", "EventParser.exe", "FAMEH32.exe", "FCDBLog.exe", "FCH32.exe", "FPAVServer.exe", "FProtTray.exe",
            "FSCUIF.exe", "FSHDLL32.exe", "FSM32.exe", "FSMA32.exe", "FSMB32.exe", "FWCfg.exe", "FireSvc.exe", "FireTray.exe", "FirewallGUI.exe", "ForceField.exe",
            "FortiProxy.exe", "FortiTray.exe", "FortiWF.exe", "FrameworkService.exe", "FreeProxy.exe", "GDFirewallTray.exe", "GDFwSvc.exe", "HWAPI.exe", "ISNTSysMonitor.exe",
            "ISSVC.exe", "ISWMGR.exe", "ITMRTSVC.exe", "ITMRT_SupportDiagnostics.exe", "ITMRT_TRACE.exe", "IcePack.exe", "IdsInst.exe", "InoNmSrv.exe", "InoRT.exe",
            "InoRpc.exe", "InoTask.exe", "InoWeb.exe", "IsntSmtp.exe", "KABackReport.exe", "KANMCMain.exe", "KAVFS.EXE", "KAVStart.exe", "KLNAGENT.EXE", "KMailMon.exe",
            "KNUpdateMain.exe", "KPFWSvc.exe", "KSWebShield.exe", "KVMonXP.exe", "KVMonXP_2.exe", "KVSrvXP.exe", "KWSProd.exe", "KWatch.exe", "KavAdapterExe.exe",
            "KeyPass.exe", "KvXP.exe", "LUALL.EXE", "LWDMServer.exe", "LockApp.exe", "LockAppHost.exe", "LogGetor.exe", "MCSHIELD.EXE", "MCUI32.exe", "MSASCui.exe",
            "ManagementAgentNT.exe", "McAfeeDataBackup.exe", "McEPOC.exe", "McEPOCfg.exe", "McNASvc.exe", "McProxy.exe", "McScript_InUse.exe", "McWCE.exe",
            "McWCECfg.exe", "Mcshield.exe", "Mctray.exe", "MgntSvc.exe", "MpCmdRun.exe", "MpfAgent.exe", "MpfSrv.exe", "MsMpEng.exe", "NAIlgpip.exe", "NAVAPSVC.EXE",
            "NAVAPW32.EXE", "NCDaemon.exe", "NIP.exe", "NJeeves.exe", "NLClient.exe", "NMAGENT.EXE", "NOD32view.exe", "NPFMSG.exe", "NPROTECT.EXE", "NRMENCTB.exe",
            "NSMdtr.exe", "NTRtScan.exe", "NVCOAS.exe", "NVCSched.exe", "NavShcom.exe", "Navapsvc.exe", "NaveCtrl.exe", "NaveLog.exe", "NaveSP.exe", "Navw32.exe",
            "Navwnt.exe", "Nip.exe", "Njeeves.exe", "Npfmsg2.exe", "Npfsvice.exe", "NscTop.exe", "Nvcoas.exe", "Nvcsched.exe", "Nymse.exe", "OLFSNT40.EXE", "OMSLogManager.exe",
            "ONLINENT.exe", "ONLNSVC.exe", "OfcPfwSvc.exe", "PASystemTray.exe", "PAVFNSVR.exe", "PAVSRV51.exe", "PNmSrv.exe", "POPROXY.EXE", "POProxy.exe", "PPClean.exe",
            "PPCtlPriv.exe", "PQIBrowser.exe", "PSHost.exe", "PSIMSVC.EXE", "PXEMTFTP.exe", "PadFSvr.exe", "Pagent.exe", "Pagentwd.exe", "PavBckPT.exe", "PavFnSvr.exe",
            "PavPrSrv.exe", "PavProt.exe", "PavReport.exe", "Pavkre.exe", "PcCtlCom.exe", "PcScnSrv.exe", "PccNTMon.exe", "PccNTUpd.exe", "PpPpWallRun.exe", "PrintDevice.exe",
            "ProUtil.exe", "PsCtrlS.exe", "PsImSvc.exe", "PwdFiltHelp.exe", "Qoeloader.exe", "RAVMOND.exe", "RAVXP.exe", "RNReport.exe", "RPCServ.exe", "RSSensor.exe",
            "RTVscan.exe", "RapApp.exe", "Rav.exe", "RavAlert.exe", "RavMon.exe", "RavMonD.exe", "RavService.exe", "RavStub.exe", "RavTask.exe", "RavTray.exe", "RavUpdate.exe",
            "RavXP.exe", "RealMon.exe", "Realmon.exe", "RedirSvc.exe", "RegMech.exe", "ReporterSvc.exe", "RouterNT.exe", "Rtvscan.exe", "SAFeService.exe", "SAService.exe",
            "SAVAdminService.exe", "SAVFMSESp.exe", "SAVMain.exe", "SAVScan.exe", "SCANMSG.exe", "SCANWSCS.exe", "SCFManager.exe", "SCFService.exe", "SCFTray.exe",
            "SDTrayApp.exe", "SEVINST.EXE", "SMEX_ActiveUpdate.exe", "SMEX_Master.exe", "SMEX_RemoteConf.exe", "SMEX_SystemWatch.exe", "SMSECtrl.exe", "SMSELog.exe",
            "SMSESJM.exe", "SMSESp.exe", "SMSESrv.exe", "SMSETask.exe", "SMSEUI.exe", "SNAC.EXE", "SNAC.exe", "SNDMon.exe", "SNDSrvc.exe", "SPBBCSvc.exe", "SPIDERML.EXE",
            "SPIDERNT.EXE", "SSM.exe", "SSScheduler.exe", "SVCharge.exe", "SVDealer.exe", "SVFrame.exe", "SVTray.exe", "SWNETSUP.EXE", "SavRoam.exe", "SavService.exe",
            "SavUI.exe", "ScanMailOutLook.exe", "SeAnalyzerTool.exe", "SemSvc.exe", "SescLU.exe", "SetupGUIMngr.exe", "SiteAdv.exe", "Smc.exe", "SmcGui.exe", "SnHwSrv.exe",
            "SnICheckAdm.exe", "SnIcon.exe", "SnSrv.exe", "SnicheckSrv.exe", "SpIDerAgent.exe", "SpntSvc.exe", "SpyEmergency.exe", "SpyEmergencySrv.exe", "StOPP.exe",
            "StWatchDog.exe", "SymCorpUI.exe", "SymSPort.exe", "TBMon.exe", "TFGui.exe", "TFService.exe", "TFTray.exe", "TFun.exe", "TIASPN~1.EXE", "TSAnSrf.exe", "TSAtiSy.exe",
            "TScutyNT.exe", "TSmpNT.exe", "TmListen.exe", "TmPfw.exe", "Tmntsrv.exe", "Traflnsp.exe", "TrapTrackerMgr.exe", "UPSCHD.exe", "UcService.exe", "UdaterUI.exe",
            "UmxAgent.exe", "UmxCfg.exe", "UmxFwHlp.exe", "UmxPol.exe", "Up2date.exe", "UpdaterUI.exe", "UrlLstCk.exe", "UserActivity.exe", "UserAnalysis.exe", "UsrPrmpt.exe",
            "V3Medic.exe", "V3Svc.exe", "VPC32.exe", "VPDN_LU.exe", "VPTray.exe", "VSStat.exe", "VsStat.exe", "VsTskMgr.exe", "WEBPROXY.EXE", "WFXCTL32.EXE", "WFXMOD32.EXE",
            "WFXSNT40.EXE", "WebProxy.exe", "WebScanX.exe", "WinRoute.exe", "WrSpySetup.exe", "ZLH.exe", "Zanda.exe", "ZhuDongFangYu.exe", "Zlh.exe", "_avp32.exe", "_avpcc.exe",
            "_avpm.exe", "aAvgApi.exe", "aawservice.exe", "acaif.exe", "acctmgr.exe", "ackwin32.exe", "aclient.exe", "adaware.exe", "advxdwin.exe", "aexnsagent.exe",
            "aexsvc.exe", "aexswdusr.exe", "aflogvw.exe", "afwServ.exe", "agentsvr.exe", "agentw.exe", "ahnrpt.exe", "ahnsd.exe", "ahnsdsv.exe", "alertsvc.exe", "alevir.exe",
            "alogserv.exe", "alsvc.exe", "alunotify.exe", "aluschedulersvc.exe", "amon9x.exe", "amswmagt.exe", "anti-trojan.exe", "antiarp.exe", "antivirus.exe", "ants.exe",
            "aphost.exe", "apimonitor.exe", "aplica32.exe", "aps.exe", "apvxdwin.exe", "arr.exe", "ashAvast.exe", "ashBug.exe", "ashChest.exe", "ashCmd.exe", "ashDisp.exe",
            "ashEnhcd.exe", "ashLogV.exe", "ashMaiSv.exe", "ashPopWz.exe", "ashQuick.exe", "ashServ.exe", "ashSimp2.exe", "ashSimpl.exe", "ashSkPcc.exe", "ashSkPck.exe",
            "ashUpd.exe", "ashWebSv.exe", "ashdisp.exe", "ashmaisv.exe", "ashserv.exe", "ashwebsv.exe", "asupport.exe", "aswDisp.exe", "aswRegSvr.exe", "aswServ.exe",
            "aswUpdSv.exe", "aswUpdsv.exe", "aswWebSv.exe", "aswupdsv.exe", "atcon.exe", "atguard.exe", "atro55en.exe", "atupdater.exe", "atwatch.exe", "atwsctsk.exe",
            "au.exe", "aupdate.exe", "aupdrun.exe", "aus.exe", "auto-protect.nav80try.exe", "autodown.exe", "autotrace.exe", "autoup.exe", "autoupdate.exe", "avEngine.exe",
            "avadmin.exe", "avcenter.exe", "avconfig.exe", "avconsol.exe", "ave32.exe", "avengine.exe", "avesvc.exe", "avfwsvc.exe", "avgam.exe", "avgamsvr.exe", "avgas.exe",
            "avgcc.exe", "avgcc32.exe", "avgcsrvx.exe", "avgctrl.exe", "avgdiag.exe", "avgemc.exe", "avgfws8.exe", "avgfws9.exe", "avgfwsrv.exe", "avginet.exe", "avgmsvr.exe",
            "avgnsx.exe", "avgnt.exe", "avgregcl.exe", "avgrssvc.exe", "avgrsx.exe", "avgscanx.exe", "avgserv.exe", "avgserv9.exe", "avgsystx.exe", "avgtray.exe", "avguard.exe",
            "avgui.exe", "avgupd.exe", "avgupdln.exe", "avgupsvc.exe", "avgvv.exe", "avgw.exe", "avgwb.exe", "avgwdsvc.exe", "avgwizfw.exe", "avkpop.exe", "avkserv.exe",
            "avkservice.exe", "avkwctl9.exe", "avltmain.exe", "avmailc.exe", "avmcdlg.exe", "avnotify.exe", "avnt.exe", "avp.exe", "avp32.exe", "avpcc.exe", "avpdos32.exe",
            "avpexec.exe", "avpm.exe", "avpncc.exe", "avps.exe", "avptc32.exe", "avpupd.exe", "avscan.exe", "avsched32.exe", "avserver.exe", "avshadow.exe", "avsynmgr.exe",
            "avwebgrd.exe", "avwin.exe", "avwin95.exe", "avwinnt.exe", "avwupd.exe", "avwupd32.exe", "avwupsrv.exe", "avxmonitor9x.exe", "avxmonitornt.exe", "avxquar.exe",
            "backweb.exe", "bargains.exe", "basfipm.exe", "bd_professional.exe", "bdagent.exe", "bdc.exe", "bdlite.exe", "bdmcon.exe", "bdss.exe", "bdsubmit.exe", "beagle.exe",
            "belt.exe", "bidef.exe", "bidserver.exe", "bipcp.exe", "bipcpevalsetup.exe", "bisp.exe", "blackd.exe", "blackice.exe", "blink.exe", "blss.exe", "bmrt.exe",
            "bootconf.exe", "bootwarn.exe", "borg2.exe", "bpc.exe", "bpk.exe", "brasil.exe", "bs120.exe", "bundle.exe", "bvt.exe", "bwgo0000.exe", "ca.exe", "caav.exe",
            "caavcmdscan.exe", "caavguiscan.exe", "caf.exe", "cafw.exe", "caissdt.exe", "capfaem.exe", "capfasem.exe", "capfsem.exe", "capmuamagt.exe", "casc.exe",
            "casecuritycenter.exe", "caunst.exe", "cavrep.exe", "cavrid.exe", "cavscan.exe", "cavtray.exe", "ccApp.exe", "ccEvtMgr.exe", "ccLgView.exe", "ccProxy.exe",
            "ccSetMgr.exe", "ccSetmgr.exe", "ccSvcHst.exe", "ccap.exe", "ccapp.exe", "ccevtmgr.exe", "cclaw.exe", "ccnfagent.exe", "ccprovsp.exe", "ccproxy.exe", "ccpxysvc.exe",
            "ccschedulersvc.exe", "ccsetmgr.exe", "ccsmagtd.exe", "ccsvchst.exe", "ccsystemreport.exe", "cctray.exe", "ccupdate.exe", "cdp.exe", "cfd.exe", "cfftplugin.exe",
            "cfgwiz.exe", "cfiadmin.exe", "cfiaudit.exe", "cfinet.exe", "cfinet32.exe", "cfnotsrvd.exe", "cfp.exe", "cfpconfg.exe", "cfpconfig.exe", "cfplogvw.exe",
            "cfpsbmit.exe", "cfpupdat.exe", "cfsmsmd.exe", "checkup.exe", "cka.exe", "clamscan.exe", "claw95.exe", "claw95cf.exe", "clean.exe", "cleaner.exe", "cleaner3.exe",
            "cleanpc.exe", "cleanup.exe", "click.exe", "cmdagent.exe", "cmdinstall.exe", "cmesys.exe", "cmgrdian.exe", "cmon016.exe", "comHost.exe", "connectionmonitor.exe",
            "control_panel.exe", "cpd.exe", "cpdclnt.exe", "cpf.exe", "cpf9x206.exe", "cpfnt206.exe", "crashrep.exe", "csacontrol.exe", "csinject.exe", "csinsm32.exe",
            "csinsmnt.exe", "csrss_tc.exe", "ctrl.exe", "cv.exe", "cwnb181.exe", "cwntdwmo.exe", "cz.exe", "datemanager.exe", "dbserv.exe", "dbsrv9.exe", "dcomx.exe",
            "defalert.exe", "defscangui.exe", "defwatch.exe", "deloeminfs.exe", "deputy.exe", "diskmon.exe", "divx.exe", "djsnetcn.exe", "dllcache.exe", "dllreg.exe",
            "doors.exe", "doscan.exe", "dpf.exe", "dpfsetup.exe", "dpps2.exe", "drwagntd.exe", "drwatson.exe", "drweb.exe", "drweb32.exe", "drweb32w.exe", "drweb386.exe",
            "drwebcgp.exe", "drwebcom.exe", "drwebdc.exe", "drwebmng.exe", "drwebscd.exe", "drwebupw.exe", "drwebwcl.exe", "drwebwin.exe", "drwupgrade.exe", "dsmain.exe",
            "dssagent.exe", "dvp95.exe", "dvp95_0.exe", "dwengine.exe", "dwhwizrd.exe", "dwwin.exe", "ecengine.exe", "edisk.exe", "efpeadm.exe", "egui.exe", "ekrn.exe",
            "elogsvc.exe", "emet_agent.exe", "emet_service.exe", "emsw.exe", "engineserver.exe", "ent.exe", "era.exe", "esafe.exe", "escanhnt.exe", "escanv95.exe",
            "esecagntservice.exe", "esecservice.exe", "esmagent.exe", "espwatch.exe", "etagent.exe", "ethereal.exe", "etrustcipe.exe", "evpn.exe", "evtProcessEcFile.exe",
            "evtarmgr.exe", "evtmgr.exe", "exantivirus-cnet.exe", "exe.avxw.exe", "execstat.exe", "expert.exe", "explore.exe", "f-agnt95.exe", "f-prot.exe", "f-prot95.exe",
            "f-stopw.exe", "fameh32.exe", "fast.exe", "fch32.exe", "fih32.exe", "findviru.exe", "firesvc.exe", "firetray.exe", "firewall.exe", "fmon.exe", "fnrb32.exe",
            "fortifw.exe", "fp-win.exe", "fp-win_trial.exe", "fprot.exe", "frameworkservice.exe", "frminst.exe", "frw.exe", "fsaa.exe", "fsaua.exe", "fsav.exe", "fsav32.exe",
            "fsav530stbyb.exe", "fsav530wtbyb.exe", "fsav95.exe", "fsavgui.exe", "fscuif.exe", "fsdfwd.exe", "fsgk32.exe", "fsgk32st.exe", "fsguidll.exe", "fsguiexe.exe",
            "fshdll32.exe", "fsm32.exe", "fsma32.exe", "fsmb32.exe", "fsorsp.exe", "fspc.exe", "fspex.exe", "fsqh.exe", "fssm32.exe", "fwinst.exe", "gator.exe", "gbmenu.exe",
            "gbpoll.exe", "gcascleaner.exe", "gcasdtserv.exe", "gcasinstallhelper.exe", "gcasnotice.exe", "gcasserv.exe", "gcasservalert.exe", "gcasswupdater.exe",
            "generics.exe", "gfireporterservice.exe", "ghost_2.exe", "ghosttray.exe", "giantantispywaremain.exe", "giantantispywareupdater.exe", "gmt.exe", "guard.exe",
            "guarddog.exe", "guardgui.exe", "hacktracersetup.exe", "hbinst.exe", "hbsrv.exe", "hipsvc.exe", "hotactio.exe", "hotpatch.exe", "htlog.exe", "htpatch.exe",
            "hwpe.exe", "hxdl.exe", "hxiul.exe", "iamapp.exe", "iamserv.exe", "iamstats.exe", "ibmasn.exe", "ibmavsp.exe", "icepack.exe", "icload95.exe", "icloadnt.exe",
            "icmon.exe", "icsupp95.exe", "icsuppnt.exe", "idle.exe", "iedll.exe", "iedriver.exe", "iface.exe", "ifw2000.exe", "igateway.exe", "inetlnfo.exe", "infus.exe",
            "infwin.exe", "inicio.exe", "init.exe", "inonmsrv.exe", "inorpc.exe", "inort.exe", "inotask.exe", "intdel.exe", "intren.exe", "iomon98.exe", "isPwdSvc.exe",
            "isUAC.exe", "isafe.exe", "isafinst.exe", "issvc.exe", "istsvc.exe", "jammer.exe", "jdbgmrg.exe", "jedi.exe", "kaccore.exe", "kansgui.exe", "kansvr.exe",
            "kastray.exe", "kav.exe", "kav32.exe", "kavfs.exe", "kavfsgt.exe", "kavfsrcn.exe", "kavfsscs.exe", "kavfswp.exe", "kavisarv.exe", "kavlite40eng.exe",
            "kavlotsingleton.exe", "kavmm.exe", "kavpers40eng.exe", "kavpf.exe", "kavshell.exe", "kavss.exe", "kavstart.exe", "kavsvc.exe", "kavtray.exe", "kazza.exe",
            "keenvalue.exe", "kerio-pf-213-en-win.exe", "kerio-wrl-421-en-win.exe", "kerio-wrp-421-en-win.exe", "kernel32.exe", "killprocesssetup161.exe", "kis.exe",
            "kislive.exe", "kissvc.exe", "klnacserver.exe", "klnagent.exe", "klserver.exe", "klswd.exe", "klwtblfs.exe", "kmailmon.exe", "knownsvr.exe", "kpf4gui.exe",
            "kpf4ss.exe", "kpfw32.exe", "kpfwsvc.exe", "krbcc32s.exe", "kvdetech.exe", "kvolself.exe", "kvsrvxp.exe", "kvsrvxp_1.exe", "kwatch.exe", "kwsprod.exe",
            "kxeserv.exe", "launcher.exe", "ldnetmon.exe", "ldpro.exe", "ldpromenu.exe", "ldscan.exe", "leventmgr.exe", "livesrv.exe", "lmon.exe", "lnetinfo.exe",
            "loader.exe", "localnet.exe", "lockdown.exe", "lockdown2000.exe", "log_qtine.exe", "lookout.exe", "lordpe.exe", "lsetup.exe", "luall.exe", "luau.exe",
            "lucallbackproxy.exe", "lucoms.exe", "lucomserver.exe", "lucoms~1.exe", "luinit.exe", "luspt.exe", "makereport.exe", "mantispm.exe", "mapisvc32.exe",
            "masalert.exe", "massrv.exe", "mcafeefire.exe", "mcagent.exe", "mcappins.exe", "mcconsol.exe", "mcdash.exe", "mcdetect.exe", "mcepoc.exe", "mcepocfg.exe",
            "mcinfo.exe", "mcmnhdlr.exe", "mcmscsvc.exe", "mcods.exe", "mcpalmcfg.exe", "mcpromgr.exe", "mcregwiz.exe", "mcscript.exe", "mcscript_inuse.exe", "mcshell.exe",
            "mcshield.exe", "mcshld9x.exe", "mcsysmon.exe", "mctool.exe", "mctray.exe", "mctskshd.exe", "mcuimgr.exe", "mcupdate.exe", "mcupdmgr.exe", "mcvsftsn.exe",
            "mcvsrte.exe", "mcvsshld.exe", "mcwce.exe", "mcwcecfg.exe", "md.exe", "mfeann.exe", "mfevtps.exe", "mfin32.exe", "mfw2en.exe", "mfweng3.02d30.exe",
            "mgavrtcl.exe", "mgavrte.exe", "mghtml.exe", "mgui.exe", "minilog.exe", "mmod.exe", "monitor.exe", "monsvcnt.exe", "monsysnt.exe", "moolive.exe",
            "mostat.exe", "mpcmdrun.exe", "mpf.exe", "mpfagent.exe", "mpfconsole.exe", "mpfservice.exe", "mpftray.exe", "mps.exe", "mpsevh.exe", "mpsvc.exe", "mrf.exe",
            "mrflux.exe", "msapp.exe", "msascui.exe", "msbb.exe", "msblast.exe", "mscache.exe", "msccn32.exe", "mscifapp.exe", "mscman.exe", "msconfig.exe", "msdm.exe",
            "msdos.exe", "msiexec16.exe", "mskagent.exe", "mskdetct.exe", "msksrver.exe", "msksrvr.exe", "mslaugh.exe", "msmgt.exe", "msmpeng.exe", "msmsgri32.exe",
            "msscli.exe", "msseces.exe", "mssmmc32.exe", "msssrv.exe", "mssys.exe", "msvxd.exe", "mu0311ad.exe", "mwatch.exe", "myagttry.exe", "n32scanw.exe", "nSMDemf.exe",
            "nSMDmon.exe", "nSMDreal.exe", "nSMDsch.exe", "naPrdMgr.exe", "nav.exe", "navap.navapsvc.exe", "navapsvc.exe", "navapw32.exe", "navdx.exe", "navlu32.exe",
            "navnt.exe", "navstub.exe", "navw32.exe", "navwnt.exe", "nc2000.exe", "ncinst4.exe", "MSASCuiL.exe", "MBAMService.exe", "mbamtray.exe", "CylanceSvc.exe",
            "ndd32.exe", "ndetect.exe", "neomonitor.exe", "neotrace.exe", "neowatchlog.exe", "netalertclient.exe",
            "netarmor.exe", "netcfg.exe", "netd32.exe", "netinfo.exe", "netmon.exe", "netscanpro.exe", "netspyhunter-1.2.exe", "netstat.exe", "netutils.exe", "networx.exe",
            "ngctw32.exe", "ngserver.exe", "nip.exe", "nipsvc.exe", "nisoptui.exe", "nisserv.exe", "nisum.exe", "njeeves.exe", "nlsvc.exe", "nmain.exe", "nod32.exe",
            "nod32krn.exe", "nod32kui.exe", "normist.exe", "norton_internet_secu_3.0_407.exe", "notstart.exe", "npf40_tw_98_nt_me_2k.exe", "npfmessenger.exe",
            "npfmntor.exe", "npfmsg.exe", "nprotect.exe", "npscheck.exe", "npssvc.exe", "nrmenctb.exe", "nsched32.exe", "nscsrvce.exe", "nsctop.exe", "nsmdtr.exe",
            "nssys32.exe", "nstask32.exe", "nsupdate.exe", "nt.exe", "ntcaagent.exe", "ntcadaemon.exe", "ntcaservice.exe", "ntrtscan.exe", "ntvdm.exe", "ntxconfig.exe",
            "nui.exe", "nupgrade.exe", "nvarch16.exe", "nvc95.exe", "nvcoas.exe", "nvcsched.exe", "nvsvc32.exe", "nwinst4.exe", "nwservice.exe", "nwtool16.exe", "nymse.exe",
            "oasclnt.exe", "oespamtest.exe", "ofcdog.exe", "ofcpfwsvc.exe", "okclient.exe", "olfsnt40.exe", "ollydbg.exe", "onsrvr.exe", "op_viewer.exe", "opscan.exe",
            "optimize.exe", "ostronet.exe", "otfix.exe", "outpost.exe", "outpostinstall.exe", "outpostproinstall.exe", "paamsrv.exe", "padmin.exe", "pagent.exe",
            "pagentwd.exe", "panixk.exe", "patch.exe", "pavbckpt.exe", "pavcl.exe", "pavfires.exe", "pavfnsvr.exe", "pavjobs.exe", "pavkre.exe", "pavmail.exe",
            "pavprot.exe", "pavproxy.exe", "pavprsrv.exe", "pavsched.exe", "pavsrv50.exe", "pavsrv51.exe", "pavsrv52.exe", "pavupg.exe", "pavw.exe", "pccNT.exe",
            "pccclient.exe", "pccguide.exe", "pcclient.exe", "pccnt.exe", "pccntmon.exe", "pccntupd.exe", "pccpfw.exe", "pcctlcom.exe", "pccwin98.exe", "pcfwallicon.exe",
            "pcip10117_0.exe", "pcscan.exe", "pctsAuxs.exe", "pctsGui.exe", "pctsSvc.exe", "pctsTray.exe", "pdsetup.exe", "pep.exe", "periscope.exe", "persfw.exe",
            "perswf.exe", "pf2.exe", "pfwadmin.exe", "pgmonitr.exe", "pingscan.exe", "platin.exe", "pmon.exe", "pnmsrv.exe", "pntiomon.exe", "pop3pack.exe", "pop3trap.exe",
            "poproxy.exe", "popscan.exe", "portdetective.exe", "portmonitor.exe", "powerscan.exe", "ppinupdt.exe", "ppmcativedetection.exe", "pptbc.exe", "ppvstop.exe",
            "pqibrowser.exe", "pqv2isvc.exe", "prevsrv.exe", "prizesurfer.exe", "prmt.exe", "prmvr.exe", "programauditor.exe", "proport.exe", "protectx.exe", "psctris.exe",
            "psh_svc.exe", "psimreal.exe", "psimsvc.exe", "pskmssvc.exe", "pspf.exe", "purge.exe", "pview.exe", "pviewer.exe", "pxemtftp.exe", "pxeservice.exe",
            "qclean.exe", "qconsole.exe", "qdcsfs.exe", "qoeloader.exe", "qserver.exe", "rapapp.exe", "rapuisvc.exe", "ras.exe", "rasupd.exe", "rav7.exe", "rav7win.exe",
            "rav8win32eng.exe", "ravmon.exe", "ravmond.exe", "ravstub.exe", "ravxp.exe", "ray.exe", "rb32.exe", "rcsvcmon.exe", "rcsync.exe", "realmon.exe", "reged.exe",
            "remupd.exe", "reportsvc.exe", "rescue.exe", "rescue32.exe", "rfwmain.exe", "rfwproxy.exe", "rfwsrv.exe", "rfwstub.exe", "rnav.exe", "rrguard.exe", "rshell.exe",
            "rsnetsvr.exe", "rstray.exe", "rtvscan.exe", "rtvscn95.exe", "rulaunch.exe", "saHookMain.exe", "safeboxtray.exe", "safeweb.exe", "sahagent.exescan32.exe",
            "sav32cli.exe", "save.exe", "savenow.exe", "savroam.exe", "savscan.exe", "savservice.exe", "sbserv.exe", "scam32.exe", "scan32.exe", "scan95.exe", "scanexplicit.exe",
            "scanfrm.exe", "scanmailoutlook.exe", "scanpm.exe", "schdsrvc.exe", "schupd.exe", "scrscan.exe", "seestat.exe", "serv95.exe", "setloadorder.exe",
            "setup_flowprotector_us.exe", "setupguimngr.exe", "setupvameeval.exe", "sfc.exe", "sgssfw32.exe", "sh.exe", "shellspyinstall.exe", "shn.exe", "showbehind.exe",
            "shstat.exe", "siteadv.exe", "smOutlookPack.exe", "smc.exe", "smoutlookpack.exe", "sms.exe", "smsesp.exe", "smss32.exe", "sndmon.exe", "sndsrvc.exe",
            "soap.exe", "sofi.exe", "softManager.exe", "spbbcsvc.exe", "spf.exe", "sphinx.exe", "spideragent.exe", "spiderml.exe", "spidernt.exe", "spiderui.exe",
            "spntsvc.exe", "spoler.exe", "spoolcv.exe", "spoolsv32.exe", "spyxx.exe", "srexe.exe", "srng.exe", "srvload.exe", "srvmon.exe", "ss3edit.exe", "sschk.exe",
            "ssg_4104.exe", "ssgrate.exe", "st2.exe", "stcloader.exe", "stinger.exe", "stopp.exe", "stwatchdog.exe", "supftrl.exe", "support.exe", "supporter5.exe",
            "svcGenericHost", "svcharge.exe", "svchostc.exe", "svchosts.exe", "svcntaux.exe", "svdealer.exe", "svframe.exe", "svtray.exe", "swdsvc.exe", "sweep95.exe",
            "sweepnet.sweepsrv.sys.swnetsup.exe", "sweepsrv.exe", "swnetsup.exe", "swnxt.exe", "swserver.exe", "symlcsvc.exe", "symproxysvc.exe", "symsport.exe", "symtray.exe",
            "symwsc.exe", "sysdoc32.exe", "sysedit.exe", "sysupd.exe", "taskmo.exe", "taumon.exe", "tbmon.exe", "tbscan.exe", "tc.exe", "tca.exe", "tclproc.exe", "tcm.exe",
            "tdimon.exe", "tds-3.exe", "tds2-98.exe", "tds2-nt.exe", "teekids.exe", "tfak.exe", "tfak5.exe", "tgbob.exe", "titanin.exe", "titaninxp.exe", "tmas.exe",
            "tmlisten.exe", "tmntsrv.exe", "tmpfw.exe", "tmproxy.exe", "tnbutil.exe", "tpsrv.exe", "tracesweeper.exe", "trickler.exe", "trjscan.exe", "trjsetup.exe",
            "trojantrap3.exe", "trupd.exe", "tsadbot.exe", "tvmd.exe", "tvtmd.exe", "udaterui.exe", "undoboot.exe", "unvet32.exe", "updat.exe", "updtnv28.exe", "upfile.exe",
            "upgrad.exe", "uplive.exe", "urllstck.exe", "usergate.exe", "usrprmpt.exe", "utpost.exe", "v2iconsole.exe", "v3clnsrv.exe", "v3exec.exe", "v3imscn.exe",
            "vbcmserv.exe", "vbcons.exe", "vbust.exe", "vbwin9x.exe", "vbwinntw.exe", "vcsetup.exe", "vet32.exe", "vet95.exe", "vetmsg.exe", "vettray.exe", "vfsetup.exe",
            "vir-help.exe", "virusmdpersonalfirewall.exe", "vnlan300.exe", "vnpc3000.exe", "vpatch.exe", "vpc32.exe", "vpc42.exe", "vpfw30s.exe", "vprosvc.exe",
            "vptray.exe", "vrv.exe", "vrvmail.exe", "vrvmon.exe", "vrvnet.exe", "vscan40.exe", "vscenu6.02d30.exe", "vsched.exe", "vsecomr.exe", "vshwin32.exe", "vsisetup.exe",
            "vsmain.exe", "vsmon.exe", "vsserv.exe", "vsstat.exe", "vstskmgr.exe", "vswin9xe.exe", "vswinntse.exe", "vswinperse.exe", "w32dsm89.exe", "w9x.exe",
            "watchdog.exe", "webdav.exe", "webproxy.exe", "webscanx.exe", "webtrap.exe", "webtrapnt.exe", "wfindv32.exe", "wfxctl32.exe", "wfxmod32.exe", "wfxsnt40.exe",
            "whoswatchingme.exe", "wimmun32.exe", "win-bugsfix.exe", "winactive.exe", "winmain.exe", "winnet.exe", "winppr32.exe", "winrecon.exe", "winroute.exe", "winservn.exe",
            "winssk32.exe", "winstart.exe", "winstart001.exe", "wintsk32.exe", "winupdate.exe", "wkufind.exe", "wnad.exe", "wnt.exe", "wradmin.exe", "wrctrl.exe",
            "wsbgate.exe", "wssfcmai.exe", "wupdater.exe", "wupdt.exe", "wyvernworksfirewall.exe", "xagt.exe", "xagtnotif.exe", "xcommsvr.exe", "xfilter.exe", "xpf202en.exe",
            "zanda.exe", "zapro.exe", "zapsetup3001.exe", "zatutor.exe", "zhudongfangyu.exe", "zlclient.exe", "zlh.exe", "zonalm2601.exe", "zonealarm.exe", "cb.exe",
            "MsMpEng.exe", "MsSense.exe", "CSFalconService.exe", "CSFalconContainer.exe", "redcloak.exe", "OmniAgent.exe","CrAmTray.exe","AmSvc.exe","minionhost.exe",
            "PylumLoader.exe","CrsSvc.exe"};

        public static readonly string[] Admin = { "MobaXterm.exe", "bash.exe", "git-bash.exe", "mmc.exe", "Code.exe", "notepad++.exe", "notepad.exe", "cmd.exe",
            "drwatson.exe", "DRWTSN32.EXE", "drwtsn32.exe", "dumpcap.exe", "ethereal.exe", "filemon.exe", "idag.exe", "idaw.exe", "k1205.exe", "loader32.exe",
            "netmon.exe", "netstat.exe", "netxray.exe", "NmWebService.exe", "nukenabber.exe", "portmon.exe", "powershell.exe", "PRTG Traffic Gr.exe",
            "PRTG Traffic Grapher.exe", "prtgwatchdog.exe", "putty.exe", "regmon.exe", "SystemEye.exe", "taskman.exe", "TASKMGR.EXE", "tcpview.exe", "Totalcmd.exe",
            "TrafMonitor.exe", "windbg.exe", "winobj.exe", "wireshark.exe", "WMonAvNScan.exe", "WMonAvScan.exe", "WMonSrv.exe","regedit.exe", "regedit32.exe",
            "accesschk.exe", "accesschk64.exe", "AccessEnum.exe", "ADExplorer.exe", "ADInsight.exe", "adrestore.exe", "Autologon.exe", "Autoruns.exe", "Autoruns64.exe",
            "autorunsc.exe", "autorunsc64.exe", "Bginfo.exe", "Bginfo64.exe", "Cacheset.exe", "Clockres.exe", "Clockres64.exe", "Contig.exe", "Contig64.exe",
            "Coreinfo.exe", "ctrl2cap.exe", "Dbgview.exe", "Desktops.exe", "disk2vhd.exe", "diskext.exe", "diskext64.exe", "Diskmon.exe", "DiskView.exe", "du.exe",
            "du64.exe", "efsdump.exe", "FindLinks.exe", "FindLinks64.exe", "handle.exe", "handle64.exe", "hex2dec.exe", "hex2dec64.exe", "junction.exe", "junction64.exe",
            "ldmdump.exe", "Listdlls.exe", "Listdlls64.exe", "livekd.exe", "livekd64.exe", "LoadOrd.exe", "LoadOrd64.exe", "LoadOrdC.exe", "LoadOrdC64.exe",
            "logonsessions.exe", "logonsessions64.exe", "movefile.exe", "movefile64.exe", "notmyfault.exe", "notmyfault64.exe", "notmyfaultc.exe", "notmyfaultc64.exe",
            "ntfsinfo.exe", "ntfsinfo64.exe", "pagedfrg.exe", "pendmoves.exe", "pendmoves64.exe", "pipelist.exe", "pipelist64.exe", "portmon.exe", "procdump.exe",
            "procdump64.exe", "procexp.exe", "procexp64.exe", "Procmon.exe", "PsExec.exe", "PsExec64.exe", "psfile.exe", "psfile64.exe", "PsGetsid.exe", "PsGetsid64.exe",
            "PsInfo.exe", "PsInfo64.exe", "pskill.exe", "pskill64.exe", "pslist.exe", "pslist64.exe", "PsLoggedon.exe", "PsLoggedon64.exe", "psloglist.exe", "pspasswd.exe",
            "pspasswd64.exe", "psping.exe", "psping64.exe", "PsService.exe", "PsService64.exe", "psshutdown.exe", "pssuspend.exe", "pssuspend64.exe", "RAMMap.exe",
            "RegDelNull.exe", "RegDelNull64.exe", "regjump.exe", "ru.exe", "ru64.exe", "sdelete.exe", "sdelete64.exe", "ShareEnum.exe", "ShellRunas.exe", "sigcheck.exe",
            "sigcheck64.exe", "streams.exe", "streams64.exe", "strings.exe", "strings64.exe", "sync.exe", "sync64.exe", "Sysmon.exe", "Sysmon64.exe", "Tcpvcon.exe",
            "Tcpview.exe", "Testlimit.exe", "Testlimit64.exe", "vmmap.exe", "Volumeid.exe", "Volumeid64.exe", "whois.exe", "whois64.exe", "Winobj.exe", "ZoomIt.exe",
            "KeePass.exe", "1Password.exe", "lastpass.exe"};

        // Thanks Harley!
        // https://github.com/harleyQu1nn/AggressorScripts/blob/master/EDR.cna
        public static readonly string[] Edr = {"CiscoAMPCEFWDriver.sys", "CiscoAMPHeurDriver.sys", "cbstream.sys", "cbk7.sys", "Parity.sys", "libwamf.sys",
            "LRAgentMF.sys", "BrCow_x_x_x_x.sys", "brfilter.sys", "BDSandBox.sys", "TRUFOS.SYS", "AVC3.SYS", "Atc.sys", "AVCKF.SYS", "bddevflt.sys", "gzflt.sys",
            "bdsvm.sys", "hbflt.sys", "cve.sys", "psepfilter.sys", "cposfw.sys", "dsfa.sys", "medlpflt.sys", "epregflt.sys", "TmFileEncDmk.sys", "tmevtmgr.sys",
            "TmEsFlt.sys", "fileflt.sys", "SakMFile.sys", "SakFile.sys", "AcDriver.sys", "TMUMH.sys", "hfileflt.sys", "TMUMS.sys", "MfeEEFF.sys", "mfprom.sys",
            "hdlpflt.sys", "swin.sys", "mfehidk.sys", "mfencoas.sys", "epdrv.sys", "carbonblackk.sys", "csacentr.sys", "csaenh.sys", "csareg.sys", "csascr.sys",
            "csaav.sys", "csaam.sys", "esensor.sys", "fsgk.sys", "fsatp.sys", "fshs.sys", "eaw.sys", "im.sys", "csagent.sys", "rvsavd.sys", "dgdmk.sys", "atrsdfw.sys",
            "mbamwatchdog.sys", "edevmon.sys", "SentinelMonitor.sys", "edrsensor.sys", "ehdrv.sys", "HexisFSMonitor.sys", "CyOptics.sys", "CarbonBlackK.sys",
            "CyProtectDrv32.sys", "CyProtectDrv64.sys", "CRExecPrev.sys", "ssfmonm.sys", "CybKernelTracker.sys", "SAVOnAccess.sys", "savonaccess.sys", "sld.sys",
            "aswSP.sys", "FeKern.sys", "klifks.sys", "klifaa.sys", "Klifsm.sys", "mfeaskm.sys", "mfencfilter.sys", "WFP_MRT.sys", "groundling32.sys", "SAFE-Agent.sys",
            "groundling64.sys", "avgtpx86.sys", "avgtpx64.sys", "pgpwdefs.sys", "GEProtection.sys", "diflt.sys", "sysMon.sys", "ssrfsf.sys", "emxdrv2.sys", "reghook.sys",
            "spbbcdrv.sys", "bhdrvx86.sys", "bhdrvx64.sys", "SISIPSFileFilter.sys", "symevent.sys", "VirtualAgent.sys", "vxfsrep.sys", "VirtFile.sys", "SymAFR.sys",
            "symefasi.sys", "symefa.sys", "symefa64.sys", "SymHsm.sys", "evmf.sys", "GEFCMP.sys", "VFSEnc.sys", "pgpfs.sys", "fencry.sys", "symrg.sys", "cfrmd.sys",
            "cmdccav.sys", "cmdguard.sys", "CmdMnEfs.sys", "MyDLPMF.sys", "PSINPROC.SYS", "PSINFILE.SYS", "amfsm.sys", "amm8660.sys", "amm6460.sys"};

    }
}