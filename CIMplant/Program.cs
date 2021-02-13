using CommandLine;
using CommandLine.Text;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Management;
using System.Reflection;
using System.Runtime.InteropServices;
using Execute;
using System.Management.Automation;
using System.Management.Automation.Remoting;

namespace CIMplant
{
    public class SubCommandException : Exception
    {
        public SubCommandException(string message)
            : base(message) { }
    }

    public class CaughtByAvException : Exception
    {
        public CaughtByAvException(string message)
            : base(message) { }
    }

    public class ExtraCommandException : Exception
    {
    }

    public class ServiceUnknownException : Exception
    {
        public ServiceUnknownException(string message)
            : base(message) { }
    }

    public class RektDebugFilePath : Exception
    {
        public RektDebugFilePath(string message)
            : base(message) { }
    }

    public class ProcessCommandException : Exception
    {
        public ProcessCommandException(string message)
            : base(message) { }
    }

    internal class Program
    {
        private static readonly string[] CommandArray =
        {
            "cat", "copy", "download", "ls", "search", "upload", "command_exec", "process_kill", "process_start",
            "ps", "active_users", "basic_info", "drive_list", "ifconfig", "installed_programs", "logoff", "reboot",
            "restart", "power_off", "shutdown",
            "vacant_system", "logon_events", "command_exec", "disable_wdigest", "enable_wdigest", "disable_winrm",
            "enable_winrm",
            "reg_mod", "reg_create", "reg_delete", "remote_posh", "sched_job", "service_mod"
        };

        public class Options
        {
            public static Options Instance { get; set; }

            // Command line options
            [Option('v', "verbose", Required = false, HelpText = "Set output to verbose")]
            public bool Verbose { get; set; }

            [Option('u', "username", Required = false, HelpText = "Specify a username to use")]
            public string Username { get; set; }

            [Option('p', "password", Required = false, HelpText = "Specify a password to use", Default = null)]
            public string Password { get; set; }

            [Option('d', "domain", Required = false, HelpText = "Specify a domain", Default = "WORKGROUP")]
            public string Domain { get; set; }

            [Option('s', "system", Group = "Required", Required = true, HelpText = "Specify a system to target",
                Default = "localhost")]
            public string System { get; set; }

            [Option('n', "namespace", Required = false, HelpText = "Specify a namespace to use",
                Default = "root\\cimv2")]
            public string NameSpace { get; set; }

            [Option('c', "command", Group = "Command", Required = true,
                HelpText = "Specify a command to run, run program with just '--show-commands' for a list of commands")]
            public string Command { get; set; }

            [Option('r', "reset", Group = "Command", Required = true,
                HelpText =
                    "Reset the DebugFilePath property back to the Windows default in the event of any execution errors")]
            public bool Reset { get; set; }

            [Option('e', "execute", Required = false,
                HelpText =
                    "Specify a command-line command to execute and receive the output for (use double quotes \"command\" for complex commands)")]
            public string Execute { get; set; }

            [Option('f', "file", Group = "Required", Required = true,
                HelpText = "Specify a remote or local file to cat/download/copy/search for/execute ps1/etc.",
                Default = null)]
            public string File { get; set; }

            [Option("cmdlet", Group = "Required", Required = true,
                HelpText = "Specify a cmdlet to run and obtain the results for", Default = null)]
            public string Cmdlet { get; set; }

            [Option("fileto", Group = "Required", Required = true, HelpText = "Specify a name to copy the file to",
                Default = null)]
            public string FileTo { get; set; }

            [Option("directory", Group = "Required", Required = true, HelpText = "Specify a directory to list/search",
                Default = null)]
            public string Directory { get; set; }

            [Option("regkey", Group = "Required", Required = true,
                HelpText =
                    "Specify a registry key to create/delete/modify (ex: HKLM\\SYSTEM\\CurrentControlSet\\Control\\SecurityProviders\\WDigest)",
                Default = null)]
            public string RegistryKey { get; set; }

            [Option("regsubkey", Group = "Required", Required = true,
                HelpText = "Specify a registry subkey to create/delete/modify (ex: UseLogonCredential)",
                Default = null)]
            public string RegistrySubkey { get; set; }

            [Option("regval", Group = "Required", Required = true,
                HelpText = "Specify a registry data value to create/delete/modify (ex: \"1\" for REG_DWORD)",
                Default = null)]
            public string RegistryValue { get; set; }

            [Option("regvaltype", Group = "Required", Required = true,
                HelpText =
                    "Specify a registry data type to create/delete/modify (case insensitive, ex: REG_DWORD, reg_binary, or reg_sz)",
                Default = null)]
            public string RegistryValueType { get; set; }

            [Option("service", Group = "Required", Required = true,
                HelpText = "Specify a service name to create/delete/start/stop", Default = null)]
            public string Service { get; set; }

            [Option("servicebin", Group = "Required", Required = true,
                HelpText = "Specify a service binary while creating a new service", Default = null)]
            public string ServiceBin { get; set; }

            [Option("process", Group = "Required", Required = true,
                HelpText = "Specify a process name or handle to kill or start (wildcards accepted for name)",
                Default = null)]
            public string Process { get; set; }

            [Option("wmi", Required = false, HelpText = "Use WMI (DCOM) to connect to the remote system instead of CIM/MI (WSMan)",
                Default = false)]
            public bool Wmi { get; set; }

            [Option("provider", Required = false, HelpText = "Use InstallUtil to register a WMI provider (Not Currently Working)",
                Default = false)]
            public bool Provider { get; set; }

            [Option("show-commands", Group = "Command", Required = true,
                HelpText = "Displays a list of available commands")]
            public bool ShowCommands { get; set; }

            [Option("show-examples", Group = "Command", Required = true,
                HelpText = "Displays examples for all available commands")]
            public bool ShowExamples { get; set; }

            [Option("no-banner", Group = "Command", Required = false,
                HelpText = "Disables that gorgeous ASCII art (probably should never use this)")]
            public bool NoBanner { get; set; }

            [Option("test", Group = "Command", Required = false,
                HelpText = "Tests all commands with a specified username/password/system (or against the localhost)")]
            public bool Test { get; set; }
        }

        private static void DisplayHelp<T>(ParserResult<T> result, IEnumerable<Error> errs)
        {
            HelpText helpText = HelpText.AutoBuild(result, h =>
            {
                h.AdditionalNewLineAfterOption = false;
                h.Heading = "WMI C# Version 0.1"; //change header
                h.Copyright = ""; //change copyright text
                h.AutoVersion = false;
                return HelpText.DefaultParsingErrorsHandler(result, h);
            }, e => e);
            Console.WriteLine(helpText);
            System.Environment.Exit(1);
        }

        private static void RunTestCases(Options options)
        {
            Commander commander = new Commander();
            var connector = new Connector();
            var planter = new Planter(commander, connector);
            planter.Connector = new Connector(options.Wmi, planter);

            try
            {
                
                foreach (var command in CommandArray)
                {
                    commander = command == null ? new Commander() : new Commander(command);

                    if (commander.Method == null)
                        commander.Method = commander.Command;

                    if (planter.Commander.Command != null)
                        Messenger.GoodMessage("[+] Results from " + planter.Commander.Command + ":\n");

                    object result = null;

                    // Block to set the specific Type for WMI/CIM command
                    Type type = options.Wmi ? typeof(ExecuteWmi) : typeof(ExecuteCim);
                    MethodInfo method = type.GetMethod((planter.Commander.Method ?? planter.Commander.Command) ?? string.Empty);

                    // Create an instance of the type
                    object instance = Activator.CreateInstance(type);

                    // Create parameter object
                    object[] stringMethodParams = { planter };

                    result = method.Invoke(instance, stringMethodParams);



                }
            }

            catch (Exception e)
            {
                ExceptionLogging.SendErrorToText(e);

            }
        }

        private static void Main(string[] args)
        {
            Messenger.ImportantAsciiArt();
            Stopwatch watch = new Stopwatch();
            watch.Start();

            const string defaultDebugFilePath = "%SystemRoot%\\MEMORY.DMP";

            // Parse arguments passed
            Parser parser = new Parser(with =>
            {
                with.CaseInsensitiveEnumValues = true;
                with.CaseSensitive = false;
                with.HelpWriter = null;
            });

            ParserResult<Options> parserResult = parser.ParseArguments<Options>(args);
            parserResult.WithParsed<Options>(o => { Options.Instance = o; })
                .WithNotParsed(errs => DisplayHelp(parserResult, errs));
            Options options = Options.Instance;
            Console.WriteLine();

            // Print all commands before doing anything if that's what the user wants
            if (options.ShowCommands)
                Messenger.GetCommands();

            if (options.ShowExamples)
                Messenger.GetExamples();

            if (options.Test)
            {
                RunTestCases(options);
                Console.WriteLine("Test cases completed");
                System.Environment.Exit(0);
            }

            //Need a separate namespace for certain commands that deal with registry get/set
            //if (options.Command.ToLower() == "disable_wdigest")
            //{
            //    //Console.WriteLine("New namespace due to command");
            //    NameSpace = @"stdRegProv";
            //}


            //////////
            // Block to instantiate the Commander class (houses all command information and checks for required vals)
            //////////
            Commander commander = new Commander();
            if (options.Command != null && CommandArray.Any(options.Command.ToLower().Contains) || options.Reset)
            {
                try
                {
                    commander = options.Command == null ? new Commander() : new Commander(options.Command);

                    if (commander.Method == null)
                        commander.Method = commander.Command;
                }

                catch (ArgumentNullException e)
                {
                    Messenger.ErrorMessage(
                        $"[-] Error: The parameter '{e.ParamName.ToLower()}' cannot be null (--{e.ParamName.ToLower()})");
                    System.Environment.Exit(1);
                }

                catch (SubCommandException e)
                {
                    if (e.Message == "service_mod")
                        Messenger.ErrorMessage(
                            $"[-] Error: the subcommand for '{e.Message}' is incorrect. It must be list, create, start, or delete");
                    System.Environment.Exit(1);
                }

                catch (ExtraCommandException)
                {
                    Messenger.ErrorMessage(
                        "[-] Error: Please only specify one command to execute using the -c or --command flag");
                    System.Environment.Exit(1);
                }

                catch (ProcessCommandException)
                {
                    Messenger.ErrorMessage(
                        "[-] Error: Please specify a process or handle to kill (wildcards accepted for name, ex: --process note* or --process 5384)");
                }

                catch (Exception e)
                {
                    Console.WriteLine($"Exception {e.Message} Trace {e.StackTrace}");
                }
            }

            else
            {
                Messenger.ErrorMessage("\n[-] Incorrenct command used. Try one of these:\n");
                Messenger.GetCommands();
                System.Environment.Exit(1);
            }

            //////////
            // Block to instantiate the Planter class (houses all info about the target system)
            //////////
            var connector = new Connector();
            var planter = new Planter(commander, connector);

            //////////
            // Block to create the connection to either WMI or CIM and fallback to the other
            //////////
            try
            {
                planter.Connector = new Connector(options.Wmi, planter);

                // We can use && since one will always start as null
                if (planter.Connector.ConnectedCimSession == null && planter.Connector.ConnectedWmiSession == null)
                {
                    options.Wmi = !options.Wmi;
                    Messenger.ErrorMessage("[-] Issue with using the selected protocol, falling back to the other");
                    planter.Connector = new Connector(options.Wmi, planter);
                   
                    if (planter.Connector.ConnectedCimSession == null && planter.Connector.ConnectedWmiSession == null)
                    {
                        Messenger.ErrorMessage("[-] ERROR: Unable to connect using either CIM or WMI.");
                        System.Environment.Exit(1);
                    }
                }
            }

            catch (COMException)
            {
                Messenger.ErrorMessage("\n[-] ERROR: Cannot connect to remote system, due to firewall or it being offline!");
                System.Environment.Exit(1);
            }

            catch (UnauthorizedAccessException e)
            {
                Messenger.ErrorMessage("\n[-] ERROR: Access is denied, check the account you are using!");
                Console.WriteLine(e);
                System.Environment.Exit(1);
            }

            catch (ManagementException e)
            {
                Messenger.ErrorMessage($"\n[-] ERROR: {e.Message}");
                Console.WriteLine(e);
                System.Environment.Exit(1);
            }

            catch (Exception e)
            {
                Messenger.ErrorMessage("\n[-] ERROR: Something else went wrong, try a different protocol maybe");
                Console.WriteLine(e);
                System.Environment.Exit(1);
            }

            //////////
            // Block to reset the DebugFilePath if needed (generally only for testing)
            //////////
            if (commander.Reset)
            {
                if (!string.IsNullOrEmpty(options.Command))
                {
                    Console.WriteLine("Please don't specify -r/--reset with -c/--command, the reset is redundant");
                    System.Environment.Exit(0);
                }

                try
                {
                    if (!options.Wmi)
                        ExecuteCim.SetOsRecovery(planter.Connector.ConnectedCimSession, defaultDebugFilePath);
                    else
                        ExecuteWmi.SetOsRecovery(planter.Connector.ConnectedWmiSession, defaultDebugFilePath);
                    Console.WriteLine("\nDebugFilePath set back to the default Windows value\n");
                    System.Environment.Exit(0);
                }

                catch (RektDebugFilePath)
                {
                    // Good Sir or Madame reading this,
                    // This only happens if something goes really, really wrong when using CIM.
                    // We need to use WMI to reset the DebugFilePath if it's too large (above 512KB) unless we want to go through
                    // A ton of mods to increase the maxEnvelopeSize within an administrative console
                    // This should rarely happen but there's really no way around it
                    Messenger.WarningMessage("[*] Something bad happened when resetting the DebugFilePath property. Using 'sudo'...");
                    try
                    {
                        planter.Connector.ConnectedCimSession.Close();
                    }
                    catch
                    {
                        //pass
                    }

                    planter.Connector = new Connector(true, planter);
                    ExecuteWmi.SetOsRecovery(planter.Connector.ConnectedWmiSession, defaultDebugFilePath);
                    System.Environment.Exit(0);
                }

                catch (Exception e)
                {
                    Messenger.ErrorMessage("[-] Issue resetting DebugFilePath\n\n");
                    Console.WriteLine(e);
                }
            }

            // We'll want this for DG check
            //if (options.Wmi == true)
            //{
            //    // Unpack the Tuple
            //    wmiScope = planter.Connector.ConnectedWmiSession.Item1;
            //    dgScope = planter.Connector.ConnectedWmiSession.Item2;
            //}

            // Block to check for the existence of Device Guard. If enabled, use PowerShell until I can find a better solution
            // If not enabled, install a WMI provider and use that method. This can be forced with the --provider flag
            //bool credguard = false;
            //if (options.Provider == false)
            //{
            //    try
            //    {
            //        credguard = !options.Wmi == true ? GetDeviceGuard.CheckDgCim(planter.Connector.ConnectedCimSession) : GetDeviceGuard.CheckDgWmi(dgScope, planter.System);
            //    }
            //    catch
            //    {
            //        Messenger.ErrorMessage("[-] Error when grabbing Device Guard info");
            //    }
            //}

            if (planter.Commander.Command != null)
                Messenger.GoodMessage("[+] Results from " + planter.Commander.Command + ":\n");


            ////////
            // Reflection Block
            ////////

            object result = null;

            // Block to set the specific Type for WMI/CIM command
            Type type = options.Wmi ? typeof(ExecuteWmi) : typeof(ExecuteCim);
            MethodInfo method = type.GetMethod(planter.Commander.Method ?? planter.Commander.Command);

            // Create an instance of the type
            object instance = Activator.CreateInstance(type);
            
            // Create parameter object
            object[] stringMethodParams = { planter };

            try
            {
                result = method.Invoke(instance, stringMethodParams);
            }

            catch (TargetInvocationException e)
            {
                if (e.InnerException.Message == planter.Commander.Service)
                {
                    Messenger.ErrorMessage(
                        $"[-] Error: The service name {planter.Commander.Service} not valid, please ensure it's a valid service name (case sensitive)");
                    System.Environment.Exit(1);
                }
            }

            catch (TimeoutException)
            {
                Console.WriteLine("timeout hit");
            }

            catch (PropertyNotFoundException)
            {
                Messenger.ErrorMessage("[-] Registry key does not exist or another issue occurred");
            }

            catch (FormatException e)
            {
                Messenger.ErrorMessage("[-] The registry value for subkey " + planter.Commander.RegSubKey +
                                       " is not in the correct format\n");
                Console.WriteLine("Full error:\n" + e);
            }

            catch (RektDebugFilePath)
            {
                // Good Sir or Madame reading this,
                // This only happens if something goes really, really wrong when using CIM.
                // We need to use WMI to reset the DebugFilePath if it's too large (above 512KB) unless we want to go through
                // A ton of mods to increase the maxEnvelopeSize within an administrative console
                // This should rarely happen but there's really no way around it
                Messenger.WarningMessage("[*] Something bad happened when resetting the DebugFilePath property. Using 'sudo'...");
                try
                {
                    planter.Connector.ConnectedCimSession.Close();
                }
                catch
                {
                    //pass
                }

                planter.Connector = new Connector(true, planter);
                ExecuteWmi.SetOsRecovery(planter.Connector.ConnectedWmiSession, defaultDebugFilePath);
                Console.WriteLine("Successfully reset the DebugFilePath property");
                System.Environment.Exit(0);
            }

            catch (PSRemotingTransportException)
            {
                // Pass, but we already caught above
            }

            catch (CaughtByAvException e)
            {
                Messenger.ErrorMessage("[-] Error: Issues with PowerShell script, it may have been flagged by AV");
                Console.WriteLine(e);
            }

            catch (Exception e)
            {
                //Console.WriteLine("Exception {0} Trace {1}", e.Message, e.StackTrace);
                Console.WriteLine(e);
            }

            ////////
            // End Reflection Block
            ////////

            if (result == null)
            {
                Messenger.ErrorMessage("\n[-] Issue running command after connecting");
                return;
            }

            // Just in case, close the CIM session
            if (!options.Wmi)
            {
                planter.Connector.ConnectedCimSession?.Close();
            }
            else
            {
                planter.Connector.ConnectedWmiSession = null;
            }

            Messenger.GoodMessage("\n\n[+] Successfully completed " + options.Command + " command");
            watch.Stop();
            Console.WriteLine("Execution time: " + watch.ElapsedMilliseconds / 1000 + " Seconds");
        }
    }
}
