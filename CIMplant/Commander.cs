using System;
using System.Collections.Generic;
using System.Linq;
using CommandLine;
using CommandLine.Text;

namespace CIMplant
{
    public class Commander
    {
        public string Command, Execute, File, Cmdlet, FileTo, Directory, RegKey, RegSubKey,
            RegVal, RegValType, Service, ServiceBin, Method, Process;
        public bool Reset, NoPS;
        
        private readonly string[] _shutdown = { "logoff", "reboot", "restart", "power_off", "shutdown" };
        private readonly string[] _fileCommand = { "cat", "copy", "download", "ls", "search", "upload" };

        private readonly string[] _lateralMovement = { "command_exec", "disable_wdigest", "enable_wdigest", "disable_winrm", "enable_winrm",
            "remote_posh", "sched_job", "service_mod" };

        private readonly string[] _registryModify = { "reg_mod", "reg_create", "reg_delete" };
        private readonly string[] _serviceSubCommand = { "list", "start", "stop", "create", "delete" };
        private readonly string[] _processCommand = { "ps", "process_kill", "process_start" };

        public Commander()
        {
            this.RegKey = Options.Instance.RegistryKey;
            this.RegSubKey = Options.Instance.RegistrySubkey;
            this.RegVal = Options.Instance.RegistryValue;
            this.Execute = Options.Instance.Execute;
            this.RegValType = Options.Instance.RegistryValueType;
            this.Service = Options.Instance.Service;
            this.ServiceBin = Options.Instance.ServiceBin;
            this.Cmdlet = Options.Instance.Cmdlet;
            this.File = Options.Instance.File;
            this.FileTo = Options.Instance.FileTo;
            this.Directory = Options.Instance.Directory;
            this.Reset = Options.Instance.Reset;
            this.Process = Options.Instance.Process;
            this.Method = null;
            this.NoPS = Options.Instance.NoPS;
        }

        public Commander(string command)
        : this()
        {
            this.Command = command;
            ParseCommands();
        }

        private void ParseCommands()
        {
            if (Command == "cat")
                _ = File ?? throw new ArgumentNullException(nameof(File));

            if (_shutdown.Any(Command.Contains))
            {
                //This method is the same for all 3 types of commands in the shutdown array, so use this one and pass it the command
                this.Method = "Win32Shutdown";
            }

            if (_registryModify.Any(Command.Contains))
            {
                this.Method = "registry_mod";
                switch (Command)
                {
                    case "reg_mod" when RegKey != null &&
                                        RegSubKey != null &&
                                        RegVal != null:
                        Method = "registry_mod";
                        break;
                    case "reg_delete" when RegKey != null &&
                                           RegSubKey != null:
                        Method = "registry_mod";
                        break;
                    case "reg_create" when RegKey != null &&
                                           RegSubKey != null &&
                                           RegVal != null &&
                                           RegValType != null:
                        Method = "registry_mod";
                        break;
                    default:
                    {
                        _ = RegKey ?? throw new ArgumentNullException(nameof(RegKey));
                        _ = RegSubKey ?? throw new ArgumentNullException(nameof(RegSubKey));
                        _ = RegVal ?? throw new ArgumentNullException(nameof(RegVal));
                        _ = RegValType ?? throw new ArgumentNullException(nameof(RegValType));
                    }
                        break;
                }
            }

            if (Command == "service_mod")
            {
                // All commands need this val so check first
                _ = Execute ?? throw new ArgumentNullException(nameof(Execute));

                if (_serviceSubCommand.Any(Execute.Contains))
                {
                    switch (Execute)
                    {
                        case "list":
                            break;
                        case "create" when Service != null &&
                                           ServiceBin != null:
                            break;
                        case "start" when Service != null:
                            break;
                        case "delete" when Service != null:
                            break;
                        default:
                        {
                            _ = Service ?? throw new ArgumentNullException(nameof(Service));
                            _ = ServiceBin ?? throw new ArgumentNullException(nameof(ServiceBin));
                        }
                            break;
                    }
                }
                else
                {
                    throw new SubCommandException("service_mod");
                }
            }

            if (_processCommand.Any(Command.Contains))
            {
                switch (Command)
                {
                    case "ps":
                        break;
                    case "process_kill" when Process != null:
                        break;
                    case "process_start" when Process != null:
                        break;
                    default:
                    {
                        _ = Process ?? throw new ArgumentNullException(nameof(Process));
                    }
                        break;
                }
            }

            if (_fileCommand.Any(Command.Contains))
            {
                switch (Command)
                {
                    case "ls":
                        _ = Directory ?? throw new ArgumentNullException(nameof(Directory));
                        break;
                    case "cat" when File != null:
                        break;
                    case "copy" when File != null &&
                                     FileTo != null:
                        break;
                    case "download" when File != null:
                        break;
                    case "search" when File != null:
                        _ = Directory ?? throw new ArgumentNullException(nameof(Directory));
                        break;
                    case "upload" when File != null &&
                                     FileTo != null:
                        break;
                    default:
                    {
                        _ = File ?? throw new ArgumentNullException(nameof(File));
                        _ = FileTo ?? throw new ArgumentNullException(nameof(FileTo));
                        _ = Directory ?? throw new ArgumentNullException(nameof(Directory));
                    } break;
                }
            }

            if (_lateralMovement.Any(Command.Contains))
            {
                switch (Command)
                {
                    case "command_exec" when Execute != null:
                        break;
                    case "remote_posh":
                        _ = File ?? throw new ArgumentNullException(nameof(File));
                        break;
                    case "disable_wdigest":
                    case "enable_wdigest":
                    case "disable_winrm":
                    case "enable_winrm":
                        break;
                    default:
                        _ = Execute ?? throw new ArgumentNullException(nameof(Execute));
                        break;
                }
            }
        }

        public static readonly string[] CommandArray =
        {
            "cat", "copy", "download", "ls", "search", "upload", "command_exec", "process_kill", "process_start",
            "ps", "active_users", "basic_info", "drive_list", "ifconfig", "installed_programs", "logoff", "reboot",
            "restart", "power_off", "shutdown",
            "vacant_system", "logon_events", "command_exec", "disable_wdigest", "enable_wdigest", "disable_winrm",
            "enable_winrm",
            "reg_mod", "reg_create", "reg_delete", "remote_posh", "sched_job", "service_mod", "edr_query"
        };

        public static void DisplayHelp<T>(ParserResult<T> result, IEnumerable<Error> errs)
        {
            HelpText helpText = HelpText.AutoBuild(result, h =>
            {
                h.AdditionalNewLineAfterOption = false;
                h.Heading = "WMI C# Version 0.2"; //change header
                h.Copyright = ""; //change copyright text
                h.AutoVersion = false;
                return HelpText.DefaultParsingErrorsHandler(result, h);
            }, e => e);
            Console.WriteLine(helpText);
            System.Environment.Exit(1);
        }

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
           
            [Option("nops", Required = false, HelpText = "Do not allow any PowerShell execution (will die before)",
                Default = false)]
            public bool NoPS { get; set; }

            [Option("show-commands", Group = "Command", Required = true,
                HelpText = "Displays a list of available commands")]
            public bool ShowCommands { get; set; }

            [Option("show-examples", Group = "Command", Required = true,
                HelpText = "Displays examples for all available commands")]
            public bool ShowExamples { get; set; }

            [Option("no-banner", Group = "Command", Required = false,
                HelpText = "Disables that gorgeous ASCII art (probably should never use this)", Default = false)]
            public bool NoBanner { get; set; }

            [Option("test", Group = "Command", Required = false,
                HelpText = "Tests all commands with a specified username/password/system (or against the localhost)")]
            public bool Test { get; set; }
        }
    }
}