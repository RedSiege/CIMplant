using System;
using System.Linq;

namespace CIMplant
{
    public class Commander
    {
        public string Command, Execute, File, Cmdlet, FileTo, Directory, RegKey, RegSubKey,
            RegVal, RegValType, Service, ServiceBin, Method, Process;
        public bool Reset;
        
        private readonly string[] _shutdown = { "logoff", "reboot", "restart", "power_off", "shutdown" };
        private readonly string[] _fileCommand = { "cat", "copy", "download", "ls", "search", "upload" };

        private readonly string[] _lateralMovement = { "command_exec", "disable_wdigest", "enable_wdigest", "disable_winrm", "enable_winrm",
            "remote_posh", "sched_job", "service_mod" };

        private readonly string[] _registryModify = { "reg_mod", "reg_create", "reg_delete" };
        private readonly string[] _serviceSubCommand = { "list", "start", "stop", "create", "delete" };
        private readonly string[] _processCommand = { "ps", "process_kill", "process_start" };

        public Commander()
        {
            this.RegKey = Program.Options.Instance.RegistryKey;
            this.RegSubKey = Program.Options.Instance.RegistrySubkey;
            this.RegVal = Program.Options.Instance.RegistryValue;
            this.Execute = Program.Options.Instance.Execute;
            this.RegValType = Program.Options.Instance.RegistryValueType;
            this.Service = Program.Options.Instance.Service;
            this.ServiceBin = Program.Options.Instance.ServiceBin;
            this.Cmdlet = Program.Options.Instance.Cmdlet;
            this.File = Program.Options.Instance.File;
            this.FileTo = Program.Options.Instance.FileTo;
            this.Directory = Program.Options.Instance.Directory;
            this.Reset = Program.Options.Instance.Reset;
            this.Process = Program.Options.Instance.Process;
            this.Method = null;
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
    }
}