using System;
using System.Collections.Generic;
using System.IO;

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
                        "download - Download a file from the targeted machine",
                        "ls - File/Directory listing of a specific directory",
                        "search - Search for a file on a user-specified drive",
                        "upload - Upload a file to the targeted machine"
                    }
                },
                {"Lateral Movement Facilitation", new string[] {"command_exec - Run a command line command and receive the output",
                        "disable_wdigest - Sets the registry value for UseLogonCredential to zero",
                        "enable_wdigest - Adds registry value UseLogonCredential",
                        "disable_winrm - Disables WinRM on the targeted system",
                        "enable_winrm - Enables WinRM on the targeted system",
                        "reg_mod - Modify the registry on the targeted machine",
                        "reg_create - Create the registry value on the targeted machine",
                        "reg_delete - Delete the registry on the targeted machine",
                        "remote_posh - Run a PowerShell script on a remote machine and receive the output",
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
                        "vacant_system - Determine if a user is away from the system"
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
                        @"vacant_system , -c vacant_system , -c vacant_system -s 192.168.64.4 -u test -p 1"
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

    }
}