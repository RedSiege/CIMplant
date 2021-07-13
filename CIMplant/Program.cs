﻿using CommandLine;
using CommandLine.Text;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Management;
using System.Reflection;
using System.Runtime.InteropServices;
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


        private static void RunTestCases(Commander.Options options)
        {
            Commander commander = new Commander();
            var connector = new Connector();
            var planter = new Planter(commander, connector);
            planter.Connector = new Connector(options.Wmi, planter);

            try
            {
                foreach (var command in Commander.CommandArray)
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

            ParserResult<Commander.Options> parserResult = parser.ParseArguments<Commander.Options>(args);
            parserResult.WithParsed<Commander.Options>(o => { Commander.Options.Instance = o; })
                .WithNotParsed(errs => Commander.DisplayHelp(parserResult, errs));
            Commander.Options options = Commander.Options.Instance;
            Console.WriteLine();

            if (!options.NoBanner)
                Messenger.ImportantAsciiArt();

            // Print all commands before doing anything if that's what the user wants
            if (options.ShowCommands)
                Messenger.GetCommands();

            if (options.ShowExamples)
                Messenger.GetExamples();

            if (options.Test)
            {
                Console.WriteLine("Test method not currently supported");
                //RunTestCases(options);
                //Console.WriteLine("Test cases completed");
                System.Environment.Exit(0);
            }

            //////////
            // Block to instantiate the Commander class (houses all command information and checks for required vals)
            //////////
            Commander commander = new Commander();
            if (options.Command != null && Commander.CommandArray.Any(options.Command.ToLower().Contains) || options.Reset)
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

                if (!options.Wmi) //using CIM
                {

                }

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
                else if (e.InnerException.Message == "System.Management.Automation.PropertyNotFoundException")
                {
                    Messenger.ErrorMessage("[-] Registry key does not exist or another issue occurred");
                    System.Environment.Exit(1);
                }
            }

            catch (TimeoutException)
            {
                Console.WriteLine("timeout hit");
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
