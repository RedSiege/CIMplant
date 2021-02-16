<p align="center">
  <img width="360" height="340" src="https://github.com/FortyNorthSecurity/CIMplant/raw/main/Extras/cimplant_logo_letters.png">
</p>

# CIMplant

C# port of WMImplant which uses either CIM or WMI to query remote systems. It can use provided credentials or the current user's session.

Note: Some commands will use PowerShell in combination with WMI, denoted with ** in the `--show-commands` command.

## Introduction

CIMplant is a C# rewrite and expansion on [@christruncer](https://twitter.com/christruncer)'s [WMImplant](https://github.com/FortyNorthSecurity/WMImplant). It allows you to gather data about a remote system, execute commands, exfil data, and more. The tool allows connections using Windows Management Instrumentation, [WMI](https://docs.microsoft.com/en-us/windows/win32/wmisdk/about-wmi), or Common Interface Model, [CIM](https://www.dmtf.org/standards/cim) ; well more accurately Windows Management Infrastructure, [MI](https://docs.microsoft.com/en-us/previous-versions/windows/desktop/wmi_v2/windows-management-infrastructure). CIMplant requires local administrator permissions on the target system.

## Setup:

It's probably easiest to use the built version under Releases, just note that it is compiled in Debug mode. If you want to build the solution yourself, follow the steps below.

1. Load CIMplant.sln into Visual Studio
2. Go to Build at the top and then Build Solution if no modifications are wanted

## Usage

```
CIMplant.exe --help
CIMplant.exe --show-commands
CIMplant.exe --show-examples
CIMplant.exe -s [remote IP address] -c cat -f c:\users\user\desktop\file.txt
CIMplant.exe -s [remote IP address] -u [username] -d [domain] -p [password] -c cat -f c:\users\test\desktop\file.txt
CIMplant.exe -s [remote IP address] -u [username] -d [domain] -p [password] -c command_exec --execute "dir c:\\"
```
### Some Helpful Commands

![image](https://github.com/FortyNorthSecurity/CIMplant/raw/main/Extras/CIMplant-Help.gif)

### Some Example Usage Commands

![image](https://github.com/FortyNorthSecurity/CIMplant/raw/main/Extras/CIMplant-Usage.gif)

### Cobalt Strike Execute-Assembly

I wanted to code CIMplant in a way that would allow usage through execute-assembly so everything is packed into one executable and loaded reflectively. You should be able to run all commands through beacon without issue. Enjoy!

![image](https://github.com/FortyNorthSecurity/CIMplant/raw/main/Extras/CIMplant-CS-Usage.gif)

## Important Files

1. Program.cs
> This is the brains of the operation, the driver for the program.

2. Connector.cs
> This is where the initial CIM/WMI connections are made and passed to the rest of the application

2. ExecuteWMI.cs
> All function code for the WMI commands

3. ExecuteCIM.cs
> All function code for the CIM (MI) commands

## Detection

Of course, the first thing we'll want to be aware of is the initial WMI or CIM connection. In general, WMI uses DCOM as a communication protocol whereas CIM uses WSMan (or, WinRM). This can be modified for CIM, and is in CIMplant, but let's just go over the default values for now. For DCOM, the first thing we can do is look for initial TCP connections over **port 135**. The connecting and receiving systems will then decide on a new, very high port to use so that will vary drastically. For WSMan, the initial TCP connection is over **port 5985**.

Next, you'll want to look at the Microsoft-Windows-WMI-Activity/Trace event log in the Event Viewer. Search for **Event ID 11** and filter on the IsLocal property if possible. You can also look for **Event ID 1295** within the Microsoft-Windows-WinRM/Analytic log.

Finally, you'll want to look for any modifications to the **DebugFilePath** property with the **Win32_OSRecoveryConfiguration** class. More detailed information about detection can be found at Part 1 of our blog series here: [CIMplant Part 1: Detection of a C# Implementation of WMImplant](https://fortynorthsecurity.com/blog/cimplant-part-1-detections/)
