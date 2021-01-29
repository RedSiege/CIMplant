using System;
using Microsoft.Management.Infrastructure;
using Microsoft.Management.Infrastructure.Options;
using System.Management;
using System.Security;

namespace CIMplant
{
    public class Connector
    {
        public CimSession ConnectedCimSession;
        public ManagementScope ConnectedWmiSession;
        public string SystemToConn { get; set; }
        public string Username { get; set; }
        public SecureString Password { get; set; }
        public string Domain { get; set; }

        public Connector()
        {

        }

        public Connector(bool wmi, Planter planter)
        {
            if (wmi)
            {
                this.ConnectedWmiSession = DoWmiConnection(planter);
            }
            else
            {
                this.ConnectedCimSession = DoCimConnection(planter);
            }
        }

        private CimSession DoCimConnection(Planter planter)
        {
            //Block for connecting to the remote system and returning a CimSession object
            SystemToConn = planter.System;
            Domain = planter.Domain;
            Username = planter.User;
            Password = planter.Password;
            WSManSessionOptions sessionOptions = new WSManSessionOptions();
            CimSession connectedCimSession;

            if (SystemToConn == null)
                SystemToConn = "localhost";
            if (Username == null)
                Username = Environment.UserName;

            switch (SystemToConn)
            {
                case "127.0.0.1":
                case "localhost":
                    Messenger.GoodMessage("[+] Connecting to local CIM instance using " + Username + "...");
                    break;
                default:
                    Messenger.GoodMessage("[+] Connecting to remote CIM instance using " + Username + "...");
                    break;
            }

            if (!string.IsNullOrEmpty(Password?.ToString()))
            {
                // create Credentials
                CimCredential credentials = new CimCredential(PasswordAuthenticationMechanism.Default, Domain, Username, Password);
                sessionOptions.AddDestinationCredentials(credentials);
                sessionOptions.MaxEnvelopeSize = 256000; // Not sure how else to get around this
                connectedCimSession = CimSession.Create(SystemToConn, sessionOptions);
            }

            else
            {
                DComSessionOptions options = new DComSessionOptions {Impersonation = ImpersonationType.Impersonate};
                connectedCimSession = CimSession.Create(SystemToConn, options);
            }

            // Test connection to make sure we're connected
            if (!connectedCimSession.TestConnection())
                return null;

            Messenger.GoodMessage("[+] Connected\n");
            return connectedCimSession;
        }

        //private Tuple<ManagementScope, ManagementScope> DoWmiConnection(Planter planter)
        private ManagementScope DoWmiConnection(Planter planter)
        {
            //Block for connecting to the remote system and returning a ManagementScope object
            SystemToConn = planter.System;
            Domain = planter.Domain;
            Username = planter.User;
            Password = planter.Password;

            ConnectionOptions options = new ConnectionOptions();

            if (SystemToConn == null)
                SystemToConn = "localhost";
            if (Username == null)
                Username = Environment.UserName;

            switch (SystemToConn)
            {
                case "127.0.0.1":
                case "localhost":
                    Messenger.GoodMessage("[+] Connecting to local WMI instance using " + Username + "...");
                    break;
                default:
                    Messenger.GoodMessage("[+] Connecting to remote WMI instance using " + Username + "...");
                    break;
            }

            if (!string.IsNullOrEmpty(Password?.ToString()))
            {
                options.Username = Username;
                options.SecurePassword = Password;
                options.Authority = "ntlmdomain:" + Domain;
                options.Impersonation = ImpersonationLevel.Impersonate;
                options.EnablePrivileges = true; // This may be ok for all or may not, need to verify

            }
            else 
            {
                options.Impersonation = ImpersonationLevel.Impersonate;
                options.EnablePrivileges = true;
            }

            ManagementScope scope = new ManagementScope(@"\\" + SystemToConn + @"\root\cimv2", options);
            //ManagementScope deviceguard = new ManagementScope(@"\\" + System + @"\root\Microsoft\Windows\DeviceGuard", options);
            // Need to create a second MS object since we use a separate namespace. 
            //! Need to find a more elegant solution to this!

            scope.Connect();
            //deviceguard.Connect();

            // We'll need this when we get the provider going so we can check for DG
            //ManagementScope deviceScope = scope.Clone();
            //deviceScope.Path = new ManagementPath(@"\\" + System + @"\root\Microsoft\Windows\DeviceGuard");
            //if (GetDeviceGuard.CheckDgWmi(scope, planter.System))
            //    Console.WriteLine("deviceguard enabled");

            Messenger.GoodMessage("[+] Connected\n");
            //return Tuple.Create(scope, deviceguard);
            return scope;
        }
    }
}