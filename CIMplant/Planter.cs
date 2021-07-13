using System.Security;

namespace CIMplant
{
    public class Planter
    {
        public string System, NameSpace, Domain, User;
        public SecureString Password;
        public Commander Commander;
        public Connector Connector;

        public Planter()
        {
            this.System = Commander.Options.Instance.System;
            this.Domain = Commander.Options.Instance.Domain;
            this.User = Commander.Options.Instance.Username;
            this.Password = CreateSecuredString(Commander.Options.Instance.Password);
            this.NameSpace = Commander.Options.Instance.NameSpace;
        }

        public Planter(Commander commander, Connector connector)
            : this()
        {
            this.Commander = commander;
            this.Connector = connector;
        }

        public static SecureString CreateSecuredString(string pw)
        {
            SecureString secureString = new SecureString();
            if (string.IsNullOrEmpty(pw))
                return null;
            foreach (char c in pw)
                secureString.AppendChar(c);
            return secureString;
        }
    }
}