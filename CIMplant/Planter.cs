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
            this.System = Program.Options.Instance.System;
            this.Domain = Program.Options.Instance.Domain;
            this.User = Program.Options.Instance.Username;
            this.Password = CreateSecuredString(Program.Options.Instance.Password);
            this.NameSpace = Program.Options.Instance.NameSpace;
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