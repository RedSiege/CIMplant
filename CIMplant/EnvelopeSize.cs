using System;
using CIMplant;
using Microsoft.Management.Infrastructure;
using Microsoft.Win32;

namespace Execute
{
    public class EnvelopeSize
    {
        public static void GetLocalMaxEnvelopeSize(int envelopeSize)
        {
            Messenger.WarningMessage("[*] Setting the MaxEnvelopeSizeKB on the local system to " + envelopeSize);

            try
            {
                using (RegistryKey key =
                    Registry.CurrentUser.CreateSubKey(@"SOFTWARE\Microsoft\Windows\CurrentVersion\WSMAN\Client"))
                {
                    key.SetValue("maxEnvelopeSize", Convert.ToUInt32(envelopeSize), RegistryValueKind.DWord);
                }
            }
            catch (Exception e)
            {
                Messenger.ErrorMessage(
                    $"[-] Error: Unable to create local runspace to change maxEnvelopeSizeKB.\n");
                Console.WriteLine(e);
            }
        }

        public static void GetMaxEnvelopeSize(string envelopeSize, CimSession cimSession)
        {
            Messenger.WarningMessage("[*] Setting the MaxEnvelopeSizeKB on the remote system to " + envelopeSize);

            CimMethodResult result = RegistryMod.SetRegistryCim(regMethod: "SetDWORDValue", defKey: 0x80000002,
                regSubKey: @"SOFTWARE\Microsoft\Windows\CurrentVersion\WSMAN\Client",
                regSubKeyValue: "maxEnvelopeSize", data: envelopeSize, cimSession);
            if (Convert.ToUInt32(result.ReturnValue.Value.ToString()) == 0)
            {
            }
            else
            {
                Console.WriteLine("Issues setting maxEnvelopeSize");
            }
        }

        public static void SetLocalMaxEnvelopeSize(int envelopeSize)
        {
            Messenger.WarningMessage("[*] Setting the MaxEnvelopeSizeKB on the local system to " + envelopeSize);

            try
            {
                using (RegistryKey key =
                    Registry.CurrentUser.CreateSubKey(@"SOFTWARE\Microsoft\Windows\CurrentVersion\WSMAN\Client"))
                {
                    key.SetValue("maxEnvelopeSize", Convert.ToUInt32(envelopeSize), RegistryValueKind.DWord);
                }
            }
            catch (Exception e)
            {
                Messenger.ErrorMessage(
                    $"[-] Error: Unable to create local runspace to change maxEnvelopeSizeKB.\n");
                Console.WriteLine(e);
            }
        }

        public static void SetMaxEnvelopeSize(string envelopeSize, CimSession cimSession)
        {
            Messenger.WarningMessage("[*] Setting the MaxEnvelopeSizeKB on the remote system to " + envelopeSize);

            CimMethodResult result = RegistryMod.SetRegistryCim(regMethod: "SetDWORDValue", defKey: 0x80000002,
                regSubKey: @"SOFTWARE\Microsoft\Windows\CurrentVersion\WSMAN\Client",
                regSubKeyValue: "maxEnvelopeSize", data: envelopeSize, cimSession);
            if (Convert.ToUInt32(result.ReturnValue.Value.ToString()) == 0)
            {
            }
            else
            {
                Console.WriteLine("Issues setting maxEnvelopeSize");
            }
        }
    }
}