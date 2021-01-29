using System;
using System.Collections.Generic;
using Microsoft.Management.Infrastructure;
using System.Management;

namespace CIMplant
{
    public class GetDeviceGuard
    {
        public static bool CheckDgCim(CimSession cimSession)
        {
            // Use this method to check for the existence of Device Guard on a system. This isn't a perfect remedy but should work just fine

            var osQuery = @"SELECT * FROM Win32_DeviceGuard";
            IEnumerable<CimInstance> queryInstance =
                cimSession.QueryInstances(@"root\Microsoft\Windows\DeviceGuard", "WQL", osQuery);

            foreach (CimInstance cimObject in queryInstance)
            {
                if (Convert.ToInt32(cimObject.CimInstanceProperties["CodeIntegrityPolicyEnforcementStatus"].Value) ==
                    2 || Convert.ToInt32(cimObject.CimInstanceProperties["UsermodeCodeIntegrityPolicyEnforcementStatus"]
                        .Value) == 2)
                {
                    Messenger.ErrorMessage("[-] Device Guard enabled! Using PowerShell with Constrained Language Mode\n");
                    return true;
                }
            }

            return false;
        }

        public static bool CheckDgWmi(ManagementScope wmiSession)
        {
            // Use this method to check for the existence of Device Guard on a system. This isn't a perfect remedy but should work just fine

            ObjectQuery query = new ObjectQuery(@"SELECT * FROM Win32_DeviceGuard");
            ManagementObjectSearcher searcher = new ManagementObjectSearcher(wmiSession, query);
            ManagementObjectCollection queryCollection = searcher.Get();

            foreach (var o in queryCollection)
            {
                var queryObj = (ManagementObject)o;

                //foreach (var item in (dynamic)((queryObj["AvailableSecurityProperties"])))
                //{
                //    Console.WriteLine(item.ToString());
                //}

                if (Convert.ToInt32(queryObj["CodeIntegrityPolicyEnforcementStatus"]) == 2 ||
                    Convert.ToInt32(queryObj["UsermodeCodeIntegrityPolicyEnforcementStatus"]) == 2)
                {
                    Messenger.ErrorMessage("[-] Device Guard enabled! Using PowerShell with Constrained Language Mode\n");
                    return true;
                }
            }

            return false;
        }
    }
}