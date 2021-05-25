using System;
using System.Management;
using System.Reflection;
using CIMplant;
using Microsoft.Management.Infrastructure;

namespace Execute
{
    public class RegistryMod
    {
        public static string Namespace = @"root\cimv2";
        public static CimMethodResult CheckRegistryCim(string regMethod, uint defKey, string regSubKey, string regSubKeyValue, CimSession cs)
        {
            // Block to be used to check the registry for specific values (before modifying or deleting)
            CimMethodParametersCollection cimParams = new CimMethodParametersCollection
            {
                CimMethodParameter.Create("hDefKey", defKey, CimFlags.In),
                CimMethodParameter.Create("sSubKeyName", regSubKey, CimFlags.In),
                CimMethodParameter.Create("sValueName", regSubKeyValue, CimFlags.In)
            };

            CimMethodResult results = cs.InvokeMethod(new CimInstance("StdRegProv", Namespace), regMethod, cimParams);

            return results;
        }

        public static string CheckRegistryTypeCim(uint defKey, string regSubKey, string regSubKeyValue, CimSession cs)
        {
            // Block to be used to check the registry for specific values (before modifying or deleting)
            const int REG_SZ = 1;
            const int REG_EXPAND_SZ = 2;
            const int REG_BINARY = 3;
            const int REG_DWORD = 4;
            const int REG_MULTI_SZ = 7;
            int type;

            CimMethodParametersCollection cimParams = new CimMethodParametersCollection
            {
                CimMethodParameter.Create("hDefKey", defKey, CimFlags.In),
                CimMethodParameter.Create("sSubKeyName", regSubKey, CimFlags.In),
            };

            CimMethodResult results = cs.InvokeMethod(new CimInstance("StdRegProv", Namespace), "EnumValues", cimParams);

            //Hacky way to get the type from the returned arrays
            try
            {
                type = ((int[])results.OutParameters["Types"].Value)[Array.IndexOf((string[])results.OutParameters["sNames"].Value, regSubKeyValue)];
            }
            catch (TargetInvocationException e)
            {
                Console.WriteLine(e);
                throw;
            }

            switch (type)
            {
                case REG_SZ:
                    return "REG_SZ";
                case REG_EXPAND_SZ:
                    return "REG_EXPAND_SZ";
                case REG_BINARY:
                    return "REG_BINARY";
                case REG_DWORD:
                    return "REG_DWORD";
                case REG_MULTI_SZ:
                    return "REG_MULTI_SZ";
            }
            return null;
        }

        public static CimMethodResult SetRegistryCim(string regMethod, uint defKey, string regSubKey, string regSubKeyValue, string data, CimSession cs)
        {
            // Block to be used to set the registry for specific values
            CimMethodParametersCollection cimParamsSetReg = new CimMethodParametersCollection
            {
                CimMethodParameter.Create("hDefKey", defKey, CimFlags.In),
                CimMethodParameter.Create("sSubKeyName", regSubKey, CimFlags.In),
                CimMethodParameter.Create("sValueName", regSubKeyValue, CimFlags.In)
            };

            switch (regMethod)
            {
                // Need diff values for different methods
                case "SetStringValue":
                    cimParamsSetReg.Add(CimMethodParameter.Create("sValue", data, CimFlags.In));
                    break;
                case "SetDWORDValue":
                    cimParamsSetReg.Add(CimMethodParameter.Create("uValue", Convert.ToUInt32(data), CimFlags.In));
                    break;
            }

            CimMethodResult results = cs.InvokeMethod(new CimInstance("StdRegProv", Namespace), regMethod, cimParamsSetReg);

            return results;
        }

        public static bool DeleteRegistryCim(uint defKey, string regSubKey, string regSubKeyValue, CimSession cs)
        {
            // Block to be used to delete the registry for specific values

            CimMethodParametersCollection cimParams = new CimMethodParametersCollection
            {
                CimMethodParameter.Create("hDefKey", defKey, CimFlags.In),
                CimMethodParameter.Create("sSubKeyName", regSubKey, CimFlags.In),
                CimMethodParameter.Create("sValueName", regSubKeyValue, CimFlags.In)
            };

            CimMethodResult results = cs.InvokeMethod(new CimInstance("StdRegProv", Namespace), "DeleteValue", cimParams);

            if (Convert.ToUInt32(results.ReturnValue.Value) == 0)
                return true;
            
            Messenger.ErrorMessage("[-] Error deleting key");
            Console.WriteLine("\nFull key provided: " + regSubKey + "\n" + "Value provided: " + regSubKeyValue);
            return false;
        }

        public static ManagementBaseObject CheckRegistryWmi(string regMethod, uint defKey, string regSubKey, string regSubKeyValue, ManagementClass mc)
        {
            // Block to be used to check the registry for specific values (before modifying or deleting)
            ManagementBaseObject inParamsSet = mc.GetMethodParameters(regMethod);
            inParamsSet["hDefKey"] = defKey;
            inParamsSet["sSubKeyName"] = regSubKey;
            inParamsSet["sValueName"] = regSubKeyValue;

            ManagementBaseObject outParamsSet = mc.InvokeMethod(regMethod, inParamsSet, null);

            //if (outParamsSet != null && Convert.ToUInt32(outParamsSet["ReturnValue"]) == 0)
            return outParamsSet;
            
            //Messenger.ErrorMessage("[-] Registry key not valid, not modifying or deleting");
            //Console.WriteLine("\nFull key provided: " + regSubKey + "\n" + "Value provided: " + regSubKeyValue);
            //return false;
        }

        public static string CheckRegistryTypeWmi(uint defKey, string regSubKey, string regSubKeyValue, ManagementClass mc)
        {
            // Block to be used to check the registry for specific values (before modifying or deleting)
            const int regSz = 1;
            const int regExpandSz = 2;
            const int regBinary = 3;
            const int regDword = 4;
            const int regMultiSz = 7;

            // Obtain in-parameters for the method
            ManagementBaseObject inParams = mc.GetMethodParameters("EnumValues");

            // Add the input parameters.
            inParams["hDefKey"] = defKey;
            inParams["sSubKeyName"] = regSubKey;

            // Execute the method and obtain the return values.
            ManagementBaseObject outParams = mc.InvokeMethod("EnumValues", inParams, null);

            //Hacky way to get the type from the returned arrays
            int type = ((int[])outParams["Types"])[Array.IndexOf((string[])outParams.Properties["sNames"].Value, regSubKeyValue)];

            switch (type)
            {
                case regSz:
                    return "REG_SZ";
                case regExpandSz:
                    return "REG_EXPAND_SZ";
                case regBinary:
                    return "REG_BINARY";
                case regDword:
                    return "REG_DWORD";
                case regMultiSz:
                    return "REG_MULTI_SZ";
            }
            return null;
        }

        public static ManagementBaseObject SetRegistryWmi(string regMethod, uint defKey, string regSubKey, string regSubKeyValue, string data, ManagementClass mc)
        {
            // Block to be used to set the registry for specific values
            ManagementBaseObject inParamsSet = mc.GetMethodParameters(regMethod);
            inParamsSet["hDefKey"] = defKey;
            inParamsSet["sSubKeyName"] = regSubKey;
            inParamsSet["sValueName"] = regSubKeyValue;

            switch (regMethod)
            {
                // Need diff values for different methods
                case "SetStringValue":
                    inParamsSet["sValue"] = data;
                    break;
                case "SetDWORDValue":
                    inParamsSet["uValue"] = data;
                    break;
            }

            ManagementBaseObject outParamsSet = mc.InvokeMethod(regMethod, inParamsSet, null);

            //if (outParamsSet != null && Convert.ToUInt32(outParamsSet["ReturnValue"]) == 0)
            return outParamsSet;
            
            //Messenger.ErrorMessage("[-] Error modifying key");
            //Console.WriteLine("\nFull key provided: " + regSubKey + "\n" + "Value provided: " + regSubKeyValue);
            //return false;
        }

        public static bool DeleteRegistryWmi(uint defKey, string regSubKey, string regSubKeyValue, ManagementClass mc)
        {
            // Block to be used to set the registry for specific values
            ManagementBaseObject inParamsSet = mc.GetMethodParameters("DeleteValue");
            inParamsSet["hDefKey"] = defKey;
            inParamsSet["sSubKeyName"] = regSubKey;
            inParamsSet["sValueName"] = regSubKeyValue;

            ManagementBaseObject outParamsSet = mc.InvokeMethod("DeleteValue", inParamsSet, null);

            if (outParamsSet != null && Convert.ToUInt32(outParamsSet["ReturnValue"]) == 0)
                return true;
            
            Messenger.ErrorMessage("[-] Error deleting key");
            Console.WriteLine("\nFull key provided: " + regSubKey + "\n" + "Value provided: " + regSubKeyValue);
            return false;
        }
    }
}