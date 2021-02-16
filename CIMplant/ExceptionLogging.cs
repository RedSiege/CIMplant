using System;
using System.Globalization;
using System.IO;

namespace CIMplant
{
    /// <summary>  
    /// Summary description for ExceptionLogging
    /// Borrowed from https://www.c-sharpcorner.com/UploadFile/0c1bb2/logging-excetion-to-text-file/
    /// </summary>  
    public static class ExceptionLogging
    {

        private static string _errorlineNo, _errormsg, _extype, _exurl, _hostIp, _errorLocation;

        public static void SendErrorToText(Exception ex)
        {
            var line = Environment.NewLine + Environment.NewLine;

            _errorlineNo = ex.StackTrace.Substring(ex.StackTrace.Length - 7, 7);
            _errormsg = ex.GetType().Name;
            _extype = ex.GetType().ToString();
            _errorLocation = ex.Message;

            try
            {
                string filepath = Directory.GetCurrentDirectory();  //Text File Path

                if (!Directory.Exists(filepath))
                    Directory.CreateDirectory(filepath);
                
                filepath = filepath + DateTime.Today.ToString("dd-MM-yy") + ".txt";   //Text File Name
                
                if (!File.Exists(filepath))
                    File.Create(filepath).Dispose();

                using (StreamWriter sw = File.AppendText(filepath))
                {
                    string error = "Log Written Date:" + " " + DateTime.Now.ToString(CultureInfo.InvariantCulture) + line + "Error Line No :" + " " + _errorlineNo + line + "Error Message:" + " " + _errormsg + line + "Exception Type:" + " " + _extype + line + "Error Location :" + " " + _errorLocation + line + " Error Page Url:" + " " + _exurl + line + "User Host IP:" + " " + _hostIp + line;
                    sw.WriteLine("-----------Exception Details on " + " " + DateTime.Now.ToString(CultureInfo.InvariantCulture) + "-----------------");
                    sw.WriteLine("-------------------------------------------------------------------------------------");
                    sw.WriteLine(line);
                    sw.WriteLine(error);
                    sw.WriteLine("--------------------------------*End*------------------------------------------");
                    sw.WriteLine(line);
                    sw.Flush();
                    sw.Close();
                }
            }
            catch (Exception e)
            {
                e.ToString();
            }
        }
    }
}