using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace SqlScriptExec

{
    public partial class Options
    {
        public class Connections
        {
            public static List <string> Servers {get; set;}
            public static string SelectedServer { get; set; }
            public static string ConnectionString { get; set; }
            public static bool WindowsAuthentication { get; set; }
            public static string UserName { get; set; }
            public static string Password { get; set; }
        }

        public class Log
        {
            public static bool UseDefaultLogPath { get; set; }
            public static string LogPath { get; set; }
            public static bool CreateExcelLog { get; set; }
            public static bool LogErrorsOnly { get; set; }
            public static string LogFileName { get; set; }            //TODO: Make ReadOnly?
        }

        public class Execution
        {
            public static bool IncludeSubFolders { get; set; }
            public static bool WarnBeforeRunning { get; set; }
            public static bool ExecuteScriptsUntilFinished { get; set; }
            public static int StopAfterErrorCount { get; set; }
            public static bool ConsecutiveErrors { get; set; }
            //public static string ScriptsToExecutePath { get; set; }

            private static string _ScriptsToExecutePath;
            public static string ScriptsToExecutePath
            {
                get { return _ScriptsToExecutePath; }
                set
                {
                    _ScriptsToExecutePath = value;

                    // Set LogName when Script Path is given
                    string tempLogName = "";
                    string dirName = new DirectoryInfo(_ScriptsToExecutePath.Trim()).Name;

                    if (string.IsNullOrWhiteSpace(_ScriptsToExecutePath))
                    {
                        dirName = "log";
                    }
                    tempLogName = dirName + "_" + DateTime.Now.ToString("yyyyMMddHHmmss");

                    Log.LogFileName = tempLogName;
                }
            }

        }

        public static void SetDefaults()
        {
            //TODO : Set defaults above

            //Connections
            Connections.ConnectionString = "";
            Connections.WindowsAuthentication = true;
            Connections.UserName = "";
            Connections.Password = "";
            
            //Log
            Log.UseDefaultLogPath = true;
            Log.LogPath = "";
            Log.CreateExcelLog = true;
            Log.LogErrorsOnly = false;

            //Execution
            Execution.IncludeSubFolders = true;
            Execution.WarnBeforeRunning = true;
            Execution.ExecuteScriptsUntilFinished = true;
            Execution.StopAfterErrorCount = 10;
            Execution.ConsecutiveErrors = false;
            Execution.ScriptsToExecutePath = "";
        }


        /*-----------------------------------------------------
        Validate Option Values
        -----------------------------------------------------*/
        public static string ValidateOptions(bool runtime)
        {
            string errorMessage = "";
            if (Connections.WindowsAuthentication == false)
            //if (Connections.WindowsAuthentication == false)
            {
                //Blank UserName when WindowsAuth not used
                if (string.IsNullOrWhiteSpace(Connections.UserName))
                {
                    errorMessage += "You must provide UserName\r\n";
                }
                //Blank Password when WindowsAuth not used
                if (string.IsNullOrWhiteSpace(Connections.Password))
                {
                    errorMessage += "You must provide Password\r\n";
                }
            }

            //Blank Log Path
            if (Log.UseDefaultLogPath == false && string.IsNullOrWhiteSpace(Log.LogPath))
            {
                errorMessage += "You must provide a Log Path\r\n";
            }

            //Stop after xx errors recieved
            if (Execution.ExecuteScriptsUntilFinished == false)
            {
                if (Execution.StopAfterErrorCount < 1)
                {
                    errorMessage += "Stop after Error Count is Invalid\r\n";
                }
            }

            //Only evaluate these at runtime
            if (runtime == true)
            {
                //Blank Connection String
                if (string.IsNullOrWhiteSpace(Connections.ConnectionString))
                {
                    errorMessage += "Connection String is blank\r\n";
                }

                //Blank Scripts to Execute Path
                if (string.IsNullOrWhiteSpace(Execution.ScriptsToExecutePath))
                {
                    errorMessage += "Scripts to execute Path is blank\r\n";
                }
            }

            return errorMessage;
        }

     }
}