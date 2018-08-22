using System;
using System.Collections.Generic;
using System.Configuration;
using System.Data;
using System.Data.SqlClient;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using ClosedXML.Excel;

namespace SqlScriptExec
{
    public static class Helper
    {

        /*-----------------------------------------------------
        Load Options Class from Property Settings
        -----------------------------------------------------*/
        public static void LoadOptionValues()
        {
            //These values get entered by the user and are not saved to Properties file (TODO: Possibly save in future.  Encrypt credentials)
            //Options.Connections.UserName
            //Options.Connections.Password
            //Options.Connections.SelectedServer
            //Options.Execution.ScriptsToExecutePath
            Options.Connections.WindowsAuthentication = Properties.Settings.Default.WindowsAuthentication;
            Options.Connections.ConnectionString = Helper.GetConnectionString(Options.Connections.SelectedServer, "master");
            
            Options.Log.UseDefaultLogPath = Properties.Settings.Default.UseDefaultLogPath;
            Options.Log.LogPath = Properties.Settings.Default.LogPath;
            Options.Log.CreateExcelLog = Properties.Settings.Default.CreateExcelLog;
            Options.Log.LogErrorsOnly = Properties.Settings.Default.LogErrorsOnly;

            Options.Execution.IncludeSubFolders = Properties.Settings.Default.IncludeSubFolders;
            Options.Execution.WarnBeforeRunning = Properties.Settings.Default.WarnBeforeRunning;
            Options.Execution.ExecuteScriptsUntilFinished = Properties.Settings.Default.ExecuteScriptsUntilFinished;
            Options.Execution.StopAfterErrorCount = Properties.Settings.Default.StopAfterErrorCount;
            Options.Execution.ConsecutiveErrors = Properties.Settings.Default.ConsecutiveErrors;
        }

        /*-----------------------------------------------------
        SaveProperty - from Options Class
        -----------------------------------------------------*/
        public static void SavePropertySettings()
        {
            Properties.Settings.Default.WindowsAuthentication = Options.Connections.WindowsAuthentication;
            Properties.Settings.Default.UseDefaultLogPath = Options.Log.UseDefaultLogPath;
            Properties.Settings.Default.LogPath = Options.Log.LogPath;
            Properties.Settings.Default.CreateExcelLog = Options.Log.CreateExcelLog;
            Properties.Settings.Default.LogErrorsOnly = Options.Log.LogErrorsOnly;
            Properties.Settings.Default.IncludeSubFolders = Options.Execution.IncludeSubFolders;
            Properties.Settings.Default.WarnBeforeRunning = Options.Execution.WarnBeforeRunning;
            Properties.Settings.Default.ExecuteScriptsUntilFinished = Options.Execution.ExecuteScriptsUntilFinished;
            Properties.Settings.Default.StopAfterErrorCount = Options.Execution.StopAfterErrorCount;
            Properties.Settings.Default.ConsecutiveErrors = Options.Execution.ConsecutiveErrors;

            Properties.Settings.Default.Save();
        }

        /*-----------------------------------------------------
        Reset Default Values
        -----------------------------------------------------*/
        public static void ResetDefaultValues()
        {
            Properties.Settings.Default.Reset();
            LoadOptionValues();
        }


        /*-----------------------------------------------------
        GetConnectionString - wrapper
        -----------------------------------------------------*/
        public static string GetConnectionString(string datasource, string intialCatalog)
        {
            string userName =Options.Connections.UserName;
            string password = Options.Connections.Password;
            string connectionString;

            if (Properties.Settings.Default.WindowsAuthentication == true)
            {
                connectionString = Helper.BuildConnectionString(datasource, intialCatalog);
            }
            else
            {
                //Without Windows Authentication, UserName and Password are required
                if (string.IsNullOrWhiteSpace(userName) || string.IsNullOrWhiteSpace(password))
                {
                    connectionString = "";
                    //MessageBox.Show("UserName and Password are required when not using Windows Authentication.", "Message", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
                else
                {
                    connectionString = Helper.BuildConnectionString(datasource, intialCatalog, userName, password);
                }
            }
            return connectionString;
        }

        /*-----------------------------------------------------
        GetConnectionStringSafeForDisplay - Removes UserName and Password
        -----------------------------------------------------*/
        public static string GetConnectionStringSafeForDisplay(string cn)
        {
            SqlConnectionStringBuilder builder = new SqlConnectionStringBuilder(cn);
            builder.Remove("User ID");
            builder.Remove("Password");
            return builder.ConnectionString;
        }

        /*-----------------------------------------------------
        Build Connection String - NO UserName, Password
        -----------------------------------------------------*/
        private static string BuildConnectionString(string dataSource, string intialCatalog)
        {
            string cn = "";

            // Retrieve the partial connection string named databaseConnection from the application's app.config
            ConnectionStringSettings settings = ConfigurationManager.ConnectionStrings["master"];

            if (null != settings)
            {
                // Retrieve the partial connection string.
                string connectString = settings.ConnectionString;
                //Console.WriteLine("Original: {0}", connectString);

                // Create a new SqlConnectionStringBuilder based on the
                // partial connection string retrieved from the config file.
                SqlConnectionStringBuilder builder = new SqlConnectionStringBuilder(connectString);

                // Remove UnNeeded values
                builder.Remove("User ID");
                builder.Remove("Password");

                // Supply the additional values.
                builder.DataSource = dataSource;
                builder.InitialCatalog = intialCatalog;
                builder.IntegratedSecurity = true;
                //Console.WriteLine("Modified: {0}", builder.ConnectionString);

                cn = builder.ConnectionString;
            }
            return cn;
        }

        /*-----------------------------------------------------
        Build Connection String (Overload) - WITH UserName, Password
        -----------------------------------------------------*/
        private static string BuildConnectionString(string dataSource, string intialCatalog, string userName, string userPassword)
        {
            string cn = "";

            // Retrieve the partial connection string named databaseConnection from the application's app.config
            ConnectionStringSettings settings = ConfigurationManager.ConnectionStrings["master"];

            if (null != settings)
            {
                // Retrieve the partial connection string.
                string connectString = settings.ConnectionString;
                //Console.WriteLine("Original: {0}", connectString);

                // Create a new SqlConnectionStringBuilder based on the
                // partial connection string retrieved from the config file.
                SqlConnectionStringBuilder builder = new SqlConnectionStringBuilder(connectString);

                // Supply the additional values.
                builder.IntegratedSecurity = false;
                builder.DataSource = dataSource;
                builder.InitialCatalog = intialCatalog;
                builder.UserID = userName;
                builder.Password = userPassword;
                //Console.WriteLine("Modified: {0}", builder.ConnectionString);
                cn = builder.ConnectionString;
            }
            return cn;
        }


        public static void WriteExcelFile(DataTable dataTable, string fileName, string tableName)
        {
            // use ClosedXML to write to excel
            using (var book = new XLWorkbook(XLEventTracking.Disabled))
            {
                book.Worksheets.Add(dataTable, tableName);
                book.SaveAs(fileName);
            }

        }


    }
}
