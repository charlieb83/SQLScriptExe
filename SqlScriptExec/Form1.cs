using System;
using System.ComponentModel;
using System.Data;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using System.IO;
using System.Data.SqlClient;
using Microsoft.SqlServer.Management.Smo;
using Microsoft.SqlServer.Management.Common;

//Add to App.Config startup:    useLegacyV2RuntimeActivationPolicy="true"
namespace SqlScriptExec
{
    public partial class Form1 : Form
    {
        //TODO: Move to Helper?
        BackgroundWorker backgroundWorker1;
        DataTable excelDataTable = new DataTable();     

        public Form1()
        {
            InitializeComponent();

            //TEMP
            //TODO: Need to fix ConnectionString when Use Auth is NOT checked and No UserName is given
            //TODO: Error happens when use windows auth is false and no username or password given
            textBoxServer.Text = @"192.168.1.120\CHAZ_SQLSERVER";
            textBoxScriptPath.Text = @"C:\Users\cbach\Documents\VisualStudioTestingFiles\SQLScripts";
            //TEMP

            //Reset form
            ResetRuntimeStats();
            SetFormOptionValues();  //Load Option Values

            //Setup DataTable for Excel
            excelDataTable.Clear();
            excelDataTable.Columns.Add("Server");
            excelDataTable.Columns.Add("File");
            excelDataTable.Columns.Add("Status");
            excelDataTable.Columns.Add("Date");
            excelDataTable.Columns.Add("Message");

            // Create background worker thread
            backgroundWorker1 = new BackgroundWorker();
            backgroundWorker1.DoWork += new DoWorkEventHandler(backgroundWorker1_DoWork);
            backgroundWorker1.ProgressChanged += new ProgressChangedEventHandler(backgroundWorker1_ProgressChanged);
            backgroundWorker1.RunWorkerCompleted += new RunWorkerCompletedEventHandler(backgroundWorker1_RunWorkerCompleted);
            backgroundWorker1.WorkerReportsProgress = true;
            backgroundWorker1.WorkerSupportsCancellation = true;

        }

        /*-----------------------------------------------------
        Sets values on form from Option Class
        -----------------------------------------------------*/
        private void SetFormOptionValues()
        {
            Options.Connections.SelectedServer = textBoxServer.Text;
            Options.Execution.ScriptsToExecutePath = textBoxScriptPath.Text;

            Helper.LoadOptionValues();
            labelUseWindowsAuthentication.Text = Options.Connections.WindowsAuthentication.ToString();
            labelCreateExcelLog.Text = Options.Log.CreateExcelLog.ToString();
            labelLogErrorsOnly.Text = Options.Log.LogErrorsOnly.ToString();
            labelIncludeSubfolders.Text = Options.Execution.IncludeSubFolders.ToString();
            labelExecuteUntilFinished.Text = Options.Execution.ExecuteScriptsUntilFinished.ToString();

            if (Options.Execution.ExecuteScriptsUntilFinished == true)
            {
                labelStopAfterErrorCount.Text = "--";
            }
            else
            {
                labelStopAfterErrorCount.Text = Options.Execution.StopAfterErrorCount.ToString();
            }
        }


        /*-----------------------------------------------------
        Browse Clicked
        -----------------------------------------------------*/
        private void buttonBrowse_Click(object sender, EventArgs e)
        {
            if (folderBrowserDialog1.ShowDialog() == DialogResult.OK)
            {
                textBoxScriptPath.Text = folderBrowserDialog1.SelectedPath;
                Options.Execution.ScriptsToExecutePath = textBoxScriptPath.Text.Trim();
            }
        }


        /*-----------------------------------------------------
        textBoxStatus
        -----------------------------------------------------*/
        public void SetTextBoxStatus(string text)
        {
            if (InvokeRequired)
                Invoke(new Action<string>(SetTextBoxStatus), text);
            else
                textBoxStatus.Text = text;
        }

        public void AppendTextBoxStatus(string text)
        {
            if (InvokeRequired)
                Invoke(new Action<string>(AppendTextBoxStatus), text);
            else
                textBoxStatus.AppendText(text);
        }

        public void AppendTextBoxStatusWithTimeStamp(string text)
        {
            if (InvokeRequired)
                Invoke(new Action<string>(AppendTextBoxStatusWithTimeStamp), text);
            else
                textBoxStatus.AppendText(DateTime.Now.ToString("yyyyMMddHHmmss") + "--" + text);
        }

        /*-----------------------------------------------------
        SET lablelScriptsInQueue
        -----------------------------------------------------*/
        public void SetScriptsInQueueText(string text)
        {
            if (InvokeRequired)
                Invoke(new Action<string>(SetScriptsInQueueText), text);
            else
                labelScriptsInQueue.Text = text;
        }

        /*-----------------------------------------------------
        SET lablelScriptsUpdated
        -----------------------------------------------------*/
        public void SetScriptsUpdatedText(string text)
        {
            if (InvokeRequired)
                Invoke(new Action<string>(SetScriptsUpdatedText), text);
            else
                labelScriptsUpdated.Text = text;
        }

        /*-----------------------------------------------------
        SET lablelScriptsInQueue
        -----------------------------------------------------*/
        public void SetScriptsErrorText(string text)
        {
            if (InvokeRequired)
                Invoke(new Action<string>(SetScriptsErrorText), text);
            else
                labelScriptErrors.Text = text;
        }

        /*-----------------------------------------------------
        Reset ProgressBar
        -----------------------------------------------------*/
        private void ResetRuntimeStats()
        {
            //Reset Stats
            labelScriptsInQueue.Text = "0";
            labelScriptsUpdated.Text = "0";
            labelScriptErrors.Text = "0";

            //Reset progressbar
            circularProgressBar1.Value = 0;
            circularProgressBar1.Minimum = 0;
            circularProgressBar1.Maximum = 100;
            circularProgressBar1.Text = "0%";
        }


        //TODO: Default Path of Log should be where Run Script is

        /*-----------------------------------------------------
        Run Clicked
        -----------------------------------------------------*/
        private void buttonRun_Click(object sender, EventArgs e)
        {
            if(Options.Execution.WarnBeforeRunning==true)
            {
                DialogResult result = MessageBox.Show("You are about to execute SQL Scripts."+ Environment.NewLine + "You can Cancel while running but IT WILL NOT Rollback scripts that already executed." + Environment.NewLine + "Are you sure you want to continue?", "Message", MessageBoxButtons.YesNo, MessageBoxIcon.Exclamation, MessageBoxDefaultButton.Button2);
                if (result == DialogResult.No)
                {
                    this.AppendTextBoxStatus("User Chickened Out....." + "\r\n");
                    return;
                }
            }
            string errorMessage = "";

            Options.Connections.SelectedServer = textBoxServer.Text.Trim();
            Options.Execution.ScriptsToExecutePath = textBoxScriptPath.Text.Trim();
            SetFormOptionValues();  //Set values in Options Class
            textBoxStatus.Clear();
            excelDataTable.Clear();

            this.AppendTextBoxStatusWithTimeStamp( "Validating Options.....");
            //Don't run if not valid
            errorMessage = Options.ValidateOptions(true);
            if (!string.IsNullOrWhiteSpace(errorMessage))
            {
                this.AppendTextBoxStatus("Failed\r\n");
                MessageBox.Show(errorMessage, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }
            this.AppendTextBoxStatus("Passed\r\n");

            this.AppendTextBoxStatusWithTimeStamp("Checking Server Connection.....");
            //Check server connection
            if (IsServerConnected(Options.Connections.ConnectionString, false) == false)
            {
                this.AppendTextBoxStatus("Failed\r\n");
                return;
            }
            this.AppendTextBoxStatus("Passed\r\n");

            //TODO: disable buttons and re-enable after done
            ResetRuntimeStats();
            
            if (backgroundWorker1.IsBusy != true)
            {
                this.AppendTextBoxStatusWithTimeStamp("Starting Worker Thread.....\r\n");
                backgroundWorker1.RunWorkerAsync();
            }
               
            buttonCancel.Enabled = true;
            buttonRun.Enabled = false;
        }



        /*-----------------------------------------------------
        BackgroundWorker DoWork - Main Thread
        -----------------------------------------------------*/
        private void backgroundWorker1_DoWork(object sender, DoWorkEventArgs e)
        {
            string sqlScriptFile;
            string scriptsPath = "";
            string serverConnection = "";
            int totalScriptCount = 0;
            int scriptsUpdated = 0;
            int errorCount = 0;
            int currentCount = 0;
            int consecutiveErrorCount = 0;
            string exceptionMessage = "";
            string status = "";

            DateTime startTime = DateTime.Now;   //.ToString("yyyyMMddHHmmss");

            StringBuilder errorMessages = new StringBuilder();

            SearchOption option = Options.Execution.IncludeSubFolders ? SearchOption.AllDirectories : SearchOption.TopDirectoryOnly;

            this.Invoke((MethodInvoker)delegate ()
            {
                serverConnection = textBoxServer.Text.Trim();      //TODO: Set in Options Class - CurrentServer
                scriptsPath = textBoxScriptPath.Text.Trim();
            });

            //Get ConnectionString
            string sqlConnectionString = Options.Connections.ConnectionString; //Helper.GetConnectionString(Options.Connections.SelectedServer, "master");
            //TODO: Move this validation to button click
            //TODO: Remove Temp Count
            if (sqlConnectionString == "")
            {
                return;
            }
            //string sqlConnectionString = @"Server=" + serverConnection + "; Database=master; User ID=john.smith; Password=123456;";

            // Get number of files
            totalScriptCount = Directory.GetFiles(scriptsPath, "*", option).Length;
            SetScriptsInQueueText( totalScriptCount.ToString() );

            int step = circularProgressBar1.Maximum / totalScriptCount;

            using (StreamWriter w = File.CreateText(Options.Log.LogPath + @"\" + Options.Log.LogFileName + @".log"))
            {
                //Log Header
                Log(GetLogHeader(), w, false);
           
                using (SqlConnection connection = new SqlConnection(sqlConnectionString))
                {
                    Server server = new Server(new ServerConnection(connection));

                    //Loop through folders and files order by name
                    foreach (string file in Directory.EnumerateFiles(scriptsPath, "*.sql", option).OrderBy(f => f))
                    //foreach (string file in Directory.EnumerateFiles(scriptsPath, "*.sql", option) )
                    {
                        currentCount++;
                        sqlScriptFile = File.ReadAllText(file);
                        this.AppendTextBoxStatusWithTimeStamp(file);
                        exceptionMessage = "";

                        try
                        {
                            server.ConnectionContext.ExecuteNonQuery(sqlScriptFile);
                            this.AppendTextBoxStatus("  -- Complete" + "\r\n");
                            scriptsUpdated++;
                            consecutiveErrorCount = 0;  //Reset Consecutive Error count

                            //Log Success if LogErrorsOnly is false
                            if (Options.Log.LogErrorsOnly==false)
                            {
                                Log(file + "  -- Complete", w, true);
                            }
                        }

                        catch (Exception ex)
                        {
                            //Console.WriteLine(ex.InnerException.Message.ToString());
                            //Console.WriteLine(ex.GetBaseException().Message.ToString());
                            exceptionMessage = ex.InnerException.Message.ToString();
                            this.AppendTextBoxStatus("  -- Error: " + exceptionMessage + "\r\n");
                            errorCount++;
                            consecutiveErrorCount++;

                            //Log Error
                            Log(file + "-- Error: " + exceptionMessage, w, true);
                        }
                        //Compile values into an Array (ProgressStepValue, ScriptsUpdated, ErrorCount)
                        string[] arr1 = new string[] { step.ToString(), scriptsUpdated.ToString(), errorCount.ToString() };
                        backgroundWorker1.ReportProgress(0, arr1);
                        //backgroundWorker1.ReportProgress(step);   //orig

                        //Add to Datatable if Create Excel Log is True
                        if (Options.Log.CreateExcelLog == true)
                        {
                            if (string.IsNullOrWhiteSpace(exceptionMessage))
                            {
                                exceptionMessage = "Complete";
                                status = "C";       //TODO: Enums???
                            }
                            else
                            {
                                status = "E";
                            }
                            //Only Log Errors if Log Errors Only = true
                            if (status == "E" || Options.Log.LogErrorsOnly == false)
                            {
                                PopulateDataTable(Options.Connections.SelectedServer, file, status, DateTime.Now.ToString("yyyyMMddHHmmss"), exceptionMessage);
                            }
                        }
                        
                        //If Stop after xx Errors is set, check here and cancel 
                        if (Options.Execution.ExecuteScriptsUntilFinished ==false)
                        {
                            if (Options.Execution.ConsecutiveErrors == true)
                            {
                                if (consecutiveErrorCount >= Options.Execution.StopAfterErrorCount)
                                {
                                    this.AppendTextBoxStatus("Max Consecutive Errors of " + Options.Execution.StopAfterErrorCount.ToString() + " has been reached.....Canceling" + "\r\n");
                                    e.Cancel = true;
                                    break;
                                }
                            }
                            else if (errorCount >= Options.Execution.StopAfterErrorCount )
                            {
                                this.AppendTextBoxStatus("Max Errors of " + Options.Execution.StopAfterErrorCount.ToString() + " has been reached.....Canceling" + "\r\n");
                                e.Cancel = true;
                                break;
                            }
                        }

                        if (backgroundWorker1.CancellationPending)
                        {
                            // Set the e.Cancel flag so that the WorkerCompleted event
                            // knows that the process was cancelled.
                            e.Cancel = true;
                            backgroundWorker1.ReportProgress(0);
                            return;
                        }
                    }
                }
                //Finished
                //totalRuntime = startTime;
                TimeSpan span = (DateTime.Now - startTime);
                this.AppendTextBoxStatusWithTimeStamp("Execution Complete--" + "\r\n");
                this.AppendTextBoxStatus("----------Total Runtime: " + String.Format("{0} minutes, {1} seconds", span.Minutes, span.Seconds) + "\r\n");
                this.AppendTextBoxStatus("----------Total Scripts In Queue: " + totalScriptCount.ToString() + "    Total Scripts Run: " + currentCount.ToString() + "    Total Errors: " + errorCount.ToString() + "\r\n");
                //Log
                Log( GetLogFooter(totalScriptCount, scriptsUpdated, errorCount), w, false);
                if (Options.Log.CreateExcelLog == true)
                {
                    CreateExcelLog();
                }
                    
            }

        }

        /*-----------------------------------------------------
        BackgroundWorker Progress Changed
        -----------------------------------------------------*/
        private void backgroundWorker1_ProgressChanged(object sender, ProgressChangedEventArgs e)
        {
            string[] runStats = (string[])e.UserState;
            //Console.WriteLine(runStats[0].ToString());  //ProgressBar Step
            //Console.WriteLine(runStats[1].ToString());  //scriptsUpdated
            //Console.WriteLine(runStats[2].ToString());  //errorCount
            int progressStep = Int32.Parse(runStats[0]);
            SetScriptsUpdatedText( runStats[1].ToString() );
            SetScriptsErrorText ( runStats[2].ToString() );
 
            circularProgressBar1.Value += progressStep;
            circularProgressBar1.Text = circularProgressBar1.Value.ToString() + "%";
            circularProgressBar1.Update();
        }

        /*-----------------------------------------------------
        BackgroundWorker Completed
        -----------------------------------------------------*/
        private void backgroundWorker1_RunWorkerCompleted(object sender, RunWorkerCompletedEventArgs e)
        {
            circularProgressBar1.Value = 100;
            circularProgressBar1.Text = "100%";
            circularProgressBar1.Update();
            //MessageBox.Show("DONE!!!");
            buttonCancel.Enabled = false;
            buttonRun.Enabled = true;
        }


        /*-----------------------------------------------------
        Populate DataTable  - TODO: Do this better - May not want to populate entire datatable, might get too big
        -----------------------------------------------------*/
        private void PopulateDataTable(string server, string file, string status, string date, string message)
        {
            DataRow dr = excelDataTable.NewRow();

            dr[0] = server;     //Server
            dr[1] = file;       //File
            dr[2] = status;     //Status
            dr[3] = date;       //Date
            dr[4] = message;    //Message

            excelDataTable.Rows.Add(dr);//Add to end of the datatable
        }


        /*-----------------------------------------------------
        buttonTest Clicked - Test Server Connection
        -----------------------------------------------------*/
        private void buttonTest_Click(object sender, EventArgs e)
        {
            string connectionString;
            
            if (textBoxServer.Text.Trim() != "")
            {
                //Get ConnectionString
                connectionString = Helper.GetConnectionString(textBoxServer.Text.Trim(), "master");
                if (connectionString == "")
                {
                    return;
                }
                IsServerConnected(connectionString, true);
            }
        }


        /*-----------------------------------------------------
        Get Log Header - Build --TODO: Move to Class?
        -----------------------------------------------------*/
        private string GetLogHeader()
        {
            string logHeader = "";
            logHeader += "--------------------Execution Starting--------------------" + System.Environment.NewLine;
            logHeader += "--Start Time: " + DateTime.Now.ToLongDateString() + " " + DateTime.Now.ToLongTimeString() + System.Environment.NewLine;
            logHeader += "--Script To Execute Path: " + Options.Execution.ScriptsToExecutePath + System.Environment.NewLine;
            logHeader += "--Selected Server: " + Options.Connections.SelectedServer + System.Environment.NewLine;
            logHeader += "--Log Errors Only: " + Options.Log.LogErrorsOnly.ToString() + System.Environment.NewLine;
            logHeader += "--Create Excel Log: " + Options.Log.CreateExcelLog.ToString() + System.Environment.NewLine;
            logHeader += "--Include SubFolders: " + Options.Execution.IncludeSubFolders.ToString() + System.Environment.NewLine;
            logHeader += "----------------------------------------------------------";
            return logHeader;
        }

        /*-----------------------------------------------------
        Get Log Footer - Build --TODO: Move to Class?
        -----------------------------------------------------*/
        private string GetLogFooter(int totalFiles, int successfulCount, int errorCount)
        {
            string logFooter = "";
            logFooter += "--------------------Execution Complete--------------------" + System.Environment.NewLine;
            logFooter += "--End Time: " + DateTime.Now.ToLongDateString() + " " + DateTime.Now.ToLongTimeString() + System.Environment.NewLine;
            logFooter += "--Total Files Evaluated: " + totalFiles.ToString() + System.Environment.NewLine;
            logFooter += "--Total Successful: " + successfulCount.ToString() + System.Environment.NewLine;
            logFooter += "--Total Errors: " + errorCount.ToString() + System.Environment.NewLine;
            logFooter += "----------------------------------------------------------";
            return logFooter;
        }

        //TODO: Single place to generate log path/name
        //TODO: Move these to class
        /*-----------------------------------------------------
        Write to Log
        -----------------------------------------------------*/
        private void Log(string logMessage, TextWriter w, bool includeDate)
        {
            if (includeDate==true)
            {
                w.WriteLine("{0} {1} {2}", DateTime.Now.ToLongTimeString(), DateTime.Now.ToLongDateString(), logMessage);
            }
            else
            {
                w.WriteLine(logMessage);
            }
        }

        /*-----------------------------------------------------
        Test If Conncection to SQL Server Is Valid
        -----------------------------------------------------*/
        private void CreateExcelLog()
        {
            string file = Options.Log.LogPath + @"\" + Options.Log.LogFileName + @".xlsx";
            //using (DataTable table = new DataTable { TableName = "ExcelLog" })
            //{
            //    table.Columns.Add("File");
            //    table.Columns.Add("Status");
                Helper.WriteExcelFile(excelDataTable, file, "ExcelLog");
            //}
        }


        /*-----------------------------------------------------
        Test If Conncection to SQL Server Is Valid
        -----------------------------------------------------*/
        private static bool IsServerConnected(string connectionString, bool showSuccessMessage)
        {
            using (SqlConnection connection = new SqlConnection(connectionString))
            {
                try
                {
                    connection.Open();
                    if (showSuccessMessage == true)
                        MessageBox.Show("Connection Successful!", "Message", MessageBoxButtons.OK, MessageBoxIcon.None);
                    return true;
                }
                catch (SqlException ex)
                {
                    MessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    return false;
                }
            }
        }

        /*-----------------------------------------------------
        Cancel Clicked
        -----------------------------------------------------*/
        private void buttonCancel_Click(object sender, EventArgs e)
        {
            //if (backgroundWorker1.IsBusy)
            if (backgroundWorker1.WorkerSupportsCancellation == true)
            {
                backgroundWorker1.CancelAsync();
            }
        }


        /*-----------------------------------------------------
        Exit Clicked
        -----------------------------------------------------*/
        private void exitToolStripMenuItem_Click(object sender, EventArgs e)
        {
            Application.Exit();
        }

        /*-----------------------------------------------------
        Options Configuration Clicked
        -----------------------------------------------------*/
        private void configurationToolStripMenuItem_Click(object sender, EventArgs e)
        {
            OptionsConfig o = new OptionsConfig();
            o.ShowDialog();
            SetFormOptionValues();
        }

    }
}
