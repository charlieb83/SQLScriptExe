using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace SqlScriptExec
{
    public partial class OptionsConfig : Form
    {
        public OptionsConfig()
        {
            InitializeComponent();
        }

        /*-----------------------------------------------------
        Form Load
        -----------------------------------------------------*/
        private void OptionsConfig_Load(object sender, EventArgs e)
        {
            SetFormValues(sender, e);
        }

        /*-----------------------------------------------------
        Set Values on Form from Options Class
        -----------------------------------------------------*/
        private void SetFormValues(object sender, EventArgs e)
        {
            checkBoxWindowsAuthentication.Checked = Options.Connections.WindowsAuthentication;
            textBoxUserName.Text = Options.Connections.UserName;
            textBoxPassword.Text = Options.Connections.Password;
            textBoxConnectionString.Text = Helper.GetConnectionStringSafeForDisplay(Options.Connections.ConnectionString);
            checkBoxWindowsAuthentication_Click(sender, e);

            checkBoxLogDefaultPath.Checked = Options.Log.UseDefaultLogPath;
            textBoxLogPath.Text = Options.Log.LogPath;
            checkBoxLogCreateExcel.Checked = Options.Log.CreateExcelLog;
            checkBoxLogErrorsOnly.Checked = Options.Log.LogErrorsOnly;
            checkBoxLogDefaultPath_Click(sender, e);

            checkBoxIncludeSubfolders.Checked = Options.Execution.IncludeSubFolders;
            checkBoxWarnBeforeRunning.Checked = Options.Execution.WarnBeforeRunning;
            radioButtonKeepExecutingUntilFinished.Checked = Options.Execution.ExecuteScriptsUntilFinished;
            comboBoxErrorReceivedCount.Text = Options.Execution.StopAfterErrorCount.ToString();
            checkBoxConsecutiveErrors.Checked = Options.Execution.ConsecutiveErrors;

            if (!radioButtonKeepExecutingUntilFinished.Checked)
            {
                radioButtonStopAfterError.Checked = true;
            }
            radioButtonKeepExecutingUntilFinished_Click(sender, e);
            radioButtonStopAfterError_Click(sender, e);
        }

        /*-----------------------------------------------------
        Windows Authentication Clicked
        -----------------------------------------------------*/
        private void checkBoxWindowsAuthentication_Click(object sender, EventArgs e)
        {
            if (checkBoxWindowsAuthentication.Checked == true)
            {
                textBoxUserName.Enabled = false;
                textBoxPassword.Enabled = false;
            }
            else
            {
                textBoxUserName.Enabled = true;
                textBoxPassword.Enabled = true;
            }
        }

        /*-----------------------------------------------------
        Default Log Path CheckBox Clicked
        -----------------------------------------------------*/
        private void checkBoxLogDefaultPath_Click(object sender, EventArgs e)
        {
            if (checkBoxLogDefaultPath.Checked)
            {
                //TODO: Set in Options Class??
                string path = System.IO.Directory.GetCurrentDirectory();
                textBoxLogPath.Text = path;
                textBoxLogPath.Enabled = false;
                buttonLogBrowse.Enabled = false;
            }
            else
            {
                textBoxLogPath.Enabled = true;
                buttonLogBrowse.Enabled = true;
            }
        }

        /*-----------------------------------------------------
        Button Browse Log Path Clicked
        -----------------------------------------------------*/
        private void buttonLogBrowse_Click(object sender, EventArgs e)
        {
            if (folderBrowserDialog1.ShowDialog() == DialogResult.OK)
            {
                textBoxLogPath.Text = folderBrowserDialog1.SelectedPath;
            }
        }

        /*-----------------------------------------------------
        RadioButton KeepExecutingUntilFinished Clicked
        -----------------------------------------------------*/
        private void radioButtonKeepExecutingUntilFinished_Click(object sender, EventArgs e)
        {
            if (radioButtonKeepExecutingUntilFinished.Checked)
            {
                comboBoxErrorReceivedCount.Enabled = false;
                checkBoxConsecutiveErrors.Enabled = false;
            }
        }

        /*-----------------------------------------------------
        RadioButton StopAfterError Clicked
        -----------------------------------------------------*/
        private void radioButtonStopAfterError_Click(object sender, EventArgs e)
        {
            if(radioButtonStopAfterError.Checked)
            {
                comboBoxErrorReceivedCount.Enabled = true;
                checkBoxConsecutiveErrors.Enabled = true;
            }
        }

        /*-----------------------------------------------------
        OK Clicked
        -----------------------------------------------------*/
        private void buttonOK_Click(object sender, EventArgs e)
        {
            SetOptionClassValues();
            if (SettingsAreValid() == false)
            {
                return;
            }
            
            Helper.SavePropertySettings();
            this.Close();
        }

        /*-----------------------------------------------------
        Cancel Clicked
        -----------------------------------------------------*/
        private void buttonCancel_Click(object sender, EventArgs e)
        {
            this.Close();
        }

        /*-----------------------------------------------------
        Restore Defaults Clicked
        -----------------------------------------------------*/
        private void buttonRestoreDefaults_Click(object sender, EventArgs e)
        {
            Helper.ResetDefaultValues();    //Reset Defaults
            SetFormValues(sender, e);       //Reload form
        }

        /*-----------------------------------------------------
        Validate Settings
        -----------------------------------------------------*/
        private bool SettingsAreValid()
        {
            bool retVal = true;
            string errorMessage = "";
            errorMessage = Options.ValidateOptions(false);
            if (!string.IsNullOrWhiteSpace(errorMessage))
            {
                MessageBox.Show(errorMessage, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                retVal = false;
            }
            return retVal;
        }

        /*-----------------------------------------------------
        Set Current Values in Options Class
        -----------------------------------------------------*/
        private void SetOptionClassValues()
        {
            Options.Connections.WindowsAuthentication = checkBoxWindowsAuthentication.Checked;
            Options.Connections.UserName = textBoxUserName.Text;
            Options.Connections.Password = textBoxPassword.Text;

            Options.Log.UseDefaultLogPath = checkBoxLogDefaultPath.Checked;
            Options.Log.LogPath = textBoxLogPath.Text;
            Options.Log.CreateExcelLog = checkBoxLogCreateExcel.Checked;
            Options.Log.LogErrorsOnly = checkBoxLogErrorsOnly.Checked;

            Options.Execution.IncludeSubFolders = checkBoxIncludeSubfolders.Checked;
            Options.Execution.WarnBeforeRunning = checkBoxWarnBeforeRunning.Checked;
            Options.Execution.ExecuteScriptsUntilFinished = radioButtonKeepExecutingUntilFinished.Checked;
            Options.Execution.StopAfterErrorCount = Int32.Parse(comboBoxErrorReceivedCount.Text);
            Options.Execution.ConsecutiveErrors = checkBoxConsecutiveErrors.Checked;
        }

        private void buttonServerList_Click(object sender, EventArgs e)
        {
            MessageBox.Show("Coming soon!");
        }

    }
}
