namespace SqlScriptExec
{
    partial class OptionsConfig
    {
        /// <summary>
        /// Required designer variable.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        /// <summary>
        /// Clean up any resources being used.
        /// </summary>
        /// <param name="disposing">true if managed resources should be disposed; otherwise, false.</param>
        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
            {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        #region Windows Form Designer generated code

        /// <summary>
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InitializeComponent()
        {
            this.groupBox4 = new System.Windows.Forms.GroupBox();
            this.checkBoxConsecutiveErrors = new System.Windows.Forms.CheckBox();
            this.label5 = new System.Windows.Forms.Label();
            this.comboBoxErrorReceivedCount = new System.Windows.Forms.ComboBox();
            this.radioButtonStopAfterError = new System.Windows.Forms.RadioButton();
            this.radioButtonKeepExecutingUntilFinished = new System.Windows.Forms.RadioButton();
            this.buttonRestoreDefaults = new System.Windows.Forms.Button();
            this.groupBox3 = new System.Windows.Forms.GroupBox();
            this.checkBoxWarnBeforeRunning = new System.Windows.Forms.CheckBox();
            this.checkBoxIncludeSubfolders = new System.Windows.Forms.CheckBox();
            this.buttonCancel = new System.Windows.Forms.Button();
            this.buttonOK = new System.Windows.Forms.Button();
            this.groupBox2 = new System.Windows.Forms.GroupBox();
            this.checkBoxLogErrorsOnly = new System.Windows.Forms.CheckBox();
            this.checkBoxLogCreateExcel = new System.Windows.Forms.CheckBox();
            this.checkBoxLogDefaultPath = new System.Windows.Forms.CheckBox();
            this.label4 = new System.Windows.Forms.Label();
            this.buttonLogBrowse = new System.Windows.Forms.Button();
            this.textBoxLogPath = new System.Windows.Forms.TextBox();
            this.groupBox1 = new System.Windows.Forms.GroupBox();
            this.label3 = new System.Windows.Forms.Label();
            this.textBoxConnectionString = new System.Windows.Forms.TextBox();
            this.buttonServerList = new System.Windows.Forms.Button();
            this.label2 = new System.Windows.Forms.Label();
            this.textBoxPassword = new System.Windows.Forms.TextBox();
            this.label1 = new System.Windows.Forms.Label();
            this.textBoxUserName = new System.Windows.Forms.TextBox();
            this.checkBoxWindowsAuthentication = new System.Windows.Forms.CheckBox();
            this.folderBrowserDialog1 = new System.Windows.Forms.FolderBrowserDialog();
            this.groupBox4.SuspendLayout();
            this.groupBox3.SuspendLayout();
            this.groupBox2.SuspendLayout();
            this.groupBox1.SuspendLayout();
            this.SuspendLayout();
            // 
            // groupBox4
            // 
            this.groupBox4.Controls.Add(this.checkBoxConsecutiveErrors);
            this.groupBox4.Controls.Add(this.label5);
            this.groupBox4.Controls.Add(this.comboBoxErrorReceivedCount);
            this.groupBox4.Controls.Add(this.radioButtonStopAfterError);
            this.groupBox4.Controls.Add(this.radioButtonKeepExecutingUntilFinished);
            this.groupBox4.Font = new System.Drawing.Font("Microsoft Sans Serif", 7.875F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.groupBox4.Location = new System.Drawing.Point(27, 418);
            this.groupBox4.Margin = new System.Windows.Forms.Padding(2);
            this.groupBox4.Name = "groupBox4";
            this.groupBox4.Padding = new System.Windows.Forms.Padding(2);
            this.groupBox4.Size = new System.Drawing.Size(373, 70);
            this.groupBox4.TabIndex = 13;
            this.groupBox4.TabStop = false;
            this.groupBox4.Text = "Errors";
            // 
            // checkBoxConsecutiveErrors
            // 
            this.checkBoxConsecutiveErrors.AutoSize = true;
            this.checkBoxConsecutiveErrors.Location = new System.Drawing.Point(242, 39);
            this.checkBoxConsecutiveErrors.Margin = new System.Windows.Forms.Padding(2);
            this.checkBoxConsecutiveErrors.Name = "checkBoxConsecutiveErrors";
            this.checkBoxConsecutiveErrors.Size = new System.Drawing.Size(85, 17);
            this.checkBoxConsecutiveErrors.TabIndex = 4;
            this.checkBoxConsecutiveErrors.Text = "Consecutive";
            this.checkBoxConsecutiveErrors.UseVisualStyleBackColor = true;
            // 
            // label5
            // 
            this.label5.AutoSize = true;
            this.label5.Location = new System.Drawing.Point(157, 41);
            this.label5.Margin = new System.Windows.Forms.Padding(2, 0, 2, 0);
            this.label5.Name = "label5";
            this.label5.Size = new System.Drawing.Size(83, 13);
            this.label5.TabIndex = 3;
            this.label5.Text = "Errors Received";
            // 
            // comboBoxErrorReceivedCount
            // 
            this.comboBoxErrorReceivedCount.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
            this.comboBoxErrorReceivedCount.FormattingEnabled = true;
            this.comboBoxErrorReceivedCount.Items.AddRange(new object[] {
            "1",
            "2",
            "3",
            "4",
            "5",
            "6",
            "7",
            "8",
            "9",
            "10",
            "11",
            "12",
            "13",
            "14",
            "15",
            "16",
            "17",
            "18",
            "19",
            "20",
            "21",
            "22",
            "23",
            "24",
            "25",
            "26",
            "27",
            "28",
            "29",
            "30",
            "31",
            "32",
            "33",
            "34",
            "35",
            "36",
            "37",
            "38",
            "39",
            "40",
            "41",
            "42",
            "43",
            "44",
            "45",
            "46",
            "47",
            "48",
            "49",
            "50"});
            this.comboBoxErrorReceivedCount.Location = new System.Drawing.Point(114, 38);
            this.comboBoxErrorReceivedCount.Margin = new System.Windows.Forms.Padding(2);
            this.comboBoxErrorReceivedCount.Name = "comboBoxErrorReceivedCount";
            this.comboBoxErrorReceivedCount.Size = new System.Drawing.Size(41, 21);
            this.comboBoxErrorReceivedCount.TabIndex = 2;
            // 
            // radioButtonStopAfterError
            // 
            this.radioButtonStopAfterError.AutoSize = true;
            this.radioButtonStopAfterError.Location = new System.Drawing.Point(45, 40);
            this.radioButtonStopAfterError.Margin = new System.Windows.Forms.Padding(2);
            this.radioButtonStopAfterError.Name = "radioButtonStopAfterError";
            this.radioButtonStopAfterError.Size = new System.Drawing.Size(72, 17);
            this.radioButtonStopAfterError.TabIndex = 1;
            this.radioButtonStopAfterError.TabStop = true;
            this.radioButtonStopAfterError.Text = "Stop After";
            this.radioButtonStopAfterError.UseVisualStyleBackColor = true;
            this.radioButtonStopAfterError.Click += new System.EventHandler(this.radioButtonStopAfterError_Click);
            // 
            // radioButtonKeepExecutingUntilFinished
            // 
            this.radioButtonKeepExecutingUntilFinished.AutoSize = true;
            this.radioButtonKeepExecutingUntilFinished.Location = new System.Drawing.Point(45, 16);
            this.radioButtonKeepExecutingUntilFinished.Margin = new System.Windows.Forms.Padding(2);
            this.radioButtonKeepExecutingUntilFinished.Name = "radioButtonKeepExecutingUntilFinished";
            this.radioButtonKeepExecutingUntilFinished.Size = new System.Drawing.Size(201, 17);
            this.radioButtonKeepExecutingUntilFinished.TabIndex = 0;
            this.radioButtonKeepExecutingUntilFinished.TabStop = true;
            this.radioButtonKeepExecutingUntilFinished.Text = "Keep Executing Scripts Until Finished";
            this.radioButtonKeepExecutingUntilFinished.UseVisualStyleBackColor = true;
            this.radioButtonKeepExecutingUntilFinished.Click += new System.EventHandler(this.radioButtonKeepExecutingUntilFinished_Click);
            // 
            // buttonRestoreDefaults
            // 
            this.buttonRestoreDefaults.Location = new System.Drawing.Point(27, 516);
            this.buttonRestoreDefaults.Margin = new System.Windows.Forms.Padding(2);
            this.buttonRestoreDefaults.Name = "buttonRestoreDefaults";
            this.buttonRestoreDefaults.Size = new System.Drawing.Size(112, 33);
            this.buttonRestoreDefaults.TabIndex = 12;
            this.buttonRestoreDefaults.Text = "Restore Defaults";
            this.buttonRestoreDefaults.UseVisualStyleBackColor = true;
            this.buttonRestoreDefaults.Click += new System.EventHandler(this.buttonRestoreDefaults_Click);
            // 
            // groupBox3
            // 
            this.groupBox3.Controls.Add(this.checkBoxWarnBeforeRunning);
            this.groupBox3.Controls.Add(this.checkBoxIncludeSubfolders);
            this.groupBox3.Location = new System.Drawing.Point(27, 324);
            this.groupBox3.Margin = new System.Windows.Forms.Padding(2);
            this.groupBox3.Name = "groupBox3";
            this.groupBox3.Padding = new System.Windows.Forms.Padding(2);
            this.groupBox3.Size = new System.Drawing.Size(373, 86);
            this.groupBox3.TabIndex = 11;
            this.groupBox3.TabStop = false;
            this.groupBox3.Text = "Execution";
            // 
            // checkBoxWarnBeforeRunning
            // 
            this.checkBoxWarnBeforeRunning.AutoSize = true;
            this.checkBoxWarnBeforeRunning.Location = new System.Drawing.Point(15, 47);
            this.checkBoxWarnBeforeRunning.Margin = new System.Windows.Forms.Padding(2);
            this.checkBoxWarnBeforeRunning.Name = "checkBoxWarnBeforeRunning";
            this.checkBoxWarnBeforeRunning.Size = new System.Drawing.Size(129, 17);
            this.checkBoxWarnBeforeRunning.TabIndex = 0;
            this.checkBoxWarnBeforeRunning.Text = "Warn Before Running";
            this.checkBoxWarnBeforeRunning.UseVisualStyleBackColor = true;
            // 
            // checkBoxIncludeSubfolders
            // 
            this.checkBoxIncludeSubfolders.AutoSize = true;
            this.checkBoxIncludeSubfolders.Location = new System.Drawing.Point(15, 24);
            this.checkBoxIncludeSubfolders.Margin = new System.Windows.Forms.Padding(2);
            this.checkBoxIncludeSubfolders.Name = "checkBoxIncludeSubfolders";
            this.checkBoxIncludeSubfolders.Size = new System.Drawing.Size(114, 17);
            this.checkBoxIncludeSubfolders.TabIndex = 0;
            this.checkBoxIncludeSubfolders.Text = "Include Subfolders";
            this.checkBoxIncludeSubfolders.UseVisualStyleBackColor = true;
            // 
            // buttonCancel
            // 
            this.buttonCancel.Location = new System.Drawing.Point(158, 517);
            this.buttonCancel.Margin = new System.Windows.Forms.Padding(2);
            this.buttonCancel.Name = "buttonCancel";
            this.buttonCancel.Size = new System.Drawing.Size(112, 32);
            this.buttonCancel.TabIndex = 10;
            this.buttonCancel.Text = "Cancel";
            this.buttonCancel.UseVisualStyleBackColor = true;
            this.buttonCancel.Click += new System.EventHandler(this.buttonCancel_Click);
            // 
            // buttonOK
            // 
            this.buttonOK.Location = new System.Drawing.Point(289, 517);
            this.buttonOK.Margin = new System.Windows.Forms.Padding(2);
            this.buttonOK.Name = "buttonOK";
            this.buttonOK.Size = new System.Drawing.Size(112, 32);
            this.buttonOK.TabIndex = 9;
            this.buttonOK.Text = "OK";
            this.buttonOK.UseVisualStyleBackColor = true;
            this.buttonOK.Click += new System.EventHandler(this.buttonOK_Click);
            // 
            // groupBox2
            // 
            this.groupBox2.Controls.Add(this.checkBoxLogErrorsOnly);
            this.groupBox2.Controls.Add(this.checkBoxLogCreateExcel);
            this.groupBox2.Controls.Add(this.checkBoxLogDefaultPath);
            this.groupBox2.Controls.Add(this.label4);
            this.groupBox2.Controls.Add(this.buttonLogBrowse);
            this.groupBox2.Controls.Add(this.textBoxLogPath);
            this.groupBox2.Location = new System.Drawing.Point(27, 203);
            this.groupBox2.Margin = new System.Windows.Forms.Padding(2);
            this.groupBox2.Name = "groupBox2";
            this.groupBox2.Padding = new System.Windows.Forms.Padding(2);
            this.groupBox2.Size = new System.Drawing.Size(373, 116);
            this.groupBox2.TabIndex = 8;
            this.groupBox2.TabStop = false;
            this.groupBox2.Text = "Log";
            // 
            // checkBoxLogErrorsOnly
            // 
            this.checkBoxLogErrorsOnly.AutoSize = true;
            this.checkBoxLogErrorsOnly.Location = new System.Drawing.Point(15, 92);
            this.checkBoxLogErrorsOnly.Margin = new System.Windows.Forms.Padding(2);
            this.checkBoxLogErrorsOnly.Name = "checkBoxLogErrorsOnly";
            this.checkBoxLogErrorsOnly.Size = new System.Drawing.Size(98, 17);
            this.checkBoxLogErrorsOnly.TabIndex = 5;
            this.checkBoxLogErrorsOnly.Text = "Log Errors Only";
            this.checkBoxLogErrorsOnly.UseVisualStyleBackColor = true;
            // 
            // checkBoxLogCreateExcel
            // 
            this.checkBoxLogCreateExcel.AutoSize = true;
            this.checkBoxLogCreateExcel.Location = new System.Drawing.Point(15, 65);
            this.checkBoxLogCreateExcel.Margin = new System.Windows.Forms.Padding(2);
            this.checkBoxLogCreateExcel.Name = "checkBoxLogCreateExcel";
            this.checkBoxLogCreateExcel.Size = new System.Drawing.Size(321, 17);
            this.checkBoxLogCreateExcel.TabIndex = 4;
            this.checkBoxLogCreateExcel.Text = "Create Excel File Log (Seperate text log will always be created)";
            this.checkBoxLogCreateExcel.UseVisualStyleBackColor = true;
            // 
            // checkBoxLogDefaultPath
            // 
            this.checkBoxLogDefaultPath.AutoSize = true;
            this.checkBoxLogDefaultPath.Location = new System.Drawing.Point(15, 18);
            this.checkBoxLogDefaultPath.Margin = new System.Windows.Forms.Padding(2);
            this.checkBoxLogDefaultPath.Name = "checkBoxLogDefaultPath";
            this.checkBoxLogDefaultPath.Size = new System.Drawing.Size(227, 17);
            this.checkBoxLogDefaultPath.TabIndex = 3;
            this.checkBoxLogDefaultPath.Text = "Default Log Path (Current application path)";
            this.checkBoxLogDefaultPath.UseVisualStyleBackColor = true;
            this.checkBoxLogDefaultPath.Click += new System.EventHandler(this.checkBoxLogDefaultPath_Click);
            // 
            // label4
            // 
            this.label4.AutoSize = true;
            this.label4.Location = new System.Drawing.Point(12, 40);
            this.label4.Margin = new System.Windows.Forms.Padding(2, 0, 2, 0);
            this.label4.Name = "label4";
            this.label4.Size = new System.Drawing.Size(32, 13);
            this.label4.TabIndex = 2;
            this.label4.Text = "Path:";
            // 
            // buttonLogBrowse
            // 
            this.buttonLogBrowse.Location = new System.Drawing.Point(300, 33);
            this.buttonLogBrowse.Margin = new System.Windows.Forms.Padding(2);
            this.buttonLogBrowse.Name = "buttonLogBrowse";
            this.buttonLogBrowse.Size = new System.Drawing.Size(60, 25);
            this.buttonLogBrowse.TabIndex = 1;
            this.buttonLogBrowse.Text = "Browse";
            this.buttonLogBrowse.UseVisualStyleBackColor = true;
            this.buttonLogBrowse.Click += new System.EventHandler(this.buttonLogBrowse_Click);
            // 
            // textBoxLogPath
            // 
            this.textBoxLogPath.Location = new System.Drawing.Point(44, 36);
            this.textBoxLogPath.Margin = new System.Windows.Forms.Padding(2);
            this.textBoxLogPath.Name = "textBoxLogPath";
            this.textBoxLogPath.Size = new System.Drawing.Size(252, 20);
            this.textBoxLogPath.TabIndex = 0;
            // 
            // groupBox1
            // 
            this.groupBox1.Controls.Add(this.label3);
            this.groupBox1.Controls.Add(this.textBoxConnectionString);
            this.groupBox1.Controls.Add(this.buttonServerList);
            this.groupBox1.Controls.Add(this.label2);
            this.groupBox1.Controls.Add(this.textBoxPassword);
            this.groupBox1.Controls.Add(this.label1);
            this.groupBox1.Controls.Add(this.textBoxUserName);
            this.groupBox1.Controls.Add(this.checkBoxWindowsAuthentication);
            this.groupBox1.Location = new System.Drawing.Point(27, 21);
            this.groupBox1.Margin = new System.Windows.Forms.Padding(2);
            this.groupBox1.Name = "groupBox1";
            this.groupBox1.Padding = new System.Windows.Forms.Padding(2);
            this.groupBox1.Size = new System.Drawing.Size(373, 172);
            this.groupBox1.TabIndex = 7;
            this.groupBox1.TabStop = false;
            this.groupBox1.Text = "Server Connection";
            // 
            // label3
            // 
            this.label3.AutoSize = true;
            this.label3.Location = new System.Drawing.Point(12, 99);
            this.label3.Margin = new System.Windows.Forms.Padding(2, 0, 2, 0);
            this.label3.Name = "label3";
            this.label3.Size = new System.Drawing.Size(94, 13);
            this.label3.TabIndex = 7;
            this.label3.Text = "Connection String:";
            // 
            // textBoxConnectionString
            // 
            this.textBoxConnectionString.Location = new System.Drawing.Point(15, 117);
            this.textBoxConnectionString.Margin = new System.Windows.Forms.Padding(2);
            this.textBoxConnectionString.Multiline = true;
            this.textBoxConnectionString.Name = "textBoxConnectionString";
            this.textBoxConnectionString.ReadOnly = true;
            this.textBoxConnectionString.Size = new System.Drawing.Size(342, 50);
            this.textBoxConnectionString.TabIndex = 6;
            // 
            // buttonServerList
            // 
            this.buttonServerList.Location = new System.Drawing.Point(266, 23);
            this.buttonServerList.Margin = new System.Windows.Forms.Padding(2);
            this.buttonServerList.Name = "buttonServerList";
            this.buttonServerList.Size = new System.Drawing.Size(88, 69);
            this.buttonServerList.TabIndex = 5;
            this.buttonServerList.Text = "Server List";
            this.buttonServerList.UseVisualStyleBackColor = true;
            this.buttonServerList.Click += new System.EventHandler(this.buttonServerList_Click);
            // 
            // label2
            // 
            this.label2.AutoSize = true;
            this.label2.Location = new System.Drawing.Point(12, 70);
            this.label2.Margin = new System.Windows.Forms.Padding(2, 0, 2, 0);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(56, 13);
            this.label2.TabIndex = 4;
            this.label2.Text = "Password:";
            // 
            // textBoxPassword
            // 
            this.textBoxPassword.Location = new System.Drawing.Point(75, 67);
            this.textBoxPassword.Margin = new System.Windows.Forms.Padding(2);
            this.textBoxPassword.Name = "textBoxPassword";
            this.textBoxPassword.PasswordChar = '*';
            this.textBoxPassword.Size = new System.Drawing.Size(180, 20);
            this.textBoxPassword.TabIndex = 3;
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Location = new System.Drawing.Point(12, 45);
            this.label1.Margin = new System.Windows.Forms.Padding(2, 0, 2, 0);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(63, 13);
            this.label1.TabIndex = 2;
            this.label1.Text = "User Name:";
            // 
            // textBoxUserName
            // 
            this.textBoxUserName.Location = new System.Drawing.Point(75, 42);
            this.textBoxUserName.Margin = new System.Windows.Forms.Padding(2);
            this.textBoxUserName.Name = "textBoxUserName";
            this.textBoxUserName.Size = new System.Drawing.Size(180, 20);
            this.textBoxUserName.TabIndex = 1;
            // 
            // checkBoxWindowsAuthentication
            // 
            this.checkBoxWindowsAuthentication.AutoSize = true;
            this.checkBoxWindowsAuthentication.Location = new System.Drawing.Point(15, 23);
            this.checkBoxWindowsAuthentication.Margin = new System.Windows.Forms.Padding(2);
            this.checkBoxWindowsAuthentication.Name = "checkBoxWindowsAuthentication";
            this.checkBoxWindowsAuthentication.Size = new System.Drawing.Size(163, 17);
            this.checkBoxWindowsAuthentication.TabIndex = 0;
            this.checkBoxWindowsAuthentication.Text = "Use Windows Authentication";
            this.checkBoxWindowsAuthentication.UseVisualStyleBackColor = true;
            this.checkBoxWindowsAuthentication.Click += new System.EventHandler(this.checkBoxWindowsAuthentication_Click);
            // 
            // OptionsConfig
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(433, 567);
            this.Controls.Add(this.groupBox4);
            this.Controls.Add(this.buttonRestoreDefaults);
            this.Controls.Add(this.groupBox3);
            this.Controls.Add(this.buttonCancel);
            this.Controls.Add(this.buttonOK);
            this.Controls.Add(this.groupBox2);
            this.Controls.Add(this.groupBox1);
            this.Name = "OptionsConfig";
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterParent;
            this.Text = "OptionsConfig";
            this.Load += new System.EventHandler(this.OptionsConfig_Load);
            this.groupBox4.ResumeLayout(false);
            this.groupBox4.PerformLayout();
            this.groupBox3.ResumeLayout(false);
            this.groupBox3.PerformLayout();
            this.groupBox2.ResumeLayout(false);
            this.groupBox2.PerformLayout();
            this.groupBox1.ResumeLayout(false);
            this.groupBox1.PerformLayout();
            this.ResumeLayout(false);

        }

        #endregion

        private System.Windows.Forms.GroupBox groupBox4;
        private System.Windows.Forms.CheckBox checkBoxConsecutiveErrors;
        private System.Windows.Forms.Label label5;
        private System.Windows.Forms.ComboBox comboBoxErrorReceivedCount;
        private System.Windows.Forms.RadioButton radioButtonStopAfterError;
        private System.Windows.Forms.RadioButton radioButtonKeepExecutingUntilFinished;
        private System.Windows.Forms.Button buttonRestoreDefaults;
        private System.Windows.Forms.GroupBox groupBox3;
        private System.Windows.Forms.CheckBox checkBoxIncludeSubfolders;
        private System.Windows.Forms.Button buttonCancel;
        private System.Windows.Forms.Button buttonOK;
        private System.Windows.Forms.GroupBox groupBox2;
        private System.Windows.Forms.CheckBox checkBoxLogErrorsOnly;
        private System.Windows.Forms.CheckBox checkBoxLogCreateExcel;
        private System.Windows.Forms.CheckBox checkBoxLogDefaultPath;
        private System.Windows.Forms.Label label4;
        private System.Windows.Forms.Button buttonLogBrowse;
        private System.Windows.Forms.TextBox textBoxLogPath;
        private System.Windows.Forms.GroupBox groupBox1;
        private System.Windows.Forms.Label label3;
        private System.Windows.Forms.TextBox textBoxConnectionString;
        private System.Windows.Forms.Button buttonServerList;
        private System.Windows.Forms.Label label2;
        private System.Windows.Forms.TextBox textBoxPassword;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.TextBox textBoxUserName;
        private System.Windows.Forms.CheckBox checkBoxWindowsAuthentication;
        private System.Windows.Forms.CheckBox checkBoxWarnBeforeRunning;
        private System.Windows.Forms.FolderBrowserDialog folderBrowserDialog1;
    }
}