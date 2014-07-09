namespace TestCaseExport
{
    partial class FrmMain
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
            this.components = new System.ComponentModel.Container();
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(FrmMain));
            this.pictureBox1 = new System.Windows.Forms.PictureBox();
            this.groupBox1 = new System.Windows.Forms.GroupBox();
            this.comBoxTestSuite = new System.Windows.Forms.ComboBox();
            this.bsData = new System.Windows.Forms.BindingSource(this.components);
            this.testSuitesBindingSource = new System.Windows.Forms.BindingSource(this.components);
            this.lblTestSuite = new System.Windows.Forms.Label();
            this.comBoxTestPlan = new System.Windows.Forms.ComboBox();
            this.testPlansBindingSource = new System.Windows.Forms.BindingSource(this.components);
            this.lblTestPlan = new System.Windows.Forms.Label();
            this.btnTeamProject = new System.Windows.Forms.Button();
            this.txtTeamProject = new System.Windows.Forms.TextBox();
            this.lblTeamProject = new System.Windows.Forms.Label();
            this.groupBox2 = new System.Windows.Forms.GroupBox();
            this.btnFolderBrowse = new System.Windows.Forms.Button();
            this.txtSaveFile = new System.Windows.Forms.TextBox();
            this.label1 = new System.Windows.Forms.Label();
            this.btnCancel = new System.Windows.Forms.Button();
            this.btnExport = new System.Windows.Forms.Button();
            this.groupBox3 = new System.Windows.Forms.GroupBox();
            this.lblWelcomeMessage = new System.Windows.Forms.Label();
            this.btnAbout = new System.Windows.Forms.Button();
            this.saveFileDialog = new System.Windows.Forms.SaveFileDialog();
            ((System.ComponentModel.ISupportInitialize)(this.pictureBox1)).BeginInit();
            this.groupBox1.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.bsData)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.testSuitesBindingSource)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.testPlansBindingSource)).BeginInit();
            this.groupBox2.SuspendLayout();
            this.groupBox3.SuspendLayout();
            this.SuspendLayout();
            // 
            // pictureBox1
            // 
            this.pictureBox1.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
            | System.Windows.Forms.AnchorStyles.Left)));
            this.pictureBox1.Image = ((System.Drawing.Image)(resources.GetObject("pictureBox1.Image")));
            this.pictureBox1.Location = new System.Drawing.Point(10, 12);
            this.pictureBox1.Name = "pictureBox1";
            this.pictureBox1.Size = new System.Drawing.Size(235, 495);
            this.pictureBox1.TabIndex = 0;
            this.pictureBox1.TabStop = false;
            // 
            // groupBox1
            // 
            this.groupBox1.Controls.Add(this.comBoxTestSuite);
            this.groupBox1.Controls.Add(this.lblTestSuite);
            this.groupBox1.Controls.Add(this.comBoxTestPlan);
            this.groupBox1.Controls.Add(this.lblTestPlan);
            this.groupBox1.Controls.Add(this.btnTeamProject);
            this.groupBox1.Controls.Add(this.txtTeamProject);
            this.groupBox1.Controls.Add(this.lblTeamProject);
            this.groupBox1.Font = new System.Drawing.Font("Calibri", 11F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.groupBox1.Location = new System.Drawing.Point(260, 164);
            this.groupBox1.Name = "groupBox1";
            this.groupBox1.Size = new System.Drawing.Size(529, 195);
            this.groupBox1.TabIndex = 0;
            this.groupBox1.TabStop = false;
            this.groupBox1.Text = "Source";
            // 
            // comBoxTestSuite
            // 
            this.comBoxTestSuite.DataSource = this.testSuitesBindingSource;
            this.comBoxTestSuite.DisplayMember = "Title";
            this.comBoxTestSuite.ValueMember = "Title";
            this.comBoxTestSuite.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
            this.comBoxTestSuite.Font = new System.Drawing.Font("Calibri", 10F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.comBoxTestSuite.FormattingEnabled = true;
            this.comBoxTestSuite.Location = new System.Drawing.Point(11, 158);
            this.comBoxTestSuite.Name = "comBoxTestSuite";
            this.comBoxTestSuite.Size = new System.Drawing.Size(452, 23);
            this.comBoxTestSuite.TabIndex = 3;
            // 
            // bsData
            // 
            this.bsData.AllowNew = false;
            this.bsData.DataSource = typeof(TestCaseExport.Data);
            // 
            // testSuitesBindingSource
            // 
            this.testSuitesBindingSource.DataMember = "TestSuites";
            this.testSuitesBindingSource.DataSource = this.bsData;
            // 
            // lblTestSuite
            // 
            this.lblTestSuite.AutoSize = true;
            this.lblTestSuite.Font = new System.Drawing.Font("Calibri", 10F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.lblTestSuite.Location = new System.Drawing.Point(6, 138);
            this.lblTestSuite.Name = "lblTestSuite";
            this.lblTestSuite.Size = new System.Drawing.Size(66, 17);
            this.lblTestSuite.TabIndex = 0;
            this.lblTestSuite.Text = "Test Suite:";
            // 
            // comBoxTestPlan
            // 
            this.comBoxTestPlan.DataSource = this.testPlansBindingSource;
            this.comBoxTestPlan.DisplayMember = "Name";
            this.comBoxTestPlan.ValueMember = "Name";
            this.comBoxTestPlan.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
            this.comBoxTestPlan.Font = new System.Drawing.Font("Calibri", 10F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.comBoxTestPlan.FormattingEnabled = true;
            this.comBoxTestPlan.Location = new System.Drawing.Point(11, 99);
            this.comBoxTestPlan.Name = "comBoxTestPlan";
            this.comBoxTestPlan.Size = new System.Drawing.Size(452, 23);
            this.comBoxTestPlan.TabIndex = 2;
            // 
            // testPlansBindingSource
            // 
            this.testPlansBindingSource.DataMember = "TestPlans";
            this.testPlansBindingSource.DataSource = this.bsData;
            // 
            // lblTestPlan
            // 
            this.lblTestPlan.AutoSize = true;
            this.lblTestPlan.Font = new System.Drawing.Font("Calibri", 10F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.lblTestPlan.Location = new System.Drawing.Point(6, 79);
            this.lblTestPlan.Name = "lblTestPlan";
            this.lblTestPlan.Size = new System.Drawing.Size(62, 17);
            this.lblTestPlan.TabIndex = 0;
            this.lblTestPlan.Text = "Test Plan:";
            // 
            // btnTeamProject
            // 
            this.btnTeamProject.Location = new System.Drawing.Point(475, 44);
            this.btnTeamProject.Name = "btnTeamProject";
            this.btnTeamProject.Size = new System.Drawing.Size(35, 23);
            this.btnTeamProject.TabIndex = 1;
            this.btnTeamProject.Text = "...";
            this.btnTeamProject.UseVisualStyleBackColor = true;
            this.btnTeamProject.Click += new System.EventHandler(this.btnTeamProject_Click);
            // 
            // txtTeamProject
            // 
            this.txtTeamProject.DataBindings.Add(new System.Windows.Forms.Binding("Text", this.bsData, "SelectedProjectName", true));
            this.txtTeamProject.Font = new System.Drawing.Font("Calibri", 10F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.txtTeamProject.Location = new System.Drawing.Point(11, 43);
            this.txtTeamProject.Name = "txtTeamProject";
            this.txtTeamProject.ReadOnly = true;
            this.txtTeamProject.Size = new System.Drawing.Size(453, 24);
            this.txtTeamProject.TabIndex = 0;
            // 
            // lblTeamProject
            // 
            this.lblTeamProject.AutoSize = true;
            this.lblTeamProject.Font = new System.Drawing.Font("Calibri", 10F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.lblTeamProject.Location = new System.Drawing.Point(6, 21);
            this.lblTeamProject.Name = "lblTeamProject";
            this.lblTeamProject.Size = new System.Drawing.Size(87, 17);
            this.lblTeamProject.TabIndex = 0;
            this.lblTeamProject.Text = "Team-Project:";
            // 
            // groupBox2
            // 
            this.groupBox2.Controls.Add(this.btnFolderBrowse);
            this.groupBox2.Controls.Add(this.txtSaveFile);
            this.groupBox2.Controls.Add(this.label1);
            this.groupBox2.Font = new System.Drawing.Font("Calibri", 11.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.groupBox2.Location = new System.Drawing.Point(260, 374);
            this.groupBox2.Name = "groupBox2";
            this.groupBox2.Size = new System.Drawing.Size(529, 82);
            this.groupBox2.TabIndex = 0;
            this.groupBox2.TabStop = false;
            this.groupBox2.Text = "Destination";
            // 
            // btnFolderBrowse
            // 
            this.btnFolderBrowse.Font = new System.Drawing.Font("Calibri", 11.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.btnFolderBrowse.Location = new System.Drawing.Point(428, 42);
            this.btnFolderBrowse.Name = "btnFolderBrowse";
            this.btnFolderBrowse.Size = new System.Drawing.Size(82, 25);
            this.btnFolderBrowse.TabIndex = 4;
            this.btnFolderBrowse.Text = "Browse...";
            this.btnFolderBrowse.UseVisualStyleBackColor = true;
            this.btnFolderBrowse.Click += new System.EventHandler(this.btnFolderBrowse_Click);
            // 
            // txtSaveFile
            // 
            this.txtSaveFile.DataBindings.Add(new System.Windows.Forms.Binding("Text", this.bsData, "ExportFileName", true));
            this.txtSaveFile.Font = new System.Drawing.Font("Calibri", 10F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.txtSaveFile.Location = new System.Drawing.Point(11, 42);
            this.txtSaveFile.Name = "txtSaveFile";
            this.txtSaveFile.ReadOnly = true;
            this.txtSaveFile.Size = new System.Drawing.Size(411, 24);
            this.txtSaveFile.TabIndex = 7;
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Font = new System.Drawing.Font("Calibri", 10F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label1.Location = new System.Drawing.Point(6, 22);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(206, 17);
            this.label1.TabIndex = 0;
            this.label1.Text = "Specify File to Save Exported Script:";
            // 
            // btnCancel
            // 
            this.btnCancel.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right)));
            this.btnCancel.DialogResult = System.Windows.Forms.DialogResult.Cancel;
            this.btnCancel.Font = new System.Drawing.Font("Calibri", 10F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.btnCancel.Location = new System.Drawing.Point(707, 482);
            this.btnCancel.Name = "btnCancel";
            this.btnCancel.Size = new System.Drawing.Size(82, 25);
            this.btnCancel.TabIndex = 7;
            this.btnCancel.Text = "Cancel";
            this.btnCancel.UseVisualStyleBackColor = true;
            this.btnCancel.Click += new System.EventHandler(this.btnCancel_Click);
            // 
            // btnExport
            // 
            this.btnExport.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right)));
            this.btnExport.DataBindings.Add(new System.Windows.Forms.Binding("Enabled", this.bsData, "SuiteIsSelected", true, System.Windows.Forms.DataSourceUpdateMode.OnPropertyChanged));
            this.btnExport.Font = new System.Drawing.Font("Calibri", 10F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.btnExport.Location = new System.Drawing.Point(617, 482);
            this.btnExport.Name = "btnExport";
            this.btnExport.Size = new System.Drawing.Size(82, 25);
            this.btnExport.TabIndex = 6;
            this.btnExport.Text = "Export";
            this.btnExport.UseVisualStyleBackColor = true;
            this.btnExport.Click += new System.EventHandler(this.btnExport_Click);
            // 
            // groupBox3
            // 
            this.groupBox3.Controls.Add(this.lblWelcomeMessage);
            this.groupBox3.Font = new System.Drawing.Font("Calibri", 11F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.groupBox3.Location = new System.Drawing.Point(260, 10);
            this.groupBox3.Name = "groupBox3";
            this.groupBox3.Size = new System.Drawing.Size(529, 127);
            this.groupBox3.TabIndex = 0;
            this.groupBox3.TabStop = false;
            this.groupBox3.Text = "Welcome";
            // 
            // lblWelcomeMessage
            // 
            this.lblWelcomeMessage.AutoSize = true;
            this.lblWelcomeMessage.Font = new System.Drawing.Font("Calibri", 10F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.lblWelcomeMessage.Location = new System.Drawing.Point(6, 18);
            this.lblWelcomeMessage.Name = "lblWelcomeMessage";
            this.lblWelcomeMessage.Size = new System.Drawing.Size(501, 102);
            this.lblWelcomeMessage.TabIndex = 4;
            this.lblWelcomeMessage.Text = resources.GetString("lblWelcomeMessage.Text");
            // 
            // btnAbout
            // 
            this.btnAbout.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
            this.btnAbout.Font = new System.Drawing.Font("Calibri", 10F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.btnAbout.Location = new System.Drawing.Point(260, 482);
            this.btnAbout.Name = "btnAbout";
            this.btnAbout.Size = new System.Drawing.Size(82, 25);
            this.btnAbout.TabIndex = 8;
            this.btnAbout.Text = "About";
            this.btnAbout.UseVisualStyleBackColor = true;
            this.btnAbout.Click += new System.EventHandler(this.btnAbout_Click);
            // 
            // saveFileDialog
            // 
            this.saveFileDialog.DefaultExt = "xlsx";
            this.saveFileDialog.Filter = "Excel files|*.xlsx";
            this.saveFileDialog.SupportMultiDottedExtensions = true;
            // 
            // FrmMain
            // 
            this.AcceptButton = this.btnExport;
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.BackColor = System.Drawing.SystemColors.Control;
            this.CancelButton = this.btnCancel;
            this.ClientSize = new System.Drawing.Size(801, 518);
            this.Controls.Add(this.btnAbout);
            this.Controls.Add(this.groupBox3);
            this.Controls.Add(this.btnExport);
            this.Controls.Add(this.btnCancel);
            this.Controls.Add(this.groupBox2);
            this.Controls.Add(this.groupBox1);
            this.Controls.Add(this.pictureBox1);
            this.Font = new System.Drawing.Font("Calibri", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedSingle;
            this.Icon = ((System.Drawing.Icon)(resources.GetObject("$this.Icon")));
            this.MaximizeBox = false;
            this.Name = "FrmMain";
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
            this.Text = "Test Case Export to Excel (Beta)";
            ((System.ComponentModel.ISupportInitialize)(this.pictureBox1)).EndInit();
            this.groupBox1.ResumeLayout(false);
            this.groupBox1.PerformLayout();
            ((System.ComponentModel.ISupportInitialize)(this.bsData)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.testSuitesBindingSource)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.testPlansBindingSource)).EndInit();
            this.groupBox2.ResumeLayout(false);
            this.groupBox2.PerformLayout();
            this.groupBox3.ResumeLayout(false);
            this.groupBox3.PerformLayout();
            this.ResumeLayout(false);

        }

        #endregion

        private System.Windows.Forms.PictureBox pictureBox1;
        private System.Windows.Forms.GroupBox groupBox1;
        private System.Windows.Forms.ComboBox comBoxTestSuite;
        private System.Windows.Forms.Label lblTestSuite;
        private System.Windows.Forms.ComboBox comBoxTestPlan;
        private System.Windows.Forms.Label lblTestPlan;
        private System.Windows.Forms.Button btnTeamProject;
        private System.Windows.Forms.TextBox txtTeamProject;
        private System.Windows.Forms.Label lblTeamProject;
        private System.Windows.Forms.GroupBox groupBox2;
        private System.Windows.Forms.Button btnFolderBrowse;
        private System.Windows.Forms.TextBox txtSaveFile;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.Button btnCancel;
        private System.Windows.Forms.Button btnExport;
        private System.Windows.Forms.GroupBox groupBox3;
        private System.Windows.Forms.Label lblWelcomeMessage;
        private System.Windows.Forms.Button btnAbout;
        private System.Windows.Forms.BindingSource bsData;
        private System.Windows.Forms.BindingSource testPlansBindingSource;
        private System.Windows.Forms.BindingSource testSuitesBindingSource;
        private System.Windows.Forms.SaveFileDialog saveFileDialog;
    }
}

