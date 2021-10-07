namespace ExcelFileCleaner
{
    partial class Form1
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
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(Form1));
            this.button1 = new System.Windows.Forms.Button();
            this.groupBox1 = new System.Windows.Forms.GroupBox();
            this.lblUsage = new System.Windows.Forms.Label();
            this.linkLabel2 = new System.Windows.Forms.LinkLabel();
            this.linkLabel1 = new System.Windows.Forms.LinkLabel();
            this.label1 = new System.Windows.Forms.Label();
            this.pictureBox1 = new System.Windows.Forms.PictureBox();
            this.groupBox2 = new System.Windows.Forms.GroupBox();
            this.lstLog = new System.Windows.Forms.ListBox();
            this.pbMain = new System.Windows.Forms.ProgressBar();
            this.btnClose = new System.Windows.Forms.Button();
            this.chkBackup = new System.Windows.Forms.CheckBox();
            this.chkOpen = new System.Windows.Forms.CheckBox();
            this.chkLog = new System.Windows.Forms.CheckBox();
            this.statusStrip1 = new System.Windows.Forms.StatusStrip();
            this.tsLabel = new System.Windows.Forms.ToolStripStatusLabel();
            this.tsProgress = new System.Windows.Forms.ToolStripProgressBar();
            this.chkHide = new System.Windows.Forms.CheckBox();
            this.groupBox3 = new System.Windows.Forms.GroupBox();
            this.groupBox4 = new System.Windows.Forms.GroupBox();
            this.rdoCorrectAll = new System.Windows.Forms.RadioButton();
            this.rdoCorrectOnly = new System.Windows.Forms.RadioButton();
            this.rdoCorrectDefault = new System.Windows.Forms.RadioButton();
            this.correctGroup = new System.Windows.Forms.GroupBox();
            this.chkFormulas = new System.Windows.Forms.CheckBox();
            this.chkLinks = new System.Windows.Forms.CheckBox();
            this.chkRanges = new System.Windows.Forms.CheckBox();
            this.chkCells = new System.Windows.Forms.CheckBox();
            this.chkStyles = new System.Windows.Forms.CheckBox();
            this.groupBox6 = new System.Windows.Forms.GroupBox();
            this.optionsGroup = new System.Windows.Forms.GroupBox();
            this.tabControl1 = new System.Windows.Forms.TabControl();
            this.tabPage1 = new System.Windows.Forms.TabPage();
            this.tabPage2 = new System.Windows.Forms.TabPage();
            this.tabPage3 = new System.Windows.Forms.TabPage();
            this.groupBox1.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.pictureBox1)).BeginInit();
            this.groupBox2.SuspendLayout();
            this.statusStrip1.SuspendLayout();
            this.groupBox3.SuspendLayout();
            this.groupBox4.SuspendLayout();
            this.correctGroup.SuspendLayout();
            this.groupBox6.SuspendLayout();
            this.optionsGroup.SuspendLayout();
            this.tabControl1.SuspendLayout();
            this.tabPage1.SuspendLayout();
            this.tabPage2.SuspendLayout();
            this.tabPage3.SuspendLayout();
            this.SuspendLayout();
            // 
            // button1
            // 
            this.button1.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.button1.Location = new System.Drawing.Point(1090, 400);
            this.button1.Margin = new System.Windows.Forms.Padding(6);
            this.button1.Name = "button1";
            this.button1.Size = new System.Drawing.Size(208, 84);
            this.button1.TabIndex = 1;
            this.button1.Text = "Open and Clean";
            this.button1.UseVisualStyleBackColor = true;
            this.button1.Click += new System.EventHandler(this.button1_Click);
            // 
            // groupBox1
            // 
            this.groupBox1.Controls.Add(this.lblUsage);
            this.groupBox1.Controls.Add(this.linkLabel2);
            this.groupBox1.Controls.Add(this.linkLabel1);
            this.groupBox1.Controls.Add(this.label1);
            this.groupBox1.Controls.Add(this.pictureBox1);
            this.groupBox1.Location = new System.Drawing.Point(12, 11);
            this.groupBox1.Margin = new System.Windows.Forms.Padding(6);
            this.groupBox1.Name = "groupBox1";
            this.groupBox1.Padding = new System.Windows.Forms.Padding(6);
            this.groupBox1.Size = new System.Drawing.Size(1014, 483);
            this.groupBox1.TabIndex = 4;
            this.groupBox1.TabStop = false;
            // 
            // lblUsage
            // 
            this.lblUsage.AutoSize = true;
            this.lblUsage.Location = new System.Drawing.Point(64, 323);
            this.lblUsage.Margin = new System.Windows.Forms.Padding(6, 0, 6, 0);
            this.lblUsage.Name = "lblUsage";
            this.lblUsage.Size = new System.Drawing.Size(112, 25);
            this.lblUsage.TabIndex = 8;
            this.lblUsage.Text = "This tool...";
            // 
            // linkLabel2
            // 
            this.linkLabel2.AutoSize = true;
            this.linkLabel2.Location = new System.Drawing.Point(104, 250);
            this.linkLabel2.Margin = new System.Windows.Forms.Padding(6, 0, 6, 0);
            this.linkLabel2.Name = "linkLabel2";
            this.linkLabel2.Size = new System.Drawing.Size(110, 25);
            this.linkLabel2.TabIndex = 2;
            this.linkLabel2.TabStop = true;
            this.linkLabel2.Text = "linkLabel2";
            this.linkLabel2.LinkClicked += new System.Windows.Forms.LinkLabelLinkClickedEventHandler(this.linkLabel2_LinkClicked);
            // 
            // linkLabel1
            // 
            this.linkLabel1.AutoSize = true;
            this.linkLabel1.Location = new System.Drawing.Point(104, 194);
            this.linkLabel1.Margin = new System.Windows.Forms.Padding(6, 0, 6, 0);
            this.linkLabel1.Name = "linkLabel1";
            this.linkLabel1.Size = new System.Drawing.Size(110, 25);
            this.linkLabel1.TabIndex = 1;
            this.linkLabel1.TabStop = true;
            this.linkLabel1.Text = "linkLabel1";
            this.linkLabel1.LinkClicked += new System.Windows.Forms.LinkLabelLinkClickedEventHandler(this.linkLabel1_LinkClicked);
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Location = new System.Drawing.Point(64, 31);
            this.label1.Margin = new System.Windows.Forms.Padding(6, 0, 6, 0);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(112, 25);
            this.label1.TabIndex = 0;
            this.label1.Text = "This tool...";
            // 
            // pictureBox1
            // 
            this.pictureBox1.Image = global::ExcelFileCleaner.Properties.Resources.Picture1;
            this.pictureBox1.Location = new System.Drawing.Point(819, 27);
            this.pictureBox1.Margin = new System.Windows.Forms.Padding(6);
            this.pictureBox1.Name = "pictureBox1";
            this.pictureBox1.Size = new System.Drawing.Size(183, 192);
            this.pictureBox1.SizeMode = System.Windows.Forms.PictureBoxSizeMode.StretchImage;
            this.pictureBox1.TabIndex = 7;
            this.pictureBox1.TabStop = false;
            // 
            // groupBox2
            // 
            this.groupBox2.Controls.Add(this.lstLog);
            this.groupBox2.Location = new System.Drawing.Point(6, 6);
            this.groupBox2.Margin = new System.Windows.Forms.Padding(6);
            this.groupBox2.Name = "groupBox2";
            this.groupBox2.Padding = new System.Windows.Forms.Padding(6);
            this.groupBox2.Size = new System.Drawing.Size(1026, 494);
            this.groupBox2.TabIndex = 5;
            this.groupBox2.TabStop = false;
            this.groupBox2.Text = "Log:";
            // 
            // lstLog
            // 
            this.lstLog.Font = new System.Drawing.Font("Courier New", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.lstLog.FormattingEnabled = true;
            this.lstLog.ItemHeight = 24;
            this.lstLog.Location = new System.Drawing.Point(12, 41);
            this.lstLog.Margin = new System.Windows.Forms.Padding(6);
            this.lstLog.Name = "lstLog";
            this.lstLog.Size = new System.Drawing.Size(998, 412);
            this.lstLog.TabIndex = 0;
            // 
            // pbMain
            // 
            this.pbMain.Location = new System.Drawing.Point(24, 591);
            this.pbMain.Margin = new System.Windows.Forms.Padding(6);
            this.pbMain.Name = "pbMain";
            this.pbMain.Size = new System.Drawing.Size(1274, 44);
            this.pbMain.TabIndex = 0;
            // 
            // btnClose
            // 
            this.btnClose.Location = new System.Drawing.Point(1090, 494);
            this.btnClose.Margin = new System.Windows.Forms.Padding(6);
            this.btnClose.Name = "btnClose";
            this.btnClose.Size = new System.Drawing.Size(208, 84);
            this.btnClose.TabIndex = 6;
            this.btnClose.Text = "Close";
            this.btnClose.UseVisualStyleBackColor = true;
            this.btnClose.Click += new System.EventHandler(this.button2_Click);
            // 
            // chkBackup
            // 
            this.chkBackup.AutoSize = true;
            this.chkBackup.Checked = true;
            this.chkBackup.CheckState = System.Windows.Forms.CheckState.Checked;
            this.chkBackup.Location = new System.Drawing.Point(12, 39);
            this.chkBackup.Margin = new System.Windows.Forms.Padding(6);
            this.chkBackup.Name = "chkBackup";
            this.chkBackup.Size = new System.Drawing.Size(226, 29);
            this.chkBackup.TabIndex = 8;
            this.chkBackup.Text = "Backup original file";
            this.chkBackup.UseVisualStyleBackColor = true;
            // 
            // chkOpen
            // 
            this.chkOpen.AutoSize = true;
            this.chkOpen.Location = new System.Drawing.Point(12, 81);
            this.chkOpen.Margin = new System.Windows.Forms.Padding(6);
            this.chkOpen.Name = "chkOpen";
            this.chkOpen.Size = new System.Drawing.Size(280, 29);
            this.chkOpen.TabIndex = 9;
            this.chkOpen.Text = "Open file when complete";
            this.chkOpen.UseVisualStyleBackColor = true;
            // 
            // chkLog
            // 
            this.chkLog.AutoSize = true;
            this.chkLog.Location = new System.Drawing.Point(12, 41);
            this.chkLog.Margin = new System.Windows.Forms.Padding(6);
            this.chkLog.Name = "chkLog";
            this.chkLog.Size = new System.Drawing.Size(128, 29);
            this.chkLog.TabIndex = 10;
            this.chkLog.Text = "Save log";
            this.chkLog.UseVisualStyleBackColor = true;
            // 
            // statusStrip1
            // 
            this.statusStrip1.ImageScalingSize = new System.Drawing.Size(32, 32);
            this.statusStrip1.Items.AddRange(new System.Windows.Forms.ToolStripItem[] {
            this.tsLabel,
            this.tsProgress});
            this.statusStrip1.Location = new System.Drawing.Point(0, 661);
            this.statusStrip1.Name = "statusStrip1";
            this.statusStrip1.Padding = new System.Windows.Forms.Padding(2, 0, 28, 0);
            this.statusStrip1.Size = new System.Drawing.Size(1312, 22);
            this.statusStrip1.TabIndex = 11;
            this.statusStrip1.Text = "statusStrip1";
            // 
            // tsLabel
            // 
            this.tsLabel.Name = "tsLabel";
            this.tsLabel.Size = new System.Drawing.Size(30, 32);
            this.tsLabel.Text = "...";
            this.tsLabel.Visible = false;
            // 
            // tsProgress
            // 
            this.tsProgress.Name = "tsProgress";
            this.tsProgress.Size = new System.Drawing.Size(200, 16);
            this.tsProgress.Visible = false;
            // 
            // chkHide
            // 
            this.chkHide.AutoSize = true;
            this.chkHide.Location = new System.Drawing.Point(12, 84);
            this.chkHide.Margin = new System.Windows.Forms.Padding(6);
            this.chkHide.Name = "chkHide";
            this.chkHide.Size = new System.Drawing.Size(432, 29);
            this.chkHide.TabIndex = 13;
            this.chkHide.Text = "Hide logging (faster, uses more memory)";
            this.chkHide.UseVisualStyleBackColor = true;
            // 
            // groupBox3
            // 
            this.groupBox3.Controls.Add(this.chkLog);
            this.groupBox3.Controls.Add(this.chkHide);
            this.groupBox3.Location = new System.Drawing.Point(538, 186);
            this.groupBox3.Margin = new System.Windows.Forms.Padding(6);
            this.groupBox3.Name = "groupBox3";
            this.groupBox3.Padding = new System.Windows.Forms.Padding(6);
            this.groupBox3.Size = new System.Drawing.Size(460, 125);
            this.groupBox3.TabIndex = 15;
            this.groupBox3.TabStop = false;
            this.groupBox3.Text = "Log Options";
            // 
            // groupBox4
            // 
            this.groupBox4.Controls.Add(this.rdoCorrectAll);
            this.groupBox4.Controls.Add(this.rdoCorrectOnly);
            this.groupBox4.Controls.Add(this.rdoCorrectDefault);
            this.groupBox4.Controls.Add(this.correctGroup);
            this.groupBox4.Location = new System.Drawing.Point(12, 47);
            this.groupBox4.Margin = new System.Windows.Forms.Padding(6);
            this.groupBox4.Name = "groupBox4";
            this.groupBox4.Padding = new System.Windows.Forms.Padding(6);
            this.groupBox4.Size = new System.Drawing.Size(514, 419);
            this.groupBox4.TabIndex = 16;
            this.groupBox4.TabStop = false;
            this.groupBox4.Text = "Correction Options";
            // 
            // rdoCorrectAll
            // 
            this.rdoCorrectAll.AutoSize = true;
            this.rdoCorrectAll.Location = new System.Drawing.Point(14, 83);
            this.rdoCorrectAll.Margin = new System.Windows.Forms.Padding(6);
            this.rdoCorrectAll.Name = "rdoCorrectAll";
            this.rdoCorrectAll.Size = new System.Drawing.Size(143, 29);
            this.rdoCorrectAll.TabIndex = 3;
            this.rdoCorrectAll.Text = "Correct All";
            this.rdoCorrectAll.UseVisualStyleBackColor = true;
            // 
            // rdoCorrectOnly
            // 
            this.rdoCorrectOnly.AutoSize = true;
            this.rdoCorrectOnly.Location = new System.Drawing.Point(14, 127);
            this.rdoCorrectOnly.Margin = new System.Windows.Forms.Padding(6);
            this.rdoCorrectOnly.Name = "rdoCorrectOnly";
            this.rdoCorrectOnly.Size = new System.Drawing.Size(165, 29);
            this.rdoCorrectOnly.TabIndex = 1;
            this.rdoCorrectOnly.Text = "Correct only:";
            this.rdoCorrectOnly.UseVisualStyleBackColor = true;
            this.rdoCorrectOnly.CheckedChanged += new System.EventHandler(this.rdoCorrectOnly_CheckedChanged);
            // 
            // rdoCorrectDefault
            // 
            this.rdoCorrectDefault.AutoSize = true;
            this.rdoCorrectDefault.Checked = true;
            this.rdoCorrectDefault.Location = new System.Drawing.Point(14, 39);
            this.rdoCorrectDefault.Margin = new System.Windows.Forms.Padding(6);
            this.rdoCorrectDefault.Name = "rdoCorrectDefault";
            this.rdoCorrectDefault.Size = new System.Drawing.Size(412, 29);
            this.rdoCorrectDefault.TabIndex = 0;
            this.rdoCorrectDefault.TabStop = true;
            this.rdoCorrectDefault.Text = "Correct (Styles, Cells, Named Ranges)";
            this.rdoCorrectDefault.UseVisualStyleBackColor = true;
            // 
            // correctGroup
            // 
            this.correctGroup.Controls.Add(this.chkFormulas);
            this.correctGroup.Controls.Add(this.chkLinks);
            this.correctGroup.Controls.Add(this.chkRanges);
            this.correctGroup.Controls.Add(this.chkCells);
            this.correctGroup.Controls.Add(this.chkStyles);
            this.correctGroup.Enabled = false;
            this.correctGroup.Location = new System.Drawing.Point(52, 127);
            this.correctGroup.Margin = new System.Windows.Forms.Padding(6);
            this.correctGroup.Name = "correctGroup";
            this.correctGroup.Padding = new System.Windows.Forms.Padding(6);
            this.correctGroup.Size = new System.Drawing.Size(450, 275);
            this.correctGroup.TabIndex = 2;
            this.correctGroup.TabStop = false;
            // 
            // chkFormulas
            // 
            this.chkFormulas.AutoSize = true;
            this.chkFormulas.Location = new System.Drawing.Point(14, 223);
            this.chkFormulas.Margin = new System.Windows.Forms.Padding(6);
            this.chkFormulas.Name = "chkFormulas";
            this.chkFormulas.Size = new System.Drawing.Size(302, 29);
            this.chkFormulas.TabIndex = 4;
            this.chkFormulas.Text = "Remove linked formulas (*)";
            this.chkFormulas.UseVisualStyleBackColor = true;
            this.chkFormulas.CheckedChanged += new System.EventHandler(this.chkFormulas_CheckedChanged);
            // 
            // chkLinks
            // 
            this.chkLinks.AutoSize = true;
            this.chkLinks.Location = new System.Drawing.Point(14, 177);
            this.chkLinks.Margin = new System.Windows.Forms.Padding(6);
            this.chkLinks.Name = "chkLinks";
            this.chkLinks.Size = new System.Drawing.Size(201, 29);
            this.chkLinks.TabIndex = 3;
            this.chkLinks.Text = "Remove links (*)";
            this.chkLinks.UseVisualStyleBackColor = true;
            this.chkLinks.CheckedChanged += new System.EventHandler(this.chkLinks_CheckedChanged);
            // 
            // chkRanges
            // 
            this.chkRanges.AutoSize = true;
            this.chkRanges.Location = new System.Drawing.Point(14, 131);
            this.chkRanges.Margin = new System.Windows.Forms.Padding(6);
            this.chkRanges.Name = "chkRanges";
            this.chkRanges.Size = new System.Drawing.Size(334, 29);
            this.chkRanges.TabIndex = 2;
            this.chkRanges.Text = "Remove invalid named ranges";
            this.chkRanges.UseVisualStyleBackColor = true;
            // 
            // chkCells
            // 
            this.chkCells.AutoSize = true;
            this.chkCells.Location = new System.Drawing.Point(14, 84);
            this.chkCells.Margin = new System.Windows.Forms.Padding(6);
            this.chkCells.Name = "chkCells";
            this.chkCells.Size = new System.Drawing.Size(250, 29);
            this.chkCells.TabIndex = 1;
            this.chkCells.Text = "Remove unused cells";
            this.chkCells.UseVisualStyleBackColor = true;
            // 
            // chkStyles
            // 
            this.chkStyles.AutoSize = true;
            this.chkStyles.Location = new System.Drawing.Point(14, 39);
            this.chkStyles.Margin = new System.Windows.Forms.Padding(6);
            this.chkStyles.Name = "chkStyles";
            this.chkStyles.Size = new System.Drawing.Size(257, 29);
            this.chkStyles.TabIndex = 0;
            this.chkStyles.Text = "Reduce unused styles";
            this.chkStyles.UseVisualStyleBackColor = true;
            // 
            // groupBox6
            // 
            this.groupBox6.Controls.Add(this.chkBackup);
            this.groupBox6.Controls.Add(this.chkOpen);
            this.groupBox6.Location = new System.Drawing.Point(538, 47);
            this.groupBox6.Margin = new System.Windows.Forms.Padding(6);
            this.groupBox6.Name = "groupBox6";
            this.groupBox6.Padding = new System.Windows.Forms.Padding(6);
            this.groupBox6.Size = new System.Drawing.Size(460, 127);
            this.groupBox6.TabIndex = 17;
            this.groupBox6.TabStop = false;
            this.groupBox6.Text = "Basic Options";
            // 
            // optionsGroup
            // 
            this.optionsGroup.Controls.Add(this.groupBox6);
            this.optionsGroup.Controls.Add(this.groupBox4);
            this.optionsGroup.Controls.Add(this.groupBox3);
            this.optionsGroup.Location = new System.Drawing.Point(6, 11);
            this.optionsGroup.Margin = new System.Windows.Forms.Padding(6);
            this.optionsGroup.Name = "optionsGroup";
            this.optionsGroup.Padding = new System.Windows.Forms.Padding(6);
            this.optionsGroup.Size = new System.Drawing.Size(1016, 478);
            this.optionsGroup.TabIndex = 19;
            this.optionsGroup.TabStop = false;
            this.optionsGroup.Text = "Options";
            // 
            // tabControl1
            // 
            this.tabControl1.Controls.Add(this.tabPage1);
            this.tabControl1.Controls.Add(this.tabPage2);
            this.tabControl1.Controls.Add(this.tabPage3);
            this.tabControl1.Location = new System.Drawing.Point(24, 23);
            this.tabControl1.Margin = new System.Windows.Forms.Padding(6);
            this.tabControl1.Name = "tabControl1";
            this.tabControl1.SelectedIndex = 0;
            this.tabControl1.Size = new System.Drawing.Size(1054, 556);
            this.tabControl1.TabIndex = 20;
            // 
            // tabPage1
            // 
            this.tabPage1.Controls.Add(this.groupBox1);
            this.tabPage1.Location = new System.Drawing.Point(8, 39);
            this.tabPage1.Margin = new System.Windows.Forms.Padding(6);
            this.tabPage1.Name = "tabPage1";
            this.tabPage1.Padding = new System.Windows.Forms.Padding(6);
            this.tabPage1.Size = new System.Drawing.Size(1038, 509);
            this.tabPage1.TabIndex = 0;
            this.tabPage1.Text = "About";
            this.tabPage1.UseVisualStyleBackColor = true;
            // 
            // tabPage2
            // 
            this.tabPage2.Controls.Add(this.optionsGroup);
            this.tabPage2.Location = new System.Drawing.Point(8, 39);
            this.tabPage2.Margin = new System.Windows.Forms.Padding(6);
            this.tabPage2.Name = "tabPage2";
            this.tabPage2.Padding = new System.Windows.Forms.Padding(6);
            this.tabPage2.Size = new System.Drawing.Size(1038, 509);
            this.tabPage2.TabIndex = 1;
            this.tabPage2.Text = "Options";
            this.tabPage2.UseVisualStyleBackColor = true;
            // 
            // tabPage3
            // 
            this.tabPage3.Controls.Add(this.groupBox2);
            this.tabPage3.Location = new System.Drawing.Point(8, 39);
            this.tabPage3.Margin = new System.Windows.Forms.Padding(6);
            this.tabPage3.Name = "tabPage3";
            this.tabPage3.Size = new System.Drawing.Size(1038, 509);
            this.tabPage3.TabIndex = 2;
            this.tabPage3.Text = "Log";
            this.tabPage3.UseVisualStyleBackColor = true;
            // 
            // Form1
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(12F, 25F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(1312, 683);
            this.ControlBox = false;
            this.Controls.Add(this.tabControl1);
            this.Controls.Add(this.btnClose);
            this.Controls.Add(this.statusStrip1);
            this.Controls.Add(this.pbMain);
            this.Controls.Add(this.button1);
            this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedSingle;
            this.Icon = ((System.Drawing.Icon)(resources.GetObject("$this.Icon")));
            this.Margin = new System.Windows.Forms.Padding(6);
            this.Name = "Form1";
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
            this.Text = "Excel File Cleaner";
            this.Load += new System.EventHandler(this.Form1_Load);
            this.groupBox1.ResumeLayout(false);
            this.groupBox1.PerformLayout();
            ((System.ComponentModel.ISupportInitialize)(this.pictureBox1)).EndInit();
            this.groupBox2.ResumeLayout(false);
            this.statusStrip1.ResumeLayout(false);
            this.statusStrip1.PerformLayout();
            this.groupBox3.ResumeLayout(false);
            this.groupBox3.PerformLayout();
            this.groupBox4.ResumeLayout(false);
            this.groupBox4.PerformLayout();
            this.correctGroup.ResumeLayout(false);
            this.correctGroup.PerformLayout();
            this.groupBox6.ResumeLayout(false);
            this.groupBox6.PerformLayout();
            this.optionsGroup.ResumeLayout(false);
            this.tabControl1.ResumeLayout(false);
            this.tabPage1.ResumeLayout(false);
            this.tabPage2.ResumeLayout(false);
            this.tabPage3.ResumeLayout(false);
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.Button button1;
        private System.Windows.Forms.GroupBox groupBox1;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.GroupBox groupBox2;
        private System.Windows.Forms.ProgressBar pbMain;
        private System.Windows.Forms.ListBox lstLog;
        private System.Windows.Forms.Button btnClose;
        private System.Windows.Forms.PictureBox pictureBox1;
        private System.Windows.Forms.LinkLabel linkLabel2;
        private System.Windows.Forms.LinkLabel linkLabel1;
        private System.Windows.Forms.CheckBox chkBackup;
        private System.Windows.Forms.CheckBox chkOpen;
        private System.Windows.Forms.CheckBox chkLog;
        private System.Windows.Forms.StatusStrip statusStrip1;
        private System.Windows.Forms.ToolStripStatusLabel tsLabel;
        private System.Windows.Forms.ToolStripProgressBar tsProgress;
        private System.Windows.Forms.CheckBox chkHide;
        private System.Windows.Forms.GroupBox groupBox3;
        private System.Windows.Forms.GroupBox groupBox4;
        private System.Windows.Forms.RadioButton rdoCorrectOnly;
        private System.Windows.Forms.RadioButton rdoCorrectDefault;
        private System.Windows.Forms.GroupBox correctGroup;
        private System.Windows.Forms.CheckBox chkFormulas;
        private System.Windows.Forms.CheckBox chkLinks;
        private System.Windows.Forms.CheckBox chkRanges;
        private System.Windows.Forms.CheckBox chkCells;
        private System.Windows.Forms.CheckBox chkStyles;
        private System.Windows.Forms.RadioButton rdoCorrectAll;
        private System.Windows.Forms.GroupBox groupBox6;
        private System.Windows.Forms.GroupBox optionsGroup;
        private System.Windows.Forms.TabControl tabControl1;
        private System.Windows.Forms.TabPage tabPage1;
        private System.Windows.Forms.TabPage tabPage2;
        private System.Windows.Forms.TabPage tabPage3;
        private System.Windows.Forms.Label lblUsage;
    }
}

