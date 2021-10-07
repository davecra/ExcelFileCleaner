using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using System.IO.Packaging;
using System.IO;
using System.Xml;
using System.Diagnostics;
//using DocumentFormat.OpenXml.Packaging;
//using DocumentFormat.OpenXml.Spreadsheet;
//using DocumentFormat.OpenXml;

namespace ExcelFileCleaner
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
            label1.Text = "This tool is designed to reduce the size and speed up Excel OpenXml \n" +
                          "based workbooks. It will only work with the new XML file formats \n" +
                          "(XLSX and XLSM files). It cleans the files by reducing the number \n" +
                          "of styles in use and the number of activated cells.\n\n" +
                          "See the following Knowledge Base articles for more information:";
            linkLabel1.Text = "(KB213904) You receive a 'Too many different cell formats' error \n" +
                              "message in Excel";
            linkLabel2.Text = "(KB244435) How to reset the last cell in Excel";
            lblUsage.Text = "To use this application, click Options, set your desired settings \n" +
                            "then click Open and Clean to correct the Excel file (XLSX or XLSM).";
        }

        private void button1_Click(object sender, EventArgs e)
        {
            // clear list
            lstLog.Items.Clear();
            tsLabel.Text = "";
            tsLabel.Visible = false;
            tsLabel.Visible = false;
            pbMain.Value = 0;

            btnClose.Enabled = false;
            button1.Enabled = false;
            optionsGroup.Enabled = false;

            // allow the user to browse for the file first...
            OpenFileDialog ofd = new OpenFileDialog();
            ofd.Multiselect = false;
            ofd.Title = "Clean Excel Workbook";
            ofd.Filter = "Excel 2007 and 2010 Workbooks (*.xlsx)|*.xlsx" +
                         "|Excel 2007 and 2010 Macro Workbooks (*.xlsm)|*.xlsm";
            string baseFilename = "";
            if (ofd.ShowDialog() == DialogResult.OK)
            {
                tabControl1.SelectTab(2); // select the log page...

                // we got a file from the user
                if (ofd.FileName.ToLower().EndsWith(".xlsx") || ofd.FileName.ToLower().EndsWith(".xlsm"))
                {
                    try
                    {
                        if (chkBackup.Checked ||
                            MessageBox.Show("Do not want to make a backup copy of the original file?",
                                               "Backup Copy",
                                               MessageBoxButtons.YesNo,
                                               MessageBoxIcon.Asterisk) == DialogResult.Yes)
                        {
                            FileInfo fi = new FileInfo(ofd.FileName);
                            baseFilename = fi.Directory.FullName + "\\" + fi.Name.Substring(0, fi.Name.Length - fi.Extension.Length);
                            fi.CopyTo(baseFilename + "_BACKUP" + fi.Extension);
                        }
                    }
                    catch (Exception ex)
                    {
                        MessageBox.Show("The file could not be backed up. " + ex.Message,
                                    "Excel File Cleaner", MessageBoxButtons.OK, MessageBoxIcon.Error);
                        return; // STOP
                    }

                    clsCleanExcel cleaner = new clsCleanExcel();
                    cleaner.LogUpdated += new clsCleanExcel.CleanExcelEventHandler(cleaner_LogUpdated);
                    cleaner.StatusUpdated += new clsCleanExcel.CleanExcelEventHandler(cleaner_StatusUpdated);
                    cleaner.UpdateProgress += new clsCleanExcel.CleanExcelEventHandler(cleaner_UpdateProgress);
                    cleaner.UpdateSmallProgress += new clsCleanExcel.CleanExcelEventHandler2(cleaner_UpdateSmallProgress);
                    try
                    {
                        pbMain.Maximum = cleaner.getSheetCount(ofd.FileName);
                    }
                    catch { }
                    cleaner.setOptions = compileOptions();
                    lstLog.Items.Add("Starting...");
                    if (!cleaner.CleanFile(ofd.FileName))
                    {
                        MessageBox.Show("The file was unable to be cleaned and saved. An error occurred. Please see the log for more information.", 
                            "Excel File Cleaner", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    }
                    else
                    {
                        // done
                        MessageBox.Show("The file has been cleaned!","Excel File Cleaner",MessageBoxButtons.OK,MessageBoxIcon.Information);
                        // open the file if this is what the user wants
                        if (chkOpen.Checked)
                            launchLink("excel.exe@" + ofd.FileName);
                    }
                }
            }

            // now display the log if the user had hidden it for speed
            if (chkHide.Checked)
            {
                foreach (string s in offlineLog)
                    lstLog.Items.Add(s);
                lstLog.SelectedIndex = lstLog.Items.Count - 1;
            }

            // save the log
            if (chkLog.Checked)
            {
                try
                {
                    StreamWriter sw = new StreamWriter(baseFilename + "_LOG.txt");
                    foreach (string item in lstLog.Items)
                        sw.WriteLine(item);
                    sw.Close();
                }
                catch (Exception ex)
                {
                    MessageBox.Show("The log file could not be saved. " + ex.Message,
                            "Excel File Cleaner", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
            }

            btnClose.Enabled = true;
            button1.Enabled = true;
            optionsGroup.Enabled = true;
        }

        void cleaner_UpdateSmallProgress(double value)
        {
            tsProgress.Visible = true;
            tsProgress.Value = (int)(value * 100);
        }

        bool[] compileOptions()
        {
            bool[] options;
            // compile the options into an array that will be passed
            // to the cleaner class
            if (rdoCorrectDefault.Checked)
            {
                options = new bool[] { false, false, false, false, false };
                options[(int)clsCleanExcel.OPTIONS.CORRECT_STYLES] = true;
                options[(int)clsCleanExcel.OPTIONS.CORRECT_CELLS] = true;
                options[(int)clsCleanExcel.OPTIONS.CORRECT_NAMED] = true;
            }
            else if (rdoCorrectAll.Checked)
            {
                options = new bool[] { true, true, true, true, true };
            }
            else
            {
                options = new bool[] { false, false, false, false, false };
                if(chkStyles.Checked)
                    options[(int)clsCleanExcel.OPTIONS.CORRECT_STYLES] = true;
                if (chkRanges.Checked)
                    options[(int)clsCleanExcel.OPTIONS.CORRECT_NAMED] = true;
                if (chkCells.Checked)
                    options[(int)clsCleanExcel.OPTIONS.CORRECT_CELLS] = true;
                if (chkLinks.Checked)
                    options[(int)clsCleanExcel.OPTIONS.CORRECT_LINK] = true;
                if (chkFormulas.Checked)
                    options[(int)clsCleanExcel.OPTIONS.CORRECT_FORMULAS] = true;
            }
            // return options
            return options;
        }

        void cleaner_StatusUpdated(object sender, EventArgs e)
        {
            clsCleanExcel cleaner = (clsCleanExcel)sender;
            tsLabel.Text = cleaner.getStatus();
            if (tsLabel.Text.Length == 0)
                tsLabel.Visible = false;
            else
                tsLabel.Visible = true;
            this.Refresh();
            Application.DoEvents();
        }

        void cleaner_UpdateProgress(object sender, EventArgs e)
        {
            pbMain.Increment(1);
            this.Refresh();
            Application.DoEvents();
        }

        List<string> offlineLog = new List<string>();
        void cleaner_LogUpdated(object sender, EventArgs e)
        {
            clsCleanExcel cleaner = (clsCleanExcel)sender;
            string data = cleaner.getLogData();
            if (!chkHide.Checked)
            {
                lstLog.Items.Add(data);
                lstLog.SelectedIndex = lstLog.Items.Count - 1;
                this.Refresh();
                Application.DoEvents();
            }
            else
            {
                offlineLog.Add(data);
            }
        }

        private void linkLabel2_LinkClicked(object sender, LinkLabelLinkClickedEventArgs e)
        {
            launchLink("http://support.microsoft.com/kb/244435");
        }

        private void launchLink(string link)
        {
            try
            {
                System.Diagnostics.Process proc = new System.Diagnostics.Process();
                proc.EnableRaisingEvents = false;
                if (link.Contains("@"))
                {
                    proc.StartInfo.Arguments = "\"" + link.Split('@')[1] + "\"";
                    proc.StartInfo.FileName = link.Split('@')[0];
                }
                else
                {
                    proc.StartInfo.FileName = link;
                }
                proc.Start();
            }
            catch (Exception ex)
            {
                MessageBox.Show("Unable to open. " + ex.Message, "Excel File Cleaner", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }

        }

        private void Form1_Load(object sender, EventArgs e)
        {

        }

        private void linkLabel1_LinkClicked(object sender, LinkLabelLinkClickedEventArgs e)
        {
            launchLink("http://support.microsoft.com/kb/213904");
        }

        private void button2_Click(object sender, EventArgs e)
        {
            this.Close();
        }

        private void rdoCorrectOnly_CheckedChanged(object sender, EventArgs e)
        {
            correctGroup.Enabled = rdoCorrectOnly.Checked;
        }

        private void chkLinks_CheckedChanged(object sender, EventArgs e)
        {
            if (chkLinks.Checked)
            {
                MessageBox.Show("(*) WARNING: ONLY USE THIS... \n\n" +
                                "If your file hangs or fails to resolve links when opening. \n\n" +
                                "This method only removes the external reference parts from the file " +
                                "and does not resolve all cell ranges using those references. \n\n" +
                                "NOTE: Excel will have to REPAIR the file after using this method as " +
                                "it will need to resolve all referenced ranges.", "Remove Links", MessageBoxButtons.OK,
                                MessageBoxIcon.Exclamation); ;
            }
        }

        private void chkFormulas_CheckedChanged(object sender, EventArgs e)
        {
            if (chkFormulas.Checked)
            {
                MessageBox.Show("(*) WARNING: ONLY USE THIS... \n\n" +
                                "If your file hangs or fails to resolve formulas when opening. \n\n" +
                                "This method only removes the file link from the cell forulas " +
                                "and does not resolve the calculation chain using those formulas. \n\n" +
                                "NOTE: Excel will have to REPAIR the file after using this method as " +
                                "it will need to resolve the calculation chain.", "Remove Links", MessageBoxButtons.OK,
                                MessageBoxIcon.Exclamation); ;
            }
        }
    }
}
