using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;

namespace ExcelFileCleaner
{
    public partial class frmHelp : Form
    {
        string _title = "";
        string _text = "";
        public frmHelp(string title, string text)
        {
            InitializeComponent();
            _title = title;
            _text = text;
        }

        private void frmHelp_Load(object sender, EventArgs e)
        {
            this.Text = _title;
            txtHelp.Text = _text;
            txtHelp.SelectionStart = 0;
        }
    }
}
