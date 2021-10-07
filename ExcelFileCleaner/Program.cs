using System;
using System.Collections.Generic;
using System.Linq;
using System.Windows.Forms;
using System.Text;
using System.Runtime.InteropServices;

namespace ExcelFileCleaner
{
    static class Program
    {
        /// <summary>
        /// The main entry point for the application.
        /// </summary>
        [STAThread]
        static void Main(string[] args)
        {
            if (args.Length == 0)
            {
                Application.EnableVisualStyles();
                Application.SetCompatibleTextRenderingDefault(false);
                Application.Run(new Form1());
            }
            else
            {
                const string title = "Excel File Cleaner Command Line Help";
                const string help = "Excel File Cleaner Command Line Help\r\n" +
                                    "====================================\r\n\r\n" +
                                    " -?\tThis help screen\r\n" +
                                    " -d\tClean using default settings (styles, named ranges and empty cells)\r\n" +
                                    " -a\tClean all (styles, ranges, cells, links, formulas)\r\n" +
                                    " -s\tClean styles\r\n" +
                                    " -r\tClean named ranges\r\n" +
                                    " -c\tClean empty cells\r\n" +
                                    " -k\tClean external links\r\n" +
                                    " -f\tClean formulas with links\r\n" +
                                    " -l\tLog results.\r\n\r\n" +
                                    " [filename]\r\n\tFile to clean\r\n\r\n" +
                                    " [folder]\r\n\tFolder to clean\r\n\r\n" +
                                    " -x\tSearch for files in folder recursively\r\n\r\n" +
                                    "Usage:\r\n" +
                                    "------\r\n" +
                                    "To clean all Excel files in every folder on the " +
                                    "entire C drive using default settings and logging the results in a " +
                                    "file on the root of the drive:\r\n\r\n" +
                                    "\tExcelFileCleaner.exe c:\\ -x -d -l\r\n\r\n" +
                                    "To clean a single file with custom settings, no logging:\r\n\r\n" +
                                    "\tExcelFileCleaner.exe c:\\file.xlsx -s -f -k\r\n\r\n" +
                                    "To clean only the files in a folder (non recursively) with all options:\r\n\r\n" +
                                    "\tExcelFileCleaner.exe c:\\folder -a\r\n\r\n" +
                                    "Notes:\r\n" +
                                    "------\r\n" +
                                    " - You can only use -d or -a seperately, you cannot use them together.\r\n" +
                                    " - You cannot use -d/-a with -s, -r -c -k or -f.\r\n"+
                                    " - You cannot use [file] specification and [folder] together.\r\n" +
                                    " - If you use -l with [folder] all logging will be in the same file\r\n" +
                                    " - You can only use -x with a [folder].";

                if (args[0] == "-?" || args[0] == "/?")
                {
                    frmHelp hForm = new frmHelp(title, help);
                    hForm.ShowDialog();
                }
            }
        }
    }
}
