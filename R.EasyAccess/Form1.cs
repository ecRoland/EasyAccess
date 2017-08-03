using System;
using System.Configuration;
using System.Windows.Forms;
using Microsoft.Office.Interop.Excel;
using Excel = Microsoft.Office.Interop.Excel;
using Word = Microsoft.Office.Interop.Word;
using System.Runtime.InteropServices;
using System.Diagnostics;

namespace R.EasyAccess
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        // Open Excel from configuration //

        private void Proccess(string excel, string sheet)
        {
            string fileExcel = ConfigurationManager.AppSettings[excel];          
            Excel.Application xlApp;
            Excel.Workbook xldoc;
            Excel.Worksheet xlsheet;
            object misValue = System.Reflection.Missing.Value;
            xlApp = new Excel.Application();
            xldoc = xlApp.Workbooks.Open(fileExcel, 0, false, 5, "", "", false, Excel.XlPlatform.xlWindows, "", true, false, 0, true, false, false);
            // open desired sheet from configuration //
            try
            {
                xlsheet = (Excel.Worksheet)xldoc.Sheets[ConfigurationManager.AppSettings[sheet]];
                xlsheet.Visible = Excel.XlSheetVisibility.xlSheetVisible;
                xlsheet.Activate();
                xlApp.Visible = true;
                xlApp.UserControl = false;
                xlApp.DisplayAlerts = false;
            }
            // bypass opening sheet //
            catch (System.Runtime.InteropServices.COMException)
            {
                xlApp.Visible = true;
                xlApp.UserControl = false;
                xlApp.DisplayAlerts = false;
            }
        }

        // Open Word from configuration  //
        public void ProccessWord(string file)
        {
            string filepath = ConfigurationManager.AppSettings[file];
            object misvl = System.Reflection.Missing.Value;
            Word.Application wdApp = new Word.Application();
            Word.Document doc = wdApp.Documents.Open(filepath, ReadOnly: false);
            wdApp.Visible = true;
            
            /*
            Process p = new Process();
            p.StartInfo.FileName = file;
            p.Start("WINWORD.EXE", );
            */
        }

        // Open Path from configuration //
        private void Path(string path)
        {
            System.Diagnostics.Process.Start(@"Explorer.exe", ConfigurationManager.AppSettings[path]);
        }

        private void open_Click(object sender, EventArgs e)
        {
            int mode;
            if (excel.Checked)
            {
                mode = 1;
            }
            else if (word.Checked)
            {
                mode = 2;
            }
            else if (path.Checked)
            {
                mode = 3;
            }
            else
            {
                MessageBox.Show("select one of the tree options", @"Error", MessageBoxButtons.OK); // Insurance message, if you forgot to select an option!
            }

            ///// Opening Excel //////

            if (excel.Checked)
            {
                if (list.Text == "")
                {
                    MessageBox.Show("You need to select an item from list!", @"Error", MessageBoxButtons.OK); // Insurance message, if you forgot to select an item from list!
                    return;
                }
                switch (list.Text)
                {
                    case "Item-1":
                        Proccess(@"Excel-Item-1", @"Sheet_Excel-Item-1");
                        break;
                    case "Item-2":
                        Proccess(@"Excel-Item-2", @"Sheet_Excel-Item-2");
                        break;
                    case "Item-3":
                        Proccess(@"Excel-Item-3", @"Sheet_Excel-Item-3");
                        break;
                    case "Item-4":
                        Proccess(@"Excel-Item-4", @"Sheet_Excel-Item-4");
                        break;
                    case "Item-5":
                        Proccess(@"Excel-Item-5", @"Sheet_Excel-Item-5");
                        break;
                    default:
                        break;
                }
            }

            ///// Opening Word //////

            if (word.Checked)
            {
                if (list.Text == "")
                {
                    MessageBox.Show("You need to select an item from list!", @"Error", MessageBoxButtons.OK); // Insurance message, if you forgot to select an item from list!
                    return;
                }
                switch (list.Text)
                {
                    case "Item-1":
                        ProccessWord(@"Word-Item-1");
                        break;
                    case "Item-2":
                        ProccessWord(@"Word-Item-2");
                        break;
                    case "Item-3":
                        ProccessWord(@"Word-Item-3");
                        break;
                    case "Item-4":
                        ProccessWord(@"Word-Item-4");
                        break;
                    case "Item-5":
                        ProccessWord(@"Word-Item-5");
                        break;
                    default:
                        break;
                }
            }

            ///// Opening Path //////

            if (path.Checked)
            {
                if (list.Text == "")
                {
                    MessageBox.Show("You need to select an item from list!", @"Error", MessageBoxButtons.OK); // Insurance message, if you forgot to select an item from list!
                    return;
                }
                switch (list.Text)
                {
                    case "Item-1":
                        Path(@"Path-Item-1");
                        break;
                    case "Item-2":
                        Path(@"Path-Item-2");
                        break;
                    case "Item-3":
                        Path(@"Path-Item-3");
                        break;
                    case "Item-4":
                        Path(@"Path-Item-4");
                        break;
                    case "Item-5":
                        Path(@"Path-Item-5");
                        break;
                    default:
                        break;
                }
            }
        }
    }
}
