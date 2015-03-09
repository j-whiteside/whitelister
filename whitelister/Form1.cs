using System;
using System.IO;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Excel = Microsoft.Office.Interop.Excel;

namespace whitelister
{
    public partial class Form1 : Form
    {
        String path = @"C:\exceltest\";
        String filepath = @"C:\exceltest\spam.csv";
        String filepathEdit = @"C:\exceltest\spamEdit.csv";
        Excel.Application xlApp = new Excel.Application();
        Excel.Workbook xlWorkBook;
        Excel.Worksheet xlWorkSheet;
        Boolean appFailed = false;

        public Form1()
        {
            InitializeComponent();
        }

        private void Form1_Load(object sender, EventArgs e)
        {            

            fileLocation();
            if (appFailed != true)
            {
                fileModification(path);
            }
            else
            {
                logger("Program terminating");
                this.Close();
            }

        }

        public void fileLocation()
        {
            logger("Start time: " + DateTime.Now.ToString());
            logger(" ");

            if (File.Exists(path + "spam.csv") == true)
            {
                logger("File location successful!");
                File.Copy(filepath, filepathEdit, true);
                logger("File duplication successful!");
            }
            else
            {
                logger(path + "spam.csv was not found");
                appFailed = true;
            }
        }

        public void fileModification(String path)
        {
            try
            {
                int numberOfRowsInteger = File.ReadLines(filepathEdit).Count();
                //xlWorkBook = xlApp.Workbooks.Open(filepathEdit, 0, false, 2, "", "", true, Microsoft.Office.Interop.Excel.XlPlatform.xlWindows, "\t", false, false, 0, false, 1, 0);
                xlWorkBook = xlApp.Workbooks.Open(filepathEdit);
                xlWorkSheet = (Excel.Worksheet)xlWorkBook.Worksheets.get_Item(1);

                //Delete all records older than a week

                xlWorkSheet.Range[G2, 7].EntireColumn.NumberFormat = "MM/dd/yyyy";

                DateTime lastWeekDateTime = getLastWeek();

                /*
                for(int i = 1; i < numberOfRowsInteger ; i++)
                {
                    if (xlWorkSheet.Cells[i,7].Contains(lastWeekDateTime))
                    {
            
                    }
                }

                //Order by spam index
                Excel.Range rngSort = xlWorkSheet.get_Range("A2", "J" + numberOfRowsInteger);

                rngSort.Sort(rngSort.Columns[1, Type.Missing], Excel.XlSortOrder.xlAscending,
                                rngSort.Columns[2, Type.Missing], Type.Missing, Excel.XlSortOrder.xlAscending,
                                Type.Missing, Excel.XlSortOrder.xlAscending,ss
                                Excel.XlYesNoGuess.xlYes, Type.Missing, Type.Missing,
                                Excel.XlSortOrientation.xlSortColumns,
                                Excel.XlSortMethod.xlPinYin,
                                Excel.XlSortDataOption.xlSortNormal,
                                Excel.XlSortDataOption.xlSortNormal,
                                Excel.XlSortDataOption.xlSortNormal);
                rngSort = null;

                xlWorkBook.Save();
                 */
            }
            catch (Exception ex)
            {
                xlWorkBook.Close(false, Type.Missing, Type.Missing);

                throw;
            }
            finally
            {
                GC.Collect();
                GC.WaitForPendingFinalizers();

                System.Runtime.InteropServices.Marshal.FinalReleaseComObject(xlWorkBook);
                System.Runtime.InteropServices.Marshal.FinalReleaseComObject(xlApp);

            }
        }

        public DateTime getLastWeek()
        {
            return DateTime.Today.AddDays(-7);
        }

        public void emailer(String subject, String recipient, String body)
        {

        }

        public void logger(String lines)
        {
            using (System.IO.StreamWriter log = new System.IO.StreamWriter(path + "whitelistEmailerLog.txt", true))
            {
                log.WriteLine(lines);
            }
        }


    }
}
