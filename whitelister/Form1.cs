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
//using Marshal = System.Runtime.InteropServices.Marshal;

namespace whitelister
{
    public partial class Form1 : Form
    {
        String path = (@"C:\exceltest\");
        String filepath = (@"C:\exceltest\spam.csv");
        String filepathEdit = (@"C:\exceltest\spamEdit.csv");
        private Excel.Application xlApp;
        private Excel.Workbooks xlWorkBooks;
        private Excel.Workbook xlWorkBook;
        protected Excel.Sheets xlWorkSheets;
        protected Excel.Worksheet xlWorkSheet;


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

        
        protected void fileModification(String path)
        {
            logger("File modification started");
            
            //xlWorkBook = xlApp.Workbooks.Open(filepathEdit, 0, false, 2, "", "", true, Microsoft.Office.Interop.Excel.XlPlatform.xlWindows, "\t", false, false, 0, false, 1, 0);
            xlApp = new Excel.Application();

                
            xlWorkBooks = xlApp.Workbooks;
            xlWorkBook = xlWorkBooks.Open(filepathEdit, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing);
            xlWorkSheets = xlWorkBook.Worksheets;
            xlWorkSheet = (Excel.Worksheet)xlWorkSheets[1];
            xlWorkSheet.Select(Type.Missing);

            int numberOfRowsInteger = File.ReadLines(@"C:\exceltest\spam.csv").Count();

            //Delete all records older than a week

            //xlWorkSheet.Range[G2, 7].EntireColumn.NumberFormat = "MM/dd/yyyy";

            DateTime lastWeekDateTime = getLastWeek();

            textBox1.Text = numberOfRowsInteger.ToString();
                
            /*
                
            for(int i = 1; i < numberOfRowsInteger ; i++)
            {
                if (xlWorkSheet.Cells[i,7] < (lastWeekDateTime))
                {
                    Excel.Range range = xlWorkSheet.get_Range(i, Type.Missing);
                    range.Delete(Excel.XlDeleteShiftDirection.xlShiftUp);
                    logger("Row" + i + " deleted");
                }
            }

            /*
            //Order by spam index
            Excel.Range rngSort = xlWorkSheet.get_Range("A2", "J" + numberOfRowsInteger);

            rngSort.Sort(rngSort.Columns[1, Type.Missing], Excel.XlSortOrder.xlAscending,
                            rngSort.Columns[2, Type.Missing], Type.Missing, Excel.XlSortOrder.xlAscending,
                            Type.Missing, Excel.XlSortOrder.xlAscending,
                            Excel.XlYesNoGuess.xlYes, Type.Missing, Type.Missing,
                            Excel.XlSortOrientation.xlSortColumns,
                            Excel.XlSortMethod.xlPinYin,
                            Excel.XlSortDataOption.xlSortNormal,
                            Excel.XlSortDataOption.xlSortNormal,
                            Excel.XlSortDataOption.xlSortNormal);
            rngSort = null;

            xlWorkBook.Save();
                */

            
            xlWorkBook.Save();
            xlWorkBook.Close(false, Type.Missing, Type.Missing);
            xlApp.Quit();

            releaseObject(xlWorkSheet);
            releaseObject(xlWorkSheets);
            releaseObject(xlWorkBook);
            releaseObject(xlWorkBooks);
            releaseObject(xlApp);
            
        }

        protected void releaseObject(object excelObject)
        {
            try
            {
                if (excelObject != null)
                {
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(excelObject);
                    excelObject = null;
                }

            }
            catch (Exception ex)
            {
                excelObject = null;
                MessageBox.Show("Unable to release the Object " + ex.Message);
            }
            finally
            {
                GC.Collect();
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
