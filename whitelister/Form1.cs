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
using System.Diagnostics;
using CsvHelper;

namespace whitelister
{
    public partial class Form1 : Form
    {
        String path = (@"C:\exceltest\");
        String modificationMacroPath = (@"C:\exceltest\formatData.xls");
        String filepath = (@"C:\exceltest\spam.csv");
        String filepathEdit = (@"C:\exceltest\spamEdit.csv");

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
                fileModification(modificationMacroPath);
            }
            else
            {
                logger("Program terminating");
                this.Close();
            }
            spamLogParser(filepathEdit);

        }

        public void fileLocation()
        {
            logger("Start time: " + DateTime.Now.ToString());
            logger(" ");

            if (File.Exists(path + "spam.csv") == true)
            {
                logger("Spam log location successful!");
                File.Copy(filepath, filepathEdit, true);
                logger("Spam log duplication successful!");
            }
            else
            {
                logger(path + "spam.csv was not found");
                appFailed = true;
            }
        }

        public void fileModification(String path)
        {
            logger("Formatting spam log file");
            Process excel = new Process();
            excel.StartInfo.FileName = path;
            excel.Start();
            excel.WaitForExit(20000);
            logger("Spam log has been formatted");
        }

        public void spamLogParser(String path)
        {
            StreamReader file = new StreamReader(path);
            CsvReader csv = new CsvReader(file);
            IEnumerable<DataRecord> records = csv.GetRecords<DataRecord>();

            foreach (var rec in records) // Each record will be fetched and printed on the screen
            {
                String spamScore = rec.name;
                MessageBox.Show(spamScore);
                
                //Response.Write(string.Format("Name : {0}, Sex : {1}, Occupation : {2} <br/>", rec.name, rec.sex, rec.occupation));
            }
            
            file.Close();
            
            
            String sampleParse = csv.GetField(1);

            logger(sampleParse);

            //return recipient;
        }


        public class DataRecord // Test record class
        {
            public string spamScore { get; set; }
            public string Sender { get; set; }
            public string occupation { get; set; }
        }


        public void emailer(String subject, String recipient, String body)
        {

        }

        public void logger(String lines)
        {
            using (StreamWriter log = new StreamWriter(path + "whitelistEmailerLog.txt", true))
            {
                log.WriteLine(lines);
            }
        }


    }
}
