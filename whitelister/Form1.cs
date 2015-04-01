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
using CsvHelper.TypeConversion;
using CsvHelper.Configuration;

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
            CsvConfiguration config = new CsvConfiguration();
            config.IgnoreQuotes = true;
            config.IgnoreHeaderWhiteSpace = true;
            CsvReader csv = new CsvReader(file, config);
            IEnumerable<DataRecord> lines = csv.GetRecords<DataRecord>();
            String currentEmployeeEmail;

            foreach (DataRecord line in lines.Take(25)) // Each record will be fetched and printed on the screen
            {
                currentEmployeeEmail = line.RecipientAddress;
            }
            
            file.Close();
        }


        public class DataRecord // Test record class
        {
            public string SpamScore { get; set; }
            public string SenderEMailOnList { get; set; }
            public string SenderDomainOnList { get; set; }
            public string SenderAddress { get; set; }
            public string RecipientAddress { get; set; }
            public string Subject { get; set; }
            public string Date { get; set; }
        }


        public void emailComposition()
        {
            String htmlHeader, htmlBody, htmlFooter;
            htmlHeader = ;
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
