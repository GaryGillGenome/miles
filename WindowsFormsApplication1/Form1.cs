using Newtonsoft.Json;
using System;
using System.Collections.Generic;
using System.Data;
using System.Windows.Forms;
using Newtonsoft.Json.Linq;
using System.Linq;
using DevExpress.XtraPrinting;
using System.IO;
using System.ComponentModel;
using System.Globalization;
using System.Data.OleDb;
using DevExpress.ClipboardSource.SpreadsheetML;
using ExcelDataReader;
using IronPdf;
// ...

namespace WindowsFormsApplication1 {
    public partial class Form1 : Form {


        public Form1() {
            InitializeComponent();
        }

        List<Dictionary<string, string>> execSummaryList = new List<Dictionary<string, string>>();
        Dictionary<string, List<string>> chartsData = new Dictionary<string, List<string>>();
        

        private void GenerateReport(string folderName)
        {
            label3.Text = "";
            
            if (textBox1.Text.Trim() == "" || textBox2.Text.Trim() == "")
            {
                label3.Text = "Require Source and Destination folder to process this request";
                return;
            }
            string reportPath = textBox2.Text.Trim(); //@"c:\\Temp\Test.pdf";
            string sourcePath = textBox1.Text.Trim(); //@"C:\repos\miles\SFX Trade Panel\Archive r3 - May 2018 (updated) -lite\Strat Trade Files-20180424-170500";

            
            // A path to export a report.
            execSummaryList = new List<Dictionary<string, string>>();


            DataTable dtExecSummary = new DataTable(); // null;// ConvertJsonToDatatable(jsonObj);
            DataTable dtNetPnL = new DataTable();
            DataTable dtNetPips = new DataTable();
            int euTransactionCounter = 0;
            int gbTransactionCounter = 0;
            int usTransactionCounter = 0;
            double euNetPnlTotal = 0;
            double gbNetPnlTotal = 0;
            double usNetPnlTotal = 0;

            double euNetPipsTotal = 0;
            double gbNetPipsTotal = 0;
            double usNetPipsTotal = 0;
            //double netPipsTotal = 0;
            //int transactionsTotal = 0;

            if (Directory.Exists(sourcePath+"/"+folderName))
            {
                string[] files = Directory.GetFiles(sourcePath + "/" + folderName);
                if (files == null || files.Length <= 0)
                {
                    label3.Text = "No files found in the source folder";
                    return;
                }
                foreach (var f in files)
                {
                    if (f.Contains("ExecSumReport_stratmator"))
                    {
                        var parsedTxt = File.ReadAllLines(f);
                        //string[] c = { "\n" };
                        Dictionary<string, string> d = new Dictionary<string, string>();
                        foreach (var s in parsedTxt)
                        {
                            var eachLine = s;
                            var parsedCols = s.Split(',');
                            d.Add(parsedCols[0], parsedCols[1]);


                        }
                        Console.Write("File: ", d);
                        execSummaryList.Add(d);
                        //ToDataTable(execSummaryList);
                    }
                    else if (f.EndsWith("Trade.xlsx"))
                    {

                        //string[] c = { "\n" };
                        SaveAsCsv(f, f.Replace(".xlsx", ".csv"));
                        var parsedTxt = File.ReadAllLines(f.Replace(".xlsx", ".csv"));
                        Dictionary<string, string> d = new Dictionary<string, string>();
                        int lineCntr = 0;
                        foreach (var s in parsedTxt)
                        {
                            lineCntr++;
                            if (lineCntr == 1)
                            {
                                continue;
                            }

                            var eachLine = s;
                            var parsedCols = s.Split(',');
                            var symx = 0;
                            if (parsedCols[4].Contains("JPY"))
                            {
                                symx = 10000;
                            }
                            else
                            {
                                symx = 100;
                            }
                            double netPipsVal = 0;
                            if (parsedCols[2].Equals("SELL"))
                            {
                                netPipsVal = (double.Parse(parsedCols[5]) - double.Parse(parsedCols[9])) * symx;
                            }
                            else
                            {
                                netPipsVal = (double.Parse(parsedCols[9]) - double.Parse(parsedCols[5])) * symx;
                            }

                            if (parsedCols[4].Equals("EURUSDpro"))
                            {
                                euNetPnlTotal += double.Parse(parsedCols[13]);
                                euNetPipsTotal += netPipsVal;
                                euTransactionCounter++;
                            }
                            else if (parsedCols[4].Equals("GBPUSDpro"))
                            {
                                gbTransactionCounter++;
                                gbNetPnlTotal += double.Parse(parsedCols[13]);
                                gbNetPipsTotal += netPipsVal;
                            }
                            else if (parsedCols[4].Equals("USDCADpro"))
                            {
                                usTransactionCounter++;
                                usNetPnlTotal += double.Parse(parsedCols[13]);
                                usNetPipsTotal += netPipsVal;
                            }
                        }
                        var arr = new List<string>();
                        string csvContent = String.Empty;
                        List<string> chartAttrList = new List<string>();
                        //Add headers
                        csvContent += "Items,Transactions,NetPnl,NetPips" + "\n";
                        chartAttrList.Add("EURUSDpro");
                        chartAttrList.Add(euTransactionCounter.ToString());
                        chartAttrList.Add(euNetPnlTotal.ToString());
                        chartAttrList.Add(euNetPipsTotal.ToString());
                        //chartsData.Add("EURUSDpro", chartAttrList);
                        csvContent += string.Join(",", chartAttrList) + "\n";

                        chartAttrList = new List<string>();
                        chartAttrList.Add("GBPUSDpro");
                        chartAttrList.Add(gbTransactionCounter.ToString());
                        chartAttrList.Add(gbNetPnlTotal.ToString());
                        chartAttrList.Add(gbNetPipsTotal.ToString());
                        //chartsData.Add("GBPUSDpro", chartAttrList);
                        csvContent += string.Join(",", chartAttrList) + "\n";

                        chartAttrList = new List<string>();
                        chartAttrList.Add("USDCADpro");
                        chartAttrList.Add(usTransactionCounter.ToString());
                        chartAttrList.Add(usNetPnlTotal.ToString());
                        chartAttrList.Add(usNetPipsTotal.ToString());
                        //.Add("USDCADpro", chartAttrList);
                        csvContent += string.Join(",", chartAttrList) + "\n";
                        var filenameArry = f.Split('\\');
                        var filename = filenameArry[filenameArry.Length - 1];
                        StreamWriter csv = new StreamWriter(f.Replace(filename, "PieChartTotal.csv"), false);
                        csv.Write(csvContent);
                        csv.Close();

                        Console.Write("File: ", d);
                        //execSummaryList.Add(d);
                    }

                }

            }
            else
            {
                label3.Text = "Folder doesn't exist";
                return;
            }

            using (XtraReport2 report = new XtraReport2())
            {
                // Specify PDF-specific export options.
                PdfExportOptions pdfOptions = report.ExportOptions.Pdf;

                pdfOptions.PageRange = "1, 3-5";
                report.Parameters[1].Value = execSummaryList[0]["ROI"];
                CultureInfo provider = CultureInfo.InvariantCulture;
                var srcPathParsed = folderName.Split('-');
                var dateStr =  srcPathParsed[srcPathParsed.Length - 2];
                report.Parameters[0].Value = DateTime.ParseExact(dateStr, "yyyyMMdd", provider).ToLongDateString();// DateTime.Now.ToLongDateString();
                // Specify the quality of exported images.
                pdfOptions.ConvertImagesToJpeg = false;
                pdfOptions.ImageQuality = PdfJpegImageQuality.Medium;

                // Specify the PDF/A-compatibility.
                pdfOptions.PdfACompatibility = PdfACompatibility.PdfA3b;

                // The following options are not compatible with PDF/A.
                // The use of these options will result in errors on PDF validation.
                //pdfOptions.NeverEmbeddedFonts = "Tahoma;Courier New";
                //pdfOptions.ShowPrintDialogOnOpen = true;

                // If required, you can specify the security and signature options. 
                //pdfOptions.PasswordSecurityOptions
                //pdfOptions.SignatureOptions

                // If required, specify necessary metadata and attachments
                // (e.g., to produce a ZUGFeRD-compatible PDF).
                //pdfOptions.AdditionalMetadata
                //pdfOptions.Attachments

                // Specify the document options.
                pdfOptions.DocumentOptions.Application = "Test Application";
                pdfOptions.DocumentOptions.Author = "DX Documentation Team";
                pdfOptions.DocumentOptions.Keywords = "DevExpress, Reporting, PDF";
                pdfOptions.DocumentOptions.Producer = Environment.UserName.ToString();
                pdfOptions.DocumentOptions.Subject = "Document Subject";
                pdfOptions.DocumentOptions.Title = "Document Title";

                // Checks the validity of PDF export options 
                // and return a list of any detected inconsistencies.
                IList<string> result = pdfOptions.Validate();
                try
                {
                    if (result.Count > 0)
                        Console.WriteLine(String.Join(Environment.NewLine, result));
                    else
                        report.ExportToPdf(reportPath+"/DailyReport_" + folderName.Replace(" ", "") + ".pdf", pdfOptions);
                    MessageBox.Show(folderName + " folder processed in " + textBox2.Text);
                }
                catch(Exception e)
                {
                    MessageBox.Show("Error in processing the report: "+e.Message);
                    return;
                }
                
            }
            // Generate pdf from htm
            CreateHtmlToPdf(Path.Combine(sourcePath, folderName), reportPath, folderName);
        }

        public static bool SaveAsCsv(string excelFilePath, string destinationCsvFilePath)
        {

            using (var stream = new FileStream(excelFilePath, FileMode.Open, FileAccess.Read, FileShare.ReadWrite))
            {
                IExcelDataReader reader = null;
                if (excelFilePath.EndsWith(".xls"))
                {
                    reader = ExcelReaderFactory.CreateBinaryReader(stream);
                }
                else if (excelFilePath.EndsWith(".xlsx"))
                {
                    reader = ExcelReaderFactory.CreateOpenXmlReader(stream);
                }

                if (reader == null)
                    return false;

                var ds = reader.AsDataSet(new ExcelDataSetConfiguration()
                {
                    ConfigureDataTable = (tableReader) => new ExcelDataTableConfiguration()
                    {
                        UseHeaderRow = false
                    }
                });

                var csvContent = string.Empty;
                int row_no = 0;
                while (row_no < ds.Tables[0].Rows.Count)
                {
                    var arr = new List<string>();
                    for (int i = 0; i < ds.Tables[0].Columns.Count; i++)
                    {
                        arr.Add(ds.Tables[0].Rows[row_no][i].ToString());
                    }
                    row_no++;
                    csvContent += string.Join(",", arr) + "\n";
                }
                StreamWriter csv = new StreamWriter(destinationCsvFilePath, false);
                csv.Write(csvContent);
                csv.Close();
                return true;
            }
        }
    
        private void button2_Click(object sender, EventArgs e)
        {
            ChooseFolder();
        }

        public void ChooseFolder(Boolean destFolder = false)
        {
            if (folderBrowserDialog1.ShowDialog() == DialogResult.OK)
            {
                if (destFolder)
                {
                    textBox2.Text = folderBrowserDialog1.SelectedPath;
                }
                else
                {
                    textBox1.Text = folderBrowserDialog1.SelectedPath;
                }
                

            }
        }

        private void Form1_Load(object sender, EventArgs e)
        {
            //var sourcePath = ConfigurationManager.AppSettings["MySetting"];
            var settingsReader = new System.Configuration.AppSettingsReader();
            string src = settingsReader.GetValue("source", typeof(string)).ToString();
            string dest = settingsReader.GetValue("dest", typeof(string)).ToString();
            textBox1.Text = src;
            textBox2.Text = dest;

            //string destPath = Settings.Default.Properties["dest"].DefaultValue.ToString();
        }

        void loadData()
        {
            try
            {
                if (textBox1.Text.Trim() == "" || textBox2.Text.Trim() == "")
                {
                    label3.Text = "Require Source and Destination folder to process this request";
                    return;
                }
                string reportPath = textBox2.Text.Trim();
                string sourcePath = textBox1.Text.Trim();

                IDictionary<int, string> dict = new Dictionary<int, string>();
                IList<string[]> folderList = new List<string[]>();

                var foldersToProcess = Directory.EnumerateDirectories(sourcePath);
                var cntr = 0;
                foreach (var d in foldersToProcess)
                {
                    IList<string> l = new List<string>();

                    var arr = d.Split('\\');
                    var foldername = arr[arr.Length - 1];
                    string[] rowInfo = new string[3];
                    rowInfo[0] = foldername;
                    //Check this file to see if and when last processed
                    var processedStatusIndicator = sourcePath + "\\" + foldername + "\\PieChartTotal.csv";
                    if (File.Exists(processedStatusIndicator))
                    {
                        var fInfo = new FileInfo(processedStatusIndicator);
                        var splitStr = foldername.Split('-');
                        CultureInfo provider = CultureInfo.InvariantCulture;
                        var processedDate = DateTime.ParseExact(splitStr[splitStr.Length - 2], "yyyyMMdd", provider);

                        rowInfo[1] = fInfo.LastWriteTime.ToLongDateString() + " " + fInfo.LastWriteTime.ToShortTimeString();
                        rowInfo[2] = "Yes";
                    }
                    else
                    {
                        rowInfo[1] = "Not Processed";
                        rowInfo[2] = "No";
                    }

                    folderList.Add(rowInfo);
                    //l.Add(arr[arr.Length - 1]);
                    //dict.Add(cntr, foldername);
                    cntr++;
                }
                CreateDataTable(folderList);
            }
            catch (Exception ex)
            {
                MessageBox.Show("Error is loading the folder: " + ex.Message);
                return;
            }
        }

        private void button4_Click(object sender, EventArgs e)
        {

            loadData();
        }

        void CreateHtmlToPdf(string dirPath, string destpath, string foldername)
        {
            var htmlToPdf = new HtmlToPdf();  // new instance of HtmlToPdf
            var sourceFileName = "DetailedReportStatement_stratmator_9010_20180424-170000_ExecSum.htm";

            if (!File.Exists(Path.Combine(dirPath, sourceFileName)))
            {
                MessageBox.Show("Source file ( " + sourceFileName + " ) doesn't exist. Other files will be still processed");
                return;
            }
            //html to pdf
            //html to turn into pdf
            //var html = @"<h1>Hello World!</h1><br><p>This is IronPdf.</p>";

            // turn html to pdf
            //var pdf = htmlToPdf.RenderHtmlAsPdf(html);

            // save resulting pdf into file
            //pdf.SaveAs(Path.Combine(dirPath, "HtmlToPdf.Pdf"));

            //url to pdf
            // uri of the page to turn into pdf
            //var uri = new Uri("http://www.google.com/ncr");

            // turn page into pdf
            var pdf = htmlToPdf.RenderUrlAsPdf(Path.Combine(dirPath, sourceFileName));

            // save resulting pdf into file
            pdf.SaveAs(Path.Combine(destpath, "UrlToPdf_"+foldername+".Pdf"));

        }

        void CreateDataTable(IList<string[]> l)
        {
            dataGridView1.Columns.Clear();
            //create datatable and columns,
            DataTable dtable = new DataTable();
            dtable.Columns.Add(new DataColumn("Folder Name"));
            dtable.Columns.Add(new DataColumn("Processed Date"));
            dtable.Columns.Add(new DataColumn("Processed?"));
            object[] RowValues = { "", "",true };
            
            foreach (var eachRow in l)
            {
                //create new data row
                DataRow dRow;
                dRow = dtable.Rows.Add(eachRow);
                dtable.AcceptChanges();
            }
            //assign values into row object
            //now bind datatable to gridview... 
            dataGridView1.DataSource = dtable;
            dataGridView1.Columns[0].AutoSizeMode = DataGridViewAutoSizeColumnMode.AllCells;
            dataGridView1.Columns[1].AutoSizeMode = DataGridViewAutoSizeColumnMode.AllCells;
            dataGridView1.Columns[2].AutoSizeMode = DataGridViewAutoSizeColumnMode.AllCells;
            //.Width = 200;
           // dataGridView1.Columns[1].Width = 100;
            //dataGridView1.Columns[2].Width = 50;
            DataGridViewButtonColumn btn = new DataGridViewButtonColumn();
            dataGridView1.Columns.Add(btn);
            btn.HeaderText = "";
            btn.Text = "Run Report";
            btn.Name = "runReport";
            btn.UseColumnTextForButtonValue = true;
            //inisde some method or event when enabling/disabling buttons:
    //        foreach (DataGridViewRow row in dataGridView1.Rows)
    //        {
    //            //your conditon if there is data in some column:
    //            if (!string.IsNullOrEmpty(row.Cells["DataColumn"].Value).ToString()))
    //{
    //                DataGridViewButtonCell cell = row.Cells["ButtonColumnName"] as DataGridViewButtonCell;
    //                cell.Enabled = 1 ? false : true;  //if 1 - cell disabled, else enabled!
    //            }
    //        }
            dataGridView1.Refresh();
        }

        private void dataGridView1_CellContentClick(object sender, DataGridViewCellEventArgs e)
        {
            DataGridView c = (DataGridView)sender;
            if (e.ColumnIndex == 3)
            {

                try
                {
                    GenerateReport(c.Rows[e.RowIndex].Cells[0].Value.ToString());
                    loadData();
                }
                catch (Exception ex)
                {
                    MessageBox.Show("Error in processing folder " + c.Rows[e.RowIndex].Cells[0].Value + " "+ ex.Message);
                }
                
            }
            
        }

        private void dataGridView1_RowsAdded(object sender, DataGridViewRowsAddedEventArgs e)
        {
           // MessageBox.Show((e.RowIndex + 1) + "  Row Added  ");
            DataGridView dgv = (DataGridView)sender;
            //dgv.Rows[e.RowIndex].Cells[2]`
        }

        private void button3_Click(object sender, EventArgs e)
        {
            ChooseFolder(true);
        }
    }

    
}
