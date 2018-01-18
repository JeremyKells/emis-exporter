using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Configuration;
using System.Data.SqlClient;
using System.IO;
using System.Linq;
using System.Windows;
using System.Windows.Controls;

using Excel = Microsoft.Office.Interop.Excel;

namespace EmisExporter
{
    public partial class Exporter
    {
        
        void worker_DoWork(object sender, DoWorkEventArgs e)
        {
            try //Comment out this try/catch when debugging
            {
                BackgroundWorker worker = sender as BackgroundWorker;
                string year = (string)e.Argument;
                Excel.Application excelApp = new Excel.Application();

                string filePath = Path.Combine(Directory.GetCurrentDirectory(), (string)ConfigurationManager.AppSettings["UIS Template"]);

                excelApp.Workbooks.Add(filePath);

                string country = ConfigurationManager.AppSettings["Country"];
                SqlConnection emisDBConn = new SqlConnection(
                    string.Format("Data Source={0};Initial Catalog={1};User id={2};Password={3};",
                                ConfigurationManager.AppSettings["Data Source"],
                                ConfigurationManager.AppSettings["Initial Catalog"],
                                ConfigurationManager.AppSettings["User id"],
                                ConfigurationManager.AppSettings["Password"]));
                emisDBConn.Open();

                // Create Base Tables
                string BaseSQL = File.ReadAllText(@Path.Combine(Directory.GetCurrentDirectory(), "SQL", (string)ConfigurationManager.AppSettings["BaseSQLPath"]));
                BaseSQL = String.Format(BaseSQL, year);
                new SqlCommand(BaseSQL, emisDBConn).ExecuteNonQuery();


                List<Action<Excel.Application, SqlConnection, string, string>> sheets = new List<Action<Excel.Application, SqlConnection, string, string>> { };

                List<string> ssheets = ConfigurationManager.AppSettings["Sheets"].Replace(" ", string.Empty).Split(',').ToList();

                Dictionary<String, Action<Excel.Application, SqlConnection, string, string>> actionMap 
                    = new Dictionary<String, Action<Excel.Application, SqlConnection, string, string>>()
                    {
                        {"A2", sheetA2 },
                        {"A3", sheetA3 },
                        {"A5", sheetA5 },
                        {"A6", sheetA6 },
                        {"A7", sheetA7 },
                        {"A9", sheetA9 },
                        {"A10", sheetA10 },
                    };

                foreach (String sheet in ssheets)
                    sheets.Add(actionMap[sheet]);

                sheets.Reverse();  // Leaves Excel open on first sheet, and progressbar will initially update faster
                for (int i = 0; i < sheets.Count; i++)
                {
                    Action<Excel.Application, SqlConnection, string, string> fun = sheets[i];
                    fun(excelApp, emisDBConn, year, country);
                    double progress = i * (100 / sheets.Count());
                    (sender as BackgroundWorker).ReportProgress((int)progress);
                }
                excelApp.Visible = true;
                excelApp.ActiveWorkbook.SaveAs("UIS_Export.xlsx");
                e.Result = null;
            }
            catch (Exception ex)  // Change to a file based log
            {
                MessageBox.Show(ex.ToString(), "Error", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        void worker_ProgressChanged(object sender, ProgressChangedEventArgs e)
        {
            Progress.Value = e.ProgressPercentage;
        }

        void worker_RunWorkerCompleted(object sender, RunWorkerCompletedEventArgs e)
        {
            Environment.Exit(0);
        }

        private ProgressBar Progress;
        public void export(string year, ProgressBar _Progress)
        {
            Progress = _Progress;
            BackgroundWorker worker = new BackgroundWorker();
            worker.WorkerReportsProgress = true;
            worker.ProgressChanged += worker_ProgressChanged;
            worker.DoWork += worker_DoWork;
            worker.RunWorkerCompleted += worker_RunWorkerCompleted;
            worker.RunWorkerAsync(year);
        }
    }
}
