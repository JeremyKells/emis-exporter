using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Configuration;
using System.Data.SqlClient;
using System.Linq;
using System.Windows;
using System.Windows.Controls;

using Excel = Microsoft.Office.Interop.Excel;

namespace EmisExporter
{
    public partial class Exporter
    {
        //private static readonly log4net.ILog log = log4net.LogManager.GetLogger (System.Reflection.MethodBase.GetCurrentMethod().DeclaringType);
        
        void worker_DoWork(object sender, DoWorkEventArgs e)
        {
            try //Comment out this try/catch when debugging
            {
                BackgroundWorker worker = sender as BackgroundWorker;
                string year = (string)e.Argument;
                var excelApp = new Excel.Application();
                excelApp.Workbooks.Add((string)ConfigurationManager.AppSettings["UIS Template"]);

                string country = ConfigurationManager.AppSettings["Country"];
                SqlConnection temis = new SqlConnection(
                    string.Format("Data Source={0};Initial Catalog={1};User id={2};Password={3};",
                                ConfigurationManager.AppSettings["Data Source"],
                                ConfigurationManager.AppSettings["Initial Catalog"],
                                ConfigurationManager.AppSettings["User id"],
                                ConfigurationManager.AppSettings["Password"]));
                temis.Open();
                List<Action<Excel.Application, SqlConnection, string, string>> sheets;

                switch (country)
                {
                    case "NAURU":
                        sheets = new List<Action<Excel.Application, SqlConnection, string, string>> {
                            sheetA2, sheetA3, sheetA5, sheetA6, sheetA7, sheetA8, sheetA10, sheetA12
                        };
                        break;
                    default: // Tuvalu
                        sheets = new List<Action<Excel.Application, SqlConnection, string, string>> {
                            sheetA2, sheetA3, sheetA5, sheetA6, sheetA7, sheetA8, sheetA10, sheetA12
                        };
                        break;
                }

                sheets.Reverse();  // Leaves Excel open on first sheet, and progressbar will initially update faster
                for (int i = 0; i < sheets.Count; i++)
                {
                    Action<Excel.Application, SqlConnection, string, string> fun = sheets[i];
                    fun(excelApp, temis, year, country);
                    double progress = i * (100 / sheets.Count());
                    (sender as BackgroundWorker).ReportProgress((int)progress);
                }
                excelApp.Visible = true;
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
