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
        //private static readonly log4net.ILog log = log4net.LogManager.GetLogger (System.Reflection.MethodBase.GetCurrentMethod().DeclaringType);
        
        void worker_DoWork(object sender, DoWorkEventArgs e)
        {
            //try //Comment out this try/catch when debugging
            //{
                BackgroundWorker worker = sender as BackgroundWorker;
                string year = (string)e.Argument;
                var excelApp = new Excel.Application();
                excelApp.Workbooks.Add((string)ConfigurationManager.AppSettings["UIS Template"]);

                string country = ConfigurationManager.AppSettings["Country"];
                SqlConnection emisDBConn = new SqlConnection(
                    string.Format("Data Source={0};Initial Catalog={1};User id={2};Password={3};",
                                ConfigurationManager.AppSettings["Data Source"],
                                ConfigurationManager.AppSettings["Initial Catalog"],
                                ConfigurationManager.AppSettings["User id"],
                                ConfigurationManager.AppSettings["Password"]));
                emisDBConn.Open();

            // Create #StudentsTable
            SqlCommand DbDropCommand = new SqlCommand("IF OBJECT_ID('tempdb.dbo.#StudentsTable', 'U') IS NOT NULL DROP TABLE dbo.#StudentsTable;", emisDBConn);
            DbDropCommand.ExecuteNonQuery();

            SqlCommand DbCreateTableCommand
                = new SqlCommand(@"CREATE TABLE dbo.#StudentsTable (
                                    ISCED_TOP varchar(300), 
                                    ISCED varchar(300), 
                                    SCHOOLTYPE Varchar(1000), 
                                    GENDER varchar(200), 
                                    AGE int, 
                                    REPEATER varchar(1000), 
                                    CLASS decimal, 
                                    ECE varchar(600), COUNT int)",
                                    emisDBConn);
            DbCreateTableCommand.ExecuteNonQuery();

            string StudentBaseSQL = System.IO.File.ReadAllText(@ConfigurationManager.AppSettings["StudentBaseSQLPath"]);
            StudentBaseSQL = @"insert into dbo.#StudentsTable (ISCED_TOP, ISCED, SCHOOLTYPE, GENDER, AGE, REPEATER, CLASS, ECE, COUNT) " + String.Format(StudentBaseSQL, year);

            SqlCommand DbInsertCommand = new SqlCommand(StudentBaseSQL, emisDBConn);
            DbInsertCommand.ExecuteNonQuery();

            // Create #TeachersTable
            DbDropCommand = new SqlCommand("IF OBJECT_ID('tempdb.dbo.#TeacherBaseTable', 'U') IS NOT NULL DROP TABLE dbo.#TeacherBaseTable;", emisDBConn);
            DbDropCommand.ExecuteNonQuery();

            DbCreateTableCommand
                = new SqlCommand(@"CREATE TABLE dbo.#TeacherBaseTable (
                                    ISCED varchar(300), 
                                    SCHOOLTYPE Varchar(1000), 
                                    GENDER varchar(200), 
                                    QUALIFIED varchar(10), 
                                    TRAINED varchar(10), 
                                    COUNT int)",
                                    emisDBConn);
            DbCreateTableCommand.ExecuteNonQuery();

            string TeacherBaseSQL = System.IO.File.ReadAllText(@ConfigurationManager.AppSettings["TeacherBaseSQLPath"]);
            TeacherBaseSQL = @"insert into dbo.#TeacherBaseTable (ISCED, SCHOOLTYPE, GENDER, QUALIFIED, TRAINED, COUNT) " + String.Format(TeacherBaseSQL, year);

            DbInsertCommand = new SqlCommand(TeacherBaseSQL, emisDBConn);
            DbInsertCommand.ExecuteNonQuery();

            List<Action<Excel.Application, SqlConnection, string, string>> sheets;

                switch (country)
                {
                    case "SOLOMON ISLANDS":
                        sheets = new List<Action<Excel.Application, SqlConnection, string, string>> {
                            //sheetA2, sheetA3, sheetA5, sheetA6, sheetA8, sheetA7
                        };
                        break;
                    case "NAURU":
                        sheets = new List<Action<Excel.Application, SqlConnection, string, string>> {
                            //sheetA2, sheetA3, sheetA5, sheetA6, sheetA7, sheetA8, sheetA10, sheetA12
                        };
                        break;
                    case "TUVALU":
                        sheets = new List<Action<Excel.Application, SqlConnection, string, string>> {
                            //sheetA2, 
                            //sheetA3,
                            //sheetA5,
                            //sheetA6,
                            //sheetA7,
                            //sheetA9,
                            sheetA10,
                        };
                        break;
                    default:
                        MessageBox.Show("Unknown Country in AppSettings.  Please Check.");
                        excelApp.Quit();
                        Environment.Exit(0);
                        sheets = new List<Action<Excel.Application, SqlConnection, string, string>> { };
                        break;
                }

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
            //}
            //catch (Exception ex)  // Change to a file based log
            //{
            //    MessageBox.Show(ex.ToString(), "Error", MessageBoxButton.OK, MessageBoxImage.Error);
            //}
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
