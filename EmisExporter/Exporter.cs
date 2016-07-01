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
            TeacherBaseSQL = @"insert into dbo.#TeacherBaseTable (ISCED, SCHOOLTYPE, GENDER, TRAINED, QUALIFIED, COUNT) " + String.Format(TeacherBaseSQL, year);

            DbInsertCommand = new SqlCommand(TeacherBaseSQL, emisDBConn);
            DbInsertCommand.ExecuteNonQuery();

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
