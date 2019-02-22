using Microsoft.Win32;
using System;
using System.IO;
using System.Runtime.InteropServices;
using System.Threading.Tasks;
using System.Windows;
using EXCEL = Microsoft.Office.Interop.Excel;

namespace magentr
{
    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {
        public string dirNewRequest = "";
        
        public MainWindow()
        {
            InitializeComponent();
        }

        private void OnNewRequestClick(object sender, RoutedEventArgs e)
        {
            //Initiate Load New Request Form Procedure
            // Step 001 Open file
            OpenFileDialog OpenFileNew = new OpenFileDialog();
            OpenFileNew.DefaultExt = ".xlsx;.xls";
            OpenFileNew.Filter = "Excel Worksheet (.xls;.xlsx)|*.xls;*.xlsx";
            OpenFileNew.ShowDialog();
            //MessageBox.Show(OpenFileNew + " Selected.");
            dirNewRequest = OpenFileNew.FileName;
            if (dirNewRequest != "")
            {
                FetchNewRequest(dirNewRequest);
            }
            else
            {
                MessageBox.Show("No file selected.");
            }
        }

        private async void FetchNewRequest(string dirNew)
        {
            var UpdateProgressBar = new Progress<int>(value => pbarMain.Value = value);
            var SetProgressBarMax = new Progress<int>(value => pbarMain.Maximum = value);

            //This procedure fetches information from Excel
            FileInfo InputFile = new FileInfo(dirNew);
            await Task.Run(() =>
            {
                SyncVonExcel(InputFile, UpdateProgressBar, SetProgressBarMax);
            });
        }

        private void SyncVonExcel(FileInfo inputfile, 
            IProgress<int> reportProgress, 
            IProgress<int> setProgressMax)
        {
            setProgressMax.Report(100);
            reportProgress.Report(1);
            if (!inputfile.Exists) throw new Exception("M/Agent Application File Does not Exist!");
            EXCEL.Application xlApp = new EXCEL.Application(); reportProgress.Report(20);
            EXCEL.Workbooks xlWorkbooks = xlApp.Workbooks; reportProgress.Report(40);
            EXCEL.Workbook xlWbk = xlWorkbooks.Open(inputfile.FullName); reportProgress.Report(60);
            EXCEL.Worksheet xlSht = xlWbk.ActiveSheet; reportProgress.Report(80);
            EXCEL.Range xlRange = xlSht.UsedRange;
            reportProgress.Report(100);
            MessageBox.Show("Loading Completed.");
            #region ---- Fetch M/Agent Information ----
            //Calculating Total Work Load:
            int WorkLoad_Total = xlRange.Count;
            setProgressMax.Report(WorkLoad_Total);
            int WorkdLoad_Current = 0;
            //Loading 
            foreach(EXCEL.Range r in xlRange)
            {
                reportProgress.Report(++WorkdLoad_Current);
            }
            reportProgress.Report(0);
            EXCEL.Shapes xlShapes = xlSht.Shapes;

            WorkLoad_Total = xlShapes.Count;
            setProgressMax.Report(WorkLoad_Total);
            WorkdLoad_Current = 0;
            foreach(EXCEL.Shape s in xlShapes)
            {
                reportProgress.Report(++WorkdLoad_Current);
            }




            #endregion ---- Fetch M/Agent Information ----
            xlWbk.Close();
            xlWorkbooks.Close();
            xlApp.Quit();
            Marshal.ReleaseComObject(xlApp);
            Marshal.ReleaseComObject(xlWorkbooks);
            Marshal.ReleaseComObject(xlWbk);
            Marshal.ReleaseComObject(xlSht);
            Marshal.ReleaseComObject(xlRange);
        } 
    }
}
