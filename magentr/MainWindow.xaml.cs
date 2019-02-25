using Microsoft.Win32;
using System;
using System.Collections.Generic;
using System.IO;
using System.Runtime.InteropServices;
using System.Threading.Tasks;
using System.Windows;
using System.Linq;
using System.Windows.Media;
using EXCEL = Microsoft.Office.Interop.Excel;
using System.Diagnostics;

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
            DateTime timeStart = DateTime.Now;
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
            Debug.Print(string.Format("Button Click Ran for: {0}", 
                (DateTime.Now - timeStart).ToString("hh':'mm':'ss")));
        }

        private async void FetchNewRequest(string dirNew)
        {
            DateTime timeStart = DateTime.Now;

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
            DateTime timeStart = DateTime.Now;
            setProgressMax.Report(100);
            reportProgress.Report(1);
            if (!inputfile.Exists) throw new Exception("M/Agent Application File Does not Exist!");
            EXCEL.Application xlApp = new EXCEL.Application(); reportProgress.Report(20);
            EXCEL.Workbooks xlWorkbooks = xlApp.Workbooks; reportProgress.Report(40);
            EXCEL.Workbook xlWbk = xlWorkbooks.Open(inputfile.FullName); reportProgress.Report(60);
            EXCEL.Worksheet xlSht = xlWbk.ActiveSheet; reportProgress.Report(80);
            // EXCEL.Range xlRange = xlSht.UsedRange;
            reportProgress.Report(100);
            Debug.Print("Loading Completed.");

            #region ---- Fetch M/Agent Information ----
            //Calculating Total Work Load:
            //int WorkLoad_Total = xlRange.Count;
            //setProgressMax.Report(WorkLoad_Total);
            //int WorkdLoad_Current = 0;
            //Loading 
            //foreach(EXCEL.Range r in xlRange)
            //{
            //    reportProgress.Report(++WorkdLoad_Current);
            //}
            //reportProgress.Report(0);

            EXCEL.Shapes xlShapes = xlSht.Shapes;

            int WorkLoad_Total = xlShapes.Count;
            setProgressMax.Report(WorkLoad_Total);
            int WorkdLoad_Current = 0;

            IEnumerable<EXCEL.Shape> xlCheckBoxes =
                from EXCEL.Shape s in xlShapes
                where s.Name.Contains("チェック") 
                select s;

            Dictionary<string, string> dicCheckedBoxes = 
                new Dictionary<string, string>();
            setProgressMax.Report(xlCheckBoxes.Count());
            foreach(EXCEL.Shape s in xlCheckBoxes)
            {
                s.TopLeftCell.Interior.Color = EXCEL.XlRgbColor.rgbRed;
                Debug.Print("|{0,-14}|{1,-20}|{2,-20}|"
                    , s.Name
                    , (string)s.TopLeftCell.Offset[0, 1].Value
                    , ((double)s.OLEFormat.Object.Value).ToString());
                dicCheckedBoxes.Add(
                    s.TopLeftCell.Address
                    ,(string)s
                    .TopLeftCell
                    .Offset[0, 1]
                    .Value);
                reportProgress.Report(++WorkdLoad_Current);
            }

            setProgressMax.Report(0);
            CheckBoxValue(
                xlSht.Range["H32","K33"]
                , dicCheckedBoxes
                , reportProgress
                , setProgressMax
                , out string TestCheckBoxValue);
            Debug.Print(TestCheckBoxValue);

            //MessageBox.Show(TestCheckBoxValue);

            #endregion ---- Fetch M/Agent Information ----
            xlWbk.Close();
            xlWorkbooks.Close();
            xlApp.Quit();
            Marshal.ReleaseComObject(xlApp);
            Marshal.ReleaseComObject(xlWorkbooks);
            Marshal.ReleaseComObject(xlWbk);
            Marshal.ReleaseComObject(xlSht);
            //Marshal.ReleaseComObject(xlRange);
            Debug.Print(string.Format("Open Excel Async Ran for: {0}",
                (DateTime.Now - timeStart).ToString("hh':'mm':'ss")));


        }

        private void CheckBoxValue (
            EXCEL.Range RangeGroup
            , Dictionary<string, string> dicCheckedBoxes
            , IProgress<int> reportProgress
            , IProgress<int> setProgressMax
            , out string Check_Box_Value)
        {
            //Example, Range("H32:K33")
            string result = null;

            //Dictionary<EXCEL.Range, string> CheckingRange = new Dictionary<EXCEL.Range, string>();
            reportProgress.Report(0);
            setProgressMax.Report(RangeGroup.Count);
            Debug.Print("CheckBoxValue:Started");
            Debug.Print("Input Range Counter: " + RangeGroup.Count);
            int Current_Progress = 0;
            foreach (EXCEL.Range r in RangeGroup)
            {
                reportProgress.Report(++Current_Progress);
                Debug.Print("Checking Address:" + r.Address);
                if (dicCheckedBoxes.ContainsKey(r.Address))
                {

                    result = dicCheckedBoxes[r.Address];
                }
                //foreach (var k in dicCheckedBoxes.ToList())
                //{
                //    Debug.Print((string)r.Value);
                //    Debug.Print(string.Format("{0}:{1}", k.Key, k.Value));
                //}
            }

            Debug.Print("CheckBoxValue:Ended");
            Check_Box_Value = result;
        }
    }
}
