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
using System.Reflection;

namespace magentr
{
    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {
        public string dirNewRequest = "";
        private Dictionary<string, string> dictRequestRawData
            = new Dictionary<string, string>();
        private Dictionary<string, string> dictCheckBox
            = new Dictionary<string, string>();
        //private delegate void ConvertRange(EXCEL.Range TargetRange);
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
            lbxDebug.Items.Add(OpenFileNew + " Selected.");
            dirNewRequest = OpenFileNew.FileName;
            if (dirNewRequest != "")
            {
                FetchNewRequest(dirNewRequest);
            }
            else
            {
                lbxDebug.Items.Add("No file selected.");
            }
            lbxDebug.Items.Add((string.Format("Button Click Ran for: {0}", 
                (DateTime.Now - timeStart).ToString("hh':'mm':'ss"))));
        }

        private async void FetchNewRequest(string dirNew)
        {
            DateTime timeStart = DateTime.Now;

            var UpdateProgressBar = new Progress<int>(value => pbarMain.Value = value);
            var SetProgressBarMax = new Progress<int>(value => pbarMain.Maximum = value);
            var PrintDebugListBox = new Progress<string>(value => 
            {
                lbxDebug.Items.Add(DateTime.Now.ToString("hh':'mm':'ss") + " % " + value);
                svDebug.ScrollToBottom();
            });
            //This procedure fetches information from Excel
            FileInfo InputFile = new FileInfo(dirNew);
            await Task.Run(() =>
            {
                SyncVonExcel(InputFile
                    , UpdateProgressBar
                    , SetProgressBarMax
                    , PrintDebugListBox);
            });
        }

        private void SyncVonExcel(FileInfo inputfile 
            , IProgress<int> reportProgressBar
            , IProgress<int> setProgressBarMax
            , IProgress<string> printDebugListBox
            , out RequestSheet newRequest)
        {
            newRequest = new RequestSheet();
            dictRequestRawData
            = new Dictionary<string, string>();
            dictCheckBox
            = new Dictionary<string, string>();
            DateTime timeStart = DateTime.Now;
            setProgressBarMax.Report(100);
            reportProgressBar.Report(1);
            if (!inputfile.Exists) //throw new Exception("M/Agent Application File Does not Exist!");
            {
                printDebugListBox.Report("M/Agent Application File Does not Exist!");
                return;
            }
            EXCEL.Application xlApp = new EXCEL.Application();              reportProgressBar.Report(20);
            EXCEL.Workbooks xlWorkbooks = xlApp.Workbooks;                  reportProgressBar.Report(40);
            EXCEL.Workbook xlWbk = xlWorkbooks.Open(inputfile.FullName);    reportProgressBar.Report(60);
            EXCEL.Worksheet xlSht = xlWbk.ActiveSheet;                      reportProgressBar.Report(80);
            // EXCEL.Range xlRange = xlSht.UsedRange;
            reportProgressBar.Report(100);
            Debug.Print("Loading Completed.");

            #region ---- Fetch M/Agent Information ----
            printDebugListBox.Report("Beginning Fetching M/Agent Information.");
            printDebugListBox.Report("Getting Range Dictionary Delegate.");

            void RangeToDict(EXCEL.Range TargetRange)
            {
                dictRequestRawData[TargetRange.Address]
                = (string)TargetRange.Value;
            }

            //Cell Range: D5, S163
            EXCEL.Range FormArea = xlSht.Range["D5", "S163"];
            printDebugListBox.Report("Calculating Total Form Area Ranges");
            int FormAreaRangCount = FormArea.Count;
            int FormAreaCurrentCount = 0;
            printDebugListBox.Report("Total Form Area Ranges : " + FormAreaRangCount);
            setProgressBarMax.Report(FormAreaRangCount);
            reportProgressBar.Report(FormAreaCurrentCount);

            foreach (EXCEL.Range r in FormArea)
            {
                if(r.Value == null)
                {
                    RangeToDict(r);
                }
                reportProgressBar.Report(++FormAreaCurrentCount);
            }
            printDebugListBox.Report("Sync Target Area Complete.");
            //Trying to fetch form public Dictionary Object.

            EXCEL.Shapes xlShapes = xlSht.Shapes;

            int WorkLoad_Total = xlShapes.Count;
            setProgressBarMax.Report(WorkLoad_Total);
            int WorkdLoad_Current = 0;

            IEnumerable<EXCEL.Shape> xlCheckBoxes =
                from EXCEL.Shape s in xlShapes
                where s.Name.Contains("チェック") //v0.0.0.1 Need a regex to match both en and jp version.
                select s;

            //Dictionary<string, string> dicCheckedBoxes = 
            //    new Dictionary<string, string>(); //combine with the dictRequestRawData object

            setProgressBarMax.Report(xlCheckBoxes.Count());

            foreach(EXCEL.Shape s in xlCheckBoxes)
            {
                //s.TopLeftCell.Interior.Color = EXCEL.XlRgbColor.rgbRed;
                //printDebugListBox.Report(string.Format("|{0,-30}|{1,-30}|{2,-30}|"
                //    , s.Name
                //    , (string)s.TopLeftCell.Offset[0, 1].Value
                //    , ((double)s.OLEFormat.Object.Value).ToString()));
                if ((double)s.OLEFormat.Object.Value == 1)
                {
                    dictCheckBox.Add(
                        s.TopLeftCell.Address
                        , (string)s
                        .TopLeftCell
                        .Offset[0, 1]
                        .Value);
                }
                reportProgressBar.Report(++WorkdLoad_Current);
            }
            /* Testing if CheckBox Function Works
            setProgressBarMax.Report(0);
            CheckBoxValue(
                xlSht.Range["H32","K33"]
                , dicCheckedBoxes
                , reportProgressBar
                , setProgressBarMax
                , out string TestCheckBoxValue);
            Debug.Print(TestCheckBoxValue);*/

            //MessageBox.Show(TestCheckBoxValue);
            //Show Entire Value Dictionary Object:
            
            #endregion ---- Fetch M/Agent Information ----
            xlWbk.Close(false, Missing.Value, Missing.Value); //Arguments in this will cause excel to exist without saving.
            xlWorkbooks.Close();
            xlApp.Quit();
            Marshal.ReleaseComObject(xlApp);
            Marshal.ReleaseComObject(xlWorkbooks);
            Marshal.ReleaseComObject(xlWbk);
            Marshal.ReleaseComObject(xlSht);
            //Marshal.ReleaseComObject(xlRange);
            #region --------Test two Dictionary Objects---------
            /*
            setProgressBarMax.Report(dictRequestRawData.Count);
            int CurrentProgress = 0;
            foreach (KeyValuePair<string, string> k in dictRequestRawData)
            {
                printDebugListBox.Report(string.Format("{0,-7}|{1}", k.Key, k.Value));
                reportProgressBar.Report(++CurrentProgress);
            }
            setProgressBarMax.Report(dictCheckBox.Count);
            CurrentProgress = 0;
            foreach (KeyValuePair<string, string> k in dictCheckBox)
            {
                printDebugListBox.Report(string.Format("{0,-7}|{1}", k.Key, k.Value));
                reportProgressBar.Report(++CurrentProgress);
            }
            */
            #endregion --------Test two Dictionary Objects---------

            RequestColumns colH = new RequestColumns();
            RequestColumns colL = new RequestColumns();
            RequestColumns colP = new RequestColumns();



            newRequest = new RequestSheet(colH, colL, colP);
            printDebugListBox.Report("Proceedure completed.");
            reportProgressBar.Report(0);
            printDebugListBox.Report(string.Format("Open Excel Async Ran for: {0}",
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
            }

            Debug.Print("CheckBoxValue:Ended");
            Check_Box_Value = result;
        }
    }
}
