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
using System.Net;
using System.Data.OleDb;
using System.Text.RegularExpressions;

namespace magentr
{
    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {
        public string dirNewRequest = "";
        private Dictionary<string, string> dictRequestRawData //Stores Cell Info
            = new Dictionary<string, string>();
        private Dictionary<string, string> dictCheckBox       //Stores Checkbox Info
            = new Dictionary<string, string>();

        public string RequestBango = "";

        public MainWindow()
        {
            InitializeComponent();
        }

        private async void OnNewRequestClick(object sender, RoutedEventArgs e)
        {
            //Initiate Load New Request Form Procedure
            #region Open File Dialog
            DateTime timeStart = DateTime.Now;
            OpenFileDialog OpenFileNew = new OpenFileDialog();
            OpenFileNew.DefaultExt = ".xlsx;.xls";
            OpenFileNew.Filter = "Excel Worksheet (.xls;.xlsx)|*.xls;*.xlsx";
            OpenFileNew.ShowDialog();
            lbxDebug.Items.Add(OpenFileNew.FileName + " Selected.");
            dirNewRequest = OpenFileNew.FileName;
            FileInfo RequestFileInfo = null;
            try
            {
                RequestFileInfo = new FileInfo(OpenFileNew.FileName);
            }
            catch (Exception ex)
            {
                lbxDebug.Items.Add("[Warning...] Invalid File Name or File Not selected. Existing.");
                return;
            }
            RequestBango = RequestFileInfo.Name;

            #endregion Open File Dialog
            if (dirNewRequest != "")
            {
                await FetchNewRequest(dirNewRequest);
            }
            else
            {
                lbxDebug.Items.Add("No file selected.");
            }
            lbxDebug.Items.Add((string.Format("Button Click Ran for: {0}", 
                (DateTime.Now - timeStart).ToString("hh':'mm':'ss"))));
        }



        private async Task FetchNewRequest(string dirNew)
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
            await Task.Run(() => SyncVonExcel(InputFile
                    , UpdateProgressBar
                    , SetProgressBarMax
                    , PrintDebugListBox));
        }

        private void SyncVonExcel(FileInfo inputfile 
            , IProgress<int> reportProgressBar
            , IProgress<int> setProgressBarMax
            , IProgress<string> printDebugListBox)
        {
            dictRequestRawData = new Dictionary<string, string>();
            dictCheckBox = new Dictionary<string, string>();
            DateTime timeStart = DateTime.Now;
            printDebugListBox.Report("Starting SyncVonExcel Procedure.");
            setProgressBarMax.Report(100);
            reportProgressBar.Report(1);
            if (!inputfile.Exists) //throw new Exception("M/Agent Application File Does not Exist!");
            {
                printDebugListBox.Report("[Error!!!] M/Agent Application File Does not Exist! Exiting...");
                return;
            }
            printDebugListBox.Report("Loading Excel File into Memory...");
            EXCEL.Application xlApp = new EXCEL.Application();              reportProgressBar.Report(20);
            EXCEL.Workbooks xlWorkbooks = xlApp.Workbooks;                  reportProgressBar.Report(40);
            EXCEL.Workbook xlWbk = xlWorkbooks.Open(inputfile.FullName);    reportProgressBar.Report(60);
            EXCEL.Worksheet xlSht = xlWbk.ActiveSheet;                      reportProgressBar.Report(80);
            reportProgressBar.Report(100);
            printDebugListBox.Report("Loading Completed.");
            reportProgressBar.Report(0);
            #region ---- Fetch M/Agent Information ----
            printDebugListBox.Report("Beginning Fetching M/Agent Information.");
            printDebugListBox.Report("Getting Range Dictionary Delegate.");

            void RangeToDict(EXCEL.Range TargetRange)
            {
                dictRequestRawData[TargetRange.Address]
                = Convert.ToString(TargetRange.Value);
            }

            //Cell Range: D5, S163
            printDebugListBox.Report("Assigning Worksheet Object to Target Range = D5:S163");
            EXCEL.Range FormArea = xlSht.Range["D5", "S163"]; //This is too many, Get only non null ones.
            IEnumerable<EXCEL.Range> ieFilledRange =
                from EXCEL.Range r in FormArea
                where r.Value != null
                select r;
            printDebugListBox.Report("Calculating Total Form Area Ranges");
            int FormAreaRangCount = FormArea.Count;
            int WorkLoad_Total = ieFilledRange.Count();
            int WorkdLoad_Current = 0;
            printDebugListBox.Report(string.Format(
                "Total Form Area Ranges Valid/Total: {0}/{1}"
                , WorkLoad_Total, FormAreaRangCount));
            setProgressBarMax.Report(WorkLoad_Total);
            reportProgressBar.Report(WorkdLoad_Current);
            printDebugListBox.Report("Assigning Range Objects to Local Dictionary Object");
            foreach (EXCEL.Range r in ieFilledRange)
            {
                RangeToDict(r);
                reportProgressBar.Report(++WorkdLoad_Current);
            }
            printDebugListBox.Report("Sync Target Range Area Sync to Dictionary Complete.");
            //Trying to fetch form public Dictionary Object.
            printDebugListBox.Report("Assigning Worksheet Shapes to Target shapes.");
            EXCEL.Shapes xlShapes = xlSht.Shapes;

            WorkLoad_Total = xlShapes.Count;
            setProgressBarMax.Report(WorkLoad_Total);
            WorkdLoad_Current = 0;

            //Regex mCheckBox = new Regex(@"(チェック|Check Box)");

            IEnumerable<EXCEL.Shape> xlCheckBoxes =
                from EXCEL.Shape s in xlShapes
                where (s.Name.Contains("チェック") || s.Name.Contains("Check Box")) //v0.0.0.1 Need a regex to match both en and jp version.
                && (double)s.OLEFormat.Object.Value == 1 //Select only selected Value
                select s;

            WorkLoad_Total = xlCheckBoxes.Count();
            printDebugListBox.Report(string.Format(
                "Total Worksheet Shapes Valid/Total: {0}/{1}"
                , WorkLoad_Total, FormAreaRangCount));

            //Dictionary<string, string> dicCheckedBoxes = 
            //    new Dictionary<string, string>(); //combine with the dictRequestRawData object

            setProgressBarMax.Report(WorkLoad_Total);

            foreach(EXCEL.Shape s in xlCheckBoxes)
            {
                dictCheckBox.Add(
                    s.TopLeftCell.Address
                    , (string)s.TopLeftCell
                    .Offset[0, 1].Value);
                reportProgressBar.Report(++WorkdLoad_Current);
            }
            
            #endregion ---- Fetch M/Agent Information ----
            xlWbk.Close(false, Missing.Value, Missing.Value); //Arguments in this will cause excel to exist without saving.
            xlWorkbooks.Close();
            xlApp.Quit();
            Marshal.ReleaseComObject(xlApp);
            Marshal.ReleaseComObject(xlWorkbooks);
            Marshal.ReleaseComObject(xlWbk);
            Marshal.ReleaseComObject(xlSht);
            //Marshal.ReleaseComObject(xlRange);
            #region --------Test two Dictionary Objects (Not Used)---------
            /* Below Are Listing of all contents of the two objects.
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
            //Generating Object Count Report.
            string ReportDictionaryCount = "For {0}, there are {1} elements.";
            printDebugListBox.Report(
                string.Format(ReportDictionaryCount
                , "dictRequestRawData"
                , dictRequestRawData.Count));
            printDebugListBox.Report(
                string.Format(ReportDictionaryCount
                , "dictCheckBox"
                , dictCheckBox.Count));
            #endregion --------Test two Dictionary Objects---------
            OleDbConnectionStringBuilder connSB = new OleDbConnectionStringBuilder();
            connSB.Provider = "Microsoft.ACE.OLEDB.12.0";
            connSB.DataSource = @"C:\Users\MoChen\source\repos\magentr\magentr\magentr.accdb";
            string connString = connSB.ToString();
            #region Connect to Database with Connection String

            using (OleDbConnection conn = new OleDbConnection(connString))
            {
                conn.Open();
                OleDbCommand SelectRequest = new OleDbCommand(
                    "SELECT tbRequestForm.* FROM tbRequestForm WHERE tbRequestForm.;", conn);

                OleDbCommand InsertRequest = new OleDbCommand(
                    @"INSERT INTO tbRequestForm 
                   (RequestBango, RequestFileName
                    , DateApplied, Applier
                    , Email, Phone, Approver, Comment) 
                  Values (@requestBango, @requestFileName                                                                                                                                
                    , @dateApplied, @applier
                    , @email ,@phone, @approver
                    , @comment);", conn);

                InsertRequest.CommandType = System.Data.CommandType.Text;
                InsertRequest.Parameters.AddWithValue("@requestBango", RequestBango.Substring(0, 15));
                InsertRequest.Parameters.AddWithValue("@requestFileName", RequestBango);
                InsertRequest.Parameters.AddWithValue("@dateApplied", ValidDate("$H$7"));
                InsertRequest.Parameters.AddWithValue("@applier", ValidDic("$H$8"));
                InsertRequest.Parameters.AddWithValue("@email", ValidDic("$H$9"));
                InsertRequest.Parameters.AddWithValue("@phone", ValidDic("$H$10"));
                InsertRequest.Parameters.AddWithValue("@approver", ValidDic("$H$11"));
                InsertRequest.Parameters.AddWithValue("@comment", ValidDic("$E$161"));
                printDebugListBox.Report(InsertRequest.CommandText);
                try
                {
                    int RowsAffected = InsertRequest.ExecuteNonQuery();
                    printDebugListBox.Report("Request Table Successful, Rows Affected: " + RowsAffected);
                }
                catch(OleDbException ex)
                {
                    printDebugListBox.Report(ex.Message);
                }
                InsertRequest.Parameters.Clear(); 

            }

            printDebugListBox.Report(SyncColumn(connString, "H", "J"));
            printDebugListBox.Report(SyncColumn(connString, "L", "N"));
            printDebugListBox.Report(SyncColumn(connString, "P", "R"));
            
            #endregion  Connect to Database with Connection String

            printDebugListBox.Report("Proceedure completed.");
            reportProgressBar.Report(0);
            printDebugListBox.Report(string.Format("Open Excel Async Ran for: {0}",
                (DateTime.Now - timeStart).ToString("hh':'mm':'ss")));
        }

        private string CheckBoxValue (
            string dictRange)
        {
            //Example, Range("H32:K33") => Regex = @"\$[H-K]\$(32|33)"
            //Generate dictRange Regular Expression
            Regex rxValidateRange = new Regex(@"[A-Z]+\d+\:[A-Z]+\d+");
            if (!rxValidateRange.Match(dictRange).Success)
                throw new Exception("String do not Match format " + rxValidateRange.ToString());
            // Parse Range("H32:K33") into regex.
            string[] dictRangeSplit = dictRange.Split(':');
            string firstCol = dictRangeSplit[0].Substring(0, 1); //Doesn't work with AA and up.
            int firstRow = Convert.ToInt32(dictRangeSplit[0].Remove(0, 1));
            string secondCol = dictRangeSplit[1].Substring(0, 1); //Doesn't work with AA and up.
            int secondRow = Convert.ToInt32(dictRangeSplit[1].Remove(0, 1));
            
            //Need to generate {3}, H31:K37 = "31|32|33|34|35|36|37"
            string AllRows = "";
            for (int i = firstRow; i <= secondRow; i++)
            {
                AllRows += i.ToString() + '|';

            }//"31|32|33|34|35|36|37|"

            AllRows = AllRows.Remove(AllRows.Length - 1, 1); //"31|32|33|34|35|36|37"
            Debug.Print(AllRows);
            string rxRange = string.Format(@"\$[{0}-{1}]\$({2})"
                , firstCol
                , secondCol
                , AllRows);

            Debug.Print(rxRange);
            Regex rxRangeMatch = new Regex(rxRange);
            string result = "未選択";
            try
            {
                Debug.Print(string.Format("Testing against {0}", rxRangeMatch.ToString()));
                var EnumResult = from KeyValuePair<string, string> Checked in dictCheckBox
                                 where rxRangeMatch.IsMatch(Checked.Key)
                                 select Checked;
                Debug.Print(string.Format("Checked Box Count: {0}", EnumResult.Count()));
                switch(EnumResult.Count())
                {
                    case 0:
                        result = "未選択";
                        break;
                    case 1 :
                        result = (string)EnumResult.First().Value;
                        break;
                    default :
                        result = "無効な選択";
                        break;
                }
            }
            catch (Exception ex)
            {
                Debug.Print(ex.Message);
            }



            return result;
        }

        private string ValidDic(string KeyVal)
        {

            string result = dictRequestRawData.ContainsKey(KeyVal) ? dictRequestRawData[KeyVal] : "未入力";
            //Need to Validate "N/A" String =? @"N/?A"
            Regex rx = new Regex(@"(N/?A)");
            if (rx.IsMatch(result)) result = "未入力";
            return result;
        }
        private DateTime ValidDate(string KeyVal)
        {
            string resultstring =
                dictRequestRawData.ContainsKey(KeyVal) ? dictRequestRawData[KeyVal] : "1900-01-01";
            return DateTime.Parse(resultstring);
        }

        private string SyncColumn
            ( string ConnectionString
            , string ColumnStart
            , string ColumnFinish)
        {
            //Before Sync, Judge if any of the must fill values are invalid, 
            //if yes, direcly return Error Message without connecting to Database.
            //Rule No1: $32:^33 Cannot be "Not Selected"
            //Rule No2: $51 must be non-empty.
            //We first judge these values, in the sync we can directly use these string variables
            string RegisterType = CheckBoxValue(ColumnStart + "32:" + ColumnFinish + "33");
            string PhysicalHostPRI = ValidDic("$" + ColumnStart + "$51");
            if (RegisterType.Contains("選択") || PhysicalHostPRI.Length < 8)
                return "[Warning...] Either Apply Type not Selected, or no primary host specified. Aborting Sync.";
            using (OleDbConnection conn = new OleDbConnection(ConnectionString))
            {
                conn.Open();
                var InsertRequest = new OleDbCommand(
@"INSERT INTO tbAgents (
rlnFileName,  rlnBango,  ApplyType,  ChangePoint,  SIer,  ServerPIC,  SystemID,  SystemName
,  SystemSubName,  NetworkLocation,  NetworkArea,  ServerVIP,  ServerPRI,  ServerSEC
,  MStMACommunicationPort,  MA_InstallDate,  MS_Connection,  JobStartDate,  JobCount
,  HasCallorder,  HasFirewall,  MA_Version,  IsFirstTime,  IsProduction,  TestDoneDate
,  CostFrom,  CostFromSystemName,  CostFromSubSystemName,  HasSundayJobs,  HasRelatedSystems
,  RelatedSystemID,  RelatedSystemName,  RelatedSystemSubName,  RelatedSystemDatacenter
,  MAtMSCommunicationPort,  MSVIP,  MSPRI,  MSSEC 
) VALUES (
@rlnfileName, @rlnbango, @applyType, @changePoint, @sIer, @serverPIC, @systemID, @systemName
, @systemSubName, @networkLocation, @networkArea, @serverVIP, @serverPRI, @serverSEC
, @mStMACommunicationPort, @mA_InstallDate, @mS_Connection, @jobStartDate, @jobCount
, @hasCallorder, @hasFirewall, @mA_Version, @isFirstTime, @isProduction, @testDoneDate
, @costFrom, @costFromSystemName, @costFromSubSystemName, @hasSundayJobs, @hasRelatedSystems
, @relatedSystemID, @relatedSystemName, @relatedSystemSubName, @relatedSystemDatacenter
, @mAtMSCommunicationPort, @mSVIP, @mSPRI, @mSSEC
);", conn);

                InsertRequest.Parameters.AddWithValue("@requestFileName", RequestBango);
                InsertRequest.Parameters.AddWithValue("@requestBango", RequestBango.Substring(0, 15));

                InsertRequest.Parameters.AddWithValue("@applyType", RegisterType);

                InsertRequest.Parameters.AddWithValue("@changePoint", CheckBoxValue(ColumnStart + "34:" + ColumnFinish + "36"));
                InsertRequest.Parameters.AddWithValue("@sIer", ValidDic("$" + ColumnStart + "$37"));
                InsertRequest.Parameters.AddWithValue("@serverPIC", ValidDic("$" + ColumnStart + "$38"));
                InsertRequest.Parameters.AddWithValue("@systemID", ValidDic("$" + ColumnStart + "$39"));
                InsertRequest.Parameters.AddWithValue("@systemName", ValidDic("$" + ColumnStart + "$40"));
                InsertRequest.Parameters.AddWithValue("@systemSubName", ValidDic("$" + ColumnStart + "$41"));

                InsertRequest.Parameters.AddWithValue("@networkLocation", CheckBoxValue(ColumnStart + "42:" + ColumnFinish + "43"));

                InsertRequest.Parameters.AddWithValue("@networkArea", CheckBoxValue(ColumnStart + "44:" + ColumnFinish + "47"));
                InsertRequest.Parameters.AddWithValue("@serverVIP", ValidDic("$" + ColumnStart + "$49"));
                InsertRequest.Parameters.AddWithValue("@serverPRI", PhysicalHostPRI);
                InsertRequest.Parameters.AddWithValue("@serverSEC", ValidDic("$" + ColumnStart + "$64"));
                InsertRequest.Parameters.AddWithValue("@mStMACommunicationPort", ValidDic("$" + ColumnStart + "$77"));
                InsertRequest.Parameters.AddWithValue("@mA_InstallDate", ValidDate("$" + ColumnStart + "$78"));
                InsertRequest.Parameters.AddWithValue("@mS_Connection", ValidDate("$" + ColumnStart + "$79"));
                InsertRequest.Parameters.AddWithValue("@jobStartDate", ValidDate("$" + ColumnStart + "$80"));
                InsertRequest.Parameters.AddWithValue("@jobCount", ValidDic("$" + ColumnStart + "$81"));
                InsertRequest.Parameters.AddWithValue("@hasCallorder", CheckBoxValue(ColumnStart + "82:" + ColumnFinish + "82"));
                InsertRequest.Parameters.AddWithValue("@hasFirewall", CheckBoxValue(ColumnStart + "83:" + ColumnFinish + "83"));
                InsertRequest.Parameters.AddWithValue("@mA_Version", ValidDic("$" + ColumnStart + "$84"));
                InsertRequest.Parameters.AddWithValue("@isFirstTime", CheckBoxValue(ColumnStart + "85:" + ColumnFinish + "85"));
                InsertRequest.Parameters.AddWithValue("@isProduction", CheckBoxValue(ColumnStart + "86:" + ColumnFinish + "86"));
                InsertRequest.Parameters.AddWithValue("@testDoneDate", ValidDate("$" + ColumnStart + "$87"));
                InsertRequest.Parameters.AddWithValue("@costFrom", CheckBoxValue(ColumnStart + "88:" + ColumnFinish + "88"));
                InsertRequest.Parameters.AddWithValue("@costFromSystemName", ValidDic("$" + ColumnStart + "$89"));
                InsertRequest.Parameters.AddWithValue("@costFromSubSystemName", ValidDic("$" + ColumnStart + "$90"));
                InsertRequest.Parameters.AddWithValue("@hasSundayJobs", CheckBoxValue(ColumnStart + "91:" + ColumnFinish + "91"));
                InsertRequest.Parameters.AddWithValue("@hasRelatedSystems", CheckBoxValue(ColumnStart + "92:" + ColumnFinish + "92"));
                InsertRequest.Parameters.AddWithValue("@relatedSystemID", ValidDic("$" + ColumnStart + "$93"));
                InsertRequest.Parameters.AddWithValue("@relatedSystemName", ValidDic("$" + ColumnStart + "$94"));
                InsertRequest.Parameters.AddWithValue("@relatedSystemSubName", ValidDic("$" + ColumnStart + "$95"));
                InsertRequest.Parameters.AddWithValue("@relatedSystemDatacenter", ValidDic("$" + ColumnStart + "$96"));
                InsertRequest.Parameters.AddWithValue("@mAtMSCommunicationPort", ValidDic("$" + ColumnStart + "$97"));
                InsertRequest.Parameters.AddWithValue("@mSVIP", ValidDic("$" + ColumnStart + "$98"));
                InsertRequest.Parameters.AddWithValue("@mSPRI", ValidDic("$" + ColumnStart + "$99"));
                InsertRequest.Parameters.AddWithValue("@mSSEC", ValidDic("$" + ColumnStart + "$100"));

                try
                {
                    int RowsAffected = InsertRequest.ExecuteNonQuery();
                    return "Agent Table Successful, Rows Affected: " + RowsAffected;
                }
                catch (OleDbException ex)
                {
                    return "Agent Table Sync Failed:" + ex.Message;
                }

            }

        }
    }
}
