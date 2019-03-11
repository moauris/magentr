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
using System.Data;

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

        private string connString = "";
        private DateTime lasttime = DateTime.Now;
        private void printDebug(string ThisName, string message)
        {
            //example: [2019-01-01 22:33:55] <~.OnNewRequestClick> % Some Message about Debug Info.
            
            string messageformat = "[{0} | {3}] <~.{1}> % {2}";
            string TimeStamp = DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss");
            //var st = new StackTrace();
            double dur = (DateTime.Now - lasttime).TotalSeconds;
            string OutputMessage = string.Format(messageformat
                , TimeStamp, ThisName, message, Math.Round(dur, 4));
            Debug.Print(OutputMessage);
            lasttime = DateTime.Now;

        }

        public MainWindow()
        {
            InitializeComponent();
        }

        private async void OnNewRequestClick(object sender, RoutedEventArgs e)
        {
            string st = "OnNewRequestClick(object sender, RoutedEventArgs e){}";
            //Initiate Load New Request Form Procedure
            #region Open File Dialog
            DateTime timeStart = DateTime.Now;
            OpenFileDialog OpenFileNew = new OpenFileDialog();
            OpenFileNew.DefaultExt = ".xlsx;.xls";
            OpenFileNew.Filter = "Excel Worksheet (.xls;.xlsx)|*.xls;*.xlsx";
            OpenFileNew.ShowDialog();
            printDebug(st, OpenFileNew.FileName + " Selected.");
            dirNewRequest = OpenFileNew.FileName;
            FileInfo RequestFileInfo = null;
            try
            {
                RequestFileInfo = new FileInfo(OpenFileNew.FileName);
            }
            catch (Exception ex)
            {
                printDebug(st, "[Warning...] Invalid File Name or File Not selected. Existing.");
                printDebug(st, ex.Message);
                return;
            }
            RequestBango = RequestFileInfo.Name;

            #endregion Open File Dialog
            if (dirNewRequest != "")
            {
                //Start Procedure when fetched file is not null.
                OleDbConnectionStringBuilder connSB = new OleDbConnectionStringBuilder();
                connSB.Provider = "Microsoft.ACE.OLEDB.12.0";
                connSB.DataSource = @"C:\Users\MoChen\source\repos\magentr\magentr\magentr.accdb";
                connString = connSB.ToString();
                printDebug(st, "Target Dir is not Empty, judging if this file is already synced.");
                if (await CheckFileExist(RequestFileInfo.Name))
                {
                    printDebug(st, "File Already Synced");
                    return;
                }
                await FetchNewRequest(dirNewRequest);
            }
            else
            {
                printDebug(st, "No file selected.");
            }
            printDebug(st, string.Format("Button Click Ran for: {0}", 
                (DateTime.Now - timeStart).ToString("hh':'mm':'ss")));
        }

        private async Task<bool> CheckFileExist(string FileName)
        {
            string st = "CheckFileExist(string FileName){ }";
            bool isExist = false;
            await Task.Run(() =>
            {
                using (OleDbConnection conn = new OleDbConnection(connString))
                {
                    conn.Open();
                    OleDbCommand SelectSQL = new OleDbCommand(
                        @"SELECT tbRequestForm.* FROM tbRequestForm WHERE tbRequestForm.RequestFileName = @param1 ", conn);
                    SelectSQL.Parameters.AddWithValue("@param1", FileName);
                    OleDbDataReader reader = SelectSQL.ExecuteReader();
                    printDebug(st, "Execute Reader Content");
                    while (reader.Read())
                    {
                        Debug.Print("Record find: ID={0}.", reader[0].ToString());
                    }
                    isExist = reader.HasRows;
                    printDebug(st, string.Format("Does the row exist? {0}", isExist));
                    printDebug(st, "Closing Reader Object"); reader.Close();
                    
                }
            });
            return isExist;
        }


        private async Task FetchNewRequest(string dirNew)
        {
            string st = "FetchNewRequest(string dirNew){ }";
            DateTime timeStart = DateTime.Now;
            var UpdateProgressBar = new Progress<int>(value => {
                if (value >= 0)
                {
                    pbarMain.IsIndeterminate = false;
                    pbarMain.Value = value;
                }
                else
                {
                    pbarMain.IsIndeterminate = true;
                }
                

            });
            var SetProgressBarMax = new Progress<int>(value => pbarMain.Maximum = value);
            var PrintDebugListBox = new Progress<string>(value => 
            {
                printDebug(st, value);
                svDebug.ScrollToBottom();
            });
            //This procedure fetches information from Excel
            FileInfo InputFile = new FileInfo(dirNew);
            printDebug(st, "Start Task: SyncVonExcel(...){ }");
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
            reportProgressBar.Report(-1);
            EXCEL.Application xlApp = new EXCEL.Application();              
            EXCEL.Workbooks xlWorkbooks = xlApp.Workbooks;                  
            EXCEL.Workbook xlWbk = xlWorkbooks.Open(inputfile.FullName);    
            EXCEL.Worksheet xlSht = xlWbk.ActiveSheet;                      
            
            printDebugListBox.Report("Loading Completed.");
            #region ---- Fetch M/Agent Information ----
            printDebugListBox.Report("Beginning Fetching M/Agent Information.");
            printDebugListBox.Report("Setting Range Dictionary Delegate.");

            void RangeToDict(EXCEL.Range TargetRange)
            {
                dictRequestRawData[TargetRange.Address]
                = Convert.ToString(TargetRange.Value);
            }

            //Cell Range: D5, S163
            printDebugListBox.Report("Assigning Worksheet Object to Target Range = D5:S163");
            //EXCEL.Range FormArea = xlSht.Range["D5", "S163"]; //This is too many, Get only non null ones.
            printDebugListBox.Report("Making IEnumerable for Filled Ranges");
            var ieFilledRange = 
                (from EXCEL.Range r in xlSht.Range["D5", "S163"]
                 where r.Value != null
                select r).ToList();
            printDebugListBox.Report("Making IEnumerable for All Checked Boxes");
            printDebugListBox.Report("Assigning Worksheet Shapes to Target shapes.");
            //EXCEL.Shapes xlShapes = xlSht.Shapes;
            var xlCheckBoxes = (
                from EXCEL.Shape s in xlSht.Shapes
                where (s.Name.Contains("チェック") || s.Name.Contains("Check Box")) //v0.0.0.1 Need a regex to match both en and jp version.
                && (double)s.OLEFormat.Object.Value == 1 //Select only selected Value
                select s).ToList();//.ToList();

            printDebugListBox.Report("Calculating Total Form Area Ranges");
            int WorkLoad_Total = ieFilledRange.Count() + xlCheckBoxes.Count();
            int WorkdLoad_Current = 0;
            setProgressBarMax.Report(WorkLoad_Total);
            reportProgressBar.Report(WorkdLoad_Current);
            printDebugListBox.Report("Assigning Range Objects to Local Dictionary Object");
            foreach (EXCEL.Range r in ieFilledRange)
            {
                RangeToDict(r);
                reportProgressBar.Report(++WorkdLoad_Current);
            }
            printDebugListBox.Report("Assigning Dictionary Object with Checkbox.");
            foreach (EXCEL.Shape s in xlCheckBoxes)
            {
                dictCheckBox.Add(
                    s.TopLeftCell.Address
                    , (string)s.TopLeftCell
                    .Offset[0, 1].Value);
                reportProgressBar.Report(++WorkdLoad_Current);
            }
            printDebugListBox.Report("Assigning Dictionary Object with Checkbox Completed. Close Workbook Application.");

            #endregion ---- Fetch M/Agent Information ----
            xlWbk.Close(false, Missing.Value, Missing.Value); //Arguments in this will cause excel to exist without saving.
            xlWorkbooks.Close();
            xlApp.Quit();
            Marshal.ReleaseComObject(xlApp);
            Marshal.ReleaseComObject(xlWorkbooks);
            Marshal.ReleaseComObject(xlWbk);
            Marshal.ReleaseComObject(xlSht);
            printDebugListBox.Report("Closed Workbook Application.");
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
            #endregion --------Test two Dictionary Objects---------
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
            printDebugListBox.Report(SyncServer(connString, "H", "J", 49));
            printDebugListBox.Report(SyncServer(connString, "H", "J", 51));
            printDebugListBox.Report(SyncServer(connString, "H", "J", 64));

            printDebugListBox.Report(SyncColumn(connString, "L", "N"));
            printDebugListBox.Report(SyncServer(connString, "L", "N", 49));
            printDebugListBox.Report(SyncServer(connString, "L", "N", 51));
            printDebugListBox.Report(SyncServer(connString, "L", "N", 64));

            printDebugListBox.Report(SyncColumn(connString, "P", "R"));
            printDebugListBox.Report(SyncServer(connString, "P", "R", 49));
            printDebugListBox.Report(SyncServer(connString, "P", "R", 51));
            printDebugListBox.Report(SyncServer(connString, "P", "R", 64));



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
            //Debug.Print(AllRows);
            string rxRange = string.Format(@"\$[{0}-{1}]\$({2})"
                , firstCol
                , secondCol
                , AllRows);

            //Debug.Print(rxRange);
            Regex rxRangeMatch = new Regex(rxRange);
            string result = "未選択";
            try
            {
                //Debug.Print(string.Format("Testing against {0}", rxRangeMatch.ToString()));
                var EnumResult = from KeyValuePair<string, string> Checked in dictCheckBox
                                 where rxRangeMatch.IsMatch(Checked.Key)
                                 select Checked;
                //Debug.Print(string.Format("Checked Box Count: {0}", EnumResult.Count()));
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

        private string SyncServer
            ( string ConnectionString
            , string ColumnStart
            , string ColumnFinish
            , int StartRow)
        {
            //Before Sync, Judge if any of the must fill values are invalid, 
            //if yes, direcly return Error Message without connecting to Database.
            //Rule No1: Row[0] must be non-empty.
            //We first judge these values, in the sync we can directly use these string variables
            // VIP = 49 + 2
            // PRI = 51 + 13
            // SEC = 64 + 13
            string HostName = ValidDic("$" + ColumnStart + "$" + StartRow);
            if (HostName.Length < 8) //Exit Directly if no invalid Hostname
                return string.Format("Invalid Hostname \"{0}\" at ${1}${2} \r\n" +
                    "Sync Terminated.", HostName, ColumnStart, StartRow);
            string BoxIndex = "";
            switch (StartRow)
            {
                case 49:
                    BoxIndex = "0";
                    break;
                case 51:
                    BoxIndex = "1";
                    break;
                case 64:
                    BoxIndex = "2";
                    break;

            }
            string VIPHost = ValidDic("$" + ColumnStart + "$" + 49);
            using (OleDbConnection conn = new OleDbConnection(ConnectionString))
            {
                conn.Open();
                var InsertRequest = new OleDbCommand(
@"INSERT INTO tbServers (
 Hostname,  IPAddress,  Maker,  Model,  CPUCount,  CPUMicoprocessor,  OS,  Version,  BitVal,  ClusterBox,  ClusterIndex
) VALUES (
@hostname, @iPAddress, @maker, @model, @cPUCount, @cPUMicoprocessor, @oS, @version, @bitVal, @clusterBox, @clusterIndex
);", conn);
                InsertRequest.Parameters.AddWithValue("@hostname", HostName);
                InsertRequest.Parameters.AddWithValue("@iPAddress", ValidDic("$" + ColumnStart + "$" + ++StartRow));
                if (StartRow == 50) //Sync to a VIP Entry Since start row has already been ++ (^), it is 49+1=50
                {
                    InsertRequest.Parameters.AddWithValue("@maker", "VIP");
                    InsertRequest.Parameters.AddWithValue("@model", "VIP");
                    InsertRequest.Parameters.AddWithValue("@cPUCount", "VIP");
                    InsertRequest.Parameters.AddWithValue("@cPUMicoprocessor", "VIP");
                    InsertRequest.Parameters.AddWithValue("@oS", "VIP");
                    InsertRequest.Parameters.AddWithValue("@version", "VIP");
                    InsertRequest.Parameters.AddWithValue("@bitVal", "VIP");
                }
                else
                {
                    
                    InsertRequest.Parameters.AddWithValue("@maker", ValidDic("$" + ColumnStart + "$" + ++StartRow));
                    InsertRequest.Parameters.AddWithValue("@model", ValidDic("$" + ColumnStart + "$" + ++StartRow));
                    InsertRequest.Parameters.AddWithValue("@cPUCount", ValidDic("$" + ColumnStart + "$" + ++StartRow));
                    InsertRequest.Parameters.AddWithValue("@cPUMicoprocessor", ValidDic("$" + ColumnStart + "$" + ++StartRow));
                    InsertRequest.Parameters.AddWithValue("@oS", CheckBoxValue(ColumnStart + StartRow + ":" + ColumnFinish + (StartRow = 3 + StartRow)));
                    InsertRequest.Parameters.AddWithValue("@version", ValidDic("$" + ColumnStart + "$" + ++StartRow));
                    InsertRequest.Parameters.AddWithValue("@bitVal", ValidDic("$" + ColumnStart + "$" + ++StartRow));

                }

                InsertRequest.Parameters.AddWithValue("@clusterBox", VIPHost);
                InsertRequest.Parameters.AddWithValue("@clusterIndex", BoxIndex);

                for (int i = 0; i < InsertRequest.Parameters.Count; i++)
                {
                    Debug.Print("{0, -10} : {1}", InsertRequest.Parameters[i].ToString(), InsertRequest.Parameters[i].Value);
                }
                Debug.Print(InsertRequest.CommandText.ToString());
                try
                {
                    int RowsAffected = InsertRequest.ExecuteNonQuery();
                    return "Server Successful, Rows Affected: " + RowsAffected;
                }
                catch (OleDbException ex)
                {
                    return "Server Sync Failed:" + ex.Message;
                }
            }
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
            //Rule No3: $98 must be non-empty.
            //We first judge these values, in the sync we can directly use these string variables
            string ThisName = "SyncColumn";
            Dictionary<string, string> tableDCMS = new Dictionary<string, string>();
            tableDCMS.Add("jcs01800", "uny30110"); tableDCMS.Add("jcs01700", "uny40110");
            tableDCMS.Add("jcs01600", "uny40310"); tableDCMS.Add("jcs01200", "uny40510");
            tableDCMS.Add("jcs01100", "uny40710"); tableDCMS.Add("jcs01500", "uny40910");
            tableDCMS.Add("jcs01300", "uny41110"); tableDCMS.Add("jcs01400", "uny41310");

            string RegisterType = CheckBoxValue(ColumnStart + "32:" + ColumnFinish + "33");
            string MS_VIP = ValidDic("$" + ColumnStart + "$98");

            string AG_VIP = ValidDic("$" + ColumnStart + "$49");
            string AG_PRI = ValidDic("$" + ColumnStart + "$51");
            string AG_SEC = ValidDic("$" + ColumnStart + "$64");

            if (RegisterType.Contains("選択") || AG_PRI.Length < 8)
                return "[Warning...] Either Apply Type not Selected, or no primary host specified. Aborting Sync.";
            if (MS_VIP.Length < 8)
                return "[Warning...] No Datacenter assigned for this host. Abort Sync.";
            string ConnectedDC = tableDCMS[MS_VIP];

            string ConnectedAG = AG_VIP.Length < 8 ? AG_PRI : AG_VIP;

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
,  MAtMSCommunicationPort,  MSVIP,  MSPRI,  MSSEC,  AgentName
) VALUES (
@rlnfileName, @rlnbango, @applyType, @changePoint, @sIer, @serverPIC, @systemID, @systemName
, @systemSubName, @networkLocation, @networkArea, @serverVIP, @serverPRI, @serverSEC
, @mStMACommunicationPort, @mA_InstallDate, @mS_Connection, @jobStartDate, @jobCount
, @hasCallorder, @hasFirewall, @mA_Version, @isFirstTime, @isProduction, @testDoneDate
, @costFrom, @costFromSystemName, @costFromSubSystemName, @hasSundayJobs, @hasRelatedSystems
, @relatedSystemID, @relatedSystemName, @relatedSystemSubName, @relatedSystemDatacenter
, @mAtMSCommunicationPort, @mSVIP, @mSPRI, @mSSEC, @agentName
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
                InsertRequest.Parameters.AddWithValue("@serverVIP", AG_VIP);
                InsertRequest.Parameters.AddWithValue("@serverPRI", AG_PRI);
                InsertRequest.Parameters.AddWithValue("@serverSEC", AG_SEC);
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
                InsertRequest.Parameters.AddWithValue("@agentName", ConnectedDC + "." + ConnectedAG);

                try
                {
                    int RowsAffected = InsertRequest.ExecuteNonQuery();
                    return "Agent Table Successful, Rows Affected: " + RowsAffected;
                }
                catch (OleDbException ex)
                {
                    return "Agent Table Sync Failed: " + ex.Message;
                }
                
            }

        }

    }

}
