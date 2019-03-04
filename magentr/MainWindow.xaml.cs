﻿using Microsoft.Win32;
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

        private void OnNewRequestClick(object sender, RoutedEventArgs e)
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
            FileInfo RequestFileInfo = new FileInfo(OpenFileNew.FileName);
            RequestBango = RequestFileInfo.Name;

            #endregion Open File Dialog
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
            , IProgress<string> printDebugListBox)
            //, out RequestSheet newRequest)
        {
            //newRequest = new RequestSheet();
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
            reportProgressBar.Report(100);
            Debug.Print("Loading Completed.");

            #region ---- Fetch M/Agent Information ----
            printDebugListBox.Report("Beginning Fetching M/Agent Information.");
            printDebugListBox.Report("Getting Range Dictionary Delegate.");

            void RangeToDict(EXCEL.Range TargetRange)
            {
                dictRequestRawData[TargetRange.Address]
                = Convert.ToString(TargetRange.Value);
            }

            //Cell Range: D5, S163
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

            foreach (EXCEL.Range r in ieFilledRange)
            {
                RangeToDict(r);
                reportProgressBar.Report(++WorkdLoad_Current);
            }
            printDebugListBox.Report("Sync Target Area Complete.");
            //Trying to fetch form public Dictionary Object.
            EXCEL.Shapes xlShapes = xlSht.Shapes;

            WorkLoad_Total = xlShapes.Count;
            setProgressBarMax.Report(WorkLoad_Total);
            WorkdLoad_Current = 0;

            IEnumerable<EXCEL.Shape> xlCheckBoxes =
                from EXCEL.Shape s in xlShapes
                where s.Name.Contains("チェック") //v0.0.0.1 Need a regex to match both en and jp version.
                && (double)s.OLEFormat.Object.Value == 1 //Select only selected Value
                select s;

            WorkLoad_Total = xlCheckBoxes.Count();
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
            #region Connect to Database with Connection String
            OleDbConnection conn = new OleDbConnection();
            conn.ConnectionString = @"Provider=Microsoft.ACE.OLEDB.12.0" +
                @";Data Source=C:\Users\MoChen\source\repos\magentr\magentr\magentr.accdb";
            conn.Open();
            #region Sync to tbRequestForm
            OleDbCommand SelectRequest = new OleDbCommand(
                "SELECT tbRequestForm.* FROM tbRequestForm WHERE tbRequestForm.;");

            OleDbCommand InsertRequest = new OleDbCommand(
                @"INSERT INTO tbRequestForm 
                   (RequestBango, RequestFileName
                    , DateApplied, Applier
                    , Email, Phone, Approver, Comment) 
                  Values (@requestBango, @requestFileName
                    , @dateApplied, @applier
                    , @email ,@phone, @approver
                    , @comment)", conn);
            //Make a procedure to sync values and return "" when keys cannot be found.
            string ValidateDict(string KeyVal)
            {
                string result = dictRequestRawData.ContainsKey(KeyVal) ? dictRequestRawData[KeyVal] : "";
                return result;
            }
            InsertRequest.CommandType = System.Data.CommandType.Text;
            InsertRequest.Parameters.AddWithValue("@requestBango", RequestBango.Substring(0,15));
            InsertRequest.Parameters.AddWithValue("@requestFileName", RequestBango);
            InsertRequest.Parameters.AddWithValue("@dateApplied"
                , DateTime.Parse(ValidateDict("$H$7")));
            InsertRequest.Parameters.AddWithValue("@applier", ValidateDict("$H$8"));
            InsertRequest.Parameters.AddWithValue("@email", ValidateDict("$H$9"));
            InsertRequest.Parameters.AddWithValue("@phone", ValidateDict("$H$10"));
            InsertRequest.Parameters.AddWithValue("@approver", ValidateDict("$H$11"));
            InsertRequest.Parameters.AddWithValue("@comment", ValidateDict("$E$161"));
            printDebugListBox.Report(InsertRequest.CommandText);
            InsertRequest.ExecuteNonQuery();
            #endregion Sync to tbRequestForm
            InsertRequest = new OleDbCommand(
                @"INSERT INTO tbAgents 
                    (RequestFileName, RequestBango, ApplyType
                    , ChangePoint, SIer, ServerPIC, SystemID
                    , SystemName, SystemSubName, NetworkLocation
                    , NetworkArea, ServerVIP, ServerPRI, ServerSEC
                    , MStMACommunicationPort, MA_InstallDate
                    , MS_Connection, MS_Connection, JobCount
                    , HasCallorder, HasFirewall, MA_Version
                    , IsFirstTime, IsProduction, TestDoneDate
                    , CostFrom, CostFromSystemName
                    , CostFromSubSystemName, HasSundayJobs
                    , HasRelatedSystems, RelatedSystemID
                    , RelatedSystemName, RelatedSystemSubName
                    , RelatedSystemDatacenter 
                    , MAtMSCommunicationPort
                    , MSVIP, MSPRI, MSSEC
                ) VALUES (
                    @requestFileName, @requestBango, @applyType
                    , @changePoint, @sIer, @serverPIC, @systemID
                    , @systemName, @systemSubName, @networkLocation
                    , @networkArea, @serverVIP, @serverPRI, @serverSEC
                    , @mStMACommunicationPort, @mA_InstallDate
                    , @mS_Connection, @mS_Connection, @jobCount
                    , @hasCallorder, @hasFirewall, @mA_Version
                    , @isFirstTime, @isProduction, @testDoneDate
                    , @costFrom, @costFromSystemName
                    , @costFromSubSystemName, @hasSundayJobs
                    , @hasRelatedSystems, @relatedSystemID
                    , @relatedSystemName, @relatedSystemSubName
                    , @relatedSystemDatacenter 
                    , @mAtMSCommunicationPort
                    , @mSVIP, @mSPRI, @mSSEC);", conn);
            Dictionary<string, string> CheckBoxGroup = new Dictionary<string, string>();
            InsertRequest.Parameters.AddWithValue("@requestFileName", RequestBango);
            InsertRequest.Parameters.AddWithValue("@requestBango", RequestBango.Substring(0, 15));

            CheckBoxGroup.Add("$H$32", "");
            CheckBoxGroup.Add("$J$32", "");
            CheckBoxGroup.Add("$H$33", "");
            string CbxVal =
                (string)(from KeyValuePair<string, string> Checked in dictCheckBox
                         from KeyValuePair<string, string> Target in CheckBoxGroup
                         where Checked.Key == Target.Key
                         select Checked).First().Value;

            InsertRequest.Parameters.AddWithValue("@applyType", CbxVal);
            /*
            InsertRequest.Parameters.AddWithValue("@changePoint");
            InsertRequest.Parameters.AddWithValue("@sIer");
            InsertRequest.Parameters.AddWithValue("@serverPIC");
            InsertRequest.Parameters.AddWithValue("@systemID");
            InsertRequest.Parameters.AddWithValue("@systemName");
            InsertRequest.Parameters.AddWithValue("@systemSubName");
            InsertRequest.Parameters.AddWithValue("@networkLocation");
            InsertRequest.Parameters.AddWithValue("@networkArea");
            InsertRequest.Parameters.AddWithValue("@serverVIP");    
            InsertRequest.Parameters.AddWithValue("@serverPRI");
            InsertRequest.Parameters.AddWithValue("@serverSEC");
            InsertRequest.Parameters.AddWithValue("@mStMACommunicationPort");
            InsertRequest.Parameters.AddWithValue("@mA_InstallDate");
            InsertRequest.Parameters.AddWithValue("@mS_Connection");
            InsertRequest.Parameters.AddWithValue("@mS_Connection");
            InsertRequest.Parameters.AddWithValue("@jobCount");
            InsertRequest.Parameters.AddWithValue("@hasCallorder");
            InsertRequest.Parameters.AddWithValue("@hasFirewall");
            InsertRequest.Parameters.AddWithValue("@mA_Version");
            InsertRequest.Parameters.AddWithValue("@isFirstTime");
            InsertRequest.Parameters.AddWithValue("@isProduction");
            InsertRequest.Parameters.AddWithValue("@testDoneDate");
            InsertRequest.Parameters.AddWithValue("@costFrom");
            InsertRequest.Parameters.AddWithValue("@costFromSystemName");
            InsertRequest.Parameters.AddWithValue("@costFromSubSystemName");
            InsertRequest.Parameters.AddWithValue("@hasSundayJobs");
            InsertRequest.Parameters.AddWithValue("@hasRelatedSystems");
            InsertRequest.Parameters.AddWithValue("@relatedSystemID");
            InsertRequest.Parameters.AddWithValue("@relatedSystemName");
            InsertRequest.Parameters.AddWithValue(" @relatedSystemSubName");
            InsertRequest.Parameters.AddWithValue("@relatedSystemDatacenter");
            InsertRequest.Parameters.AddWithValue("@mAtMSCommunicationPort");
            InsertRequest.Parameters.AddWithValue("@mSVIP");
            InsertRequest.Parameters.AddWithValue("@mSPRI");
            InsertRequest.Parameters.AddWithValue("@mSSEC");
            */
            //AdaRequest.InsertCommand = InsertCommand;
            //AdaRequest.InsertCommand.ExecuteNonQuery();

            printDebugListBox.Report(InsertRequest.CommandText);
            InsertRequest.ExecuteNonQuery();
            conn.Close();
            return;
            #endregion  Connect to Database with Connection String
            string[] col1_NetLoc = new string[4]
            {
                "$H$42","$J$42",
                "$H$43","$J$43"
            };
            string[] col1_NetAre = new string[8]
            {
                "$H$44","$J$44",
                "$H$45","$J$45",
                "$H$46","$J$46",
                "$H$47","$J$47"
            };
            string[] range1 = new string[13]
            {
                "$H$51","$H$52","$H$53","$H$54","$H$55",
                "$H$56","$H$57","$H$58","$H$59","$H$60",
                "$H$61","$H$62","$H$63"
            };
            string[] oscheck1 = new string[5]
            {
                "$H$57","$J$57",
                "$H$58","$J$58",
                "$H$59"
            };
            MAServers c1Cluster = new MAServers();
            MAServers c1VIP = 
                new MAServers(
                        dictRequestRawData["$H$49"]
                        );
            MAServers c1PRI = new MAServers
                (
                    range1
                    , oscheck1
                    , col1_NetLoc
                    , col1_NetAre
                    , dictRequestRawData
                    , dictCheckBox
                    , MAServers.agCluster.PRI
                );
            string[] range2 = new string[13]
            {
                "$H$64","$H$65","$H$66","$H$67","$H$68",
                "$H$69","$H$70","$H$71","$H$72","$H$73",
                "$H$74","$H$75","$H$76"
            };
            string[] oscheck2 = new string[5]
            {
                "$H$70","$J$70",
                "$H$71","$J$71",
                "$H$72"
            };
            MAServers c1SEC = new MAServers
                (
                    range2
                    , oscheck2
                    , col1_NetLoc
                    , col1_NetAre
                    , dictRequestRawData
                    , dictCheckBox
                    , MAServers.agCluster.SEC
                );
            c1Cluster = new MAServers(c1VIP, c1PRI, c1SEC);
            printDebugListBox.Report(c1Cluster.FullClusterInfo());
            RequestColumns colH = new RequestColumns();
            RequestColumns colL = new RequestColumns();
            RequestColumns colP = new RequestColumns();

            //newRequest = new RequestSheet(colH, colL, colP);
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
