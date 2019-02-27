using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Net;
using System.Text;
using System.Threading.Tasks;
using EXCEL = Microsoft.Office.Interop.Excel;

namespace magentr
{
    class MAObjects
    {
    }

    class RequestSheet
    {
        public RequestColumns Col1, Col2, Col3;
        //Instantiate an empty sheet object, which contains 3 column objects.
        public RequestSheet() { }
        public RequestSheet(
            RequestColumns col1
            , RequestColumns col2
            , RequestColumns col3)
        {
            Col1 = col1;
            Col2 = col2;
            Col3 = col3;
        }

    }

    class RequestColumns
    {
        //Instantiate an empty Column object, which contains column objects.
         
    }

    class MAgents
    {
        //Instantiate an empty Agent object, which contains M/Agent Information.

    }

    class MAServers
    {
        public string Hostname;
        public bool IsCluster = false;
        public string NetLoca;
        /*public enum NetworkLocation
        {
            JDC,GDC,DRC,Others
        }*/
        public string NetArea;
        /*public enum NetworkArea
        {
            SEN,SecureSEN,SecureGlobal,Global
            ,SecureLocal,DMZ,iDC,SAZ
        }*/
        public agCluster Cluster;
        public enum agCluster
        {
            VIP,PRI,SEC
        }
        public MAServers VIP = null;
        public MAServers PRI = null;
        public MAServers SEC = null;
        public IPAddress IP_Address;
        public string Maker;
        public string Model;
        public int CPU_Number;
        public int CPU_Microprocessor;
        public string OS;
        public string OSVersion;
        public int OSBit;
        public int BoxSplit;
        public int BoxThis;
        //Instantiate an empty Server object, which server information.
        
        public MAServers() { }
        public MAServers(string hostname)
        {
            this.Hostname = hostname;
        }
        public MAServers(string hostname, IPAddress ip_Address)
        {
            this.Hostname = hostname;
            IP_Address = ip_Address;
        }
        public MAServers(
            string[] Range
            , string[] OSCheck
            , string[] NetLocaCheck
            , string[] NetAreaCheck
            , Dictionary<string, string> dictSrcSheet
            , Dictionary<string, string> dictCheckBox
            , agCluster clusterIndex)
        {
            //Fetching a Server info based on Server Area Directly.
            //Examples for starting cell $H$51:$H63, or $P$64:$P$76

            Debug.Print(Range[0]);
            this.Hostname = dictSrcSheet[Range[0]];
            this.IP_Address = IPAddress
                              .Parse(dictSrcSheet[Range[1]]);
            this.Maker = dictSrcSheet[Range[2]];
            this.Model = dictSrcSheet[Range[3]];
            this.CPU_Number = Convert.ToInt32(dictSrcSheet[Range[4]]);
            this.CPU_Microprocessor = Convert.ToInt32(dictSrcSheet[Range[5]]);
            this.OS = (string)
                (from KeyValuePair<string, string> k
                 in dictCheckBox
                 from string s in OSCheck
                where s == k.Key
                select k.Value).First();

            this.NetLoca = (string)
                (from KeyValuePair<string, string> k
                 in dictCheckBox
                 from string s in NetLocaCheck
                 where s == k.Key
                 select k.Value).First();

            this.NetArea = (string)
                (from KeyValuePair<string, string> k
                 in dictCheckBox
                 from string s in NetAreaCheck
                 where s == k.Key
                 select k.Value).First();

            this.OSVersion = dictSrcSheet[Range[9]];
            this.OSBit = Convert.ToInt32(dictSrcSheet[Range[10]]);
            this.BoxSplit = Convert.ToInt32(dictSrcSheet[Range[11]]);
            this.BoxThis = Convert.ToInt32(dictSrcSheet[Range[12]]);
            this.Cluster = clusterIndex;
        }

        public MAServers(
            MAServers vip
            , MAServers pri
            , MAServers sec)
        {
            this.VIP = vip;
            this.PRI = pri;
            this.SEC = sec;
        }

        public void AddClusterMember(MAServers pri)
        {
            this.PRI = pri;
        }

        public void AddClusterMember(
            MAServers vip
            , MAServers pri
            , MAServers sec)
        {
            this.VIP = vip;
            this.PRI = pri;
            this.SEC = sec;

        }

        public string FullInfo()
        {
            StringBuilder sbInfo = new StringBuilder();
            sbInfo.AppendLine();
            sbInfo.AppendLine(new string('_', 57));
            sbInfo.AppendLine("M/Agent Info");
            sbInfo.AppendLine(new string('=', 57));
            sbInfo.AppendLine(this.NetLoca);
            sbInfo.AppendLine(this.NetArea);
            sbInfo.AppendLine(this.Hostname);
            sbInfo.AppendLine(this.IP_Address.ToString());
            sbInfo.AppendLine(this.Maker);
            sbInfo.AppendLine(this.Model);
            sbInfo.AppendLine(this.CPU_Number.ToString());
            sbInfo.AppendLine(this.CPU_Microprocessor.ToString());
            sbInfo.AppendLine(this.OS);
            sbInfo.AppendLine(this.OSVersion);
            sbInfo.AppendLine(this.OSBit.ToString());
            sbInfo.AppendLine(this.BoxSplit.ToString());
            sbInfo.AppendLine(this.BoxThis.ToString());
            sbInfo.AppendLine(this.Cluster.ToString());


            sbInfo.AppendLine(new string('-', 57));

            return sbInfo.ToString();
        }

        public string FullClusterInfo()
        {
            StringBuilder sbInfo = new StringBuilder();

            sbInfo.AppendLine(this.VIP.FullInfo());
            sbInfo.AppendLine(this.PRI.FullInfo());
            sbInfo.AppendLine(this.SEC.FullInfo());


            return sbInfo.ToString();
        }

    }
}
