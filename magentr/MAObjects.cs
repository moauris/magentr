using System;
using System.Collections.Generic;
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

    class Servers
    {
        public string Hostname;
        public bool IsCluster = false;
        public enum NetworkLocation
        {
            JDC,GDC,DRC,Others
        }
        public enum NetworkArea
        {
            SEN,SecureSEN,SecureGlobal,Global
            ,SecureLocal,DMZ,iDC,SAZ
        }
        public enum Cluster
        {
            VIP,PRI,SEC
        }
        public Servers VIP;
        public Servers PRI;
        public Servers SEC;
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

        public Servers()
        {
            //return a string that represents the object.
        }
        public Servers(string hostname)
        {
            this.Hostname = hostname;
        }
        public Servers(string hostname, IPAddress ip_Address)
        {
            this.Hostname = hostname;
            IP_Address = ip_Address;
        }
        public Servers(string starting_cell, EXCEL.Worksheet xlsheet)
        {
            //Fetching a Server info based on Server Area Directly.
            //Examples for starting cell $H$51
            EXCEL.Range StartingCell = xlsheet.Range["$H$51"];

            
            this.Hostname = StartingCell.Value as string;
            this.IP_Address = IPAddress
                .Parse(StartingCell.Offset[1, 1].Value);
            this.Maker = StartingCell.Offset[2, 0].Value as string;
            this.Model = StartingCell.Offset[3, 0].Value as string;
            this.CPU_Number = (int)StartingCell.Offset[4, 0].Value;
            this.CPU_Microprocessor = (int)StartingCell.Offset[5, 0].Value;


            
        }

    }
}
