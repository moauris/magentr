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
        public MainWindow()
        {
            InitializeComponent();
        }

        private async void OnNewRequestClick(object sender, RoutedEventArgs e)
        {
            var UpdateProgressBar = new Progress<int>(value =>
            {
                pbarMain.IsIndeterminate = true;
            });
            TaskNewRequest NewTask = new TaskNewRequest();
            await NewTask.Start(UpdateProgressBar);
            pbarMain.IsIndeterminate = false;
        }


    }

}
