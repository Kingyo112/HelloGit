using MahApps.Metro.Controls;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Shapes;
using System.IO;

namespace AttendanceManagement
{
    /// <summary>
    /// Interaction logic for StartWindow.xaml
    /// </summary>
    public partial class StartWindow : MetroWindow
    {
        public StartWindow()
        {
            InitializeComponent();
        }

        private void tilAttendance_Click(object sender, RoutedEventArgs e)
        {
            if (!File.Exists(this.txtExcelPath_Attendance.Text))
            {
                File.Copy(this.txtExcelPath_AttendanceSample.Text, this.txtExcelPath_Attendance.Text);
            }

            if (!string.IsNullOrEmpty(this.txtExcelPath_Getudo.Text)
            && !string.IsNullOrEmpty(this.txtExcelPath_Holiday.Text)
            && !string.IsNullOrEmpty(this.txtExcelPath_Attendance.Text))
            {
                MainWindow main = new MainWindow(this.txtExcelPath_Getudo.Text, this.txtExcelPath_Holiday.Text,this.txtExcelPath_Attendance.Text);
                main.Owner = this;
                main.ShowDialog();
            }
        }

        private void tilMasterM_Click(object sender, RoutedEventArgs e)
        {

        }

        private void tilAnalyze_Click(object sender, RoutedEventArgs e)
        {

        }

        private void tilSetting_Click(object sender, RoutedEventArgs e)
        {
            this.flyout.IsOpen = true;
        }

        private void MetroWindow_Loaded(object sender, RoutedEventArgs e)
        {
            this.txtExcelPath_Getudo.Text = Properties.Settings.Default.ExcelPath_Getudo;
            this.txtExcelPath_Holiday.Text = Properties.Settings.Default.ExcelPath_Holiday;
            this.txtExcelPath_Attendance.Text = Properties.Settings.Default.ExcelPath_Attendance;
            this.txtExcelPath_AttendanceSample.Text = Properties.Settings.Default.ExcelPath_AttendanceSample;

            this.tilAttendance.Focus();
        }

        private void MetroWindow_Closed(object sender, EventArgs e)
        {
            Properties.Settings.Default.ExcelPath_Attendance = this.txtExcelPath_Attendance.Text;
            Properties.Settings.Default.Save();
        }
    }
}
