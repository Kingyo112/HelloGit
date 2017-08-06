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
using System.Windows.Navigation;
using System.Windows.Shapes;
using AttendanceManagement.Data;
using MahApps.Metro.Controls.Dialogs;
using AttendanceMamagement.Logic;
using AttendanceManagement.Logic;

namespace AttendanceManagement
{
    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>
    public partial class MainWindow : MetroWindow
    {
        private DateTime KijunDate = DateTime.Now;

        private GetudoNumber GetudoNum = new GetudoNumber();

        private MonthlyAttendanceData MonthlyAttemdance = new MonthlyAttendanceData();

        private List<DispDailyAttendanceData> DispAtttendanceList = new List<DispDailyAttendanceData>();

        private List<List<DailyAttendanceData>> AttendanceDataBox = new List<List<DailyAttendanceData>>();

        private List<SummaryAttendanceData> SummaryAttendance = new List<SummaryAttendanceData>();

        private string ExcelPath_Getudo;
        private string ExcelPath_Holiday;
        private string ExcelPath_Attendance;

        public MainWindow(string getudo, string holiday, string attendance)
        {
            InitializeComponent();

            ExcelPath_Getudo = getudo;
            ExcelPath_Holiday = holiday;
            ExcelPath_Attendance = attendance;
        }

        private void InitDisplayData()
        {
            var getudo = this.GetGetudoNumber();
            this.DisplayGetudoData(getudo);
            this.DisplayAttendanceData(this.GetAttendancdData(getudo));
            this.DisplaySummaryAttendanceData(this.SummaryAttendance, getudo);
        }

        private void DisplaySummaryAttendanceData(List<SummaryAttendanceData> list, MasterGetudo getudo)
        {
            foreach (var item in list)
            {
                if (item.GetudoYYYYMM == getudo.GetudoYYYYMM)
                {
                    this.lblthisMonthProspectsOverTime.Content = this.DisplayMonthProspectsOverTime(item.OverTime);
                    this.lblthisMonthProspectsOverTime.Foreground = this.GetColor(item.Is45Over);
                    return;
                }
            }
            this.lblthisMonthProspectsOverTime.Content = this.DisplayMonthProspectsOverTime(new TimeSpan(0, 0, 0));
            this.lblthisMonthProspectsOverTime.Foreground = this.GetColor(false);
            return;
        }

        private Brush GetColor(bool is45over)
        {
            if (is45over)
            {
                return new SolidColorBrush(Colors.Red);

            }
            else
            {
                return new SolidColorBrush(Colors.Black);
            }
        }

        private void DisplayAttendanceData(List<DispDailyAttendanceData> dispatttendancelist)
        {
            this.DispAtttendanceList = dispatttendancelist;
            dgMontlyData.ItemsSource = DispAtttendanceList;
            //dgMontlyData
        }

        private List<DispDailyAttendanceData> GetAttendancdData(MasterGetudo getudorecord)
        {
            this.MonthlyAttemdance = new MonthlyAttendanceData();
            MonthlyAttemdance.ExcelPath_Holidays = this.txtExcelPath_Holiday.Text;
            MonthlyAttemdance.ExcelPath = this.txtExcelPath_Attendance.Text;
            var errorinfo = this.MonthlyAttemdance.GetInitAttendanceData(getudorecord, out AttendanceDataBox, out SummaryAttendance);
            if (errorinfo.HasError)
            {
                this.ShowMessageDialog(this, new RoutedEventArgs(), errorinfo.ErrorReason);
            }

            return this.ConvertAttendanceListtoDisp(this.GetAttendancdDataByGetudoYYYMM(AttendanceDataBox, getudorecord));
        }

        private List<DailyAttendanceData> GetAttendancdDataByGetudoYYYMM(List<List<DailyAttendanceData>> AttendanceDataBox, MasterGetudo getudorecord)
        {
            foreach (var item in AttendanceDataBox)
            {
                if (item.Count(r => r.GetudoYYYYMM == getudorecord.GetudoYYYYMM) >= 1)
                {
                    AttendanceDataBox.Remove(item);
                    return item;
                }
            }
            var newList = new List<DailyAttendanceData>();
            this.MonthlyAttemdance.GetInitMonthlyData(getudorecord, out newList);
            return newList;
        }

        private List<DispDailyAttendanceData> ConvertAttendanceListtoDisp(List<DailyAttendanceData> list)
        {
            var list_disp = new List<DispDailyAttendanceData>();
            foreach (var item in list)
            {
                var record = new DispDailyAttendanceData();
                record.ConvertToDispData(item);
                list_disp.Add(record);
            }
            return list_disp;
        }

        private MasterGetudo GetGetudoNumber()
        {
            this.GetudoNum = new Logic.GetudoNumber();
            GetudoNum.ExcelPath = this.txtExcelPath_Getudo.Text;
            var master = new Data.MasterGetudo();
            var errorinfo = this.GetudoNum.GetInitValue(out master);
            if (errorinfo.HasError)
            {
                this.ShowMessageDialog(this, new RoutedEventArgs(), errorinfo.ErrorReason);
            }
            return master;
        }

        private void DisplayGetudoData(MasterGetudo masterGetudo)
        {
            this.txtGetudoNumber.Value = masterGetudo.GetudoYYYYMM;
            this.lblGetudoFromTo.Content = masterGetudo.StartDate.ToShortDateString() + " ~ " + masterGetudo.EndDate.ToShortDateString();
        }

        private void Button_Click(object sender, RoutedEventArgs e)
        {
            this.flyout.IsOpen = true;
        }

        private void txtGetudoNumber_ValueChanged(object sender, RoutedPropertyChangedEventArgs<double?> e)
        {
            var master = new Data.MasterGetudo();
            int value = 0;
            var errorinfo = this.GetudoNum.GetConsistentValue((int)txtGetudoNumber.Value, out value, out master);
            if (errorinfo.HasError)
            {
                this.ShowMessageDialog(this, e, errorinfo.ErrorReason);
                return;
            }


            txtGetudoNumber.ValueChanged -= new RoutedPropertyChangedEventHandler<double?>(txtGetudoNumber_ValueChanged);
            this.DisplayGetudoData(master);
            txtGetudoNumber.ValueChanged += new RoutedPropertyChangedEventHandler<double?>(txtGetudoNumber_ValueChanged);


            this.AvoidanceAttendanceData(this.DispAtttendanceList);
            this.DisplayAttendanceData(this.ConvertAttendanceListtoDisp(this.GetAttendancdDataByGetudoYYYMM(this.AttendanceDataBox, master)));

            this.AvoidSummaryData(this.lblthisMonthProspectsOverTime.Content.ToString(), e.OldValue);
            this.DisplaySummaryAttendanceData(this.SummaryAttendance, master);

        }

        private void AvoidSummaryData(string overtime, double? getudoyyyymm)
        {
            if (getudoyyyymm == null) return;

            var o = new OverTime();
            var timespan_overtime = o.ConvertDispOverTime(overtime);
            int i = 0;
            if (int.TryParse(overtime.Replace(":", ""), out i))
            {
                i = i / 100;
            }

            foreach (var item in this.SummaryAttendance)
            {
                if (item.GetudoYYYYMM == (int)getudoyyyymm)
                {
                    item.OverTime = timespan_overtime;
                    item.Is45Over = i >= 45;
                    item.Is80Over = i >= 80;
                    return;
                }
            }
            var record = new SummaryAttendanceData();
            record.GetudoYYYYMM = (int)getudoyyyymm;
            record.OverTime = timespan_overtime;
            record.Is45Over = i >= 45;
            record.Is80Over = i >= 80;
            return;
        }

        private void AvoidanceAttendanceData(List<DispDailyAttendanceData> displist)
        {
            var list = new List<DailyAttendanceData>();
            foreach (var item in displist)
            {
                var record = new DailyAttendanceData();
                record.ConvertToData(item);
                list.Add(record);
            }
            this.AttendanceDataBox.Add(list);
        }

        private async void ShowMessageDialog(object sender, RoutedEventArgs e, string message)
        {
            this.ShowMessageDialog(sender, e, message, "エラー");
        }

        private async void ShowMessageDialog(object sender, RoutedEventArgs e, string message, string title)
        {
            await this.ShowMessageAsync(title, message);
        }

        private void btnSave_Click(object sender, RoutedEventArgs e)
        {
            this.AvoidanceAttendanceData(this.DispAtttendanceList);
            this.AvoidSummaryData(this.lblthisMonthProspectsOverTime.Content.ToString(), (int)txtGetudoNumber.Value);

            var errorinfo = this.MonthlyAttemdance.SaveData(this.AttendanceDataBox, this.SummaryAttendance, this.KijunDate);
            if (errorinfo.HasError)
            {
                this.ShowMessageDialog(sender, e, errorinfo.ErrorReason);
                return;
            }
            this.ShowMessageDialog(sender, e, "保存しました。\n保存後自動で画面を閉じます。", "登録確認");
            this.Close();
        }

        private void txtKijunDate_TextChanged(object sender, TextChangedEventArgs e)
        {
            if (!DateTime.TryParse(txtKijunDate.Text, out this.KijunDate))
            {
                KijunDate = DateTime.Now;
            }
            //計算

            TimeSpan overtime = this.MonthlyAttemdance.CalcOverTime(this.DispAtttendanceList, this.KijunDate);
            this.lblthisMonthProspectsOverTime.Content = this.DisplayMonthProspectsOverTime(overtime);

        }

        private object DisplayMonthProspectsOverTime(TimeSpan overtime)
        {
            int hourWithDay = 24 * overtime.Days + overtime.Hours;
            return hourWithDay.ToString() + overtime.ToString(@"\:mm");
        }

        private void dgMontlyData_CurrentCellChanged(object sender, EventArgs e)
        {
            this.RefrectionBackNumberCell(this.DispAtttendanceList);
            //計算
            TimeSpan overtime = this.MonthlyAttemdance.CalcOverTime(this.DispAtttendanceList, this.KijunDate);
            int hourWithDay = 24 * overtime.Days + overtime.Hours;
            this.lblthisMonthProspectsOverTime.Content = hourWithDay.ToString() + overtime.ToString(@"\:mm");
            this.lblthisMonthProspectsOverTime.Foreground = this.GetColor(hourWithDay >= 45);

        }

        public List<DispDailyAttendanceData> RefrectionBackNumberCell(List<DispDailyAttendanceData> list)
        {
            foreach (var item in list)
            {
                var sub = new DailyAttendanceData();
                sub.ConvertToData(item);
                item.PlanMMSS_Start = sub.PlanMMSS_Start;
                item.PlanMMSS_END = sub.PlanMMSS_END;
                item.ResultMMSS_Start = sub.ResultMMSS_Start;
                item.ResultMMSS_END = sub.ResultMMSS_END;
            }
            return list;
        }

        /// <summary>
        /// クリップボード貼り付け
        /// </summary>
        /// <param name="dataGrid"></param>
        private void pasteClipboard(DataGrid dataGrid)
        {
            try
            {
                // 張り付け開始位置設定
                var startRowIndex = dataGrid.ItemContainerGenerator.IndexFromContainer(
                    (DataGridRow)dataGrid.ItemContainerGenerator.ContainerFromItem
                    (dataGrid.CurrentCell.Item));
                var startColIndex = dataGrid.SelectedCells[0].Column.DisplayIndex;


                // クリップボード文字列から行を取得
                var pasteRows = ((string)Clipboard.GetData(DataFormats.Text)).Replace("\r", "")
                    .Split(new string[] { "\n" }, StringSplitOptions.RemoveEmptyEntries);

                var maxRowCount = pasteRows.Count();
                for (int rowCount = 0; rowCount < maxRowCount; rowCount++)
                {
                    var rowIndex = startRowIndex + rowCount;

                    // タブ区切りでセル値を取得
                    var pasteCells = pasteRows[rowCount].Split('\t');

                    // 選択位置から列数繰り返す
                    var maxColCount = Math.Min(pasteCells.Count(), dataGrid.Columns.Count - startColIndex);
                    for (int colCount = 0; colCount < maxColCount; colCount++)
                    {
                        var column = dataGrid.Columns[colCount + startColIndex];

                        // 貼り付け
                        column.OnPastingCellClipboardContent(dataGrid.Items[rowIndex], pasteCells[colCount]);
                        //if (!byteFiled.Contains(column.Header.ToString()))
                        //{
                        //    column.OnPastingCellClipboardContent(dataGrid.Items[rowIndex], pasteCells[colCount]);
                        //}
                    }
                }

                // 選択位置復元
                dataGrid.CurrentCell = new DataGridCellInfo(
                    dataGrid.Items[startRowIndex], dataGrid.Columns[startColIndex]);
            }
            catch
            {

            }

        }

        private void dgMontlyData_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.Key == Key.V)
            {
                if (Keyboard.Modifiers == ModifierKeys.Control)
                {
                    var dataGrid = sender as DataGrid;
                    if (dataGrid != null)
                    {
                        pasteClipboard(dataGrid);
                    }
                }
            }
        }

        private void MetroWindow_Loaded(object sender, RoutedEventArgs e)
        {
            this.txtExcelPath_Getudo.Text = ExcelPath_Getudo;
            this.txtExcelPath_Holiday.Text = ExcelPath_Holiday;
            this.txtExcelPath_Attendance.Text = ExcelPath_Attendance;
            this.txtKijunDate.Text = DateTime.Now.ToShortDateString();
            this.KijunDate = DateTime.Now;
            this.InitDisplayData();
        }

        private void MetroWindow_SizeChanged(object sender, SizeChangedEventArgs e)
        {
            var a = this.Height;
            this.dgMontlyData.Height = a - 175;
        }
    }
}
