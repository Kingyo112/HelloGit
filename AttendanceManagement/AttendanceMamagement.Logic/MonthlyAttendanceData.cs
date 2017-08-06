using AttendanceManagement.Data;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.IO;
using NPOI.HSSF.UserModel;
using NPOI.HSSF.Util;
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;


namespace AttendanceMamagement.Logic
{
    public class MonthlyAttendanceData
    {

        private List<DailyAttendanceData> DailyAttendanceList = new List<DailyAttendanceData>();

        public string ExcelPath { get; set; }

        public string ExcelPath_Holidays { get; set; }

        public MonthlyAttendanceData()
        {

        }

        public ErrorInfo GetInitAttendanceData(MasterGetudo getudorecord, out List<List<DailyAttendanceData>> attendanceDataBox , out List<SummaryAttendanceData> summaryAttendanceList)
        {
            attendanceDataBox = new List<List<DailyAttendanceData>>();
            summaryAttendanceList = new List<SummaryAttendanceData>();
            var errorinfo = new ErrorInfo();

            if (getudorecord.GetudoYYYYMM == 0)
            {
                errorinfo.HasError = true;
                errorinfo.ErrorReason = "月度の値が不正です。";
                return errorinfo;
            }

            FileStream file;
            errorinfo = ExcelOperation.OpenExcelFile(ExcelPath, out file);
            if (errorinfo.HasError)
            {
                return errorinfo;
            }
            errorinfo = GetAttendanceDataList(file, out attendanceDataBox, out summaryAttendanceList);
            if (errorinfo.HasError)
            {
                return errorinfo;
            }
            errorinfo = ExcelOperation.ReleaseExcelFile(file);

            if (!attendanceDataBox.Any()
                || attendanceDataBox.Count(x => x.Count(r => r.GetudoYYYYMM == getudorecord.GetudoYYYYMM) != 0) == 0)
            {
                var list = new List<DailyAttendanceData>();
                errorinfo = GetInitMonthlyData(getudorecord, out list);
                attendanceDataBox.Add(list);
            }

            return errorinfo;
        }

        private ErrorInfo GetAttendanceDataList(FileStream file, out List<List<DailyAttendanceData>> attendanceDataBox, out List<SummaryAttendanceData> summaryAttendanceList)
        {
            var errorinfo = new ErrorInfo();
            attendanceDataBox = new List<List<DailyAttendanceData>>();
            summaryAttendanceList = new List<SummaryAttendanceData>();
            try
            {
                IWorkbook workbook = WorkbookFactory.Create(file);
                errorinfo = GetHasRecordGetudoYYYYMM(workbook, out summaryAttendanceList);
                if (errorinfo.HasError)
                {
                    return errorinfo;
                }
                foreach (var item in summaryAttendanceList)
                {
                    var attendanceDataList = new List<DailyAttendanceData>();
                    errorinfo = GetAttendanceData(workbook, item.GetudoYYYYMM, out attendanceDataList);
                    if (attendanceDataList.Count == 0) continue;
                    attendanceDataBox.Add(attendanceDataList);
                }
            }
            catch (Exception ex)
            {
                errorinfo.HasError = true;
                errorinfo.ErrorReason = ex.Message;
            }
            return errorinfo;

        }

        private ErrorInfo GetAttendanceData(IWorkbook workbook, int item, out List<DailyAttendanceData> attendanceDataList)
        {
            var errorinfo = new ErrorInfo();
            attendanceDataList = new List<DailyAttendanceData>();
            try
            {
                var sheet = workbook.GetSheet(item.ToString());
                if (sheet == null)
                {
                    return errorinfo;
                }
                int lastRow = sheet.LastRowNum;
                for (int i = 1; i <= lastRow; i++)
                {
                    var record = new DailyAttendanceData();
                    IRow row = sheet.GetRow(i);
                    ICell cell = row == null ? null : row.GetCell(0);
                    int intyyyymm = cell == null ? 0 : (int)cell.NumericCellValue;
                    if (intyyyymm != 0)
                    {
                        record.GetudoYYYYMM = intyyyymm;
                        ICell cell_date = row == null ? null : row.GetCell(1);
                        string string_date = cell_date == null ? string.Empty : cell_date.DateCellValue.ToShortDateString();
                        DateTime date;
                        DateTime.TryParse(string_date, out date);
                        record.Day = date;
                        ICell cell_daykbn = row == null ? null : row.GetCell(2);
                        double daykbn = cell_daykbn == null ? 0 : cell_daykbn.NumericCellValue;
                        record.DayKBN = (EnumManager.DayKBN)((int)daykbn);

                        ICell cell_PlanMMSS_Start = row == null ? null : row.GetCell(3);
                        string plammmss_st = cell_PlanMMSS_Start == null ? null : cell_PlanMMSS_Start.StringCellValue;
                        record.PlanMMSS_Start = string.IsNullOrEmpty(plammmss_st) ? (int?)null : int.Parse(plammmss_st);
                        ICell cell_PlanMMSS_End = row == null ? null : row.GetCell(4);
                        string plammmss_end = cell_PlanMMSS_End == null ? null : cell_PlanMMSS_End.StringCellValue;
                        record.PlanMMSS_END = string.IsNullOrEmpty(plammmss_end) ? (int?)null : int.Parse(plammmss_end);
                        ICell cell_ResultMMSS_Start = row == null ? null : row.GetCell(5);
                        string resultmmss_st = cell_ResultMMSS_Start == null ? null : cell_ResultMMSS_Start.StringCellValue;
                        record.ResultMMSS_Start = string.IsNullOrEmpty(resultmmss_st) ? (int?)null : int.Parse(resultmmss_st);
                        ICell cell_ResultMMSS_End = row == null ? null : row.GetCell(6);
                        string resultmmss_end = cell_ResultMMSS_End == null ? null : cell_ResultMMSS_End.StringCellValue;
                        record.ResultMMSS_END = string.IsNullOrEmpty(resultmmss_end) ? (int?)null : int.Parse(resultmmss_end);

                        ICell cell_Bikou = row == null ? null : row.GetCell(7);
                        record.Bikou = cell_Bikou == null ? string.Empty : cell_Bikou.StringCellValue;

                        attendanceDataList.Add(record);
                    }
                }

            }
            catch (Exception ex)
            {
                errorinfo.HasError = true;
                errorinfo.ErrorReason = ex.Message;
            }

            return errorinfo;
        }

        private ErrorInfo GetHasRecordGetudoYYYYMM(IWorkbook workbook, out List<SummaryAttendanceData> GetudoYYYYMMArray)
        {
            GetudoYYYYMMArray = new List<SummaryAttendanceData>();
            var errorinfo = new ErrorInfo();
            try
            {
                var sheet = workbook.GetSheet("Master");
                int lastRow = sheet.LastRowNum;
                var list = new List<int>();
                for (int i = 1; i <= lastRow; i++)
                {
                    var record = new SummaryAttendanceData();
                    var otime = new OverTime();
                    IRow row = sheet.GetRow(i);
                    ICell cell = row == null ? null : row.GetCell(0);
                    int intyyyymm = cell == null ? 0 : (int)cell.NumericCellValue;
                    if (intyyyymm != 0)
                    {
                        record.GetudoYYYYMM = intyyyymm;
                        ICell cell_overtime = row == null ? null : row.GetCell(1);
                        string string_overtime = cell_overtime == null ? string.Empty : cell_overtime.StringCellValue;
                        record.OverTime = otime.ConvertOverTime(string_overtime);
                        ICell cell_is45over = row == null ? null : row.GetCell(2);
                        bool bool_is45over = cell_is45over == null ? false : cell_is45over.BooleanCellValue;
                        record.Is45Over = bool_is45over;
                        ICell cell_is80over = row == null ? null : row.GetCell(3);
                        bool bool_is80over = cell_is80over == null ? false : cell_is80over.BooleanCellValue;
                        record.Is80Over = bool_is80over;
                        GetudoYYYYMMArray.Add(record);
                    }
                }

            }
            catch (Exception ex)
            {
                errorinfo.HasError = true;
                errorinfo.ErrorReason = ex.Message;
            }
            return errorinfo;
        }

        public ErrorInfo GetInitMonthlyData(MasterGetudo getudorecord, out List<DailyAttendanceData> monthlyattendancelist)
        {
            monthlyattendancelist = new List<DailyAttendanceData>();
            var errorinfo = new ErrorInfo();

            if (getudorecord.GetudoYYYYMM == 0)
            {
                errorinfo.HasError = true;
                errorinfo.ErrorReason = "月度の値が不正です。";
                return errorinfo;
            }

            //祝日マスタの読み込み
            var holidaylist = new List<MasterHoliday>();
            var holidays = new Holidays();
            holidays.ExcelPath = ExcelPath_Holidays;
            errorinfo = holidays.GetHolidaysByGetudoRecord(getudorecord, out holidaylist);
            if (errorinfo.HasError)
            {
                return errorinfo;
            }

            for (DateTime date = getudorecord.StartDate; date <= getudorecord.EndDate; date = date.AddDays(1))
            {

                var attendancerecord = new DailyAttendanceData();
                attendancerecord.GetudoYYYYMM = getudorecord.GetudoYYYYMM;
                attendancerecord.Day = date;

                var holiday = holidaylist.Find(x => x.Day == date);
                if (holiday != null)
                {
                    attendancerecord.DayKBN = EnumManager.DayKBN.HolyDay;
                    attendancerecord.Bikou = holiday.HolidayName;
                }
                else
                {
                    attendancerecord.DayKBN = this.GetDayKBN(date);
                }
                monthlyattendancelist.Add(attendancerecord);
            }
            return errorinfo;
        }

        private EnumManager.DayKBN GetDayKBN(DateTime date)
        {
            EnumManager.DayKBN daykbn = EnumManager.DayKBN.WeekDay;
            switch (date.DayOfWeek)
            {
                case DayOfWeek.Monday:
                case DayOfWeek.Tuesday:
                case DayOfWeek.Wednesday:
                case DayOfWeek.Thursday:
                case DayOfWeek.Friday:
                    daykbn = EnumManager.DayKBN.WeekDay;
                    break;
                case DayOfWeek.Saturday:
                case DayOfWeek.Sunday:
                    daykbn = EnumManager.DayKBN.WeekEnd;
                    break;
                default:
                    break;
            }

            return daykbn;
        }


        public ErrorInfo SaveData(List<List<DailyAttendanceData>> list,List<SummaryAttendanceData> list_sum , DateTime kijundate)
        {
            string AvoidPath = string.Empty;
            var errorinfo = ExcelOperation.DeleteExcelFile(this.ExcelPath, out AvoidPath);
            if (errorinfo.HasError)
            {
                return errorinfo;
            }

            IWorkbook workbook = new XSSFWorkbook();
            using (var fs = new FileStream(this.ExcelPath, FileMode.OpenOrCreate))
            {
                errorinfo = this.SaveSumAttendanceData(list_sum, workbook, out workbook);

                foreach (var item in list)
                {
                    int getudoyyyymm = item.Max(x => x.GetudoYYYYMM);
                    errorinfo = this.SaveExcelSheet(getudoyyyymm, item, workbook, out workbook);
                    if (errorinfo.HasError)
                    {
                        return errorinfo;
                    }
                }
                workbook.Write(fs);

            }

            File.Delete(AvoidPath);
            return errorinfo;
        }

        private ErrorInfo SaveSumAttendanceData(List<SummaryAttendanceData> list, IWorkbook workbook_old, out IWorkbook workbook)
        {
            workbook = workbook_old;
            var errorinfo = new ErrorInfo();
            var otime = new OverTime();
            try
            {
                var sheet = workbook.CreateSheet("Master");
                int rowcount = list.Count;
                this.GetHeaderAttendanceData_Master(sheet);
                for (int i = 0; i < rowcount; i++)
                {
                    var record = new DailyAttendanceData();
                    IRow row = sheet.CreateRow(i + 1);
                    ICell cell_GetudoYYYYMM = row.CreateCell(0);
                    cell_GetudoYYYYMM.SetCellValue(list[i].GetudoYYYYMM);
                    ICell cell_OverTime = row.CreateCell(1);
                    string overtime = otime.ConvertOverTimeToString(list[i].OverTime);
                    cell_OverTime.SetCellValue(overtime);
                    ICell cell_IS45Over = row.CreateCell(2);
                    cell_IS45Over.SetCellValue(list[i].Is45Over);
                    ICell cell_IS80Over = row.CreateCell(3);
                    cell_IS80Over.SetCellValue(list[i].Is80Over);
                }

            }
            catch (Exception ex)
            {
                errorinfo.HasError = true;
                errorinfo.ErrorReason = ex.Message;
            }

            return errorinfo;
        }

        private void GetHeaderAttendanceData_Master(ISheet sheet)
        {
            IRow row = sheet.CreateRow(0);
            ICell cell_GetudoYYYYMM = row.CreateCell(0);
            cell_GetudoYYYYMM.SetCellValue("Master");
            ICell cell_Day = row.CreateCell(1);
            cell_Day.SetCellValue("OverTime");
            ICell cell_DayKBN = row.CreateCell(2);
            cell_DayKBN.SetCellValue("Is45Over");
            ICell cell_PlanMMSS_Start = row.CreateCell(3);
            cell_PlanMMSS_Start.SetCellValue("Is80Over");
        }

        private ErrorInfo SaveExcelSheet(int getudoyyyymm, List<DailyAttendanceData> item, IWorkbook workbook_old, out IWorkbook workbook)
        {
            workbook = workbook_old;
            var errorinfo = new ErrorInfo();
            try
            {
                var sheet = workbook.CreateSheet(getudoyyyymm.ToString());
                int rowcount = item.Count;
                this.GetHeaderAttendanceData(sheet);
                for (int i = 0; i < rowcount; i++)
                {
                    var record = new DailyAttendanceData();
                    IRow row = sheet.CreateRow(i + 1);
                    ICell cell_GetudoYYYYMM = row.CreateCell(0);
                    cell_GetudoYYYYMM.SetCellValue(item[i].GetudoYYYYMM);
                    ICell cell_Day = row.CreateCell(1);
                    cell_Day.SetCellValue(item[i].Day);
                    ICell cell_DayKBN = row.CreateCell(2);
                    cell_DayKBN.SetCellValue((int)item[i].DayKBN);
                    ICell cell_PlanMMSS_Start = row.CreateCell(3);
                    cell_PlanMMSS_Start.SetCellValue(item[i].PlanMMSS_Start.ToString());
                    ICell cell_PlanMMSS_END = row.CreateCell(4);
                    cell_PlanMMSS_END.SetCellValue(item[i].PlanMMSS_END.ToString());
                    ICell cell_ResultMMSS_Start = row.CreateCell(5);
                    cell_ResultMMSS_Start.SetCellValue(item[i].ResultMMSS_Start.ToString());
                    ICell cell_ResultMMSS_END = row.CreateCell(6);
                    cell_ResultMMSS_END.SetCellValue(item[i].ResultMMSS_END.ToString());
                    ICell cell_Bikou = row.CreateCell(7);
                    cell_Bikou.SetCellValue(item[i].Bikou);
                }

            }
            catch (Exception ex)
            {
                errorinfo.HasError = true;
                errorinfo.ErrorReason = ex.Message;
            }

            return errorinfo;

        }

        private void GetHeaderAttendanceData(ISheet sheet)
        {
            IRow row = sheet.CreateRow(0);
            ICell cell_GetudoYYYYMM = row.CreateCell(0);
            cell_GetudoYYYYMM.SetCellValue("GetudoYYYYMM");
            ICell cell_Day = row.CreateCell(1);
            cell_Day.SetCellValue("Day");
            ICell cell_DayKBN = row.CreateCell(2);
            cell_DayKBN.SetCellValue("DayKBN");
            ICell cell_PlanMMSS_Start = row.CreateCell(3);
            cell_PlanMMSS_Start.SetCellValue("PlanMMSS_Start");
            ICell cell_PlanMMSS_END = row.CreateCell(4);
            cell_PlanMMSS_END.SetCellValue("PlanMMSS_END");
            ICell cell_ResultMMSS_Start = row.CreateCell(5);
            cell_ResultMMSS_Start.SetCellValue("ResultMMSS_Start");
            ICell cell_ResultMMSS_END = row.CreateCell(6);
            cell_ResultMMSS_END.SetCellValue("ResultMMSS_END");
            ICell cell_Bikou = row.CreateCell(7);
            cell_Bikou.SetCellValue("Bikou");
        }

        public TimeSpan CalcOverTime(List<DispDailyAttendanceData> list, DateTime KijunDate)
        {
            var ts = new TimeSpan(0, 0, 0);
            foreach (var item in list)
            {
                switch (item.DayKBN)
                {
                    case EnumManager.DayKBN.WeekDay:
                        ts = ts + this.CalcTimeSpan_WeekDay(item, KijunDate);
                        break;
                    case EnumManager.DayKBN.WeekEnd:
                    case EnumManager.DayKBN.HolyDay:
                    case EnumManager.DayKBN.PaidVacation:
                        ts = ts + this.CalcTimeSpan_NotWeekDay(item, KijunDate);
                        break;
                    default:
                        break;
                }
            }
            return ts;
        }

        private TimeSpan CalcTimeSpan_WeekDay(DispDailyAttendanceData item, DateTime KijunDate)
        {
            if (item.Day >= KijunDate)
            {
                return this.CalcTimeSpan_WeekDay(item.PlanMMSS_Start, item.PlanMMSS_END);
            }
            else
            {
                return this.CalcTimeSpan_WeekDay(item.ResultMMSS_Start, item.ResultMMSS_END);
            }
        }

        private TimeSpan CalcTimeSpan_NotWeekDay(DispDailyAttendanceData item, DateTime KijunDate)
        {
            if (item.Day >= KijunDate)
            {
                return this.CalcTimeSpan_NotWeekDay(item.PlanMMSS_Start, item.PlanMMSS_END);
            }
            else
            {
                return this.CalcTimeSpan_NotWeekDay(item.ResultMMSS_Start, item.ResultMMSS_END);
            }
        }

        private TimeSpan CalcTimeSpan_WeekDay(int? start, int? end)
        {
            int hh;
            int mm;
            var ts_sta = new TimeSpan(0, 0, 0);
            var ts_end = new TimeSpan(0, 0, 0);

            hh = start != null ? (int)start / 100 : 0;
            if (hh <= 0)
            {
                hh = 0;
                mm = 0;
            }
            else
            {
                mm = (int)start - hh * 100;
                ts_sta = new TimeSpan(9 - hh <= 0 ? 0 : 9 - hh, (-1) * mm, 0);
            }

            hh = end != null ? (int)end / 100 : 0;
            if (hh <= 0)
            {
                hh = 0;
                mm = 0;
            }
            else
            {
                mm = (int)end - hh * 100;
                ts_end = new TimeSpan(hh - 18 <= 0 ? 0 : hh - 18, mm, 0);
            }
            return ts_sta + ts_end;
        }

        private TimeSpan CalcTimeSpan_NotWeekDay(int? start, int? end)
        {
            int hh;
            int mm;

            if (start != null && end != null)
            {
                hh = (int)(start / 100);
                mm = (int)start - hh * 100;
                var ts_sta = new TimeSpan(hh, mm, 0);
                hh = (int)(end / 100);
                mm = (int)end - hh * 100;
                var ts_end = new TimeSpan(hh, mm, 0);

                return ts_end - ts_sta;
            }

            return new TimeSpan(0, 0, 0);
        }

    }
}
