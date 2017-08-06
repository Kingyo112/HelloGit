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


namespace AttendanceMamagement.Logic
{
    class Holidays
    {
        private List<MasterHoliday> HolidayList = new List<MasterHoliday>();

        public string ExcelPath { get; set; }

        public Holidays()
        {

        }

        public ErrorInfo GetHolidaysByGetudoRecord(MasterGetudo getudorecord, out List<MasterHoliday> holidays)
        {
            holidays = HolidayList;
            var errorinfo = GetHolidays(ExcelPath);
            if(errorinfo.HasError)
            {
                return errorinfo;
            }

            holidays = HolidayList.Where(x => x.GetudoYYYYMM == getudorecord.GetudoYYYYMM).ToList();
            if (holidays == null)
            {
                holidays = new List<MasterHoliday>();
            }
            return errorinfo;

        }

        private ErrorInfo GetHolidays(string ExcelPath)
        {
            FileStream file;
            ErrorInfo errorinfo = new ErrorInfo();
            errorinfo = ExcelOperation.OpenExcelFile(ExcelPath, out file);
            if (errorinfo.HasError)
            {
                return errorinfo;
            }
            errorinfo = GetHolidayList(file, out HolidayList);
            if (errorinfo.HasError)
            {
                return errorinfo;
            }
            errorinfo = ExcelOperation.ReleaseExcelFile(file);
            return errorinfo;
        }

        private ErrorInfo GetHolidayList(FileStream file, out List<MasterHoliday> HolidayList)
        {
            HolidayList = new List<MasterHoliday>();
            try
            {
                IWorkbook workbook = WorkbookFactory.Create(file);
                var sheet = workbook.GetSheetAt(0);
                int lastRow = sheet.LastRowNum;
                for (int i = 1; i <= lastRow; i++)
                {
                    var holidayrecord = new MasterHoliday();
                    IRow row = sheet.GetRow(i);
                    ICell cell = row == null ? null : row.GetCell(0);
                    int intyyyymm = cell == null ? 0 : (int)cell.NumericCellValue;
                    if (intyyyymm != 0)
                    {
                        holidayrecord.GetudoYYYYMM = intyyyymm;
                        ICell cell_date = row == null ? null : row.GetCell(1);
                        string string_date = cell_date == null ? string.Empty : cell_date.DateCellValue.ToShortDateString();
                        DateTime date;
                        DateTime.TryParse(string_date, out date);
                        holidayrecord.Day = date;
                        ICell cell_holidayname = row == null ? null : row.GetCell(2);
                        string string_enddate = cell_holidayname == null ? string.Empty : cell_holidayname.StringCellValue;
                        holidayrecord.HolidayName = string_enddate;
                        HolidayList.Add(holidayrecord);
                    }
                    else
                    {
                        return new ErrorInfo();
                    }
                }

            }
            catch (Exception ex)
            {
                var errorinfo = new ErrorInfo();
                errorinfo.HasError = true;
                errorinfo.ErrorReason = ex.Message;
            }
            return new ErrorInfo();
        }

    }
}
