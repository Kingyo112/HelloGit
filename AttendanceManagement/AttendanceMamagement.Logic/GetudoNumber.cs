using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using AttendanceManagement.Data;
using NPOI.HSSF.UserModel;
using NPOI.HSSF.Util;
using NPOI.SS.UserModel;
using System.IO;
using AttendanceMamagement.Logic;

namespace AttendanceManagement.Logic
{
    public class GetudoNumber
    {
        private List<MasterGetudo> GetudoNumList = new List<MasterGetudo>();

        public string ExcelPath { get; set; }

        public GetudoNumber()
        {

        }

        public ErrorInfo GetInitValue(out MasterGetudo Getudonumrecord)
        {
            return GetGetudoNum(ExcelPath, DateTime.Now, out Getudonumrecord);
        }

        public ErrorInfo GetConsistentValue(int getudoyyyymm, out int newgetudoyyyymm, out MasterGetudo Getudonumrecord)
        {
            newgetudoyyyymm = GetNewGetudoNumber(getudoyyyymm);

            return GetGetudoNumByInt(newgetudoyyyymm, out Getudonumrecord);
        }

        private int GetNewGetudoNumber(int value)
        {
            var i = value / 100;
            int c = value - i * 100;
            switch ((EnumManager.Month)c)
            {
                case EnumManager.Month.Bottom:
                    value = (i - 1) * 100 + 12;
                    break;
                case EnumManager.Month.January:
                case EnumManager.Month.February:
                case EnumManager.Month.March:
                case EnumManager.Month.April:
                case EnumManager.Month.May:
                case EnumManager.Month.June:
                case EnumManager.Month.July:
                case EnumManager.Month.August:
                case EnumManager.Month.September:
                case EnumManager.Month.October:
                case EnumManager.Month.November:
                case EnumManager.Month.December:
                    break;
                case EnumManager.Month.Top:
                    value = (i + 1) * 100 + 1;
                    break;
                default:
                    break;
            }
            return value;
        }

        private ErrorInfo GetGetudoNumByInt(int getudoyyyymm, out MasterGetudo Getudonumrecord)
        {
            foreach (var item in GetudoNumList)
            {
                if (item.GetudoYYYYMM == getudoyyyymm)
                {
                    Getudonumrecord = item;
                    return new ErrorInfo();
                }
            }
            Getudonumrecord = new MasterGetudo();
            var errorInfo = new ErrorInfo();
            errorInfo.HasError = true;
            errorInfo.ErrorReason = "月度マスタ未登録です";
            return errorInfo;
        }

        private ErrorInfo GetGetudoNum(string ExcelPath, DateTime date, out MasterGetudo Getudonumrecoed)
        {
            FileStream file;
            ErrorInfo errorinfo = new ErrorInfo();
            Getudonumrecoed = new MasterGetudo();
            errorinfo = ExcelOperation.OpenExcelFile(ExcelPath, out file);
            if (errorinfo.HasError)
            {
                return errorinfo;
            }
            errorinfo = GetGetudoNumList(file, out GetudoNumList);
            if (errorinfo.HasError)
            {
                return errorinfo;
            }
            errorinfo = ExcelOperation.ReleaseExcelFile(file);


            if (!GetudoNumList.Any())
            {
                errorinfo.HasError = true;
                errorinfo.ErrorReason = "月度マスタの取得に失敗しました";
                return errorinfo;
            }

            
            var getudonumrecord = GetudoNumList.Find(x => x.StartDate <= date && x.EndDate >= date);
            if (getudonumrecord == null)
            {
                errorinfo.HasError = true;
                errorinfo.ErrorReason = "月度マスタに当月のデータがありません";
            }
            else
            {
                Getudonumrecoed = getudonumrecord;
            }
            return errorinfo;
        }

        private ErrorInfo GetGetudoNumList(FileStream file, out List<MasterGetudo> GetudonumList)
        {
            GetudonumList = new List<MasterGetudo>();
            try
            {
                IWorkbook workbook = WorkbookFactory.Create(file);
                var sheet = workbook.GetSheetAt(0);
                int lastRow = sheet.LastRowNum;
                for (int i = 1; i <= lastRow; i++)
                {
                    var getudorecord = new MasterGetudo();
                    IRow row = sheet.GetRow(i);
                    ICell cell = row == null ? null : row.GetCell(0);
                    int intyyyymm = cell == null ? 0 : (int)cell.NumericCellValue;
                    if (intyyyymm != 0)
                    {
                        getudorecord.GetudoYYYYMM = intyyyymm;
                        ICell cell_startdate = row == null ? null : row.GetCell(1);
                        string string_startdate = cell_startdate == null ? string.Empty : cell_startdate.DateCellValue.ToShortDateString();
                        DateTime startdate;
                        DateTime.TryParse(string_startdate, out startdate);
                        getudorecord.StartDate = startdate;
                        ICell cell_enddate = row == null ? null : row.GetCell(2);
                        string string_enddate = cell_enddate == null ? string.Empty : cell_enddate.DateCellValue.ToShortDateString();
                        DateTime enddate;
                        DateTime.TryParse(string_enddate, out enddate);
                        getudorecord.EndDate = enddate;
                        GetudonumList.Add(getudorecord);
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
