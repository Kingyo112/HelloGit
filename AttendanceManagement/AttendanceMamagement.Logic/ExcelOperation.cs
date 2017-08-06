using AttendanceManagement.Data;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.IO;

namespace AttendanceMamagement.Logic
{
    public class ExcelOperation
    {
        public static ErrorInfo OpenExcelFile(string ExcelPath, out FileStream file)
        {
            var errorinfo = new ErrorInfo();
            file = null;

            if (!File.Exists(ExcelPath))
            {
                errorinfo.HasError = true;
                errorinfo.ErrorReason = "ファイルが見つかりません。";
                return errorinfo;
            }

            var guidValue = Guid.NewGuid();
            string tempfile = ExcelPath.Replace(".xlsx", guidValue.ToString() + ".xlsx");
            try
            {
                File.Copy(ExcelPath, tempfile, true);
                file = new FileStream(tempfile, FileMode.Open);
            }
            catch (Exception ex)
            {
                errorinfo.HasError = true;
                errorinfo.ErrorReason = ex.Message;
            }
            return errorinfo;
        }

        public static ErrorInfo ReleaseExcelFile(FileStream file)
        {
            try
            {
                if (File.Exists(file.Name))
                {
                    File.Delete(file.Name);
                    file.Dispose();
                }
            }
            catch (Exception ex)
            {
                var errorinfo = new ErrorInfo();
                errorinfo.HasError = true;
                errorinfo.ErrorReason = ex.Message;
                return errorinfo;
            }
            return new ErrorInfo();
        }

        public static ErrorInfo DeleteExcelFile(string ExcelPath ,out string AvoidExcelPath)
        {
            var errorinfo = new ErrorInfo();
            AvoidExcelPath = string.Empty;

            if (!File.Exists(ExcelPath))
            {
                errorinfo.HasError = true;
                errorinfo.ErrorReason = "ファイルが見つかりません。";
                return errorinfo;
            }

            var guidValue = Guid.NewGuid();
            string tempfile = ExcelPath.Replace(".xlsx", guidValue.ToString());
            try
            {
                File.Copy(ExcelPath, tempfile, true);
                AvoidExcelPath = tempfile;
                File.Delete(ExcelPath);
            }
            catch (Exception ex)
            {
                errorinfo.HasError = true;
                errorinfo.ErrorReason = ex.Message;
            }
            return errorinfo;
        }

    }
}
