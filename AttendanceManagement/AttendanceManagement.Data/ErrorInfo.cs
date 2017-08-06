using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace AttendanceManagement.Data
{
    public class ErrorInfo
    {
        /// <summary>
        /// エラー
        /// </summary>
        public bool HasError { get; set; }
        /// <summary>
        /// エラー理由
        /// </summary>
        public string ErrorReason { get; set; }

        public ErrorInfo()
        {
            HasError = false;
            ErrorReason = string.Empty;                
        }

    }
}
