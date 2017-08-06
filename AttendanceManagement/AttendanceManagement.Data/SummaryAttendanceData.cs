using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace AttendanceManagement.Data
{
    public class SummaryAttendanceData
    {
        public int GetudoYYYYMM { get; set; }

        public TimeSpan OverTime { get; set; }
        
        public bool Is45Over { get; set; }
        
        public bool Is80Over { get; set; }

        public SummaryAttendanceData()
        {
            GetudoYYYYMM = 0;
            OverTime = new TimeSpan(0,0,0);
            Is45Over = false;
            Is80Over = false;
        }
    }
}
