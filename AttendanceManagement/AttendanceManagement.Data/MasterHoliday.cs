using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace AttendanceManagement.Data
{
    public class MasterHoliday
    {
        public int GetudoYYYYMM { get; set; }

        public DateTime Day { get; set; }

        public string HolidayName { get; set; }

        public MasterHoliday()
        {
            GetudoYYYYMM = 0;
            Day = DateTime.Now;
            HolidayName = string.Empty;
        }
    }
}
