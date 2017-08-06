using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace AttendanceMamagement.Logic
{
    public class OverTime
    {
        public OverTime()
        {

        }

        public TimeSpan ConvertOverTime(string string_overtime)
        {
            int i;
            if (int.TryParse(string_overtime, out i))
            {
                int d = i / 10000;
                int h = (i - d * 10000) / 100;
                int m = (i - d * 10000 - h * 100);
                return new TimeSpan(d, h, m, 0);
            }
            return new TimeSpan(0, 0, 0);
        }

        public TimeSpan ConvertDispOverTime(string string_overtime)
        {
            int i;
            if (int.TryParse(string_overtime.Replace(":",""), out i))
            {
                int d = (i / 100) / 24;
                int h = (i / 100) % 24;
                int m = (i - d * 2400 - h * 100);
                return new TimeSpan(d, h, m, 0);
            }
            return new TimeSpan(0, 0, 0);
        }


        public string ConvertOverTimeToString(TimeSpan overtime)
        {
            int hourWithDay = 24 * overtime.Days + overtime.Hours;
            string minutes;
            if(overtime.Minutes < 10)
            {
                minutes = "0" + overtime.Minutes.ToString();
            }
            else
            {
                minutes = overtime.Minutes.ToString();
            }
            return hourWithDay.ToString() + minutes;   
        }

    }
}
