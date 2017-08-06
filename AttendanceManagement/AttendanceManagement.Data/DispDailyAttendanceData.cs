using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace AttendanceManagement.Data
{
    public class DispDailyAttendanceData 
    {
        public int GetudoYYYYMM { get; set; }

        public DateTime Day { get; set; }

        public string WeekofDayName { get; set; }

        public EnumManager.DayKBN DayKBN { get; set; }
        
        public string DayKBNName { get; set; }

        public int? PlanMMSS_Start { get; set; }

        public int? PlanMMSS_END { get; set; }

        public int? ResultMMSS_Start { get; set; }

        public int? ResultMMSS_END { get; set; }

        public string Disp_PlanMMSS_Start { get; set; }

        public string Disp_PlanMMSS_END { get; set; }

        public string Disp_ResultMMSS_Start { get; set; }

        public string Disp_ResultMMSS_END { get; set; }

        public string Bikou { get; set; }

        public DispDailyAttendanceData()
        {
            GetudoYYYYMM = 0;
            Day = DateTime.Now;
            WeekofDayName = string.Empty;
            DayKBN = EnumManager.DayKBN.WeekDay;
            DayKBNName = string.Empty;
            PlanMMSS_Start = null;
            PlanMMSS_END = null;
            ResultMMSS_Start = null;
            ResultMMSS_END = null;
            Disp_PlanMMSS_Start = null;
            Disp_PlanMMSS_END = null;
            Disp_ResultMMSS_Start = null;
            Disp_ResultMMSS_END = null;
            Bikou = string.Empty;
        }

        public bool ConvertToDispData(DailyAttendanceData data)
        {
            try
            {
                GetudoYYYYMM = data.GetudoYYYYMM;
                Day = data.Day;
                WeekofDayName = EnumManager.GetWeekofDayName(data.Day);
                DayKBN = data.DayKBN;
                DayKBNName = EnumManager.GetDayKBNName(data.DayKBN);
                PlanMMSS_Start = data.PlanMMSS_Start;
                PlanMMSS_END = data.PlanMMSS_END;
                ResultMMSS_Start = data.ResultMMSS_Start;
                ResultMMSS_END = data.ResultMMSS_END;
                Disp_PlanMMSS_Start = this.ConvertToHHMM(data.PlanMMSS_Start);
                Disp_PlanMMSS_END = this.ConvertToHHMM(data.PlanMMSS_END);
                Disp_ResultMMSS_Start = this.ConvertToHHMM(data.ResultMMSS_Start);
                Disp_ResultMMSS_END = this.ConvertToHHMM(data.ResultMMSS_END);
                Bikou = data.Bikou;
            }
            catch(Exception ex)
            {
                return false;
            }
            return true;
        }

        private string ConvertToHHMM(int? hhmm)
        {
            string hh = string.Empty;
            string mm = string.Empty;
            int h = hhmm == null ? 0 : (int)hhmm / 100;
            int m = hhmm == null ? 0 : (int)(hhmm - h * 100);
            this.ZeroFilled(m, out mm);
            if(this.ZeroFilled(h ,out hh))
            {
                return hh + ":" + mm;
            }
            return string.Empty;
        }

        private bool ZeroFilled(int h ,out string hh)
        {
            hh = string.Empty;
            if (h >= 10)
            {
                hh = h.ToString();
            }
            else if(h == 0)
            {
                hh = "00";
                return false;
            }
            else
            {
                hh = "0" + h.ToString();
            }

            return true;
        }
    }
}
