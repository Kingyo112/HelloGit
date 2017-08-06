using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace AttendanceManagement.Data
{
    public class DailyAttendanceData
    {
        public int GetudoYYYYMM { get; set; }

        public DateTime Day { get; set; }

        public EnumManager.DayKBN DayKBN { get; set; }

        public int? PlanMMSS_Start { get; set; }

        public int? PlanMMSS_END { get; set; }

        public int? ResultMMSS_Start { get; set; }

        public int? ResultMMSS_END { get; set; }

        public string Bikou { get; set; }

        public DailyAttendanceData()
        {
            GetudoYYYYMM = 0;
            Day = DateTime.Now;
            DayKBN = EnumManager.DayKBN.WeekDay;
            PlanMMSS_Start = null;
            PlanMMSS_END = null;
            ResultMMSS_Start = null;
            ResultMMSS_END = null;
            Bikou = string.Empty;
        }

        public bool ConvertToData(DispDailyAttendanceData data)
        {
            try
            {
                GetudoYYYYMM = data.GetudoYYYYMM;
                Day = data.Day;
                DayKBN = data.DayKBN;
                PlanMMSS_Start = this.ConvertToIntHHDD(data.Disp_PlanMMSS_Start);
                PlanMMSS_END = this.ConvertToIntHHDD(data.Disp_PlanMMSS_END);
                ResultMMSS_Start = this.ConvertToIntHHDD(data.Disp_ResultMMSS_Start);
                ResultMMSS_END = this.ConvertToIntHHDD(data.Disp_ResultMMSS_END);
                Bikou = data.Bikou;
            }
            catch
            {
                return false;
            }
            return true;
        }

        private int? ConvertToIntHHDD(string p)
        {
            var s = p.Replace(":",string.Empty);
            int i;
            if(int.TryParse(s,out i))
            {
                return i;
            }

            return null;
        }



    }
}
