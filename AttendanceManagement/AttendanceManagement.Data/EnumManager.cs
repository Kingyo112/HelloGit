using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace AttendanceManagement.Data
{
    public class EnumManager
    {
        public enum Month : int
        {
            Bottom = 0,
            January = 1,
            February = 2,
            March = 3,
            April = 4,
            May = 5,
            June = 6,
            July = 7,
            August = 8,
            September = 9,
            October = 10,
            November = 11,
            December = 12,
            Top = 13,
        }

        public enum DayKBN : byte
        {
            WeekDay = 0,
            WeekEnd = 1,
            HolyDay = 2,
            PaidVacation = 3,
        }


        public static string GetDayKBNName(DayKBN dayKBN)
        {
            string daykbnname;
            switch (dayKBN)
            {
                case DayKBN.WeekDay:
                    daykbnname = "平日";
                    break;
                case DayKBN.WeekEnd:
                    daykbnname = "休日";
                    break;
                case DayKBN.HolyDay:
                    daykbnname = "祝日";
                    break;
                case DayKBN.PaidVacation:
                    daykbnname = "有給";
                    break;
                default:
                    daykbnname = string.Empty;
                    break;
            }
            return daykbnname;
        }

        public static string GetWeekofDayName(DateTime dateTime)
        {
            string dayofweek;
            switch (dateTime.DayOfWeek)
            {
                case DayOfWeek.Friday:
                    dayofweek = "(金)";
                    break;
                case DayOfWeek.Monday:
                    dayofweek = "(月)";
                    break;
                case DayOfWeek.Saturday:
                    dayofweek = "(土)";
                    break;
                case DayOfWeek.Sunday:
                    dayofweek = "(日)";
                    break;
                case DayOfWeek.Thursday:
                    dayofweek = "(木)";
                    break;
                case DayOfWeek.Tuesday:
                    dayofweek = "(火)";
                    break;
                case DayOfWeek.Wednesday:
                    dayofweek = "(水)";
                    break;
                default:
                    dayofweek = string.Empty;
                    break;
            }
            return dayofweek;
        }

    }
}
