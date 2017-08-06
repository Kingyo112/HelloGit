using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace AttendanceManagement.Data
{
    public class MasterGetudo
    {
        public int GetudoYYYYMM { get; set; }

        public DateTime StartDate { get; set; }

        public DateTime EndDate { get; set; }

        public MasterGetudo()
        {
            GetudoYYYYMM = 0;
            StartDate = DateTime.Now;
            EndDate = DateTime.Now;
        }
    }
}
