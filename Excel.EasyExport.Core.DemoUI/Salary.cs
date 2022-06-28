using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;

namespace Excel.EasyExport.Core.DemoUI
{
    public class Salary
    {
        public string Name { get; set; }
        public double BaseSalary { get; set; }
        public int FullAttendance { get; set; }
        public decimal Transportation { get; set; }
        public short Meal { get; set; }

        public DateTime? CreateTime { get; set; }

        public bool IsFullAttendance { get; set; }
    }
}
