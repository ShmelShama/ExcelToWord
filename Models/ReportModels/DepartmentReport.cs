using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ExcelToWord.Models.ReportModels
{
    public class DepartmentReport
    {
        public string Name { get; set; }
        public int JobsCount { get; set; }
        public List< EmployeeReport> EmployeeReports { get; set; }
    }
}
