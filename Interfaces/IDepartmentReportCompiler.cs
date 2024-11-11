using ExcelToWord.Models.ReportModels;
using ExcelToWord.Models;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ExcelToWord.Interfaces
{
    public interface IDepartmentReportCompiler:IReportCompiler
    {
        IEnumerable<DepartmentReport> CreateReport(IEnumerable<Department> departments, IEnumerable<Employee> employees, IEnumerable<Job> jobs);
    }
}
