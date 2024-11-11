using DocumentFormat.OpenXml.Bibliography;
using ExcelToWord.Interfaces;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ExcelToWord.Models.ReportModels
{
    public class DepartmentReportCompiler:IDepartmentReportCompiler
    {
        public DepartmentReportCompiler() { }

        public IEnumerable<DepartmentReport> CreateReport(IEnumerable<Department> departments, IEnumerable<Employee> employees, IEnumerable<Job> jobs)
        {
            if(departments==null|| employees==null || jobs==null) return null;
            var departmentsList = departments.ToList();
            var employeesList = employees.ToList();
            var jobsList = jobs.ToList();
            List<DepartmentReport> departmentReports = new List<DepartmentReport>(departments.Count());
           

            foreach (var department in departmentsList)
            {
                var departmentEmployees = employees.Where(e => e.DepartmentId == department.Id);
                var employeesDictionary = new Dictionary<long, EmployeeReport>(departmentEmployees.Count());
                foreach (var employee in departmentEmployees)
                {
                    string Name = string.IsNullOrEmpty(employee.Name) ? "" : $"{employee.Name.First()}.";
                    string Patronymic = string.IsNullOrEmpty(employee.Patronymic) ? "" : $"{employee.Name.First()}.";
                    employeesDictionary.Add(employee.Id, new EmployeeReport()
                    {
                        Name = $"{employee.Surname} {Name}{Patronymic}",
                        JobsCount = 0
                    });
                }
                var departmentJobs = jobs.Where(j => j.EmployeeIds.Any(e => employeesDictionary.ContainsKey(e)));
                foreach (var job in departmentJobs)
                {
                    foreach (long employeeId in job.EmployeeIds)
                    {
                        employeesDictionary[employeeId].JobsCount++;
                    }
                }

                departmentReports.Add(new DepartmentReport()
                {
                    Name = department.Name,
                    JobsCount = departmentJobs.Count(),
                    EmployeeReports = employeesDictionary.Values.ToList(),
                });

            } 
            return departmentReports;
        }
    }
}
