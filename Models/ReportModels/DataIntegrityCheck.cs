using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ExcelToWord.Models.ReportModels
{
    public class DataIntegrityCheck
    {
        public DataIntegrityCheck() { }

        public string Message { get; private set; } = "Данные не проверены";

        public bool Status { get; private set; } = false;

        public bool CheckData(IEnumerable<Department> departments, IEnumerable<Employee> employees, IEnumerable<Job> jobs)
        {
            Message = string.Empty;
            if (!CheckBase(departments)
                || !CheckBase(employees)
                || !CheckBase(jobs))
            { return Status=false; }    

            foreach(Employee employee in employees)
            {
                if(!departments.Any(d=>d.Id==employee.DepartmentId))
                {
                    Message = $"Отсутствуют данные по отделу для сотрудника {employee.Id}";
                    return Status=false;
                }
            }
            foreach (Job job in jobs)
            {
                foreach( long employeeId in job.EmployeeIds)
                {
                    if (!employees.Any(d => d.Id == employeeId))
                    {
                        Message = $"Отсутствуют данные по сотруднику для задачи {job.Id}";
                        return Status = false;
                    }
                }
                
            }
            Message = "Данные проверены";
            return Status=true;
        }

        private bool CheckBase(IEnumerable<Base> models)
        {
            if (!models.Any())
            {
                Status = false;
                Message = "Отсутствуют данные";
                return false;
            }
            if (models.GroupBy(m => m.Id).Any(i => i.Count() > 1))
            {
                Status = false;
                Message = "Дублирование Id";
            }
            return true;
        }


    }
}
