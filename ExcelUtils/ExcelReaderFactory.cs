using ExcelToWord.Interfaces;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ExcelToWord.ExcelUtils
{
    public class ExcelReaderFactory:IReaderFactory
    {
        public IEmployeeReader CreateEmployeeReader()
        {
           return new EmployeeExcelReader();
        }

        public IDepartmentReader CreateDepartmentReader()
        {
            return new DepartmentExcelReader();
        }

        public IJobReader CreateJobReader()
        {
            return new JobExcelReader();
        }

    }
}
