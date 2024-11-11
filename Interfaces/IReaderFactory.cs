using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ExcelToWord.Interfaces
{
    public interface IReaderFactory
    {
        IEmployeeReader CreateEmployeeReader();
        IDepartmentReader CreateDepartmentReader();
        IJobReader CreateJobReader();
    }
}
