using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using ExcelToWord.Models;
namespace ExcelToWord.Interfaces
{
    public interface IDepartmentReader:IReader
    {
        IEnumerable<Department> GetData();
    }
}
