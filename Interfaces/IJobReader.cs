using ExcelToWord.Models;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ExcelToWord.Interfaces
{
    public interface IJobReader: IReader
    {
        IEnumerable<Job> GetData();
    }
}
