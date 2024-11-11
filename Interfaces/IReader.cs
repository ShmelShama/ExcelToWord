using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ExcelToWord.Interfaces
{
    public interface IReader
    {
        bool ReadDataFromFile(string filePath);
        string Message { get; }

        bool Status { get; }
    }
}
