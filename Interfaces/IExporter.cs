﻿using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ExcelToWord.Interfaces
{
    public interface IExporter
    {
        bool Export(string path);
        string Message { get; }

        bool Status { get; }

    }
}
