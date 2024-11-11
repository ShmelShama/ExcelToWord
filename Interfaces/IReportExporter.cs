using ExcelToWord.Models.ReportModels;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ExcelToWord.Interfaces
{
    public interface IReportExporter:IExporter
    {
         void SetExportData(IEnumerable<DepartmentReport> exportData);

    }
}
