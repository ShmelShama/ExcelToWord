using ExcelToWord.Interfaces;
using ExcelToWord.Models.ReportModels;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ExcelToWord.WordUtils
{
    public class ReportWordExporter : IReportExporter
    {
        public List<DepartmentReport> ExportData;
        public void SetExportData(IEnumerable<DepartmentReport> exportData)
        {
            ExportData = exportData.ToList();
        }
        public bool Export(string path)
        {
            throw new NotImplementedException();
        }

        
    }
}
