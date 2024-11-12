using ExcelToWord.Interfaces;
using ExcelToWord.Models.ReportModels;
using Word = Microsoft.Office.Interop.Word;
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

        public string Message { get; private set; }

        public bool Status { get; private set; }

        public void SetExportData(IEnumerable<DepartmentReport> exportData)
        {
            ExportData = exportData?.ToList();
        }
        public bool Export(string path)
        {
            if(ExportData == null)
            {
                Message = "Отсутствуют данные";
                return false;
            }
            var wordApp = new Word.Application();
            Word.Document doc = wordApp.Documents.Add();

            //Creating title
            Word.Paragraph paragraphTitle = doc.Paragraphs.Add();
            Word.Range rangeTitle = paragraphTitle.Range;
            rangeTitle.Font.Size = 14;
            rangeTitle.Text = "Отчет по загрузке";
            rangeTitle.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter;
            rangeTitle.InsertParagraphAfter();

            //Creating table
            Word.Paragraph tableParagraph = doc.Paragraphs.Add();
            Word.Range rangeTable= tableParagraph.Range;
            int rows = ExportData.Count+1;
            ExportData.ForEach(d => rows += d.EmployeeReports.Count);
            Word.Table table = doc.Tables.Add(rangeTable, rows, 2);
            table.Borders.InsideLineStyle = table.Borders.OutsideLineStyle = Word.WdLineStyle.wdLineStyleSingle ;

            //Fill table with data
            table.Rows[1].Range.Font.Color = Word.WdColor.wdColorWhite;
            table.Rows[1].Range.Font.Bold = 1;
            table.Rows[1].Range.Shading.BackgroundPatternColor = Word.WdColor.wdColorGray55;
            table.Cell(1, 1).Range.Text = "Отдел";
            table.Cell(1, 2).Range.Text = "Количество задач";
            int i = 2;
            foreach(var exportData in ExportData)
            {
                table.Rows[i].Range.Font.Bold = 1;
                table.Rows[i].Range.Shading.BackgroundPatternColor = Word.WdColor.wdColorGray20;
                table.Cell(i,1).Range.Text = exportData.Name;
                table.Cell(i,2).Range.Text = exportData.JobsCount.ToString();
                i++;
                foreach(var employee in exportData.EmployeeReports)
                {
                    table.Cell(i, 1).Range.Text = employee.Name;
                    table.Cell(i, 2).Range.Text = employee.JobsCount.ToString();
                    i++;
                }
            }
           
            try
            {
                doc.SaveAs2(path);
            }
            catch
            {
                Message = "Не удалось сохранить файл";
                doc.Close(false);
                wordApp.Quit();
                return Status = false;
        
                
                
            }
            Message = $"Отчет успешно создан {path}";
            doc.Close(false);
            wordApp.Quit();
            return Status=true;
            
        }

        
    }
}
