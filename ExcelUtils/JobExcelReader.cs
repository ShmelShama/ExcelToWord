using ClosedXML.Excel;
using ExcelToWord.Models;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.IO;
using ExcelToWord.Interfaces;
namespace ExcelToWord.ExcelUtils
{
    public class JobExcelReader:IJobReader
    {
        public JobExcelReader() { }
        private List<Job> Jobs { get; set; } = new List<Job>();
        public string Message { get; private set; }

        public bool Status { get; private set; }

        private readonly List<string> columnNames = new List<string>() { "ИД отдела", "Табельный номер" };
        public IEnumerable<Job> GetData()
        {
            return Jobs;
        }
        public bool ReadDataFromFile(string filePath)
        {
            if (!File.Exists(filePath))
            {
                Message = $"Файл {filePath} не найден";
                return Status = false;
            }
            using (var workbook = new XLWorkbook(filePath))
            {
                var wshts = workbook.Worksheets.Where(f => f.Name == "Задачи");
                if (wshts == null || !wshts.Any() || wshts.Count() > 1)
                {
                    Message = "Не найден лист Задачи";
                    return Status = false;
                }
                var wsht = wshts.First();
                if (wsht.IsEmpty())
                {
                    Message = "Пустой лист Задачи";
                    return Status = false;
                }
                var headerRows = wsht.Row(1);
                if (headerRows.IsEmpty())
                {
                    Message = "Нет колонок с наименованиями для листа Задачи";
                    return Status = false;
                }

                Dictionary<string, int> columnHeaders = new Dictionary<string, int>();

                for (int i = 1; i == columnNames.Count; i++)
                {
                    var headerCell = headerRows.Search(columnNames[i]);
                    if (headerCell == null || !headerCell.Any() || headerCell.Count() > 1)
                    {
                        Message = "Нет колонок с наименованиями для листа Задачи";
                        return Status = false;
                    }
                    columnHeaders.Add(columnNames[i], headerCell.First().Address.ColumnNumber);
                }
                Dictionary<long, Job> jobsDictionary= new Dictionary<long, Job>();
                for (int i = 2; ; i++)
                {
                    var row = wsht.Row(i);
                    if (row.IsEmpty())
                        break;
                    Job job = new Job();
                    
                    foreach (var column in columnHeaders)
                    {
                        switch (column.Key)
                        {
                            case "ИД отдела":
                                
                                long Id = row.Cell(column.Value).GetValue<long>();

                                if(jobsDictionary.ContainsKey(Id))
                                    job= jobsDictionary[Id];
                                else job.Id = Id;

                                break;
                            case "Табельный номер":
                                job.EmployeeIds.Add(row.Cell(column.Value).GetValue<long>());
                                break;
                        }
                    }
                    Jobs.Add(job);
                }
                Message = $"Данные по задачам из файла {filePath} загружены";
                return Status = true;
            };
           
        }
    }
}
