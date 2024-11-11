using ExcelToWord.Models;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.IO;
using ExcelToWord.Interfaces;
using Excel = Microsoft.Office.Interop.Excel;


namespace ExcelToWord.ExcelUtils
{
    public class JobExcelReader:IJobReader
    {
        public JobExcelReader() { }
        private List<Job> Jobs { get; set; } = new List<Job>();
        public string Message { get; private set; }

        public bool Status { get; private set; }

        private readonly List<string> columnNames = new List<string>() { "ИД задачи", "Табельный номер" };
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
            Jobs = new List<Job>();
            var excelApp = new Excel.Application();
            var wb = excelApp.Workbooks.Open(filePath);
            Excel.Worksheet wsht = null;
            bool hasWsht = false;
            foreach (Excel.Worksheet sheet in wb.Worksheets)
            {
                if (sheet.Name == "Задачи")
                {
                    wsht = sheet;
                    hasWsht = true;
                };
            }
            if (!hasWsht)
            {
                Message = "Не найден лист Задачи";
                wb.Close(false);
                excelApp.Quit();
                return Status = false;
            }

            for (int i = 0; i < columnNames.Count; i++)
            {
                string value = wsht.Cells[1, i + 1].Value;
                if (!value.Contains(columnNames[i]))
                {
                    Message = "Неcовпадение колонок в листе Задачи";
                    wb.Close(false);
                    excelApp.Quit();
                    return Status = false;
                }

            }


            for (int i = 2; ; i++)
            {

                string valueId = wsht.Cells[i, 1].Value?.ToString();
                if (string.IsNullOrEmpty(valueId))
                    break;

                long id;
                if (!long.TryParse(valueId, out id))
                {

                    Message = "Поврежденные данные в листе Задачи";
                    wb.Close(false);
                    excelApp.Quit();
                    return Status = false;

                }

                string valueEmployeeId = wsht.Cells[i, 2].Value?.ToString();
                long employeeId;
                if (!long.TryParse(valueEmployeeId, out employeeId))
                {

                    Message = "Поврежденные данные в листе Задачи";
                    wb.Close(false);
                    excelApp.Quit();
                    return Status = false;

                }

                var job =Jobs.FirstOrDefault(t => t.Id == id);
                if(job!=null)
                {
                    job.EmployeeIds.Add(employeeId);
                }
                else
                {
                    Jobs.Add(new Job()
                    {
                        Id = id,
                        EmployeeIds = new List<long>() { employeeId },
                    });
                }

            }

            Message = $"Данные по задачам из файла {filePath} загружены";
            excelApp.Quit();
            return Status = true;

           
           
        }
    }
}
