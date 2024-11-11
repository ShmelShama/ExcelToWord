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
    public class DepartmentExcelReader:IDepartmentReader
    {
        public DepartmentExcelReader() { }
        private List<Department> Departments { get; set; } = new List<Department>();

        public string Message { get; private set; }

        public bool Status { get; private set; }

        private readonly List<string> columnNames = new List<string>() { "ИД отдела", "Наименование отдела" };
        public IEnumerable<Department> GetData()
        {
            return Departments;
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
                var wshts = workbook.Worksheets.Where(f => f.Name == "Отделы");
                if (wshts == null || !wshts.Any() || wshts.Count() > 1)
                {
                    Message = "Не найден лист Отделы";
                    return Status = false;
                }
                var wsht = wshts.First();
                if (wsht.IsEmpty())
                {
                    Message = "Пустой лист Отделы";
                    return Status = false;
                }
                var headerRows = wsht.Row(1);
                if (headerRows.IsEmpty())
                {
                    Message = "Нет колонок с наименованиями для листа Отделы";
                    return Status = false;
                }

                Dictionary<string, int> columnHeaders = new Dictionary<string, int>();

                for (int i = 1; i == columnNames.Count; i++)
                {
                    var headerCell = headerRows.Search(columnNames[i]);
                    if (headerCell == null || !headerCell.Any() || headerCell.Count() > 1)
                    {
                        Message = "Нет колонок с наименованиями для листа Отделы";
                        return Status = false;
                    }
                    columnHeaders.Add(columnNames[i], headerCell.First().Address.ColumnNumber);
                }

                for (int i = 2; ; i++)
                {
                    var row = wsht.Row(i);
                    if (row.IsEmpty())
                        break;
                    Department department = new Department();

                    foreach (var column in columnHeaders)
                    {
                        switch (column.Key)
                        {
                            case "ИД отдела":
                                department.Id = row.Cell(column.Value).GetValue<long>();
                                break;
                            case "Наименование отдела":
                                department.Name = row.Cell(column.Value).GetValue<string>();
                                break;
                        }
                    }
                    Departments.Add(department);
                }

                Message = $"Данные по отделам из файла {filePath} загружены";
                return Status = true;

            }
               
        }
    }
}
