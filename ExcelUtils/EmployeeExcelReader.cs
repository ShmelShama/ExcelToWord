using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.IO;
using ClosedXML.Excel;
using ExcelToWord.Models;
using ExcelToWord.Interfaces;
using static System.Windows.Forms.VisualStyles.VisualStyleElement;
using System.Windows.Forms;

namespace ExcelToWord.ExcelUtils
{
    public class EmployeeExcelReader : IEmployeeReader
    {
        public EmployeeExcelReader() { }
        private List<Employee> Employees { get; set; } = new List<Employee>();
        public string Message { get; private set; }

        public bool Status { get; private set; }

        private readonly List<string> columnNames = new List<string>() { "Табельный номер", "Фамилия", "Имя", "Отчество", "ДатаРождения", "Отдел" };
        public IEnumerable<Employee> GetData() 
        { 
            return Employees; 
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
                var wshts = workbook.Worksheets.Where(f => f.Name == "Сотрудники");
                if (wshts == null || !wshts.Any() || wshts.Count() > 1)
                {
                    Message = "Не найден лист Сотрудники";
                    return Status = false;
                }
                var wsht = wshts.First();
                if (wsht.IsEmpty())
                {
                    Message = "Пустой лист Сотрудники";
                    return Status = false;
                }
                var headerRows = wsht.Row(1);
                if (headerRows.IsEmpty())
                {
                    Message = "Нет колонок с наименованиями для листа Сотрудники";
                    return Status = false;
                }

                Dictionary<string, int> columnHeaders = new Dictionary<string, int>();

                for (int i = 1; i == columnNames.Count; i++)
                {
                    var headerCell = headerRows.Search(columnNames[i]);
                    if (headerCell == null || !headerCell.Any() || headerCell.Count() > 1)
                    {
                        Message = "Нет колонок с наименованиями для листа Сотрудники";
                        return Status = false;
                    }
                    columnHeaders.Add(columnNames[i], headerCell.First().Address.ColumnNumber);
                }

                for (int i = 2; ; i++)
                {
                    var row = wsht.Row(i);
                    if (row.IsEmpty())
                        break;
                    Employee employee = new Employee();

                    foreach (var column in columnHeaders)
                    {
                        switch (column.Key)
                        {
                            case "Табельный номер":
                                employee.Id = row.Cell(column.Value).GetValue<long>();
                                break;
                            case "Фамилия":
                                employee.Surname = row.Cell(column.Value).GetValue<string>();
                                break;
                            case "Имя":
                                employee.Name = row.Cell(column.Value).GetValue<string>();
                                break;
                            case "Отчество":
                                employee.Patronymic = row.Cell(column.Value).GetValue<string>();
                                break;
                            case "ДатаРождения":
                                employee.BirthDay = row.Cell(column.Value).GetValue<DateTime>();
                                break;
                            case "Отдел":
                                employee.DepartmentId = row.Cell(column.Value).GetValue<int>();
                                break;

                        }
                    }
                    Employees.Add(employee);
                }
                Message = $"Данные по сотрудникам из файла {filePath} загружены";
                return Status = true;
            }
              
        }
    }
}
