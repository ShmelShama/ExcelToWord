using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.IO;
using ExcelToWord.Models;
using ExcelToWord.Interfaces;

using System.Windows.Forms;

using Excel = Microsoft.Office.Interop.Excel;

namespace ExcelToWord.ExcelUtils
{
    public class EmployeeExcelReader : IEmployeeReader
    {
        public EmployeeExcelReader() { }
        private List<Employee> Employees { get; set; } = new List<Employee>();
        public string Message { get; private set; }

        public bool Status { get; private set; }

        private readonly List<string> columnNames = new List<string>() { "Табельный номер", "Фамилия", "Имя", "Отчество", "Дата рождения", "Отдел" };
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
            Employees = new List<Employee>();
            var excelApp = new Excel.Application();
            var wb = excelApp.Workbooks.Open(filePath);
            Excel.Worksheet wsht = null;
            bool hasWsht = false;
            foreach (Excel.Worksheet sheet in wb.Worksheets)
            {
                if (sheet.Name == "Сотрудники")
                {
                    wsht = sheet;
                    hasWsht = true;
                };
            }
            if (!hasWsht)
            {
                Message = "Не найден лист Сотрудники";
                wb.Close(false);
                excelApp.Quit();
                return Status = false;
            }

            for (int i = 0; i < columnNames.Count; i++)
            {
                string value = wsht.Cells[1, i + 1].Value;
                if (!value.Contains( columnNames[i]))
                {
                    Message = "Неcовпадение колонок в листе Сотрудники";
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

                    Message = "Поврежденные данные в листе Сотрудники";
                    wb.Close(false);
                    excelApp.Quit();
                    return Status = false;

                }

                string valueSurname = wsht.Cells[i, 2].Value?.ToString();
                string valueName = wsht.Cells[i, 3].Value?.ToString();
                string valuePatronymic = wsht.Cells[i, 4].Value?.ToString();

                string valueBirthday = wsht.Cells[i, 5].Value?.ToString();
                DateTime birthday;
                if (!DateTime.TryParse(valueBirthday, out birthday))
                {

                    wb.Close(false);
                    Message = "Поврежденные данные в листе Сотрудники";
                    excelApp.Quit();
                    return Status = false;

                }

                string valueDepartmentId = wsht.Cells[i, 6].Value?.ToString();
                long departmentId;
                if (!long.TryParse(valueDepartmentId, out departmentId))
                {
                    Message = "Поврежденные данные в листе Сотрудники";
                    wb.Close(false);
                    excelApp.Quit();
                    return Status = false;

                }

                Employees.Add(new Employee()
                {
                    Id = id,
                    Surname = valueSurname ?? string.Empty,
                    Name = valueName ?? string.Empty,
                    Patronymic = valuePatronymic ?? string.Empty,
                    BirthDay = birthday,
                    DepartmentId = departmentId,
                });




            }

            Message = $"Данные по сотрудниками из файла {filePath} загружены";
            wb.Close(false);
            excelApp.Quit();
            return Status = true;

            
              
        }
    }
}
