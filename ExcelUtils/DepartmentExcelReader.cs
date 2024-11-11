using ExcelToWord.Models;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.IO;
using Excel = Microsoft.Office.Interop.Excel;
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
            Departments = new List<Department>();
            var excelApp = new Excel.Application();
            var wb = excelApp.Workbooks.Open(filePath);
            Excel.Worksheet wsht = null ;
            bool hasWsht = false;
            foreach(Excel.Worksheet sheet in wb.Worksheets)
            {
                if (sheet.Name == "Отделы")
                {
                    wsht=sheet;
                    hasWsht = true;
                };
            }
            if (!hasWsht)
            {
                Message = "Не найден лист Отделы";
                wb.Close(false);
                excelApp.Quit();
                return Status = false;
            }
            
            for(int i=0; i<columnNames.Count;i++ )
            {
                string value = wsht.Cells[1,i+1].Value;
                if(!value.Contains(columnNames[i]))
                {
                    Message = "Неcовпадение колонок в листе Отделы";
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

                    Message = "Поврежденные данные в листе Отделы";
                    wb.Close(false);
                    excelApp.Quit();
                    return Status = false;

                }

                string valueName = wsht.Cells[i, 2].Value.ToString();

                Departments.Add( new Department()
                {
                    Id = id,
                    Name = valueName ?? string.Empty,
                });




            }

            Message = $"Данные по отделам из файла {filePath} загружены";
            wb.Close(false);
            excelApp.Quit();
            return Status= true;


               
        }
    }
}
