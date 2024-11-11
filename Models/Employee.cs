using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Annotations;
namespace ExcelToWord.Models
{
    public class Employee:Base
    {
        public Employee() { }
        public string Name { get; set; } = string.Empty;
        public string Surname { get; set; } = string.Empty;
        public string Patronymic { get; set; } = string.Empty;
        public string FullName
        {
            get
            {
                return $"{Surname} {Name} {FullName}";
            }
        }
        public DateTime BirthDay { get; set; }
        public int DepartmentId { get; set; }
        
    }
}
