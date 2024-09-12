using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;

namespace match_employee.Models
{
    public class ExcelEmployee
    {
        public int ExcelRow { get; set; }
        public string FirstName { get; set; }
        public string LastName { get; set; }
        public string Gender { get; set; }
        public DateTime DateOfBirth { get; set; }
        public string[] NumbersFromDB {  get; set; }
    }
}
