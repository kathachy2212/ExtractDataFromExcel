using System;

namespace ExtractDataFromExcel.Models
{
    public class Employee
    {
        public int Id { get; set; }  // Primary key
        public string Name { get; set; }
        public string Email { get; set; }
        public string Phone { get; set; }
        public decimal Salary { get; set; }
        public int Age { get; set; }
        public DateTime JoiningDate { get; set; }
    }
}
