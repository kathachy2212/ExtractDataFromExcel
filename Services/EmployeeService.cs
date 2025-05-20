// Services/EmployeeService.cs
using ClosedXML.Excel;
using ExtractDataFromExcel.Data;
using ExtractDataFromExcel.Models;
using Microsoft.AspNetCore.Http;
using System;
using System.Collections.Generic;
using System.IO;
using System.Threading.Tasks;

namespace ExtractDataFromExcel.Services
{
    public class EmployeeService : IEmployeeService
    {
        private readonly AppDbContext _context;

        public EmployeeService(AppDbContext context)
        {
            _context = context;
        }

        public async Task UploadExcelAsync(IFormFile file)
        {
            using var stream = new MemoryStream();
            await file.CopyToAsync(stream);
            using var workbook = new XLWorkbook(stream);
            var worksheet = workbook.Worksheet(1);
            var employees = new List<Employee>();

            for (int row = 2; row <= worksheet.LastRowUsed().RowNumber(); row++)
            {
                var emp = new Employee
                {
                    Name = worksheet.Cell(row, 1).GetString(),
                    Email = worksheet.Cell(row, 2).GetString(),
                    Phone = worksheet.Cell(row, 3).GetString(),
                    Salary = Convert.ToDecimal(worksheet.Cell(row, 4).GetDouble()),
                    Age = worksheet.Cell(row, 5).GetValue<int>(),
                    JoiningDate = worksheet.Cell(row, 6).GetDateTime()
                };
                employees.Add(emp);
            }

            _context.Employees.AddRange(employees);
            await _context.SaveChangesAsync();
        }

        public async Task AddEmployeeAsync(Employee employee)
        {
            _context.Employees.Add(employee);
            await _context.SaveChangesAsync();
        }

        public async Task<List<Employee>> GetAllEmployeesAsync()
        {
            return await Task.FromResult(_context.Employees.ToList());
        }
    }
}
