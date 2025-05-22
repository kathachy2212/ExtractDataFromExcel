// Services/IEmployeeService.cs
using ExtractDataFromExcel.Models;
using Microsoft.AspNetCore.Http;
using System.Collections.Generic;
using System.Threading.Tasks;

namespace ExtractDataFromExcel.Services
{
    public interface IEmployeeService
    {
        Task UploadExcelAsync(IFormFile file);
        Task AddEmployeeAsync(Employee employee);
        Task<List<Employee>> GetAllEmployeesAsync();
        Task BulkInsertEmployeesAsync(List<Employee> employees);
        Task<bool> UpdateEmployeeAsync(int id, Employee updatedEmployee);
        Task<bool> DeleteEmployeeAsync(int id);
        Task<Employee> GetEmployeeByIdAsync(int id);

    }
}
