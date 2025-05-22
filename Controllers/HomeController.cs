using ClosedXML.Excel;
using ExtractDataFromExcel.Models;
using ExtractDataFromExcel.Services;
using Microsoft.AspNetCore.Mvc;
using System.Diagnostics;

namespace ExtractDataFromExcel.Controllers
{
    public class HomeController : Controller
    {
        // Simulated in-memory database
        private readonly IEmployeeService _employeeService;

        public HomeController(IEmployeeService employeeService)
        {
            _employeeService = employeeService;
        }

        public async Task<IActionResult> Index()
        {
            var employees = await _employeeService.GetAllEmployeesAsync();
            return View(employees);
        }

        public IActionResult Create()
        {
            return View();
        }

        [HttpPost]
        public async Task<IActionResult> Create(Employee employee)
        {
            if (ModelState.IsValid)
            {
                await _employeeService.AddEmployeeAsync(employee);
                TempData["Message"] = "Employee added successfully.";
                return RedirectToAction("Index");
            }

            return View(employee);
        }

        public async Task<IActionResult> Edit(int id)
        {
            var employee = await _employeeService.GetEmployeeByIdAsync(id);

            if (employee == null)
            {
                TempData["Error"] = "Employee not found.";
                return RedirectToAction("Index");
            }

            ViewBag.IsEdit = true;
            return View("Create", employee);
        }

        [HttpPost]
        public async Task<IActionResult> Edit(int id, Employee employee)
        {
            if (ModelState.IsValid)
            {
                var updateResult = await _employeeService.UpdateEmployeeAsync(id, employee);
                if (updateResult)
                {
                    TempData["Message"] = "Employee updated successfully.";
                    return RedirectToAction("Index");
                }
                else
                {
                    TempData["Error"] = "Employee not found.";
                    return RedirectToAction("Index");
                }
            }

            ViewBag.Index = id;
            return View("Create", employee);
        }


        public async Task<IActionResult> Delete(int id)
        {
            var deleteResult = await _employeeService.DeleteEmployeeAsync(id);
            if (deleteResult)
            {
                TempData["Message"] = "Employee deleted successfully.";
            }
            else
            {
                TempData["Error"] = "Invalid employee ID.";
            }

            return RedirectToAction("Index");
        }


        public IActionResult Upload()
        {
            return View();
        }

        [HttpPost]
        public async Task<IActionResult> Upload(IFormFile file)
        {
            if (file == null || file.Length == 0)
            {
                TempData["Error"] = "Please select a valid file.";
                return RedirectToAction("Upload");
            }

            var extension = Path.GetExtension(file.FileName).ToLower();
            var employees = new List<Employee>();

            try
            {
                using var stream = new MemoryStream();
                await file.CopyToAsync(stream);

                if (extension == ".xlsx")
                {
                    using var workbook = new XLWorkbook(stream);
                    var worksheet = workbook.Worksheets.First();
                    var rows = worksheet.RangeUsed().RowsUsed().Skip(1); // skip header

                    foreach (var row in rows)
                    {
                        try
                        {
                            var employee = new Employee
                            {
                                Name = row.Cell(1).GetValue<string>()?.Trim(),
                                Email = row.Cell(2).GetValue<string>()?.Trim(),
                                Phone = row.Cell(3).GetValue<string>()?.Trim(),
                                Salary = row.Cell(4).GetValue<decimal>(),
                                Age = row.Cell(5).GetValue<int>(),
                                JoiningDate = row.Cell(6).GetDateTime()
                            };

                            if (!string.IsNullOrEmpty(employee.Name) && !string.IsNullOrEmpty(employee.Email))
                            {
                                employees.Add(employee);
                            }
                        }
                        catch (Exception ex)
                        {
                            Console.WriteLine($"Row parse error: {ex.Message}");
                        }
                    }
                }
                else if (extension == ".csv")
                {
                    stream.Position = 0;
                    using var reader = new StreamReader(stream);
                    bool isFirstRow = true;

                    while (!reader.EndOfStream)
                    {
                        var line = await reader.ReadLineAsync();
                        if (isFirstRow)
                        {
                            isFirstRow = false;
                            continue;
                        }

                        var values = line.Split(',');
                        if (values.Length < 6) continue;

                        try
                        {
                            var employee = new Employee
                            {
                                Name = values[0].Trim(),
                                Email = values[1].Trim(),
                                Phone = values[2].Trim(),
                                Salary = decimal.TryParse(values[3], out var salary) ? salary : 0,
                                Age = int.TryParse(values[4], out var age) ? age : 0,
                                JoiningDate = DateTime.TryParse(values[5], out var date) ? date : DateTime.MinValue
                            };

                            if (!string.IsNullOrEmpty(employee.Name) && !string.IsNullOrEmpty(employee.Email))
                            {
                                employees.Add(employee);
                            }
                        }
                        catch (Exception ex)
                        {
                            Console.WriteLine($"CSV row parse error: {ex.Message}");
                        }
                    }
                }
                else
                {
                    TempData["Error"] = "Unsupported file format. Please upload .xlsx or .csv.";
                    return RedirectToAction("Upload");
                }

                if (employees.Any())
                {
                    await _employeeService.BulkInsertEmployeesAsync(employees);
                    TempData["Message"] = "File uploaded and data saved successfully.";
                }
                else
                {
                    TempData["Error"] = "No valid employee records found in the file.";
                }

                return RedirectToAction("Index");
            }
            catch (Exception ex)
            {
                TempData["Error"] = $"An error occurred while processing the file: {ex.Message}";
                return RedirectToAction("Upload");
            }
        }

    }
}
