using ClosedXML.Excel;
using ExtractDataFromExcel.Models;
using Microsoft.AspNetCore.Mvc;
using System.Diagnostics;

namespace ExtractDataFromExcel.Controllers
{
    public class HomeController : Controller
    {
        // Simulated in-memory database
        private static List<Employee> _employees = new List<Employee>();

        public IActionResult Index()
        {
            return View(_employees);
        }

        public IActionResult Create()
        {
            return View();
        }

        [HttpPost]
        public IActionResult Create(Employee employee)
        {
            if (ModelState.IsValid)
            {
                _employees.Add(employee);
                TempData["Message"] = "Employee added successfully.";
                return RedirectToAction("Index");
            }

            return View(employee);
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

            try
            {
                using var stream = new MemoryStream();
                await file.CopyToAsync(stream);

                if (extension == ".xlsx")
                {
                    using var workbook = new XLWorkbook(stream);
                    var worksheet = workbook.Worksheets.First();
                    var rows = worksheet.RangeUsed().RowsUsed().Skip(1); // Skip header

                    foreach (var row in rows)
                    {
                        if (row.CellsUsed().Count() < 6) continue;

                        var employee = new Employee
                        {
                            Name = row.Cell(1).GetValue<string>(),
                            Email = row.Cell(2).GetValue<string>(),
                            Phone = row.Cell(3).GetValue<string>(),
                            Salary = row.Cell(4).GetValue<decimal>(),
                            Age = row.Cell(5).GetValue<int>(),
                            JoiningDate = row.Cell(6).GetValue<DateTime>()
                        };

                        _employees.Add(employee);
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
                            isFirstRow = false; // skip header
                            continue;
                        }

                        var values = line.Split(',');
                        if (values.Length < 6) continue;

                        var employee = new Employee
                        {
                            Name = values[0],
                            Email = values[1],
                            Phone = values[2],
                            Salary = decimal.TryParse(values[3], out var salary) ? salary : 0,
                            Age = int.TryParse(values[4], out var age) ? age : 0,
                            JoiningDate = DateTime.TryParse(values[5], out var date) ? date : DateTime.MinValue
                        };

                        _employees.Add(employee);
                    }
                }
                else
                {
                    TempData["Error"] = "Unsupported file format. Please upload .xlsx or .csv.";
                    return RedirectToAction("Upload");
                }

                TempData["Message"] = "File uploaded and data extracted successfully.";
                return RedirectToAction("Index");
            }
            catch (Exception ex)
            {
                TempData["Error"] = $"An error occurred: {ex.Message}";
                return RedirectToAction("Upload");
            }
        }

        public IActionResult Edit(int id)
        {
            if (id >= 0 && id < _employees.Count)
            {
                var employee = _employees[id];
                ViewBag.Index = id;
                return View("Create", employee);
            }

            TempData["Error"] = "Invalid employee ID.";
            return RedirectToAction("Index");
        }

        [HttpPost]
        public IActionResult Edit(int id, Employee employee)
        {
            if (ModelState.IsValid && id >= 0 && id < _employees.Count)
            {
                _employees[id] = employee;
                TempData["Message"] = "Employee updated successfully.";
                return RedirectToAction("Index");
            }

            ViewBag.Index = id;
            return View("Create", employee);
        }

        public IActionResult Delete(int id)
        {
            if (id >= 0 && id < _employees.Count)
            {
                _employees.RemoveAt(id);
                TempData["Message"] = "Employee deleted successfully.";
            }
            else
            {
                TempData["Error"] = "Invalid employee ID.";
            }

            return RedirectToAction("Index");
        }

        public IActionResult Privacy() => View();

        [ResponseCache(Duration = 0, Location = ResponseCacheLocation.None, NoStore = true)]
        public IActionResult Error() =>
            View(new ErrorViewModel { RequestId = Activity.Current?.Id ?? HttpContext.TraceIdentifier });
    }
}
