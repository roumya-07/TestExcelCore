using Microsoft.AspNetCore.Http;
using Microsoft.AspNetCore.Mvc;
using Microsoft.Extensions.Logging;
using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Threading.Tasks;
using TestExcelCore.Models;
using TestExcelCore.Repository;

namespace TestExcelCore.Controllers
{
    public class HomeController : Controller
    {
        private readonly ILogger<HomeController> _logger;
        IEmployeeRepo employeeRepo;
        public HomeController(ILogger<HomeController> logger, IEmployeeRepo emp)
        {
            _logger = logger;
            employeeRepo = emp;
        }

        public async Task<IActionResult> Index()
        {
            List<EmployeeDetails> obj = await employeeRepo.Get();
            return View(obj);
        }
        [HttpPost]
        public IActionResult Index(IFormFile File)
        {
            try
            {
                if (File != null && File.Length > 0)
                {
                    string fileExtension = File.ContentType.Trim();
                    if ((fileExtension == "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet") || (fileExtension == "application/vnd.ms-excel"))
                    {

                        var filename = File.FileName + DateTime.Now.ToString("yymmssfff");
                        var path = Path.Combine(
                                    Directory.GetCurrentDirectory(), "wwwroot/Files",
                                    filename);

                        using (var stream = new FileStream(path, FileMode.Create))
                        {
                            File.CopyToAsync(stream);
                        }

                        var fileinfo = new FileInfo(path);
                        ExcelPackage.LicenseContext = LicenseContext.NonCommercial;

                        using (var package = new ExcelPackage(fileinfo))
                        {
                            var workbook = package.Workbook;
                            var worksheet = workbook.Worksheets.First();
                            int colCount = worksheet.Dimension.End.Column;  //get Column Count
                            int rowCount = worksheet.Dimension.End.Row;
                            List<Employee> employees = new List<Employee>();


                            for (int row = 2; row <= rowCount; row++)
                            {
                                Employee employee = new Employee();
                                for (int col = 1; col <= colCount; col++)
                                {
                                    if (!string.IsNullOrEmpty(worksheet.Cells[row, col].Value?.ToString().Trim()))
                                    {
                                        if (col == 1)
                                        {
                                            employee.Id = worksheet.Cells[row, col].Value?.ToString().Trim();
                                        }
                                        else if (col == 2)
                                        {
                                            employee.Name = worksheet.Cells[row, col].Value?.ToString().Trim();
                                        }
                                        else if (col == 3)
                                        {
                                            employee.Email = worksheet.Cells[row, col].Value?.ToString().Trim();
                                        }
                                        else if (col == 4)
                                        {
                                            employee.Mobile = worksheet.Cells[row, col].Value?.ToString().Trim();
                                        }
                                        else if (col == 5)
                                        {
                                            employee.Specialization = worksheet.Cells[row, col].Value?.ToString().Trim();
                                        }
                                    }

                                }
                                if (!string.IsNullOrWhiteSpace(employee.Id))
                                {
                                    employees.Add(employee);
                                }
                            }
                            if (employees != null)
                                employeeRepo.AddEmployee(employees);
                        }

                    }
                }
            }
            catch (Exception ex)
            {

            }

            return RedirectToAction("Index");

        }
        public IActionResult Privacy()
        {
            return View();
        }

        [ResponseCache(Duration = 0, Location = ResponseCacheLocation.None, NoStore = true)]
        public IActionResult Error()
        {
            return View(new ErrorViewModel { RequestId = Activity.Current?.Id ?? HttpContext.TraceIdentifier });
        }
    }
}
