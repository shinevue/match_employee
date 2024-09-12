using match_employee.Models;
using Microsoft.AspNetCore.Http;
using Microsoft.AspNetCore.Mvc;
using Microsoft.AspNetCore.Routing.Constraints;
using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Threading.Tasks;

namespace match_employee.Controllers
{
    public class EmployeeController : Controller
    {
        // private readonly AppDbContext _context;
        private readonly EmployeeService _employeeService;
        
        public EmployeeController(EmployeeService employeeService)
        {
            _employeeService = employeeService;
        }

        public async Task<IActionResult> Index()
        {
            return View(_employeeService.GetAll());
        }

        //Upload and Compare employees excel file with Entire data
        [HttpPost]
        public async Task<IActionResult> Upload(IFormFile file)
        {
            if (file != null && file.Length > 0)
            {
                var filePath = Path.Combine(Path.GetTempPath(), file.FileName);

                using (var stream = new FileStream(filePath, FileMode.Create))
                {
                    await file.CopyToAsync(stream);
                }

                var excelEmployees = _employeeService.ReadExcelFile(filePath);
                var potentialDuplicates = _employeeService.FindPotentialDuplicates(excelEmployees);

                if (potentialDuplicates.Any())
                {
                    _employeeService.UpdateExcelFile(filePath, potentialDuplicates);
                }
                TempData["filePath"] = filePath;
                TempData["fileName"] = "Updated_" + file.FileName;
                return RedirectToAction("Results", new { fileName = file.FileName });
            }
            return Redirect("Results");
        }

        //Upload excel file to MSSQL database table as HR employees Table
        [HttpPost]
        public async Task<IActionResult> UploadEmployeeData(string type, IFormFile file)
        {
            if(type == "delete")
            {
                _employeeService.DeleteFromDatabase();
                return RedirectToAction("Index");
            }
            
            TempData["validationMsg"] = "";

            if (file != null && file.Length > 0)
            {
                var filePath = Path.Combine(Path.GetTempPath(), file.FileName);

                using (var stream = new FileStream(filePath, FileMode.Create))
                {
                    await file.CopyToAsync(stream);
                }

                var Employees = _employeeService.ReadExcelFile(filePath, false, true);

                _employeeService.SaveToDatabase(Employees);
            } 
            else
            {
                TempData["validationMsg"] = "No file is selected or selected invalid excel file.";
            }
            
            return RedirectToAction("Index");
        }

        //Read table data from specific excel file
        [HttpPost]
        public async Task<IActionResult> GetExcelView(IFormFile file)
        {
            if (file != null && file.Length > 0)
            {
                var filePath = Path.Combine(Path.GetTempPath(), file.FileName);
                
                using (var stream = new FileStream(filePath, FileMode.Create))
                {
                    await file.CopyToAsync(stream);
                }
                var Employees = _employeeService.ReadExcelFile(filePath);

                return Json(Employees);
            }

            return Json(new { });
        }
        [HttpGet]
        public IActionResult DownloadFile(string fileName)
        {
            var filePath = Path.Combine(Path.GetTempPath(), fileName);
            return File(System.IO.File.ReadAllBytes(filePath), "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", $"Updated_{fileName}");
        }
        
        public IActionResult Results(string fileName)
        {
            ViewData["fileName"] = fileName;
            var filePath = Path.Combine(Path.GetTempPath(), fileName);
            var Employees = _employeeService.ReadExcelFile(filePath, true);

            return View(Employees);
        }
    }
    /*
     public class EmployeeController : Controller
    {
        private readonly DbContext _context;

        public EmployeeController(DbContext context)
        {
            _context = context;
        }

        [HttpGet]
        public IActionResult Upload()
        {
            return View();
        }

        [HttpPost]
        public async Task<IActionResult> Upload(IFormFile file)
        {
            if (file == null || file.Length == 0)
                return BadRequest("No file uploaded.");

            var employeesFromExcel = new List<Employee>();

            // Set the LicenseContext property directly in your code before using any EPPlus functionality
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
            
            using (var package = new ExcelPackage(file.OpenReadStream()))
            {
                var worksheet = package.Workbook.Worksheets.First();
                var rowCount = worksheet.Dimension.Rows;

                for (int row = 2; row <= rowCount; row++)
                {
                    var firstName = worksheet.Cells[row, 1].Text;
                    var lastName = worksheet.Cells[row, 2].Text;
                    var gender = worksheet.Cells[row, 3].Text;
                    var dateOfBirth = DateTime.Parse(worksheet.Cells[row, 4].Text);

                    employeesFromExcel.Add(new Employee { FirstName = firstName, LastName = lastName, Gender = gender, DateOfBirth = dateOfBirth });
                }
            }

            var employees = _context.Employee
                .AsEnumerable() // Switch to client-side evaluation
                .Where(e => employeesFromExcel
                .Any(excelEmp =>
                    excelEmp.FirstName == e.FirstName &&
                    excelEmp.LastName == e.LastName &&
                    excelEmp.Gender == e.Gender &&
                    excelEmp.DateOfBirth == e.DateOfBirth))
                .ToList();

            return View("Results", employees);
        }
    }
     
     */

}
