using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Security.Cryptography.X509Certificates;
using System.Threading.Tasks;
using Microsoft.AspNetCore.Http;
using Microsoft.AspNetCore.Mvc;
using Microsoft.EntityFrameworkCore;
using Microsoft.EntityFrameworkCore.Metadata.Internal;
using OfficeOpenXml;
using static System.Runtime.InteropServices.JavaScript.JSType;

namespace match_employee.Models
{
    public class EmployeeService
    {
        private readonly AppDbContext _context;

        public EmployeeService(AppDbContext context)
        {
            _context = context;
        }

        public List<Employee> GetAll()
        {
            return _context.Employee
                .AsEnumerable()
                .OrderBy (e => int.Parse(e.Number))
                .ToList<Employee>();
        }
        
        public List<Employee> ReadExcelFile(string filePath, bool isResult=false, bool isDB = false)
        {
            // Set the LicenseContext property directly in your code before using any EPPlus functionality
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;

            var employees = new List<Employee>();

            using (var package = new ExcelPackage(new FileInfo(filePath)))
            {
                var worksheetCount = package.Workbook.Worksheets.Count;

                if (worksheetCount > 0)
                {
                    var worksheet = package.Workbook.Worksheets[0];
                    int rowCount = worksheet.Dimension.End.Row;

                    for (int row = 2; row <= rowCount; row++)
                    {
                        bool isRowEmpty = true; // Assume the row is empty
                        for (int col = 1; col <= worksheet.Dimension.End.Column; col++)
                        {
                            var cellValue = worksheet.Cells[row, col].Value;
                            if (cellValue != null)
                            {
                                isRowEmpty = false; // Found a non-empty cell
                                break; // No need to check further
                            }
                        }
                        //check if specific row is empty or not
                        if (isRowEmpty)
                        {
                            continue;
                        }

                        employees.Add(new Employee
                        {
                            Number = isDB ? worksheet.Cells[row, 1].Text : (row - 1).ToString(),
                            FirstName = worksheet.Cells[row, isDB ? 2 : 1].Text,
                            LastName = worksheet.Cells[row, isDB ? 3 : 2].Text,
                            Gender = worksheet.Cells[row, 4].Text,
                            DateOfBirth = worksheet.Cells[row, 5].GetCellValue<DateTime>(),
                            NumberOfEmployees = isResult ? worksheet.Cells[row, 6].Text : "",
                            ExcelRow = row,
                        });
                    }
                }
            }
            
            return employees;
        }

        public void UpdateExcelFile(string filePath, List<ExcelEmployee> potentialDuplicates)
        {
            // Set the LicenseContext property directly in your code before using any EPPlus functionality
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;

            using (var package = new ExcelPackage(new FileInfo(filePath)))
            {
                var worksheet = package.Workbook.Worksheets[0];
                int rowCount = worksheet.Dimension.Rows;

                // Add a header for potential duplicates
                worksheet.Cells[1, 6].Value = "Employee_Number_From_DB";
                //worksheet.Cells[1, 26].Value = "first_name";
                //worksheet.Cells[1, 27].Value = "last_name";
                //worksheet.Cells[1, 28].Value = "gender";
                //worksheet.Cells[1, 29].Value = "dateofbirth";
                //worksheet.Cells[1, 30].Value = "placeofbirth";

                // int row = 2; // Start after header row
                foreach (var emp in potentialDuplicates)
                {
                    worksheet.Cells[emp.ExcelRow+1, 6].Value = string.Join(",", emp.NumbersFromDB); //$"{emp.FirstName} {emp.LastName} ({emp.Gender}, {emp.DateOfBirth.ToShortDateString()})";
                    //worksheet.Cells[emp.ExcelRow.Value, 26].Value = emp.FirstName;
                    //worksheet.Cells[emp.ExcelRow.Value, 27].Value = emp.LastName;
                    //worksheet.Cells[emp.ExcelRow.Value, 28].Value = emp.Gender;
                    ////worksheet.Cells[2, dc, rowCount + 2, dc].Style.Numberformat.Format = "mm/dd/yyyy hh:mm:ss AM/PM";
                    //worksheet.Cells[emp.ExcelRow.Value, 29].Style.Numberformat.Format = "mm/dd/yyyy";
                    //worksheet.Cells[emp.ExcelRow.Value, 29].Value = emp.DateOfBirth;
                    //worksheet.Cells[emp.ExcelRow.Value, 30].Value = emp.PlaceOfBirth;
                    // row++;
                }

                package.Save();
            }
        }

        public List< ExcelEmployee > FindPotentialDuplicates(List<Employee> excelEmployees)
        {
            var potentialDuplicates = new List<ExcelEmployee>();

            foreach (var excelEmp in excelEmployees)
            {
                var dbEmp = _context.Employee
                    .FromSqlRaw($"SELECT * FROM Employee WHERE SOUNDEX(FirstName+' '+LastName) = " +
                                $"SOUNDEX('{excelEmp.FirstName + ' ' + excelEmp.LastName}')")
                    .AsEnumerable()
                    .Where(e => e.CompareEmployee(excelEmp))
                    .ToList<Employee>();

                if (dbEmp != null)
                {                    
                    potentialDuplicates.Add(new ExcelEmployee
                    {
                        NumbersFromDB = dbEmp.Select(e => e.Number).ToArray(),
                        ExcelRow = int.Parse(excelEmp.Number),
                    });
                }

            }

            return potentialDuplicates;
        }

        //Save data loaded from HR employees excel file to MSSQL database
        public void SaveToDatabase(List<Employee> excelEmployees) 
        {
            _context.Employee.AddRange((IEnumerable<Employee>)excelEmployees);
            _context.SaveChanges();
        }

        //Delete Current HR employees Tables existing on MSSQL Local DB
        public void DeleteFromDatabase()
        {
            _context.Database.ExecuteSqlRaw("DELETE FROM Employee");
        }
    }
}
