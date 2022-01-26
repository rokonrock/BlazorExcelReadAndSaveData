using BlazorExcelReadAndSaveData.Context;
using BlazorExcelReadAndSaveData.Data;
using BlazorExcelReadAndSaveData.IRepository;
using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Threading.Tasks;

namespace BlazorExcelReadAndSaveData.Repository
{
    public class EmployeeRepository : IEmployeeRepository
    {
        DatabaseContext _context;
        public EmployeeRepository(DatabaseContext context)
        {
            _context = context;
        }
        public List<Employee> SaveExcel(string fileName)
        {
            var FilePath = $"{Directory.GetCurrentDirectory()}{@"\wwwroot\files\"}" + "\\" + fileName;
            FileInfo fileInfo = new FileInfo(FilePath);
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
            using(ExcelPackage package = new ExcelPackage(fileInfo))
            {
                ExcelWorksheet worksheet = package.Workbook.Worksheets.FirstOrDefault();
                int colCount = worksheet.Dimension.End.Column;
                int rowCount = worksheet.Dimension.End.Row;
                for (int row = 2; row <= rowCount; row++)
                {
                    Employee emp = new Employee();
                    for (int col = 1; col <= colCount; col++)
                    {
                        if (col == 1) emp.FirstName = worksheet.Cells[row, col].Value.ToString();
                        if (col == 2) emp.LastName = worksheet.Cells[row, col].Value.ToString();
                    }

                    _context.Add(emp);
                    _context.SaveChanges();

                }
            }

            return _context.Employees.ToList();
        }
    }
}
