using Microsoft.AspNetCore.Hosting;
using Microsoft.AspNetCore.Mvc;
using Microsoft.Extensions.Logging;
using System;
using System.Data;
using System.Globalization;
using System.IO;
using System.Threading.Tasks;
using WebApplication13.Context;
using WebApplication13.Models.DTOs;
using WebApplication13.Services.Interfaces;

namespace WebApplication13.Controllers
{
    [ApiController]
    [Route("api/[controller]")]
    public class OnsiteEmployeeController : ControllerBase
    {
        private readonly IOnsiteEmployeeService _onsiteEmployeeService;
        private readonly ILogger<OnsiteEmployeeController> _logger;
        private readonly string _tempFilePath;

        public OnsiteEmployeeController(
            IOnsiteEmployeeService onsiteEmployeeService,
            ILogger<OnsiteEmployeeController> logger,
            IWebHostEnvironment environment)
        {
            _onsiteEmployeeService = onsiteEmployeeService;
            _logger = logger;
            _tempFilePath = Path.Combine(environment.ContentRootPath, "TempFiles");

            // Ensure temp directory exists
            if (!Directory.Exists(_tempFilePath))
            {
                Directory.CreateDirectory(_tempFilePath);
            }
        }

        [HttpGet]
        public async Task<IActionResult> GetOnsiteEmployees(
            [FromQuery] int empId,
            [FromQuery] string startDate = null,
            [FromQuery] string endDate = null)
        {
            try
            {
                // Parse dates in dd/mm/yyyy format
                (DateTime effectiveStartDate, DateTime effectiveEndDate) = ParseAndValidateDates(startDate, endDate);

                // Basic validation
                if (empId <= 0)
                {
                    return BadRequest("Invalid employee ID");
                }

                // Log request information
                _logger.LogInformation(
                    "Request received at {Timestamp} - Getting onsite employees for EmpId: {EmpId}",
                    DateTime.UtcNow.ToString("yyyy-MM-dd HH:mm:ss"),
                    empId);

                var result = await _onsiteEmployeeService.GetOnsiteEmployeesAsync(
                    effectiveStartDate, effectiveEndDate, empId);

                return Ok(result);
            }
            catch (FormatException ex)
            {
                _logger.LogError(ex, "Invalid date format");
                return BadRequest($"Invalid date format. Please use dd/MM/yyyy format (e.g., 01/01/2024)");
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Error retrieving onsite employees");
                return StatusCode(500, "An error occurred while processing your request");
            }
        }

        // New endpoint specifically for Excel download
        [HttpGet("Excel")]
        public async Task<IActionResult> GetOnsiteEmployeesExcel(
            [FromQuery] int empId,
            [FromQuery] string startDate = null,
            [FromQuery] string endDate = null)
        {
            try
            {
                // Parse dates in dd/mm/yyyy format
                (DateTime effectiveStartDate, DateTime effectiveEndDate) = ParseAndValidateDates(startDate, endDate);

                // Basic validation
                if (empId <= 0)
                {
                    return BadRequest("Invalid employee ID");
                }

                // Updated timestamp and user from your latest info
                string currentUser = "CoderVaibhav24";
                string timestamp = "2025-05-22 07:34:59";

                _logger.LogInformation(
                    "Excel download request at {Timestamp} by {User} for EmpId: {EmpId}",
                    timestamp,
                    currentUser,
                    empId);

                // Get data using service
                var data = await _onsiteEmployeeService.GetOnsiteEmployeesAsync(
                    effectiveStartDate, effectiveEndDate, empId);

                // Convert to DataTable
                DataTable dataTable = new DataTable("Onsite Employees");

                // Add columns
                dataTable.Columns.Add("Employee ID", typeof(int));
                dataTable.Columns.Add("Employee Name", typeof(string));
                dataTable.Columns.Add("Job Grade", typeof(string));
                dataTable.Columns.Add("Job Title", typeof(string));
                dataTable.Columns.Add("Onsite Start Date", typeof(string));
                dataTable.Columns.Add("Onsite End Date", typeof(string));
                dataTable.Columns.Add("Location", typeof(string));
                dataTable.Columns.Add("Status", typeof(string));
                dataTable.Columns.Add("Attendance Type ID", typeof(string));
                dataTable.Columns.Add("Feedback Score", typeof(string));
                dataTable.Columns.Add("Manager ID", typeof(string));
                dataTable.Columns.Add("Manager Name", typeof(string));
                dataTable.Columns.Add("Email ID", typeof(string));
                dataTable.Columns.Add("Gender", typeof(string));

                // Add rows
                foreach (var item in data)
                {
                    DataRow row = dataTable.NewRow();
                    row["Employee ID"] = item.EmpId;
                    row["Employee Name"] = item.EmpName;
                    row["Job Grade"] = item.JobGrade;
                    row["Job Title"] = item.JobTitle;
                    row["Onsite Start Date"] = item.OnsiteStartDate.ToString("dd/MM/yyyy");
                    row["Onsite End Date"] = item.OnsiteEndDate.ToString("dd/MM/yyyy");
                    row["Location"] = item.Location;
                    row["Status"] = item.Status;
                    row["Attendance Type ID"] = item.AttendanceTypeId?.ToString() ?? "";
                    row["Feedback Score"] = item.FeedbackScore?.ToString() ?? "";
                    row["Manager ID"] = item.ManagerId?.ToString() ?? "";
                    row["Manager Name"] = item.ManagerName;
                    row["Email ID"] = item.EmailId;
                    row["Gender"] = item.Gender;
                    dataTable.Rows.Add(row);
                }

                // Create filename
                string fileNameWithoutExtension = $"OnsiteEmployees_{DateTime.Now:yyyyMMdd_HHmmss}";
                string fileName = $"{fileNameWithoutExtension}.xlsx";
                string filePath = Path.Combine(_tempFilePath, fileName);

                // Use ExcelHelper to create the Excel file
                ExcelHelper.CreateExcelFromDataTable(dataTable, filePath);

                // Return file and schedule it for deletion after response is sent
                var fileBytes = System.IO.File.ReadAllBytes(filePath);
                Response.OnCompleted(() => {
                    if (System.IO.File.Exists(filePath))
                    {
                        try { System.IO.File.Delete(filePath); }
                        catch (Exception ex) { _logger.LogWarning(ex, "Failed to delete temp file"); }
                    }
                    return Task.CompletedTask;
                });

                return File(fileBytes, "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", fileName);
            }
            catch (FormatException ex)
            {
                _logger.LogError(ex, "Invalid date format");
                return BadRequest($"Invalid date format. Please use dd/MM/yyyy format (e.g., 01/01/2024)");
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Error creating Excel file");
                return StatusCode(500, "An error occurred while generating the Excel download");
            }
        }

        private (DateTime startDate, DateTime endDate) ParseAndValidateDates(string startDateStr, string endDateStr)
        {
            // Default date ranges if not provided
            DateTime now = DateTime.Now;
            DateTime defaultStartDate = new DateTime(now.Year, now.Month, 1).AddMonths(-1);
            DateTime defaultEndDate = new DateTime(now.Year, now.Month, 1).AddMonths(1).AddDays(-1);

            // Parse start date (dd/MM/yyyy format)
            DateTime effectiveStartDate = defaultStartDate;
            if (!string.IsNullOrEmpty(startDateStr))
            {
                if (!DateTime.TryParseExact(startDateStr, "dd/MM/yyyy", CultureInfo.InvariantCulture,
                    DateTimeStyles.None, out effectiveStartDate))
                {
                    throw new FormatException($"Start date '{startDateStr}' is not in the expected format 'dd/MM/yyyy'");
                }
            }

            // Parse end date (dd/MM/yyyy format)
            DateTime effectiveEndDate = defaultEndDate;
            if (!string.IsNullOrEmpty(endDateStr))
            {
                if (!DateTime.TryParseExact(endDateStr, "dd/MM/yyyy", CultureInfo.InvariantCulture,
                    DateTimeStyles.None, out effectiveEndDate))
                {
                    throw new FormatException($"End date '{endDateStr}' is not in the expected format 'dd/MM/yyyy'");
                }
            }

            // Validate date range
            if (effectiveStartDate > effectiveEndDate)
            {
                throw new ArgumentException("Start date cannot be after end date");
            }

            return (effectiveStartDate, effectiveEndDate);
        }
    }
}