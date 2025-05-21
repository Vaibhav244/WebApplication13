using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Threading.Tasks;
using Microsoft.Extensions.Logging;
using Newtonsoft.Json;
using WebApplication13.Models.DTOs;
using WebApplication13.Repositories.Interfaces;
using WebApplication13.Services.Interfaces;

namespace WebApplication13.Services.Implementations
{
    public class WFOReportService : IWFOReportService
    {
        private readonly IEmployeeRepository _employeeRepository;
        private readonly ILogger<WFOReportService> _logger;

        public WFOReportService(IEmployeeRepository employeeRepository, ILogger<WFOReportService> logger)
        {
            _employeeRepository = employeeRepository;
            _logger = logger;
        }

        public async Task<IEnumerable<WorkFromOfficeReportDto>> GetWorkFromOfficeReportDataAsync(
            DateTime startDate,
            DateTime endDate,
            int empId,
            bool superUser,
            bool inactive)
        {
            try
            {
                // Validate empId and check if the employee is active
                var (isValidEmpId, isActiveEmp, isSuperUser) = await _employeeRepository.ValidateEmployeeAsync(empId);

                if (!isValidEmpId)
                {
                    throw new ArgumentException($"Invalid empId: {empId}");
                }

                if (!isActiveEmp && inactive)
                {
                    throw new ArgumentException($"Inactive empId: {empId}");
                }

                if (superUser || isSuperUser)
                {
                    if (!isSuperUser)
                    {
                        throw new ArgumentException($"Not a valid super user empId: {empId}");
                    }
                }

                return await _employeeRepository.GetWorkFromOfficeReportAsync(startDate, endDate, empId, superUser, inactive);
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Error retrieving work from office report for employee {EmpId}", empId);
                throw;
            }
        }

        public async Task<byte[]> GenerateWorkFromOfficeReportExcelAsync(
            DateTime startDate,
            DateTime endDate,
            int empId,
            bool superUser,
            bool inactive)
        {
            try
            {
                var reportData = await GetWorkFromOfficeReportDataAsync(startDate, endDate, empId, superUser, inactive);

                if (reportData == null || !reportData.Any())
                {
                    throw new InvalidOperationException($"No data available for empId: {empId}");
                }

                // Convert to DataTable
                string json = JsonConvert.SerializeObject(reportData);
                DataTable dataTable = JsonConvert.DeserializeObject<DataTable>(json);

                // Generate Excel file
                string currentDateTime = DateTime.Now.ToString("yyyyMMdd_HHmmss");
                string filePath = $"wforeport_{currentDateTime}.xlsx";
                ExcelHelper.CreateExcelFromDataTable(dataTable, filePath);

                byte[] content = File.ReadAllBytes(filePath);

                // Clean up file after reading
                try
                {
                    File.Delete(filePath);
                }
                catch
                {
                    // Log but don't throw if cleanup fails
                    _logger.LogWarning("Failed to delete temporary Excel file: {FilePath}", filePath);
                }

                return content;
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Error generating work from office report Excel for employee {EmpId}", empId);
                throw;
            }
        }
    }
}