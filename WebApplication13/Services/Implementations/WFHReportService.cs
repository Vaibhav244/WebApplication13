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
    public class WFHReportService : IWFHReportService
    {
        private readonly IAttendanceRepository _attendanceRepository;
        private readonly ILogger<WFHReportService> _logger;

        public WFHReportService(IAttendanceRepository attendanceRepository, ILogger<WFHReportService> logger)
        {
            _attendanceRepository = attendanceRepository;
            _logger = logger;
        }

        public async Task<IEnumerable<WFHReportDto>> GetWFHReportDataAsync(DateTime startDate, DateTime endDate)
        {
            try
            {
                return await _attendanceRepository.GetWFHReportAsync(startDate, endDate);
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Error retrieving WFH report");
                throw;
            }
        }

        public async Task<byte[]> GenerateWFHReportExcelAsync(DateTime startDate, DateTime endDate)
        {
            try
            {
                var reportData = await GetWFHReportDataAsync(startDate, endDate);

                if (reportData == null || !reportData.Any())
                {
                    throw new InvalidOperationException("No data available for the given date range");
                }

                // Convert to DataTable
                string json = JsonConvert.SerializeObject(reportData);
                DataTable dataTable = JsonConvert.DeserializeObject<DataTable>(json);

                // Generate Excel file
                string currentDateTime = DateTime.Now.ToString("yyyyMMdd_HHmmss");
                string filePath = $"WFHReport_{currentDateTime}.xlsx";
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
                _logger.LogError(ex, "Error generating WFH report Excel");
                throw;
            }
        }
    }
}