using System;
using System.IO;
using System.Threading.Tasks;
using Microsoft.Extensions.Logging;
using Newtonsoft.Json;
using WebApplication13.Repositories.Interfaces;
using WebApplication13.Services.Interfaces;

namespace WebApplication13.Services.Implementations
{
    public class MailLogService : IMailLogService
    {
        private readonly IMailLogRepository _mailLogRepository;
        private readonly ILogger<MailLogService> _logger;

        public MailLogService(IMailLogRepository mailLogRepository, ILogger<MailLogService> logger)
        {
            _mailLogRepository = mailLogRepository;
            _logger = logger;
        }

        public async Task<byte[]> GenerateEscalatedEmailLogExcelAsync(string templateName, DateTime startDate, DateTime endDate)
        {
            try
            {
                var dataTable = await _mailLogRepository.GetEscalatedEmailLogAsync(templateName, startDate, endDate);

                if (dataTable == null || dataTable.Rows.Count == 0)
                {
                    throw new InvalidOperationException("No data available for the given template and date range");
                }

                // Generate Excel file
                string currentDateTime = DateTime.Now.ToString("yyyyMMdd_HHmmss");
                string filePath = $"EscalatedEmailLog_{currentDateTime}.xlsx";
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
                _logger.LogError(ex, "Error generating escalated email log Excel for template {TemplateName}", templateName);
                throw;
            }
        }
    }
}