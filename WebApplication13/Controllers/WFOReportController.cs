using System;
using System.Collections.Generic;
using System.Threading.Tasks;
using Microsoft.AspNetCore.Mvc;
using Microsoft.Extensions.Logging;
using WebApplication13.Models.DTOs;
using WebApplication13.Services.Interfaces;

namespace WebApplication13.Controllers
{
    [Route("api/")]
    [ApiController]
    public class WFOReportController : ControllerBase
    {
        private readonly ILogger<WFOReportController> _logger;
        private readonly IWFOReportService _wfoReportService;
        private readonly IWFHReportService _wfhReportService;
        private readonly IMailLogService _mailLogService;

        public WFOReportController(
            ILogger<WFOReportController> logger,
            IWFOReportService wfoReportService,
            IWFHReportService wfhReportService,
            IMailLogService mailLogService)
        {
            _logger = logger;
            _wfoReportService = wfoReportService;
            _wfhReportService = wfhReportService;
            _mailLogService = mailLogService;
        }

        [HttpGet("GetWorkFromOfficeReport")]
        public async Task<IActionResult> GetWorkFromOfficeReport(
            [FromQuery] DateTime startDate,
            [FromQuery] DateTime endDate,
            [FromQuery] int empId,
            [FromQuery] bool superUser,
            [FromQuery] bool inactive)
        {
            try
            {
                byte[] excelData = await _wfoReportService.GenerateWorkFromOfficeReportExcelAsync(
                    startDate, endDate, empId, superUser, inactive);

                string fileName = $"wforeport_{DateTime.Now:yyyyMMdd_HHmmss}.xlsx";
                return File(excelData, "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", fileName);
            }
            catch (ArgumentException ex)
            {
                _logger.LogWarning(ex, "Invalid argument in GetWorkFromOfficeReport");
                return StatusCode(200, ex.Message);
            }
            catch (InvalidOperationException ex)
            {
                _logger.LogWarning(ex, "Invalid operation in GetWorkFromOfficeReport");
                return StatusCode(200, ex.Message);
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Error in GetWorkFromOfficeReport");
                return StatusCode(500, "Internal server error");
            }
        }

        [HttpGet("GetWorkFromOfficeReport2")]
        public async Task<IActionResult> GetWorkFromOfficeReport2(
            [FromQuery] DateTime startDate,
            [FromQuery] DateTime endDate,
            [FromQuery] int empId,
            [FromQuery] bool superUser,
            [FromQuery] bool inactive)
        {
            try
            {
                var data = await _wfoReportService.GetWorkFromOfficeReportDataAsync(
                    startDate, endDate, empId, superUser, inactive);

                return Ok(new
                {
                    status = "success",
                    data = data
                });
            }
            catch (ArgumentException ex)
            {
                _logger.LogWarning(ex, "Invalid argument in GetWorkFromOfficeReport2");
                return StatusCode(200, ex.Message);
            }
            catch (InvalidOperationException ex)
            {
                _logger.LogWarning(ex, "Invalid operation in GetWorkFromOfficeReport2");
                return StatusCode(200, ex.Message);
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Error in GetWorkFromOfficeReport2");
                return StatusCode(500, new { message = "Internal server error" });
            }
        }

        [HttpGet("WFHReport")]
        public async Task<IActionResult> GetWFHReport(
            [FromQuery] DateTime startDate,
            [FromQuery] DateTime endDate)
        {
            try
            {
                byte[] excelData = await _wfhReportService.GenerateWFHReportExcelAsync(startDate, endDate);

                string fileName = $"WFHReport_{DateTime.Now:yyyyMMdd_HHmmss}.xlsx";
                return File(excelData, "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", fileName);
            }
            catch (InvalidOperationException ex)
            {
                _logger.LogWarning(ex, "Invalid operation in GetWFHReport");
                return StatusCode(200, ex.Message);
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Error in GetWFHReport");
                return StatusCode(500, "Internal server error");
            }
        }

        [HttpGet("WFHReport2")]
        public async Task<IActionResult> GetWFHReport2(
            [FromQuery] DateTime startDate,
            [FromQuery] DateTime endDate)
        {
            try
            {
                var data = await _wfhReportService.GetWFHReportDataAsync(startDate, endDate);

                return Ok(new
                {
                    status = "success",
                    data = data
                });
            }
            catch (InvalidOperationException ex)
            {
                _logger.LogWarning(ex, "Invalid operation in GetWFHReport2");
                return StatusCode(200, ex.Message);
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Error in GetWFHReport2");
                return StatusCode(500, new { message = "Internal server error" });
            }
        }

        [HttpGet("EscalatedEmailLog")]
        public async Task<IActionResult> GetEscalatedEmailLog(
            [FromQuery] string escalationEmailTemplate,
            [FromQuery] DateTime startDate,
            [FromQuery] DateTime endDate)
        {
            try
            {
                byte[] excelData = await _mailLogService.GenerateEscalatedEmailLogExcelAsync(
                    escalationEmailTemplate, startDate, endDate);

                string fileName = $"EscalatedEmailLog_{DateTime.Now:yyyyMMdd_HHmmss}.xlsx";
                return File(excelData, "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", fileName);
            }
            catch (InvalidOperationException ex)
            {
                _logger.LogWarning(ex, "Invalid operation in GetEscalatedEmailLog");
                return StatusCode(200, ex.Message);
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Error in GetEscalatedEmailLog");
                return StatusCode(500, "Internal server error");
            }
        }
    }
}