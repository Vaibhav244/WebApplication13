using Microsoft.Extensions.Logging;
using System;
using System.Collections.Generic;
using System.Threading.Tasks;
using YourCompany.ResourceManagement.Models;
using YourCompany.ResourceManagement.Repositories;

namespace YourCompany.ResourceManagement.Services
{
    public class OnsiteEmployeeService : IOnsiteEmployeeService
    {
        private readonly IOnsiteEmployeeRepository _onsiteEmployeeRepository;
        private readonly ILogger<OnsiteEmployeeService> _logger;

        public OnsiteEmployeeService(
            IOnsiteEmployeeRepository onsiteEmployeeRepository,
            ILogger<OnsiteEmployeeService> logger)
        {
            _onsiteEmployeeRepository = onsiteEmployeeRepository;
            _logger = logger;
        }

        public async Task<IEnumerable<OnsiteEmployeeReportDto>> GetOnsiteEmployeesAsync(
            DateTime startDate, DateTime endDate, int empId)
        {
            try
            {
                _logger.LogInformation("Service: Getting onsite employees for {EmpId} from {StartDate} to {EndDate}",
                    empId, startDate.ToString("yyyy-MM-dd"), endDate.ToString("yyyy-MM-dd"));

                return await _onsiteEmployeeRepository.GetOnsiteEmployeesAsync(startDate, endDate, empId);
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Service error retrieving onsite employees for {EmpId}", empId);
                throw;
            }
        }
    }
}