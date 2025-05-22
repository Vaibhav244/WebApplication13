using System;
using System.Collections.Generic;
using System.Threading.Tasks;
using WebApplication13.Models.DTOs;

namespace WebApplication13.Repositories.Interfaces
{
    public interface IOnsiteEmployeeRepository
    {
        Task<IEnumerable<OnsiteEmployeeReportDto>> GetOnsiteEmployeesAsync(
            DateTime startDate, DateTime endDate, int empId);
    }
}