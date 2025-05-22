using System;
using System.Collections.Generic;
using System.Threading.Tasks;
using YourCompany.ResourceManagement.Models;

namespace YourCompany.ResourceManagement.Repositories
{
    public interface IOnsiteEmployeeRepository
    {
        Task<IEnumerable<OnsiteEmployeeReportDto>> GetOnsiteEmployeesAsync(
            DateTime startDate, DateTime endDate, int empId);
    }
}