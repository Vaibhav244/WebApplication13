using System;
using System.Collections.Generic;
using System.Threading.Tasks;
using WebApplication13.Models.Entities;
using WebApplication13.Models.DTOs;

namespace WebApplication13.Repositories.Interfaces
{
    public interface IEmployeeRepository
    {
        Task<(bool IsValid, bool IsActive, bool IsSuperUser)> ValidateEmployeeAsync(int empId);
        Task<IEnumerable<WorkFromOfficeReportDto>> GetWorkFromOfficeReportAsync(
            DateTime startDate,
            DateTime endDate,
            int empId,
            bool superUser,
            bool inactive);
        Task<bool> GetEmployeeSuperUserStatusAsync(int empId);
    }
}