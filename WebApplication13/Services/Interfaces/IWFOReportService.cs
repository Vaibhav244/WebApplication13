using System;
using System.Data;
using System.Threading.Tasks;
using WebApplication13.Models.DTOs;
using System.Collections.Generic;

namespace WebApplication13.Services.Interfaces
{
    public interface IWFOReportService
    {
        Task<byte[]> GenerateWorkFromOfficeReportExcelAsync(DateTime startDate, DateTime endDate, int empId, bool superUser, bool inactive);
        Task<IEnumerable<WorkFromOfficeReportDto>> GetWorkFromOfficeReportDataAsync(DateTime startDate, DateTime endDate, int empId, bool superUser, bool inactive);
    }
}