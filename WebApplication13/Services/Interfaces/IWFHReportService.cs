using System;
using System.Data;
using System.Threading.Tasks;
using WebApplication13.Models.DTOs;
using System.Collections.Generic;

namespace WebApplication13.Services.Interfaces
{
    public interface IWFHReportService
    {
        Task<byte[]> GenerateWFHReportExcelAsync(DateTime startDate, DateTime endDate);
        Task<IEnumerable<WFHReportDto>> GetWFHReportDataAsync(DateTime startDate, DateTime endDate);
    }
}