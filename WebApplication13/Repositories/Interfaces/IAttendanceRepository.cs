using System;
using System.Collections.Generic;
using System.Threading.Tasks;
using WebApplication13.Models.DTOs;

namespace WebApplication13.Repositories.Interfaces
{
    public interface IAttendanceRepository
    {
        Task<IEnumerable<WFHReportDto>> GetWFHReportAsync(DateTime startDate, DateTime endDate);
    }
}