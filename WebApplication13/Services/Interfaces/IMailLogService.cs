using System;
using System.Threading.Tasks;

namespace WebApplication13.Services.Interfaces
{
    public interface IMailLogService
    {
        Task<byte[]> GenerateEscalatedEmailLogExcelAsync(string templateName, DateTime startDate, DateTime endDate);
    }
}