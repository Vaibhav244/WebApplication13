using System;
using System.Collections.Generic;
using System.Data;
using System.Threading.Tasks;
using WebApplication13.Models.Entities;

namespace WebApplication13.Repositories.Interfaces
{
    public interface IMailLogRepository
    {
        Task<DataTable> GetEscalatedEmailLogAsync(string templateName, DateTime startDate, DateTime endDate);
    }
}