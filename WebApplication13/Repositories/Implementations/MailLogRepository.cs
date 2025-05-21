using System;
using System.Data;
using System.Threading.Tasks;
using Microsoft.Extensions.Configuration;
using Microsoft.Extensions.Logging;
using Oracle.ManagedDataAccess.Client;
using WebApplication13.Repositories.Interfaces;

namespace WebApplication13.Repositories.Implementations
{
    public class MailLogRepository : IMailLogRepository
    {
        private readonly IConfiguration _configuration;
        private readonly ILogger<MailLogRepository> _logger;

        public MailLogRepository(IConfiguration configuration, ILogger<MailLogRepository> logger)
        {
            _configuration = configuration;
            _logger = logger;
        }

        public async Task<DataTable> GetEscalatedEmailLogAsync(string templateName, DateTime startDate, DateTime endDate)
        {
            try
            {
                var dataSet = new DataSet();
                using (var connection = new OracleConnection(_configuration.GetConnectionString("DefaultConnection")))
                using (var command = new OracleCommand())
                using (var adapter = new OracleDataAdapter(command))
                {
                    await connection.OpenAsync();
                    command.Connection = connection;
                    command.CommandText = @"SELECT Emp_ID, Mail_To, Mail_CC, updatedON 
                                    FROM PEMS_MAIL_LOG_NEW_JOINEE 
                                    WHERE template_name = :templateName 
                                    AND updatedON BETWEEN :startDate AND :endDate";
                    command.Parameters.Add(new OracleParameter("templateName", templateName));
                    command.Parameters.Add(new OracleParameter("startDate", startDate));
                    command.Parameters.Add(new OracleParameter("endDate", endDate));
                    command.CommandType = CommandType.Text;

                    adapter.Fill(dataSet);
                }

                if (dataSet.Tables.Count == 0 || dataSet.Tables[0].Rows.Count == 0)
                {
                    return new DataTable();
                }

                return dataSet.Tables[0];
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Error retrieving escalated email log for template {TemplateName}", templateName);
                throw;
            }
        }
    }
}