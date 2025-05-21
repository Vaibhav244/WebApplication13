using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Threading.Tasks;
using Microsoft.Extensions.Configuration;
using Microsoft.Extensions.Logging;
using Newtonsoft.Json;
using Oracle.ManagedDataAccess.Client;
using WebApplication13.Models.DTOs;
using WebApplication13.Repositories.Interfaces;

namespace WebApplication13.Repositories.Implementations
{
    public class AttendanceRepository : IAttendanceRepository
    {
        private readonly IConfiguration _configuration;
        private readonly ILogger<AttendanceRepository> _logger;

        public AttendanceRepository(IConfiguration configuration, ILogger<AttendanceRepository> logger)
        {
            _configuration = configuration;
            _logger = logger;
        }

        public async Task<IEnumerable<WFHReportDto>> GetWFHReportAsync(DateTime startDate, DateTime endDate)
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
                    command.CommandText = @"SELECT adetails.emp_id, EDETAILS.EMP_NAME, adetails.start_date, adetails.end_date, adetails.status, 
                                    adetails.Comments, adetails.applied_date, adetails.processed_date, adetails.approver_id, 
                                    CASE WHEN approver_id is NOT NULL THEN (SELECT emp_name FROM ts2_employee_Details WHERE emp_id = approver_id) END as ApproverName, 
                                    adetails.APPROVAL_REJECTION_REASON 
                                    FROM RM_ATTENDANCE_DETAILS adetails 
                                    INNER JOIN ts2_employee_details edetails ON adetails.emp_id = edetails.emp_id 
                                    WHERE ATTENDANCE_TYPE_ID = 2 AND LOWER(edetails.cuid) NOT LIKE 'x%' 
                                    AND START_DATE >= To_date(:startDate, 'MM/DD/YYYY') AND END_DATE <= To_date(:endDate, 'MM/DD/YYYY') 
                                    ORDER BY EMP_ID, START_DATE";

                    command.Parameters.Add(new OracleParameter("startDate", startDate.ToShortDateString()));
                    command.Parameters.Add(new OracleParameter("endDate", endDate.ToShortDateString()));
                    command.CommandType = CommandType.Text;

                    adapter.Fill(dataSet);
                }

                if (dataSet.Tables.Count == 0 || dataSet.Tables[0].Rows.Count == 0)
                {
                    return Enumerable.Empty<WFHReportDto>();
                }

                // Convert DataTable to list of DTOs
                var dataTable = dataSet.Tables[0];
                var result = dataTable.AsEnumerable().Select(row => new WFHReportDto
                {
                    EmpId = row["emp_id"] != DBNull.Value ? Convert.ToInt32(row["emp_id"]) : (int?)null,
                    EmpName = row["EMP_NAME"] != DBNull.Value ? row["EMP_NAME"].ToString() : null,
                    StartDate = row["start_date"] != DBNull.Value ? Convert.ToDateTime(row["start_date"]).ToString("yyyy-MM-dd") : null,
                    EndDate = row["end_date"] != DBNull.Value ? Convert.ToDateTime(row["end_date"]).ToString("yyyy-MM-dd") : null,
                    Status = row["status"] != DBNull.Value ? row["status"].ToString() : null,
                    Comments = row["Comments"] != DBNull.Value ? row["Comments"].ToString() : null,
                    AppliedDate = row["applied_date"] != DBNull.Value ? Convert.ToDateTime(row["applied_date"]).ToString("yyyy-MM-dd") : null,
                    ProcessedDate = row["processed_date"] != DBNull.Value ? Convert.ToDateTime(row["processed_date"]).ToString("yyyy-MM-dd") : null,
                    ApproverId = row["approver_id"] != DBNull.Value ? Convert.ToInt32(row["approver_id"]) : (int?)null,
                    ApproverName = row["ApproverName"] != DBNull.Value ? row["ApproverName"].ToString() : null,
                    ApprovalRejectionReason = row["APPROVAL_REJECTION_REASON"] != DBNull.Value ? row["APPROVAL_REJECTION_REASON"].ToString() : null
                }).ToList();

                return result;
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Error retrieving WFH report");
                throw;
            }
        }
    }
}