using Microsoft.Extensions.Configuration;
using Microsoft.Extensions.Logging;
using Oracle.ManagedDataAccess.Client;
using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Threading.Tasks;
using YourCompany.ResourceManagement.Models;

namespace YourCompany.ResourceManagement.Repositories
{
    public class OnsiteEmployeeRepository : IOnsiteEmployeeRepository
    {
        private readonly IConfiguration _configuration;
        private readonly ILogger<OnsiteEmployeeRepository> _logger;

        public OnsiteEmployeeRepository(IConfiguration configuration, ILogger<OnsiteEmployeeRepository> logger)
        {
            _configuration = configuration;
            _logger = logger;
        }

        public async Task<IEnumerable<OnsiteEmployeeReportDto>> GetOnsiteEmployeesAsync(
            DateTime startDate, DateTime endDate, int empId)
        {
            try
            {
                // Format dates as strings in Oracle's preferred format
                string formattedStartDate = startDate.ToString("yyyy-MM-dd");
                string formattedEndDate = endDate.ToString("yyyy-MM-dd");

                _logger.LogInformation("Querying for employee {EmpId} from {StartDate} to {EndDate}",
                    empId, formattedStartDate, formattedEndDate);

                // 1. Determine employee's access level first
                int accessLevel = await DetermineAccessLevelAsync(empId);

                _logger.LogInformation("User {EmpId} access level determined as {Level}", empId, accessLevel);

                // 2. Build hierarchy filter based on access level
                string hierarchyFilter = BuildHierarchyFilter(accessLevel, empId);

                // 3. Execute query with both access control and date filtering
                return await ExecuteOnsiteEmployeeQuery(hierarchyFilter, formattedStartDate, formattedEndDate, empId);
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Error retrieving onsite employees for {EmpId}", empId);
                throw;
            }
        }

        private async Task<int> DetermineAccessLevelAsync(int empId)
        {
            try
            {
                using (var connection = new OracleConnection(_configuration.GetConnectionString("DefaultConnection")))
                {
                    await connection.OpenAsync();
                    using (var command = connection.CreateCommand())
                    {
                        command.CommandText = @"
                            SELECT 
                                CASE
                                    -- Admin check (highest priority)
                                    WHEN EXISTS (SELECT 1 FROM RM_USER_ROLE WHERE EMP_ID = ed.EMP_ID AND ROLE_ID = 1) THEN 5
                                    
                                    -- Director level (job level > 9)
                                    WHEN ed.EMP_LEVEL IS NOT NULL AND TO_NUMBER(ed.EMP_LEVEL) > 9 THEN 4
                                    
                                    -- Manager level (job level > 6 and <= 9)
                                    WHEN ed.EMP_LEVEL IS NOT NULL AND TO_NUMBER(ed.EMP_LEVEL) > 6 AND TO_NUMBER(ed.EMP_LEVEL) <= 9 THEN 3
                                    
                                    -- Check if they have direct reports (should be at least manager level)
                                    WHEN EXISTS (SELECT 1 FROM TS2_EMPLOYEE_DETAILS WHERE MANAGER_ID = ed.EMP_ID) THEN 3
                                    
                                    -- Regular employee (everyone else)
                                    ELSE 1
                                END AS ACCESS_LEVEL
                            FROM 
                                TS2_EMPLOYEE_DETAILS ed
                            WHERE 
                                ed.EMP_ID = :1";

                        command.Parameters.Add(new OracleParameter { Value = empId });
                        command.CommandType = CommandType.Text;

                        var result = await command.ExecuteScalarAsync();
                        return result != null && result != DBNull.Value ? Convert.ToInt32(result) : 1;
                    }
                }
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Error determining access level for {EmpId}", empId);
                return 1; // Default to regular employee on error
            }
        }

        private string BuildHierarchyFilter(int accessLevel, int empId)
        {
            switch (accessLevel)
            {
                case 5: // Admin level
                case 4: // Director level
                    // No restrictions - can see all records
                    return "1=1";

                case 3: // Manager level
                    // Can see self, direct reports, and indirect reports
                    return @"
                        ed.EMP_ID = :emp_id 
                        OR ed.MANAGER_ID = :emp_id
                        OR EXISTS (
                            SELECT 1 
                            FROM TS2_EMPLOYEE_DETAILS e2 
                            WHERE e2.MANAGER_ID = :emp_id 
                            AND ed.MANAGER_ID = e2.EMP_ID
                        )";

                case 1: // Regular employee
                default:
                    // Can only see themselves
                    return "ed.EMP_ID = :emp_id";
            }
        }

        private async Task<IEnumerable<OnsiteEmployeeReportDto>> ExecuteOnsiteEmployeeQuery(
            string hierarchyFilter, string startDate, string endDate, int empId)
        {
            string sql = $@"
            SELECT 
                oe.EMP_ID, 
                ed.EMP_NAME, 
                ed.JOB_GRADE,
                ed.JOB_TITLE_FULL as JOB_TITLE,
                oe.ONSITE_START_DATE, 
                oe.ONSITE_END_DATE, 
                ed.CITY as LOCATION,
                ed.EMP_STATUS as STATUS,
                oe.ATTENDANCE_TYPE_ID,
                oe.FEEDBACKSCORE,
                ed.MANAGER_ID,
                (SELECT e.EMP_NAME FROM TS2_EMPLOYEE_DETAILS e WHERE e.EMP_ID = ed.MANAGER_ID) as MANAGER_NAME,
                ed.EMAIL_ID_OFF,
                ed.GENDER
            FROM 
                RM_ONSITE_EMPLOYEES oe
            JOIN 
                TS2_EMPLOYEE_DETAILS ed ON oe.EMP_ID = ed.EMP_ID
            WHERE 
                ({hierarchyFilter})
            AND 
                (oe.ONSITE_START_DATE <= TO_DATE(:end_date, 'YYYY-MM-DD') 
                 AND oe.ONSITE_END_DATE >= TO_DATE(:start_date, 'YYYY-MM-DD'))
            ORDER BY oe.ONSITE_START_DATE DESC";

            var dataSet = new DataSet();

            using (var connection = new OracleConnection(_configuration.GetConnectionString("DefaultConnection")))
            {
                await connection.OpenAsync();

                using (var command = new OracleCommand(sql, connection))
                {
                    command.BindByName = true;
                    command.Parameters.Add(new OracleParameter("emp_id", OracleDbType.Int32) { Value = empId });
                    command.Parameters.Add(new OracleParameter("start_date", OracleDbType.Varchar2) { Value = startDate });
                    command.Parameters.Add(new OracleParameter("end_date", OracleDbType.Varchar2) { Value = endDate });

                    _logger.LogInformation("Executing query with parameters: emp_id={EmpId}, start_date={StartDate}, end_date={EndDate}",
                        empId, startDate, endDate);

                    using (var adapter = new OracleDataAdapter(command))
                    {
                        adapter.Fill(dataSet);
                        _logger.LogInformation("Retrieved {Count} records",
                            dataSet.Tables.Count > 0 ? dataSet.Tables[0].Rows.Count : 0);
                    }
                }
            }

            // Process results into DTOs
            if (dataSet.Tables.Count == 0 || dataSet.Tables[0].Rows.Count == 0)
            {
                return Enumerable.Empty<OnsiteEmployeeReportDto>();
            }

            var result = new List<OnsiteEmployeeReportDto>();
            foreach (DataRow row in dataSet.Tables[0].Rows)
            {
                result.Add(new OnsiteEmployeeReportDto
                {
                    EmpId = Convert.ToInt32(row["EMP_ID"]),
                    EmpName = row["EMP_NAME"].ToString(),
                    JobGrade = row["JOB_GRADE"].ToString(),
                    JobTitle = row["JOB_TITLE"].ToString(),
                    OnsiteStartDate = Convert.ToDateTime(row["ONSITE_START_DATE"]),
                    OnsiteEndDate = Convert.ToDateTime(row["ONSITE_END_DATE"]),
                    Location = row["LOCATION"].ToString(),
                    Status = row["STATUS"].ToString(),
                    AttendanceTypeId = row["ATTENDANCE_TYPE_ID"] != DBNull.Value ? Convert.ToInt32(row["ATTENDANCE_TYPE_ID"]) : null,
                    FeedbackScore = row["FEEDBACKSCORE"] != DBNull.Value ? Convert.ToDecimal(row["FEEDBACKSCORE"]) : null,
                    ManagerId = row["MANAGER_ID"] != DBNull.Value ? Convert.ToInt32(row["MANAGER_ID"]) : null,
                    ManagerName = row["MANAGER_NAME"].ToString(),
                    EmailId = row["EMAIL_ID_OFF"].ToString(),
                    Gender = row["GENDER"].ToString()
                });
            }

            return result;
        }
    }
}