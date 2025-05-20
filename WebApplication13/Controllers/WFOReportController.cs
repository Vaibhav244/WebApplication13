using Microsoft.AspNetCore.Http;
using Microsoft.AspNetCore.Mvc;
using Microsoft.EntityFrameworkCore;
using Newtonsoft.Json;
using NPOI.SS.Formula.Functions;
using Oracle.ManagedDataAccess.Client;
using Org.BouncyCastle.Ocsp;
using System.Collections.Generic;
using System.Data;
using WebApplication13.Context;
using static Microsoft.EntityFrameworkCore.DbLoggerCategory.Database;

namespace WebApplication13.Controllers
{
    [Route("api/")]
    [ApiController]
    public class WFOReportController : ControllerBase
    {
        private readonly ILogger<WFOReportController> _logger;
        private readonly IConfiguration _configuration;
        private readonly MyDbContext _context;
        public WFOReportController(ILogger<WFOReportController> logger, IConfiguration configuration,MyDbContext context)
        {
            _logger = logger;
            _configuration = configuration;
            _context = context;
        }

        [HttpGet("GetWorkFromOfficeReport")]
        public async Task<IActionResult> GetWorkFromOfficeReport(
            [FromQuery] DateTime startDate,
            [FromQuery] DateTime endDate,
            [FromQuery] int empId,
            [FromQuery] bool superUser,
            [FromQuery] bool inactive)
        {
            try
            {
                // Validate empId and check if the employee is active
                bool isValidEmpId = false;
                bool isActiveEmp = false;
                bool isSuperUser = false;
                using (var connection = new OracleConnection(_configuration.GetConnectionString("DefaultConnection")))
                {
                    connection.Open();
                    using (var command = connection.CreateCommand())
                    {
                        command.CommandText = $"SELECT emp_status, emp_level FROM ts2_employee_details WHERE emp_id = {empId}";
                        command.CommandType = CommandType.Text;

                        using (var reader = command.ExecuteReader())
                        {
                            if (reader.Read())
                            {
                                isValidEmpId = true;
                                isActiveEmp = reader["emp_status"].ToString() == "ACTIVE";
                                isSuperUser = Convert.ToInt32(reader["emp_level"]) >= 5;
                            }
                        }
                    }
                }

                if (!isValidEmpId)
                {
                    return StatusCode(200, "Invalid empId: " + empId);
                }

                if (!isActiveEmp && inactive)
                {
                    return StatusCode(200, "Inactive empId: " + empId);
                }

                var endDateInclusive = endDate.AddDays(1);
                string filterLoggedInEmp = string.Empty;
                string filterInActiveEmp = string.Empty;

                if (superUser || isSuperUser)
                {
                    if (!isSuperUser)
                    {
                        return StatusCode(200, "Not a valid super user empId: " + empId);
                    }
                }
                else
                {
                    filterLoggedInEmp = $" AND t1.EMP_ID IN (SELECT tbl.EMP_ID FROM ts2_employee_details tbl START WITH tbl.emp_id = {empId} CONNECT BY NOCYCLE PRIOR tbl.emp_id = tbl.manager_id)";
                }
                if (!superUser)
                {
                    filterLoggedInEmp = $" AND t1.EMP_ID IN (SELECT tbl.EMP_ID FROM ts2_employee_details tbl START WITH tbl.emp_id = {empId} CONNECT BY NOCYCLE PRIOR tbl.emp_id = tbl.manager_id)";
                }

                if (inactive && !isActiveEmp)
                {
                    filterInActiveEmp = $" AND (t1.termination_date IS NULL OR " +
                        $"(t1.emp_status = 'INACTIVE' AND t1.termination_date >= TO_DATE('{startDate:MM/dd/yyyy}', 'mm/dd/yyyy') " +
                        $"AND t1.termination_date <= TO_DATE('{endDateInclusive:MM/dd/yyyy}', 'mm/dd/yyyy') AND TRUNC(access_time) <= t1.termination_date) " +
                        $"OR (t1.termination_date >= TO_DATE('{endDateInclusive:MM/dd/yyyy}', 'mm/dd/yyyy')))";
                }
                else
                {
                    filterInActiveEmp = " AND t1.emp_status = 'ACTIVE'";
                }

                string query = GenerateSwipeQuery(startDate, endDateInclusive, filterInActiveEmp, filterLoggedInEmp);

                var dataSet = new DataSet();
                using (var connection = new OracleConnection(_configuration.GetConnectionString("DefaultConnection")))
                using (var command = new OracleCommand(query, connection))
                using (var adapter = new OracleDataAdapter(command))
                {
                    await connection.OpenAsync();
                    adapter.Fill(dataSet);
                }

                // Check if the dataset is empty
                if (dataSet.Tables.Count == 0 || dataSet.Tables[0].Rows.Count == 0)
                {
                    return StatusCode(200, "No data available for empId: " + empId);
                }

                string json = JsonConvert.SerializeObject(dataSet.Tables[0]);

                // Convert JSON to DataTable
                DataTable dataTable = JsonConvert.DeserializeObject<DataTable>(json);

                // Generate Excel file with current date and time in the filename
                string currentDateTime = DateTime.Now.ToString("yyyyMMdd_HHmmss");
                string filePath = $"wforeport_{currentDateTime}.xlsx";
                ExcelHelper.CreateExcelFromDataTable(dataTable, filePath);

                byte[] content = System.IO.File.ReadAllBytes(filePath);
                return File(content, "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", filePath);

                return Ok(json);
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Error retrieving employee swipe data.");
                return StatusCode(500, "Internal server error");
            }
        }






        [HttpGet("GetWorkFromOfficeReport2")]
        public async Task<IActionResult> GetWorkFromOfficeReport2(
            [FromQuery] DateTime startDate,
            [FromQuery] DateTime endDate,
            [FromQuery] int empId,
            [FromQuery] bool superUser,
            [FromQuery] bool inactive)
        {
            try
            {
                // Validate empId and check if the employee is active
                bool isValidEmpId = false;
                bool isActiveEmp = false;
                bool isSuperUser = false;
                using (var connection = new OracleConnection(_configuration.GetConnectionString("DefaultConnection")))
                {
                    connection.Open();
                    using (var command = connection.CreateCommand())
                    {
                        command.CommandText = $"SELECT emp_status, emp_level FROM ts2_employee_details WHERE emp_id = {empId}";
                        command.CommandType = CommandType.Text;

                        using (var reader = command.ExecuteReader())
                        {
                            if (reader.Read())
                            {
                                isValidEmpId = true;
                                isActiveEmp = reader["emp_status"].ToString() == "ACTIVE";
                                isSuperUser = Convert.ToInt32(reader["emp_level"]) >= 5;
                            }
                        }
                    }
                }

                if (!isValidEmpId)
                {
                    return StatusCode(200, "Invalid empId: " + empId);
                }

                if (!isActiveEmp && inactive)
                {
                    return StatusCode(200, "Inactive empId: " + empId);
                }

                var endDateInclusive = endDate.AddDays(1);
                string filterLoggedInEmp = string.Empty;
                string filterInActiveEmp = string.Empty;

                if (superUser || isSuperUser)
                {
                    if (!isSuperUser)
                    {
                        return StatusCode(200, "Not a valid super user empId: " + empId);
                    }
                }
                else
                {
                    filterLoggedInEmp = $" AND t1.EMP_ID IN (SELECT tbl.EMP_ID FROM ts2_employee_details tbl START WITH tbl.emp_id = {empId} CONNECT BY NOCYCLE PRIOR tbl.emp_id = tbl.manager_id)";
                }

                if (!superUser)
                {
                    filterLoggedInEmp = $" AND t1.EMP_ID IN (SELECT tbl.EMP_ID FROM ts2_employee_details tbl START WITH tbl.emp_id = {empId} CONNECT BY NOCYCLE PRIOR tbl.emp_id = tbl.manager_id)";
                }

                if (inactive && !isActiveEmp)
                {
                    filterInActiveEmp = $" AND (t1.termination_date IS NULL OR " +
                    $"(t1.emp_status = 'INACTIVE' AND t1.termination_date >= TO_DATE('{startDate:MM/dd/yyyy}', 'mm/dd/yyyy') " +
                    $"AND t1.termination_date <= TO_DATE('{endDateInclusive:MM/dd/yyyy}', 'mm/dd/yyyy') AND TRUNC(access_time) <= t1.termination_date) " +
                    $"OR (t1.termination_date >= TO_DATE('{endDateInclusive:MM/dd/yyyy}', 'mm/dd/yyyy')))";
                }
                else
                {
                    filterInActiveEmp = " AND t1.emp_status = 'ACTIVE'";
                }

                string query = GenerateSwipeQuery(startDate, endDateInclusive, filterInActiveEmp, filterLoggedInEmp);

                var dataSet = new DataSet();
                using (var connection = new OracleConnection(_configuration.GetConnectionString("DefaultConnection")))
                using (var command = new OracleCommand(query, connection))
                using (var adapter = new OracleDataAdapter(command))
                {
                    await connection.OpenAsync();
                    adapter.Fill(dataSet);
                }

                // Check if the dataset is empty
                if (dataSet.Tables.Count == 0 || dataSet.Tables[0].Rows.Count == 0)
                {
                    return StatusCode(200, "No data available for empId: " + empId);
                }

                // Convert DataTable to a list of objects
                var dataTable = dataSet.Tables[0];
                var data = dataTable.AsEnumerable().Select(row => new
                {
                    MONTH = row["MONTH"].ToString(),
                    ACCESS_TIME = row["ACCESS_TIME"].ToString(),
                    FIRST_LOGIN = row["FIRST_LOGIN"].ToString(),
                    LAST_LOGOUT = row["LAST_LOGOUT"].ToString(),
                    TOTAL_DURATION = row["TOTAL_DURATION"].ToString(),
                    EMP_ID = Convert.ToInt32(row["EMP_ID"]),
                    EMP_NAME = row["EMP_NAME"].ToString(),
                    EMP_STATUS = row["EMP_STATUS"].ToString(),
                    JOB_GRADE = row["JOB_GRADE"].ToString(),
                    REPORTING_TO_EMPID = Convert.ToInt32(row["REPORTING_TO_EMPID"]),
                    REPORTING_TO = row["REPORTING_TO"].ToString(),
                    MANAGER_ID = Convert.ToInt32(row["MANAGER_ID"]),
                    MANAGER = row["MANAGER"].ToString(),
                    DIRECTOR_ID = Convert.ToInt32(row["DIRECTOR_ID"]),
                    DIRECTOR = row["DIRECTOR"].ToString(),
                    DIRECT_REPORT_ID = Convert.ToInt32(row["DIRECT_REPORT_ID"]),
                    DIRECT_REPORT = row["DIRECT_REPORT"].ToString(),
                    VP_ID = Convert.ToInt32(row["VP_ID"]),
                    VP = row["VP"].ToString(),
                    GENDER = row["GENDER"].ToString(),
                    LOCATION = row["LOCATION"].ToString(),
                    EMPLOYEE_EMAIL = row["EMPLOYEE_EMAIL"].ToString(),
                    EMPLOYEE_TYPE = row["EMPLOYEE_TYPE"].ToString(),
                    JOB_TITLE = row["JOB_TITLE"].ToString()
                }).ToList();

                return Ok(new
                {
                    status = "success",
                    data = data
                });
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Error retrieving employee swipe data.");
                return StatusCode(500, new { message = "Internal server error" });
            }
        }










        //[HttpGet("GetSwipeData")]
        //public async Task<IActionResult> GetSwipeData(
        //    [FromQuery] DateTime startDate,
        //    [FromQuery] DateTime endDate,
        //    [FromQuery] int empId,
        //    [FromQuery] bool superUser,
        //    [FromQuery] bool inactive)
        //{
        //    try
        //    {



        //        // Validate empId and check if the employee is active
        //        bool isValidEmpId = false;
        //        bool isActiveEmp = false;
        //        using (var connection = new OracleConnection(_configuration.GetConnectionString("DefaultConnection")))
        //        {
        //            connection.Open();
        //            using (var command = connection.CreateCommand())
        //            {
        //                command.CommandText = $"SELECT emp_status FROM ts2_employee_details WHERE emp_id = {empId}";
        //                command.CommandType = CommandType.Text;

        //                var result = command.ExecuteScalar();
        //                if (result != null)
        //                {
        //                    isValidEmpId = true;
        //                    isActiveEmp = result.ToString() == "ACTIVE";
        //                }
        //            }
        //        }

        //        if (!isValidEmpId)
        //        {
        //            return StatusCode(200, "Invalid empId: " + empId);
        //        }

        //        if (!isActiveEmp)
        //        {
        //            return StatusCode(200, "Inactive empId: " + empId);
        //        }


        //        var endDateInclusive = endDate.AddDays(1);
        //        string filterLoggedInEmp = string.Empty;
        //        string filterInActiveEmp = string.Empty;


        //        if (superUser)
        //        {
        //            bool isvalidsuperuser = false;
        //            using (var connection = new OracleConnection(_configuration.GetConnectionString("DefaultConnection")))
        //            {
        //                connection.Open();
        //                using (var command = connection.CreateCommand())
        //                {
        //                    command.CommandText = $"select emp_level from ts2_employee_details where emp_id={empId}";
        //                    command.CommandType = CommandType.Text;

        //                    var result = command.ExecuteScalar();
        //                    if (result != null && Convert.ToInt32(result.ToString()) >= 5)
        //                    {
        //                        isvalidsuperuser = true;
        //                    }



        //                }
        //            }


        //            if (!isvalidsuperuser)
        //            {
        //                return StatusCode(200, "not a valid super user empid= " + empId);
        //            }
        //        }

        //        if (!superUser)
        //        {
        //            filterLoggedInEmp = $" AND t1.EMP_ID IN (SELECT tbl.EMP_ID FROM ts2_employee_details tbl START WITH tbl.emp_id = {empId} CONNECT BY NOCYCLE PRIOR tbl.emp_id = tbl.manager_id)";
        //        }

        //        if (inactive)
        //        {
        //            filterInActiveEmp = $" AND (t1.termination_date IS NULL OR " +
        //                $"(t1.emp_status = 'INACTIVE' AND t1.termination_date >= TO_DATE('{startDate:MM/dd/yyyy}', 'mm/dd/yyyy') " +
        //                $"AND t1.termination_date <= TO_DATE('{endDateInclusive:MM/dd/yyyy}', 'mm/dd/yyyy') AND TRUNC(access_time) <= t1.termination_date) " +
        //                $"OR (t1.termination_date >= TO_DATE('{endDateInclusive:MM/dd/yyyy}', 'mm/dd/yyyy')))";
        //        }
        //        else
        //        {
        //            filterInActiveEmp = " AND t1.emp_status = 'ACTIVE'";
        //        }

        //        string query = GenerateSwipeQuery(startDate, endDateInclusive, filterInActiveEmp, filterLoggedInEmp);

        //        var dataSet = new DataSet();
        //        using (var connection = new OracleConnection(_configuration.GetConnectionString("DefaultConnection")))
        //        using (var command = new OracleCommand(query, connection))
        //        using (var adapter = new OracleDataAdapter(command))
        //        {
        //            await connection.OpenAsync();
        //            adapter.Fill(dataSet);
        //        }


        //        // Check if the dataset is empty
        //        if (dataSet.Tables.Count == 0 || dataSet.Tables[0].Rows.Count == 0)
        //        {
        //            return StatusCode(200, "No data available for empId: " + empId);
        //        }

        //        string json = JsonConvert.SerializeObject(dataSet.Tables[0]);


        //        // Convert JSON to DataTable
        //        DataTable dataTable = JsonConvert.DeserializeObject<DataTable>(json);

        //        // Generate Excel file with current date and time in the filename
        //        string currentDateTime = DateTime.Now.ToString("yyyyMMdd_HHmmss");
        //        string filePath = $"wforeport_{currentDateTime}.xlsx";
        //        ExcelHelper.CreateExcelFromDataTable(dataTable, filePath);

        //        byte[] content = System.IO.File.ReadAllBytes(filePath);
        //        return File(content, "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", filePath);


        //        //// Convert JSON to DataTable
        //        //DataTable dataTable = JsonConvert.DeserializeObject<DataTable>(json);

        //        //// Generate Excel file
        //        //string filePath = "employee_swipe_data.xlsx";
        //        //ExcelHelper.CreateExcelFromDataTable(dataTable, filePath);

        //        //byte[] content = System.IO.File.ReadAllBytes(filePath);
        //        //return File(content, "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", "employee_swipe_data.xlsx");




        //        return Ok(json);
        //    }
        //    catch (Exception ex)
        //    {
        //        _logger.LogError(ex, "Error retrieving employee swipe data.");
        //        return StatusCode(500, "Internal server error");
        //    }
        //}







        //[HttpGet("GetSwipeDataForExcel")]
        //public async Task<IActionResult> GetSwipeDataForExcel(
        //    [FromQuery] DateTime startDate,
        //    [FromQuery] DateTime endDate,
        //    [FromQuery] int empId,
        //    [FromQuery] bool superUser,
        //    [FromQuery] bool inactive)
        //{
        //    try
        //    {
        //        var endDateInclusive = endDate.AddDays(1);
        //        string filterLoggedInEmp = string.Empty;
        //        string filterInActiveEmp = string.Empty;

        //        if (!superUser)
        //        {
        //            filterLoggedInEmp = $" AND t1.EMP_ID IN (SELECT tbl.EMP_ID FROM ts2_employee_details tbl START WITH tbl.emp_id = {empId} CONNECT BY NOCYCLE PRIOR tbl.emp_id = tbl.manager_id)";
        //        }

        //        if (inactive)
        //        {
        //            filterInActiveEmp = $" AND (t1.termination_date IS NULL OR " +
        //                $"(t1.emp_status = 'INACTIVE' AND t1.termination_date >= TO_DATE('{startDate:MM/dd/yyyy}', 'mm/dd/yyyy') " +
        //                $"AND t1.termination_date <= TO_DATE('{endDateInclusive:MM/dd/yyyy}', 'mm/dd/yyyy') AND TRUNC(access_time) <= t1.termination_date) " +
        //                $"OR (t1.termination_date >= TO_DATE('{endDateInclusive:MM/dd/yyyy}', 'mm/dd/yyyy')))";
        //        }
        //        else
        //        {
        //            filterInActiveEmp = " AND t1.emp_status = 'ACTIVE'";
        //        }

        //        string query = GenerateSwipeQuery(startDate, endDateInclusive, filterInActiveEmp, filterLoggedInEmp);

        //        var dataSet = new DataSet();
        //        using (var connection = new OracleConnection(_configuration.GetConnectionString("DefaultConnection")))
        //        using (var command = new OracleCommand(query, connection))
        //        using (var adapter = new OracleDataAdapter(command))
        //        {
        //            await connection.OpenAsync();
        //            adapter.Fill(dataSet);
        //        }
        //        //string json = JsonConvert.SerializeObject(dataSet.Tables[0]);

        //        // Convert JSON to DataTable
        //       // DataTable dataTable = JsonConvert.DeserializeObject<DataTable>(json);

        //        // Generate Excel file
        //        using (var package = new ExcelPackage())
        //        {
        //            var worksheet = package.Workbook.Worksheets.Add("SwipeData");
        //            worksheet.Cells["A1"].LoadFromDataTable(dataSet.Tables[0], true);
        //            var stream = new MemoryStream();
        //            package.SaveAs(stream);
        //            var content = stream.ToArray();


        //            // Generate file name with current date and time
        //            string fileName = $"wfo_{DateTime.Now:yyyyMMdd_HHmmss}.xlsx";

        //            return File(content, "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", fileName);

        //        }

        //        //return Ok(json);
        //    }
        //    catch (Exception ex)
        //    {
        //        _logger.LogError(ex, "Error retrieving employee swipe data.");
        //        return StatusCode(500, "Internal server error");
        //    }
        //}

        private string GenerateSwipeQuery(DateTime startDate, DateTime endDate, string filterInActiveEmp, string filterLoggedInEmp)
        {


            string sql = @$"SELECT DISTINCT t2.swipe_month AS MONTH,
        TO_CHAR(t2.access_time, 'mm/dd/yy')   AS access_time,
          to_char(t3.first_login, 'HH24:MI:SS') first_login,
        to_char(t3.last_logout, 'HH24:MI:SS') last_logout,
        TO_CHAR(TRUNC((t3.last_logout - t3.first_login) * 24), 'FM00') || ':' ||
        TO_CHAR(TRUNC(MOD((t3.last_logout - t3.first_login) * 24 * 60, 60)), 'FM00') || ':' ||
        TO_CHAR(TRUNC(MOD((t3.last_logout - t3.first_login) * 24 * 60 * 60, 60)), 'FM00') AS total_duration,
        CAST(t1.emp_id AS INTEGER) as emp_id,
        t1.emp_name,
        t1.emp_status,
        NVL((
        CASE
           WHEN t1.job_grade LIKE '%M2FS%'
           THEN 'Others'
           WHEN t1.job_grade LIKE '%M4FS%'
           THEN 'Others'
           WHEN t1.job_grade LIKE '%N2%'
           THEN 'Others'
           WHEN t1.job_grade LIKE '%N5%'
           THEN 'Others'
           WHEN t1.job_grade LIKE '%N9%'
           THEN 'Others'
           WHEN t1.job_grade LIKE '%NS4%'
           THEN 'Others'
           WHEN t1.job_grade LIKE '%NS9A%'
           THEN 'Others'
           WHEN t1.job_grade LIKE '%SS3%'
           THEN 'Others'
           WHEN t1.job_grade LIKE '%E1%'
           THEN 'Others'
           ELSE t1.job_grade
        END), 'Others')   AS job_grade,
        CAST(t1.manager_id AS INTEGER) AS reporting_to_empid,
        CASE
                            WHEN sup.emp_name = 'CTLI' THEN
                                'NA'
                            WHEN sup.emp_name = ''     THEN
                                'NA'
                            ELSE
                                sup.emp_name
                    END                          AS reporting_to,
                    nvl((SELECT   CAST(mgr.emp_id AS INTEGER) FROM ts2_employee_details  mgr, employeesgradetolevel grade WHERE grade.job_grade = mgr.job_grade AND grade.emp_level IN(7, 8) AND ROWNUM = 1 and mgr.emp_id != t1.emp_id  START WITH mgr.emp_id = t1.emp_id CONNECT BY NOCYCLE PRIOR mgr.manager_id = mgr.emp_id), 0)                        AS manager_id,
                    nvl((SELECT mgr.emp_name        FROM ts2_employee_details  mgr, employeesgradetolevel grade WHERE  grade.job_grade = mgr.job_grade            AND grade.emp_level IN(7, 8) AND ROWNUM = 1 and mgr.emp_id != t1.emp_id START WITH mgr.emp_id = t1.emp_id CONNECT BY NOCYCLE PRIOR mgr.manager_id = mgr.emp_id), 'NA')                     AS manager,
                    nvl((SELECT CAST(mgr.emp_id AS INTEGER)        FROM ts2_employee_details  mgr, employeesgradetolevel grade WHERE  grade.job_grade = mgr.job_grade            AND grade.emp_level IN(9, 10) AND ROWNUM = 1 and mgr.emp_id != t1.emp_id START WITH mgr.emp_id = t1.emp_id CONNECT BY NOCYCLE PRIOR mgr.manager_id = mgr.emp_id), 0)     AS director_id,
                    nvl((SELECT mgr.emp_name        FROM ts2_employee_details  mgr, employeesgradetolevel grade WHERE  grade.job_grade = mgr.job_grade            AND grade.emp_level IN(9, 10) AND ROWNUM = 1 and mgr.emp_id != t1.emp_id START WITH mgr.emp_id = t1.emp_id CONNECT BY NOCYCLE PRIOR mgr.manager_id = mgr.emp_id), 'NA')     AS director,
                    nvl((
                    CASE WHEN(
                    SELECT CAST(mgr.emp_id AS INTEGER) FROM ts2_employee_details  mgr, employeesgradetolevel grade WHERE grade.job_grade = mgr.job_grade and mgr.emp_id != t1.emp_id AND(grade.emp_level IN(10))
                        AND mgr.manager_id IN(SELECT ted.emp_id FROM ts2_employee_details ted WHERE ted.manager_id in (Select emp_id from ts2_employee_details where emp_status = 'ACTIVE' and job_grade in ('E1')))
                        AND ROWNUM = 1 and mgr.emp_id != t1.emp_id START WITH mgr.emp_id = t1.emp_id CONNECT BY NOCYCLE PRIOR mgr.manager_id = mgr.emp_id) 
                        IS NOT NULL
                        THEN
                        (SELECT CAST(mgr.emp_id AS INTEGER) FROM ts2_employee_details  mgr, employeesgradetolevel grade WHERE grade.job_grade = mgr.job_grade and mgr.emp_id != t1.emp_id AND(grade.emp_level IN(10))
                        AND mgr.manager_id IN(SELECT ted.emp_id FROM ts2_employee_details ted WHERE ted.manager_id in (Select emp_id from ts2_employee_details where emp_status = 'ACTIVE' and job_grade in ('E1')))
                        AND ROWNUM = 1 and mgr.emp_id != t1.emp_id START WITH mgr.emp_id = t1.emp_id CONNECT BY NOCYCLE PRIOR mgr.manager_id = mgr.emp_id) 
                        ELSE
                    CASE WHEN(
                        SELECT CAST(mgr.emp_id AS INTEGER) FROM ts2_employee_details  mgr, employeesgradetolevel grade WHERE grade.job_grade = mgr.job_grade and mgr.emp_id != t1.emp_id AND(grade.emp_level IN(9))
                        AND mgr.manager_id IN(SELECT ted.emp_id FROM ts2_employee_details ted WHERE ted.manager_id in (Select emp_id from ts2_employee_details where emp_status = 'ACTIVE' and job_grade in ('E1')))
                        AND ROWNUM = 1 and mgr.emp_id != t1.emp_id START WITH mgr.emp_id = t1.emp_id CONNECT BY NOCYCLE PRIOR mgr.manager_id = mgr.emp_id) 
                        IS NOT NULL
                        THEN
                        (SELECT CAST(mgr.emp_id AS INTEGER) FROM ts2_employee_details  mgr, employeesgradetolevel grade WHERE grade.job_grade = mgr.job_grade and mgr.emp_id != t1.emp_id AND(grade.emp_level IN(9))
                        AND mgr.manager_id IN(SELECT ted.emp_id FROM ts2_employee_details ted WHERE ted.manager_id in (Select emp_id from ts2_employee_details where emp_status = 'ACTIVE' and job_grade in ('E1')))
                        AND ROWNUM = 1 and mgr.emp_id != t1.emp_id START WITH mgr.emp_id = t1.emp_id CONNECT BY NOCYCLE PRIOR mgr.manager_id = mgr.emp_id) 
                        ELSE
                    CASE WHEN
                        (SELECT CAST(mgr.emp_id AS INTEGER) FROM ts2_employee_details  mgr, employeesgradetolevel grade WHERE grade.job_grade = mgr.job_grade and mgr.emp_id != t1.emp_id AND(grade.emp_level IN(8))
                        AND mgr.manager_id IN(SELECT ted.emp_id FROM ts2_employee_details ted WHERE ted.manager_id in (Select emp_id from ts2_employee_details where emp_status = 'ACTIVE' and job_grade in ('E1')))
                        AND ROWNUM = 1 and mgr.emp_id != t1.emp_id START WITH mgr.emp_id = t1.emp_id CONNECT BY NOCYCLE PRIOR mgr.manager_id = mgr.emp_id)
                        IS NOT NULL
                        THEN
                        (SELECT CAST(mgr.emp_id AS INTEGER) FROM ts2_employee_details  mgr, employeesgradetolevel grade WHERE grade.job_grade = mgr.job_grade and mgr.emp_id != t1.emp_id AND(grade.emp_level IN(8))
                        AND mgr.manager_id IN(SELECT ted.emp_id FROM ts2_employee_details ted WHERE ted.manager_id in (Select emp_id from ts2_employee_details where emp_status = 'ACTIVE' and job_grade in ('E1')))
                        AND ROWNUM = 1 and mgr.emp_id != t1.emp_id START WITH mgr.emp_id = t1.emp_id CONNECT BY NOCYCLE PRIOR mgr.manager_id = mgr.emp_id)
                        ELSE
                    CASE WHEN
                        (SELECT CAST(mgr.emp_id AS INTEGER) FROM ts2_employee_details  mgr, employeesgradetolevel grade WHERE grade.job_grade = mgr.job_grade and mgr.emp_id != t1.emp_id AND(grade.emp_level IN(7))
                        AND mgr.manager_id IN(SELECT ted.emp_id FROM ts2_employee_details ted WHERE ted.manager_id in (Select emp_id from ts2_employee_details where emp_status = 'ACTIVE' and job_grade in ('E1')))
                        AND ROWNUM = 1 and mgr.emp_id != t1.emp_id START WITH mgr.emp_id = t1.emp_id CONNECT BY NOCYCLE PRIOR mgr.manager_id = mgr.emp_id)
                        IS NOT NULL
                        THEN
                        (SELECT CAST(mgr.emp_id AS INTEGER) FROM ts2_employee_details  mgr, employeesgradetolevel grade WHERE grade.job_grade = mgr.job_grade and mgr.emp_id != t1.emp_id AND(grade.emp_level IN(7))
                        AND mgr.manager_id IN(SELECT ted.emp_id FROM ts2_employee_details ted WHERE ted.manager_id in (Select emp_id from ts2_employee_details where emp_status = 'ACTIVE' and job_grade in ('E1')))
                        AND ROWNUM = 1 and mgr.emp_id != t1.emp_id START WITH mgr.emp_id = t1.emp_id CONNECT BY NOCYCLE PRIOR mgr.manager_id = mgr.emp_id)
                        ELSE
                        (SELECT CAST(mgr.emp_id AS INTEGER) FROM ts2_employee_details  mgr, employeesgradetolevel grade WHERE grade.job_grade = mgr.job_grade and mgr.emp_id != t1.emp_id AND(grade.emp_level IN(11))
                        AND mgr.manager_id IN(SELECT ted.emp_id FROM ts2_employee_details ted WHERE ted.manager_id in (Select emp_id from ts2_employee_details where emp_status = 'ACTIVE' and job_grade in ('E1')))
                        AND ROWNUM = 1 and mgr.emp_id != t1.emp_id START WITH mgr.emp_id = t1.emp_id CONNECT BY NOCYCLE PRIOR mgr.manager_id = mgr.emp_id)
                        END
                        END
                        END
                        END
                       ), 0)                      AS direct_report_id,
                    nvl((
                    CASE WHEN(
                        SELECT mgr.emp_name FROM ts2_employee_details  mgr, employeesgradetolevel grade WHERE grade.job_grade = mgr.job_grade and mgr.emp_id != t1.emp_id AND(grade.emp_level IN(10))
                        AND mgr.manager_id IN(SELECT ted.emp_id FROM ts2_employee_details ted WHERE ted.manager_id in (Select emp_id from ts2_employee_details where emp_status = 'ACTIVE' and job_grade in ('E1')))
                        AND ROWNUM = 1 and mgr.emp_id != t1.emp_id START WITH mgr.emp_id = t1.emp_id CONNECT BY NOCYCLE PRIOR mgr.manager_id = mgr.emp_id)
                        IS NOT NULL
                        THEN
                        (SELECT mgr.emp_name FROM ts2_employee_details  mgr, employeesgradetolevel grade WHERE grade.job_grade = mgr.job_grade and mgr.emp_id != t1.emp_id AND(grade.emp_level IN(10))
                        AND mgr.manager_id IN(SELECT ted.emp_id FROM ts2_employee_details ted WHERE ted.manager_id in (Select emp_id from ts2_employee_details where emp_status = 'ACTIVE' and job_grade in ('E1')))
                        AND ROWNUM = 1 and mgr.emp_id != t1.emp_id START WITH mgr.emp_id = t1.emp_id CONNECT BY NOCYCLE PRIOR mgr.manager_id = mgr.emp_id)
                        ELSE
                    CASE WHEN(
                        SELECT mgr.emp_name FROM ts2_employee_details  mgr, employeesgradetolevel grade WHERE grade.job_grade = mgr.job_grade and mgr.emp_id != t1.emp_id AND(grade.emp_level IN(9))
                        AND mgr.manager_id IN(SELECT ted.emp_id FROM ts2_employee_details ted WHERE ted.manager_id in (Select emp_id from ts2_employee_details where emp_status = 'ACTIVE' and job_grade in ('E1')))
                        AND ROWNUM = 1 and mgr.emp_id != t1.emp_id START WITH mgr.emp_id = t1.emp_id CONNECT BY NOCYCLE PRIOR mgr.manager_id = mgr.emp_id)
                        IS NOT NULL
                        THEN
                        (SELECT mgr.emp_name FROM ts2_employee_details  mgr, employeesgradetolevel grade WHERE grade.job_grade = mgr.job_grade and mgr.emp_id != t1.emp_id AND(grade.emp_level IN(9))
                        AND mgr.manager_id IN(SELECT ted.emp_id FROM ts2_employee_details ted WHERE ted.manager_id in (Select emp_id from ts2_employee_details where emp_status = 'ACTIVE' and job_grade in ('E1')))
                        AND ROWNUM = 1 and mgr.emp_id != t1.emp_id START WITH mgr.emp_id = t1.emp_id CONNECT BY NOCYCLE PRIOR mgr.manager_id = mgr.emp_id)
                        ELSE
                    CASE WHEN
                        (SELECT mgr.emp_name FROM ts2_employee_details  mgr, employeesgradetolevel grade WHERE grade.job_grade = mgr.job_grade and mgr.emp_id != t1.emp_id AND(grade.emp_level IN(8))
                        AND mgr.manager_id IN(SELECT ted.emp_id FROM ts2_employee_details ted WHERE ted.manager_id in (Select emp_id from ts2_employee_details where emp_status = 'ACTIVE' and job_grade in ('E1')))
                        AND ROWNUM = 1 and mgr.emp_id != t1.emp_id START WITH mgr.emp_id = t1.emp_id CONNECT BY NOCYCLE PRIOR mgr.manager_id = mgr.emp_id)
                        IS NOT NULL
                        THEN
                        (SELECT mgr.emp_name FROM ts2_employee_details  mgr, employeesgradetolevel grade WHERE grade.job_grade = mgr.job_grade and mgr.emp_id != t1.emp_id AND(grade.emp_level IN(8))
                        AND mgr.manager_id IN(SELECT ted.emp_id FROM ts2_employee_details ted WHERE ted.manager_id in (Select emp_id from ts2_employee_details where emp_status = 'ACTIVE' and job_grade in ('E1')))
                        AND ROWNUM = 1 and mgr.emp_id != t1.emp_id START WITH mgr.emp_id = t1.emp_id CONNECT BY NOCYCLE PRIOR mgr.manager_id = mgr.emp_id)
                        ELSE
                    CASE WHEN
                        (SELECT mgr.emp_name FROM ts2_employee_details  mgr, employeesgradetolevel grade WHERE grade.job_grade = mgr.job_grade and mgr.emp_id != t1.emp_id AND(grade.emp_level IN(7))
                        AND mgr.manager_id IN(SELECT ted.emp_id FROM ts2_employee_details ted WHERE ted.manager_id in (Select emp_id from ts2_employee_details where emp_status = 'ACTIVE' and job_grade in ('E1')))
                        AND ROWNUM = 1 and mgr.emp_id != t1.emp_id START WITH mgr.emp_id = t1.emp_id CONNECT BY NOCYCLE PRIOR mgr.manager_id = mgr.emp_id)
                        IS NOT NULL
                        THEN
                        (SELECT mgr.emp_name FROM ts2_employee_details  mgr, employeesgradetolevel grade WHERE grade.job_grade = mgr.job_grade and mgr.emp_id != t1.emp_id AND(grade.emp_level IN(7))
                        AND mgr.manager_id IN(SELECT ted.emp_id FROM ts2_employee_details ted WHERE ted.manager_id in (Select emp_id from ts2_employee_details where emp_status = 'ACTIVE' and job_grade in ('E1')))
                        AND ROWNUM = 1 and mgr.emp_id != t1.emp_id START WITH mgr.emp_id = t1.emp_id CONNECT BY NOCYCLE PRIOR mgr.manager_id = mgr.emp_id)
                        ELSE
                        (SELECT mgr.emp_name FROM ts2_employee_details  mgr, employeesgradetolevel grade WHERE grade.job_grade = mgr.job_grade and mgr.emp_id != t1.emp_id AND(grade.emp_level IN(11))
                        AND mgr.manager_id IN(SELECT ted.emp_id FROM ts2_employee_details ted WHERE ted.manager_id in (Select emp_id from ts2_employee_details where emp_status = 'ACTIVE' and job_grade in ('E1')))
                        AND ROWNUM = 1 and mgr.emp_id != t1.emp_id START WITH mgr.emp_id = t1.emp_id CONNECT BY NOCYCLE PRIOR mgr.manager_id = mgr.emp_id)
                        END
                        END
                        END
                    END
                        ), 'NA') AS direct_report,
                    nvl((SELECT CAST(mgr.emp_id AS INTEGER)        FROM ts2_employee_details  mgr, employeesgradetolevel grade WHERE  grade.job_grade = mgr.job_grade            AND grade.emp_level IN(11) AND ROWNUM = 1 and mgr.emp_id != t1.emp_id START WITH mgr.emp_id = t1.emp_id CONNECT BY NOCYCLE PRIOR mgr.manager_id = mgr.emp_id), 0)     AS Vp_id,
                    nvl((SELECT mgr.emp_name        FROM ts2_employee_details  mgr, employeesgradetolevel grade WHERE  grade.job_grade = mgr.job_grade            AND grade.emp_level IN(11) AND ROWNUM = 1 and mgr.emp_id != t1.emp_id START WITH mgr.emp_id = t1.emp_id CONNECT BY NOCYCLE PRIOR mgr.manager_id = mgr.emp_id), 'NA')     AS Vp,
                    t1.gender,
                        nvl((
                            CASE
                            WHEN t2.reader LIKE('%BNGRIIST_04-02_WORKSTATION DR OUT%') THEN
                                'Bangalore TS 2nd floor'
                            WHEN t2.reader LIKE('%BNGRIIST_04-02_WORKSTATION DR IN%') THEN
                                'Bangalore TS 2nd floor'
                            WHEN t2.reader LIKE('%BNGRII50_03-02_RECEPTION ENTRY%') THEN
                                'Bangalore TS 2nd floor'
                            WHEN t2.reader LIKE('%BNGRIIST_03-04_B-WING WORKSTATION DR IN%') THEN
                                'Bangalore TS 4th floor'
                            WHEN t2.reader LIKE('%BNGRIIST_04-04_B-WING ELECTRICAL ROOM%') THEN
                                'Bangalore TS 4th floor'
                            WHEN t2.reader LIKE('%BNGRIIST_03-04_B-WING WORKSTATION DR OUT%') THEN
                                'Bangalore TS 4th floor'
                            WHEN t2.reader LIKE('%BNGRIIST_08-04_C-WING WORKSTATION DR OUT%') THEN
                                'Bangalore TS 4th floor'
                            WHEN t2.reader LIKE('%BNGRIIST_08-04_C-WING WORKSTATION DR IN%') THEN
                                'Bangalore TS 4th floor'
                            WHEN t2.reader LIKE('%BNGRIIST_01-04_MAIN LIFT LOBBY LS DR%') THEN
                                'Bangalore TS 4th floor'
                            WHEN t2.reader LIKE('%BNGRIIST_06-04_B-WING IT ROOM%') THEN
                                'Bangalore TS 4th floor'
                            WHEN t2.reader LIKE('%BNGRIIST_05-05_WORKSTATION LS DR OUT%') THEN
                                'Bangalore TS 4th floor'
                            WHEN t2.reader LIKE('%BNGRIIST_05-05_WORKSTATION LS DR IN%') THEN
                                'Bangalore TS 4th floor'
                            WHEN t2.reader LIKE('%BNGRIIST_06-05_WORKSTATION RS DR IN%') THEN
                                'Bangalore TS 4th floor'
                            WHEN t2.reader LIKE('%SEAL_HYDERABAD_IN%') THEN
                                'Hyderabad'
                            WHEN t2.reader LIKE('%SEAL_HYDERABAD_OUT%') THEN
                                'Hyderabad'
                            WHEN t2.reader LIKE('%SEAL_CHENNAI_IN%') THEN
                                'Chennai'
                            WHEN t2.reader LIKE('%SEAL_CHENNAI_OUT%') THEN
                                'Chennai'
                            WHEN t2.reader LIKE('%NODAII50_01-05_HUB ROOM%') THEN
                                'Noida'
                            WHEN t2.reader LIKE('%NODAII50_02-05_SERVER ROOM ENTRY%') THEN
                                'Noida'
                            WHEN t2.reader LIKE('%NODAII50_03-05_IT STORE ENTRY%') THEN
                                'Noida'
                            WHEN t2.reader LIKE('%NODAII50_04-05_TELE CONFERANCE ROOM%') THEN
                                'Noida'
                            WHEN t2.reader LIKE('%NODAII50_05-05_UPS ROOM%') THEN
                                'Noida'
                            WHEN t2.reader LIKE('%NODAII50_06-05_ELECTRICAL ROOM 01%') THEN
                                'Noida'
                            WHEN t2.reader LIKE('%NODAII50_07-05_SERVICE LIFT LOBBY%') THEN
                                'Noida'
                            WHEN t2.reader LIKE('%NODAII50_08-05_MAIN LIFT LOBBY%') THEN
                                'Noida'
                            WHEN t2.reader LIKE('%NODAII50_09-05_ELECTRICAL ROOM 02%') THEN
                                'Noida'
                            WHEN t2.reader LIKE('%NODAII50_11-05_WORKSTATION DR IN%') THEN
                                'Noida'
                            WHEN t2.reader LIKE('%NODAII50_11-05_WORKSTATION DR OUT%') THEN
                                'Noida'
                            WHEN t2.reader LIKE('%NODAII50%') THEN
                                'Noida'

                            WHEN t2.reader LIKE('%NODIAII50_11-05_WORKSTATION DR IN%') THEN
                                'Noida'
                            WHEN t2.reader LIKE('%NODIAII50_07-05_SERVICE LIFT LOBBY%') THEN
                                'Noida'
                            WHEN t2.reader LIKE('%NODIAII50_01-05_HUB ROOM%') THEN
                                'Noida'
                            WHEN t2.reader LIKE('%NODIAII50_03-05_IT STORE ENTRY%') THEN
                                'Noida'
                            WHEN t2.reader LIKE('%NODIAII50_06-05_ELECTRICAL ROOM 01%') THEN
                                'Noida'
                            WHEN t2.reader LIKE('%NODIAII50_02-05_SERVER ROOM ENTRY%') THEN
                                'Noida'
                            WHEN t2.reader LIKE('%NODIAII50_05-05_UPS ROOM%') THEN
                                'Noida'
                            WHEN t2.reader LIKE('%NODIAII50_09-05_ELECTRICAL ROOM 02%') THEN
                                'Noida'
                            WHEN t2.reader LIKE('%NODIAII50_08-05_MAIN LIFT LOBBY%') THEN
                                'Noida'
                            WHEN t2.reader LIKE('%NODIAII50_11-05_WORKSTATION DR OUT%') THEN
                                'Noida'
                        END
                    ), 'Bangalore TS 4th floor') AS location,
        t1.email_id_off AS employee_email,
        t1.ps_emp_type AS employee_type,
        t1.job_title_full AS job_title
        -- t3.first_login
        --t3.last_logout
        FROM ts2_employee_details t1
        INNER JOIN(
        SELECT emp_id,
                swipe_month,
                TRUNC(access_time) AS day,
                MIN(access_time) AS first_login,
                MAX(access_time) AS last_logout
        FROM (
           SELECT emp_id,
                  TO_CHAR(TRUNC(access_time),'Mon-yyyy') AS swipe_month,
                  access_time
           FROM rm_employee_access_details acd,
                rm_emp_batch_id bat
           WHERE acd.batch_id = bat.batch_id
             AND acd.access_time BETWEEN TO_DATE('" + startDate.ToShortDateString() + "', 'mm/dd/yyyy') AND TO_DATE('" + endDate.ToShortDateString() + "', 'mm/dd/yyyy')" +
   " UNION " +
   "SELECT CAST(temp.employee_id AS INTEGER) AS emp_id, TO_CHAR(TRUNC(acd.access_time), 'Mon-yyyy') AS swipe_month, acd.access_time " +
   "FROM rm_employee_access_details acd, rm_temp_card_details temp, ts2_employee_details emp " +
   "WHERE acd.batch_id = temp.temp_card_id " +
     "AND emp.emp_id = temp.employee_id " +
     "AND acd.access_time BETWEEN TO_DATE('" + startDate.ToShortDateString() + "', 'mm/dd/yyyy') AND TO_DATE('" + endDate.ToShortDateString() + "', 'mm/dd/yyyy') " +
     "AND acd.access_time BETWEEN temp.missed_date AND TO_DATE(TO_CHAR(temp.returndate + 1, 'mm/dd/yyyy'), 'mm/dd/yyyy') " +
     "AND(TRUNC(acd.access_time) <= TRUNC(emp.termination_date) OR emp.termination_date IS NULL) " +
") " +
"GROUP BY emp_id, swipe_month, TRUNC(access_time) " +
") t3 ON t1.emp_id = t3.emp_id " +
"INNER JOIN( " +
"SELECT emp_id, swipe_month, reader, access_time " +
            "FROM(" +
   "SELECT ROW_NUMBER() OVER(PARTITION BY emp_id, TRUNC(access_time) ORDER BY access_time) AS rank, " +
    "     emp_id, swipe_month, reader, access_time FROM(" +
     "SELECT DISTINCT bat.emp_id, " +
                     "TO_CHAR(TRUNC(acd.access_time), 'Mon-yyyy') AS swipe_month, " +
                     "acd.access_time AS access_time, " +
                     "acd.batch_reader_loc AS reader " +
     "FROM rm_employee_access_details acd, " +
          "rm_emp_batch_id bat " +
     "WHERE acd.batch_id = bat.batch_id " +
       "AND acd.access_time BETWEEN TO_DATE('" + startDate.ToShortDateString() + "', 'mm/dd/yyyy') AND TO_DATE('" + endDate.ToShortDateString() + "', 'mm/dd/yyyy') " +
     "UNION " +
     "SELECT DISTINCT CAST(temp.employee_id AS INTEGER) AS emp_id, " +
                     "TO_CHAR(TRUNC(acd.access_time), 'Mon-yyyy') AS swipe_month, " +
                     "acd.access_time AS access_time, " +
                     "acd.batch_reader_loc AS reader " +
     "FROM rm_employee_access_details acd, " +
          "rm_temp_card_details temp, " +
          "ts2_employee_details emp " +
     "WHERE acd.batch_id = temp.temp_card_id " +
       "AND emp.emp_id = temp.employee_id " +
       "AND acd.access_time BETWEEN TO_DATE('" + startDate.ToShortDateString() + "', 'mm/dd/yyyy') AND TO_DATE('" + endDate.ToShortDateString() + "', 'mm/dd/yyyy') " +
       "AND acd.access_time BETWEEN temp.missed_date AND TO_DATE(TO_CHAR(temp.returndate + 1, 'mm/dd/yyyy'), 'mm/dd/yyyy') " +
       "AND(TRUNC(acd.access_time) <= TRUNC(emp.termination_date) OR emp.termination_date IS NULL) " +
   ") " +
") " +
"WHERE rank = 1 " +
") t2 ON t1.emp_id = t2.emp_id " +
"AND TRUNC(t2.access_time) = t3.day, " +
"ts2_employee_details sup, " +
"ts2_employee_details mgr " +
"WHERE t1.manager_id = sup.emp_id " +
"AND mgr.emp_id = sup.manager_id " +
"AND lower(t1.cuid) NOT LIKE('x%') " + filterInActiveEmp + " " + filterLoggedInEmp +
" ORDER BY  " +
" first_login ASC ";
            return sql;
        }


        [HttpGet("HybridWorkFromOfficeReport")]
        public async Task<IActionResult> HybridWorkFromOfficeReport(
            [FromQuery] DateTime startDate,
            [FromQuery] DateTime endDate,
            [FromQuery] int empId,
            [FromQuery] bool superUser,
            [FromQuery] bool inactive,
            [FromQuery] bool IncludeODOO)
        {
            try
            {
                // Validate empId and check if the employee is active
                bool isValidEmpId = false;
                bool isActiveEmp = false;
                bool isSuperUser = false;
                using (var connection = new OracleConnection(_configuration.GetConnectionString("DefaultConnection")))
                {
                    await connection.OpenAsync();
                    using (var command = connection.CreateCommand())
                    {
                        command.CommandText = $"SELECT emp_status, emp_level FROM ts2_employee_details WHERE emp_id = {empId}";
                        command.CommandType = CommandType.Text;

                        using (var reader = await command.ExecuteReaderAsync())
                        {
                            if (await reader.ReadAsync())
                            {
                                isValidEmpId = true;
                                isActiveEmp = reader["emp_status"].ToString() == "ACTIVE";
                                isSuperUser = Convert.ToInt32(reader["emp_level"]) >= 5;
                            }
                        }
                    }
                }

                if (!isValidEmpId)
                {
                    return StatusCode(200, "Invalid empId: " + empId);
                }

                if (!isActiveEmp && inactive)
                {
                    return StatusCode(200, "Inactive empId: " + empId);
                }

                var endDateInclusive = endDate.AddDays(1);
                string filterLoggedInEmp = string.Empty;
                string filterInActiveEmp = string.Empty;
                string filterOfficeDutyOutsideOffice = string.Empty;

                if (superUser || isSuperUser)
                {
                    if (!isSuperUser)
                    {
                        return StatusCode(200, "Not a valid super user empId: " + empId);
                    }
                }
                else
                {
                    filterLoggedInEmp = $" AND t1.EMP_ID IN (SELECT tbl.EMP_ID FROM ts2_employee_details tbl START WITH tbl.emp_id = {empId} CONNECT BY NOCYCLE PRIOR tbl.emp_id = tbl.manager_id)";
                }

                if (inactive && !isActiveEmp)
                {
                    filterInActiveEmp = $" AND (t1.termination_date IS NULL OR " +
                        $"(t1.emp_status = 'INACTIVE' AND t1.termination_date >= TO_DATE('{startDate:MM/dd/yyyy}', 'mm/dd/yyyy') " +
                        $"AND t1.termination_date <= TO_DATE('{endDateInclusive:MM/dd/yyyy}', 'mm/dd/yyyy') AND TRUNC(access_time) <= t1.termination_date) " +
                        $"OR (t1.termination_date >= TO_DATE('{endDateInclusive:MM/dd/yyyy}', 'mm/dd/yyyy')))";
                }
                else
                {
                    filterInActiveEmp = " AND t1.emp_status = 'ACTIVE'";
                }

                if (IncludeODOO)
                {
                    filterOfficeDutyOutsideOffice = "UNION Select distinct emp_id, TO_CHAR(TRUNC(attendance_date), 'Mon-yyyy') as swipe_month, attendance_date as swipe " +
                        " from(SELECT emp_id, TRUNC(start_date) + (n - 1) AS attendance_date FROM rm_attendance_details JOIN dayrange ON n <= TRUNC(end_date) - TRUNC(start_date) + 1" +
                        " WHERE attendance_type_id = 3 AND start_date >= (sysdate - 360) ORDER BY emp_id, attendance_date) " +
                        " where(attendance_date >= TO_DATE('" + startDate.ToShortDateString() + "', 'mm/dd/yyyy') AND attendance_date < TO_DATE('" + endDateInclusive.ToShortDateString() + "', 'mm/dd/yyyy')) order by swipe asc";
                }
                else
                {
                    filterOfficeDutyOutsideOffice = " ";
                }

                string query = HybridWFHQuery(startDate, endDateInclusive, filterInActiveEmp, filterLoggedInEmp, filterOfficeDutyOutsideOffice);

                var dataSet = new DataSet();
                using (var connection = new OracleConnection(_configuration.GetConnectionString("DefaultConnection")))
                using (var command = new OracleCommand(query, connection))
                using (var adapter = new OracleDataAdapter(command))
                {
                    await connection.OpenAsync();
                    adapter.Fill(dataSet);
                }

                // Check if the dataset is empty
                if (dataSet.Tables.Count == 0 || dataSet.Tables[0].Rows.Count == 0)
                {
                    return StatusCode(200, "No data available for empId: " + empId);
                }

                string json = JsonConvert.SerializeObject(dataSet.Tables[0]);

                // Convert JSON to DataTable
                DataTable dataTable = JsonConvert.DeserializeObject<DataTable>(json);

                // Generate Excel file with current date and time in the filename
                string currentDateTime = DateTime.Now.ToString("yyyyMMdd_HHmmss");
                string filePath = $"HybridWorkFromHomeReport_{currentDateTime}.xlsx";
                ExcelHelper.CreateExcelFromDataTable(dataTable, filePath);

                byte[] content = System.IO.File.ReadAllBytes(filePath);
                return File(content, "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", filePath);
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Error retrieving employee swipe data.");
                return StatusCode(500, "Internal server error");
            }
        }


        [HttpGet("HybridWorkFromOfficeReport2")]
        public async Task<IActionResult> HybridWorkFromOfficeReport2(
    [FromQuery] DateTime startDate,
    [FromQuery] DateTime endDate,
    [FromQuery] int empId,
    [FromQuery] bool superUser,
    [FromQuery] bool inactive,
    [FromQuery] bool IncludeODOO)
        {
            try
            {
                // Validate empId and check if the employee is active
                bool isValidEmpId = false;
                bool isActiveEmp = false;
                bool isSuperUser = false;
                using (var connection = new OracleConnection(_configuration.GetConnectionString("DefaultConnection")))
                {
                    await connection.OpenAsync();
                    using (var command = connection.CreateCommand())
                    {
                        command.CommandText = $"SELECT emp_status, emp_level FROM ts2_employee_details WHERE emp_id = {empId}";
                        command.CommandType = CommandType.Text;

                        using (var reader = await command.ExecuteReaderAsync())
                        {
                            if (await reader.ReadAsync())
                            {
                                isValidEmpId = true;
                                isActiveEmp = reader["emp_status"].ToString() == "ACTIVE";
                                isSuperUser = Convert.ToInt32(reader["emp_level"]) >= 5;
                            }
                        }
                    }
                }

                if (!isValidEmpId)
                {
                    return StatusCode(200, "Invalid empId: " + empId);
                }

                if (!isActiveEmp && inactive)
                {
                    return StatusCode(200, "Inactive empId: " + empId);
                }

                var endDateInclusive = endDate.AddDays(1);
                string filterLoggedInEmp = string.Empty;
                string filterInActiveEmp = string.Empty;
                string filterOfficeDutyOutsideOffice = string.Empty;

                if (superUser || isSuperUser)
                {
                    if (!isSuperUser)
                    {
                        return StatusCode(200, "Not a valid super user empId: " + empId);
                    }
                }
                else
                {
                    filterLoggedInEmp = $" AND t1.EMP_ID IN (SELECT tbl.EMP_ID FROM ts2_employee_details tbl START WITH tbl.emp_id = {empId} CONNECT BY NOCYCLE PRIOR tbl.emp_id = tbl.manager_id)";
                }

                if (inactive && !isActiveEmp)
                {
                    filterInActiveEmp = $" AND (t1.termination_date IS NULL OR " +
                        $"(t1.emp_status = 'INACTIVE' AND t1.termination_date >= TO_DATE('{startDate:MM/dd/yyyy}', 'mm/dd/yyyy') " +
                        $"AND t1.termination_date <= TO_DATE('{endDateInclusive:MM/dd/yyyy}', 'mm/dd/yyyy') AND TRUNC(access_time) <= t1.termination_date) " +
                        $"OR (t1.termination_date >= TO_DATE('{endDateInclusive:MM/dd/yyyy}', 'mm/dd/yyyy')))";
                }
                else
                {
                    filterInActiveEmp = " AND t1.emp_status = 'ACTIVE'";
                }

                if (IncludeODOO)
                {
                    filterOfficeDutyOutsideOffice = "UNION Select distinct emp_id, TO_CHAR(TRUNC(attendance_date), 'Mon-yyyy') as swipe_month, attendance_date as swipe " +
                        " from(SELECT emp_id, TRUNC(start_date) + (n - 1) AS attendance_date FROM rm_attendance_details JOIN dayrange ON n <= TRUNC(end_date) - TRUNC(start_date) + 1" +
                        " WHERE attendance_type_id = 3 AND start_date >= (sysdate - 360) ORDER BY emp_id, attendance_date) " +
                        " where(attendance_date >= TO_DATE('" + startDate.ToShortDateString() + "', 'mm/dd/yyyy') AND attendance_date < TO_DATE('" + endDateInclusive.ToShortDateString() + "', 'mm/dd/yyyy')) order by swipe asc";
                }
                else
                {
                    filterOfficeDutyOutsideOffice = " ";
                }

                string query = HybridWFHQuery(startDate, endDateInclusive, filterInActiveEmp, filterLoggedInEmp, filterOfficeDutyOutsideOffice);

                var dataSet = new DataSet();
                using (var connection = new OracleConnection(_configuration.GetConnectionString("DefaultConnection")))
                using (var command = new OracleCommand(query, connection))
                using (var adapter = new OracleDataAdapter(command))
                {
                    await connection.OpenAsync();
                    adapter.Fill(dataSet);
                }

                // Check if the dataset is empty
                if (dataSet.Tables.Count == 0 || dataSet.Tables[0].Rows.Count == 0)
                {
                    return StatusCode(200, "No data available for empId: " + empId);
                }

                // Convert DataTable to a list of objects with all required fields
                var dataTable = dataSet.Tables[0];
                var data = dataTable.AsEnumerable().Select(row => new
                {

                    EMP_ID = Convert.ToInt32(row["EMP_ID"]),
                    CUID = row["CUID"].ToString(),
                    EMP_NAME = row["EMP_NAME"].ToString(),
                    EMP_STATUS = row["EMP_STATUS"].ToString(),
                    JOB_GRADE = row["JOB_GRADE"].ToString(),
                    JOB_TITLE_FULL = row["JOB_TITLE_FULL"].ToString(),
                    CITY = row["CITY"].ToString(),
                    SUPEMPID = Convert.ToInt32(row["SUPEMPID"]),
                    SUPEMPNAME = row["SUPEMPNAME"].ToString(),
                    SECOND_REPORTING_TO_EMPID = Convert.ToInt32(row["SECOND_REPORTING_TO_EMPID"]),
                    SECOND_REPORTING_TO = row["SECOND_REPORTING_TO"].ToString(),
                    MGREMPID = Convert.ToInt32(row["MGREMPID"]),
                    MGREMPNAME = row["MGREMPNAME"].ToString(),
                    DIRECTOR_ID = Convert.ToInt32(row["DIRECTOR_ID"]),
                    DIRECTOR = row["DIRECTOR"].ToString(),
                    DIRECT_REPORT_ID = Convert.ToInt32(row["DIRECT_REPORT_ID"]),
                    DIRECT_REPORT = row["DIRECT_REPORT"].ToString(),
                    MONTH = row["MONTH"].ToString(),
                    NO_OF_DAYS = Convert.ToInt32(row["NO_OF_DAYS"])

                }).ToList();

                return Ok(new
                {
                    status = "success",
                    data = data
                });
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Error retrieving employee swipe data.");
                return StatusCode(500, new { message = "Internal server error" });
            }
        }



        private string HybridWFHQuery(DateTime startdate, DateTime endDate, string filterInActiveEmp, string filterLoggedInEmp,string filterOfficeDutyOutsideOffice)
        {


            string sql = @$"SELECT distinct
                        t1.emp_id,
                        t1.cuid,
                        t1.emp_name,
                        t1.emp_status,
                        t1.job_grade,
                        t1.job_title_full,
                        t1.city,

     sup.emp_id                   AS supempid,
                    CASE
                        WHEN sup.emp_name = 'CTLI' THEN
                            'NA'
                        WHEN sup.emp_name = ''     THEN
                            'NA'
                        ELSE
                            sup.emp_name
                    END                          AS supempname,
                    mgr.emp_id                   AS second_reporting_to_empid,
                    CASE
                        WHEN mgr.emp_name = 'CTLI' THEN
                            'NA'
                        WHEN mgr.emp_name = ''     THEN
                            'NA'
                        ELSE
                            mgr.emp_name
                    END                          AS second_reporting_to,
    nvl((
        SELECT
            mgr.emp_id
        FROM
            ts2_employee_details  mgr, employeesgradetolevel grade
        WHERE
                grade.job_grade = mgr.job_grade
            AND grade.emp_level IN(7, 8)
            AND ROWNUM = 1 and mgr.emp_id != t1.emp_id
        START WITH
            mgr.emp_id = t1.emp_id
        CONNECT BY NOCYCLE
            PRIOR mgr.manager_id = mgr.emp_id
    ), 0)                      AS mgrempid,
    nvl((
        SELECT
            mgr.emp_name
        FROM
            ts2_employee_details  mgr, employeesgradetolevel grade
        WHERE
                grade.job_grade = mgr.job_grade
            AND grade.emp_level IN(7, 8)
            AND ROWNUM = 1 and mgr.emp_id != t1.emp_id
        START WITH
            mgr.emp_id = t1.emp_id
        CONNECT BY NOCYCLE
            PRIOR mgr.manager_id = mgr.emp_id
    ), 'NA')                   AS mgrempname,
nvl((
        SELECT
            mgr.emp_id
        FROM
            ts2_employee_details  mgr, employeesgradetolevel grade
        WHERE
                grade.job_grade = mgr.job_grade
            AND grade.emp_level IN(9, 10)
            AND ROWNUM = 1  and mgr.emp_id != t1.emp_id
        START WITH
            mgr.emp_id = t1.emp_id
        CONNECT BY NOCYCLE
            PRIOR mgr.manager_id = mgr.emp_id
    ), 0)                      AS director_id,
    nvl((
        SELECT
            mgr.emp_name
        FROM
            ts2_employee_details  mgr, employeesgradetolevel grade
        WHERE
                grade.job_grade = mgr.job_grade
            AND grade.emp_level IN(9, 10)
            AND ROWNUM = 1  and mgr.emp_id != t1.emp_id
        START WITH
            mgr.emp_id = t1.emp_id
        CONNECT BY NOCYCLE
            PRIOR mgr.manager_id = mgr.emp_id
    ), 'NA')                   AS director,
    nvl((
        CASE WHEN (
        SELECT mgr.emp_id FROM ts2_employee_details  mgr, employeesgradetolevel grade WHERE grade.job_grade = mgr.job_grade and mgr.emp_id != t1.emp_id AND (grade.emp_level IN(10))
        AND mgr.manager_id IN(SELECT ted.emp_id FROM ts2_employee_details ted WHERE ted.manager_id in (Select emp_id from ts2_employee_details where emp_status = 'ACTIVE' and job_grade in ('E1')))
        AND ROWNUM = 1 and mgr.emp_id != t1.emp_id START WITH mgr.emp_id = t1.emp_id CONNECT BY NOCYCLE PRIOR mgr.manager_id = mgr.emp_id) 
        IS NOT NULL 
        THEN
        (SELECT mgr.emp_id FROM ts2_employee_details  mgr, employeesgradetolevel grade WHERE grade.job_grade = mgr.job_grade and mgr.emp_id != t1.emp_id AND (grade.emp_level IN(10))
        AND mgr.manager_id IN(SELECT ted.emp_id FROM ts2_employee_details ted WHERE ted.manager_id in (Select emp_id from ts2_employee_details where emp_status = 'ACTIVE' and job_grade in ('E1')))
        AND ROWNUM = 1 and mgr.emp_id != t1.emp_id START WITH mgr.emp_id = t1.emp_id CONNECT BY NOCYCLE PRIOR mgr.manager_id = mgr.emp_id) 
        ELSE
        CASE WHEN (
        SELECT mgr.emp_id FROM ts2_employee_details  mgr, employeesgradetolevel grade WHERE grade.job_grade = mgr.job_grade and mgr.emp_id != t1.emp_id AND (grade.emp_level IN(9))
        AND mgr.manager_id IN(SELECT ted.emp_id FROM ts2_employee_details ted WHERE ted.manager_id in (Select emp_id from ts2_employee_details where emp_status = 'ACTIVE' and job_grade in ('E1')))
        AND ROWNUM = 1 and mgr.emp_id != t1.emp_id START WITH mgr.emp_id = t1.emp_id CONNECT BY NOCYCLE PRIOR mgr.manager_id = mgr.emp_id) 
        IS NOT NULL 
        THEN
        (SELECT mgr.emp_id FROM ts2_employee_details  mgr, employeesgradetolevel grade WHERE grade.job_grade = mgr.job_grade and mgr.emp_id != t1.emp_id AND (grade.emp_level IN(9))
        AND mgr.manager_id IN(SELECT ted.emp_id FROM ts2_employee_details ted WHERE ted.manager_id in (Select emp_id from ts2_employee_details where emp_status = 'ACTIVE' and job_grade in ('E1')))
        AND ROWNUM = 1 and mgr.emp_id != t1.emp_id START WITH mgr.emp_id = t1.emp_id CONNECT BY NOCYCLE PRIOR mgr.manager_id = mgr.emp_id) 
        ELSE
        CASE WHEN
        (SELECT mgr.emp_id FROM ts2_employee_details  mgr, employeesgradetolevel grade WHERE grade.job_grade = mgr.job_grade and mgr.emp_id != t1.emp_id AND (grade.emp_level IN(8))
        AND mgr.manager_id IN(SELECT ted.emp_id FROM ts2_employee_details ted WHERE ted.manager_id in (Select emp_id from ts2_employee_details where emp_status = 'ACTIVE' and job_grade in ('E1')))
        AND ROWNUM = 1 and mgr.emp_id != t1.emp_id START WITH mgr.emp_id = t1.emp_id CONNECT BY NOCYCLE PRIOR mgr.manager_id = mgr.emp_id)
        IS NOT NULL
        THEN
        (SELECT mgr.emp_id FROM ts2_employee_details  mgr, employeesgradetolevel grade WHERE grade.job_grade = mgr.job_grade and mgr.emp_id != t1.emp_id AND (grade.emp_level IN(8))
        AND mgr.manager_id IN(SELECT ted.emp_id FROM ts2_employee_details ted WHERE ted.manager_id in (Select emp_id from ts2_employee_details where emp_status = 'ACTIVE' and job_grade in ('E1')))
        AND ROWNUM = 1 and mgr.emp_id != t1.emp_id START WITH mgr.emp_id = t1.emp_id CONNECT BY NOCYCLE PRIOR mgr.manager_id = mgr.emp_id)
        ELSE
        CASE WHEN
        (SELECT mgr.emp_id FROM ts2_employee_details  mgr, employeesgradetolevel grade WHERE grade.job_grade = mgr.job_grade and mgr.emp_id != t1.emp_id AND (grade.emp_level IN(7))
        AND mgr.manager_id IN(SELECT ted.emp_id FROM ts2_employee_details ted WHERE ted.manager_id in (Select emp_id from ts2_employee_details where emp_status = 'ACTIVE' and job_grade in ('E1')))
        AND ROWNUM = 1 and mgr.emp_id != t1.emp_id START WITH mgr.emp_id = t1.emp_id CONNECT BY NOCYCLE PRIOR mgr.manager_id = mgr.emp_id)
        IS NOT NULL
        THEN
        (SELECT mgr.emp_id FROM ts2_employee_details  mgr, employeesgradetolevel grade WHERE grade.job_grade = mgr.job_grade and mgr.emp_id != t1.emp_id AND (grade.emp_level IN(7))
        AND mgr.manager_id IN(SELECT ted.emp_id FROM ts2_employee_details ted WHERE ted.manager_id in (Select emp_id from ts2_employee_details where emp_status = 'ACTIVE' and job_grade in ('E1')))
        AND ROWNUM = 1 and mgr.emp_id != t1.emp_id START WITH mgr.emp_id = t1.emp_id CONNECT BY NOCYCLE PRIOR mgr.manager_id = mgr.emp_id)
        ELSE
        (SELECT mgr.emp_id FROM ts2_employee_details  mgr, employeesgradetolevel grade WHERE grade.job_grade = mgr.job_grade and mgr.emp_id != t1.emp_id AND (grade.emp_level IN(11))
        AND mgr.manager_id IN(SELECT ted.emp_id FROM ts2_employee_details ted WHERE ted.manager_id in (Select emp_id from ts2_employee_details where emp_status = 'ACTIVE' and job_grade in ('E1')))
        AND ROWNUM = 1 and mgr.emp_id != t1.emp_id START WITH mgr.emp_id = t1.emp_id CONNECT BY NOCYCLE PRIOR mgr.manager_id = mgr.emp_id)
        END
        END
        END 
        END
    ), 0)                      AS direct_report_id,
    nvl((
        CASE WHEN (
        SELECT mgr.emp_name FROM ts2_employee_details  mgr, employeesgradetolevel grade WHERE grade.job_grade = mgr.job_grade and mgr.emp_id != t1.emp_id AND(grade.emp_level IN(10))
        AND mgr.manager_id IN(SELECT ted.emp_id FROM ts2_employee_details ted WHERE ted.manager_id in (Select emp_id from ts2_employee_details where emp_status = 'ACTIVE' and job_grade in ('E1')))
        AND ROWNUM = 1 and mgr.emp_id != t1.emp_id START WITH mgr.emp_id = t1.emp_id CONNECT BY NOCYCLE PRIOR mgr.manager_id = mgr.emp_id)
        IS NOT NULL 
        THEN
        (SELECT mgr.emp_name FROM ts2_employee_details  mgr, employeesgradetolevel grade WHERE grade.job_grade = mgr.job_grade and mgr.emp_id != t1.emp_id AND(grade.emp_level IN(10))
        AND mgr.manager_id IN(SELECT ted.emp_id FROM ts2_employee_details ted WHERE ted.manager_id in (Select emp_id from ts2_employee_details where emp_status = 'ACTIVE' and job_grade in ('E1')))
        AND ROWNUM = 1 and mgr.emp_id != t1.emp_id START WITH mgr.emp_id = t1.emp_id CONNECT BY NOCYCLE PRIOR mgr.manager_id = mgr.emp_id)
        ELSE
        CASE WHEN (
        SELECT mgr.emp_name FROM ts2_employee_details  mgr, employeesgradetolevel grade WHERE grade.job_grade = mgr.job_grade and mgr.emp_id != t1.emp_id AND(grade.emp_level IN(9))
        AND mgr.manager_id IN(SELECT ted.emp_id FROM ts2_employee_details ted WHERE ted.manager_id in (Select emp_id from ts2_employee_details where emp_status = 'ACTIVE' and job_grade in ('E1')))
        AND ROWNUM = 1 and mgr.emp_id != t1.emp_id START WITH mgr.emp_id = t1.emp_id CONNECT BY NOCYCLE PRIOR mgr.manager_id = mgr.emp_id)
        IS NOT NULL 
        THEN
        (SELECT mgr.emp_name FROM ts2_employee_details  mgr, employeesgradetolevel grade WHERE grade.job_grade = mgr.job_grade and mgr.emp_id != t1.emp_id AND(grade.emp_level IN(9))
        AND mgr.manager_id IN(SELECT ted.emp_id FROM ts2_employee_details ted WHERE ted.manager_id in (Select emp_id from ts2_employee_details where emp_status = 'ACTIVE' and job_grade in ('E1')))
        AND ROWNUM = 1 and mgr.emp_id != t1.emp_id START WITH mgr.emp_id = t1.emp_id CONNECT BY NOCYCLE PRIOR mgr.manager_id = mgr.emp_id)
        ELSE
        CASE WHEN
        (SELECT mgr.emp_name FROM ts2_employee_details  mgr, employeesgradetolevel grade WHERE grade.job_grade = mgr.job_grade and mgr.emp_id != t1.emp_id AND(grade.emp_level IN(8))
        AND mgr.manager_id IN(SELECT ted.emp_id FROM ts2_employee_details ted WHERE ted.manager_id in (Select emp_id from ts2_employee_details where emp_status = 'ACTIVE' and job_grade in ('E1')))
        AND ROWNUM = 1 and mgr.emp_id != t1.emp_id START WITH mgr.emp_id = t1.emp_id CONNECT BY NOCYCLE PRIOR mgr.manager_id = mgr.emp_id)
        IS NOT NULL 
        THEN
        (SELECT mgr.emp_name FROM ts2_employee_details  mgr, employeesgradetolevel grade WHERE grade.job_grade = mgr.job_grade and mgr.emp_id != t1.emp_id AND(grade.emp_level IN(8))
        AND mgr.manager_id IN(SELECT ted.emp_id FROM ts2_employee_details ted WHERE ted.manager_id in (Select emp_id from ts2_employee_details where emp_status = 'ACTIVE' and job_grade in ('E1')))
        AND ROWNUM = 1 and mgr.emp_id != t1.emp_id START WITH mgr.emp_id = t1.emp_id CONNECT BY NOCYCLE PRIOR mgr.manager_id = mgr.emp_id)
        ELSE
        CASE WHEN
        (SELECT mgr.emp_name FROM ts2_employee_details  mgr, employeesgradetolevel grade WHERE grade.job_grade = mgr.job_grade and mgr.emp_id != t1.emp_id AND(grade.emp_level IN(7))
        AND mgr.manager_id IN(SELECT ted.emp_id FROM ts2_employee_details ted WHERE ted.manager_id in (Select emp_id from ts2_employee_details where emp_status = 'ACTIVE' and job_grade in ('E1')))
        AND ROWNUM = 1 and mgr.emp_id != t1.emp_id START WITH mgr.emp_id = t1.emp_id CONNECT BY NOCYCLE PRIOR mgr.manager_id = mgr.emp_id)
        IS NOT NULL 
        THEN
        (SELECT mgr.emp_name FROM ts2_employee_details  mgr, employeesgradetolevel grade WHERE grade.job_grade = mgr.job_grade and mgr.emp_id != t1.emp_id AND(grade.emp_level IN(7))
        AND mgr.manager_id IN(SELECT ted.emp_id FROM ts2_employee_details ted WHERE ted.manager_id in (Select emp_id from ts2_employee_details where emp_status = 'ACTIVE' and job_grade in ('E1')))
        AND ROWNUM = 1 and mgr.emp_id != t1.emp_id START WITH mgr.emp_id = t1.emp_id CONNECT BY NOCYCLE PRIOR mgr.manager_id = mgr.emp_id)
        ELSE
        (SELECT mgr.emp_name FROM ts2_employee_details  mgr, employeesgradetolevel grade WHERE grade.job_grade = mgr.job_grade and mgr.emp_id != t1.emp_id AND(grade.emp_level IN(11))
        AND mgr.manager_id IN(SELECT ted.emp_id FROM ts2_employee_details ted WHERE ted.manager_id in (Select emp_id from ts2_employee_details where emp_status = 'ACTIVE' and job_grade in ('E1')))
        AND ROWNUM = 1 and mgr.emp_id != t1.emp_id START WITH mgr.emp_id = t1.emp_id CONNECT BY NOCYCLE PRIOR mgr.manager_id = mgr.emp_id)
        END
        END
        END 
        END
    ), 'NA')                   AS direct_report,                       



                        t2.swipe_month as month,COALESCE(t2.no_of_days , 0) AS no_of_days
                        FROM ts2_employee_details t1
                        LEFT JOIN
                        (
                            Select emp_id, swipe_month, COUNT(*) AS no_of_days from (
                            Select distinct bat.emp_id, TO_CHAR(TRUNC(acd.ACCESS_TIME), 'Mon-yyyy') as swipe_month, TRUNC(acd.ACCESS_TIME) AS swipe from rm_employee_access_details acd, rm_emp_batch_id bat, ts2_employee_details emp 
                            where acd.batch_id = bat.batch_id and emp.emp_id = bat.emp_id
                            and (acd.access_time >= to_date('" + startdate.ToShortDateString() + "','mm/dd/yyyy') and acd.access_time <= to_date('" + endDate.ToShortDateString() + "', 'mm/dd/yyyy')) and (TRUNC(acd.access_time) <= TRUNC(emp.termination_date) or emp.termination_date is null)" +
                            "UNION" +
                            " Select distinct temp.employee_id as emp_id, TO_CHAR(TRUNC(acd.ACCESS_TIME), 'Mon-yyyy') as swipe_month, TRUNC(acd.ACCESS_TIME) AS swipe" +
                            " from rm_employee_access_details acd,rm_temp_card_details temp, ts2_employee_details emp" +
                            " where (acd.batch_id = temp.temp_card_id" +
                            " and emp.emp_id = temp.employee_id)" +
                            " and(acd.access_time >= to_date('" + startdate.ToShortDateString() + "', 'mm/dd/yyyy') and acd.access_time <= to_date('" + endDate.ToShortDateString() + "', 'mm/dd/yyyy'))" +
                            " AND acd.access_time BETWEEN temp.missed_date AND  TO_DATE(TO_CHAR(temp.returndate + 1, 'mm/dd/yyyy'), 'mm/dd/yyyy') and (TRUNC(acd.access_time) <= TRUNC(emp.termination_date) or emp.termination_date is null) " + filterOfficeDutyOutsideOffice +
                            ") group by emp_id,swipe_month" +
                        ") t2 " +
                        "ON t1.emp_id = t2.EMP_ID, ts2_employee_details sup, ts2_employee_details mgr " +
                        "where t1.manager_id = sup.emp_id and mgr.emp_id = sup.manager_id  and lower(t1.cuid) not like ('x%') " + filterInActiveEmp + " " + filterLoggedInEmp + "";
            return sql;
        }



        [HttpGet("WFHReport")]
        public async Task<IActionResult> GetWFHReport(
    [FromQuery] DateTime startDate,
    [FromQuery] DateTime endDate)
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
                    command.CommandText = @" SELECT adetails.emp_id, EDETAILS.EMP_NAME, adetails.start_date, adetails.end_date, adetails.status, " +
                                           "adetails.Comments, adetails.applied_date, adetails.processed_date, adetails.approver_id, " +
                                           "CASE WHEN approver_id is NOT NULL THEN (SELECT emp_name FROM ts2_employee_Details WHERE emp_id = approver_id) END as ApproverName,adetails.APPROVAL_REJECTION_REASON " +
                                           "FROM RM_ATTENDANCE_DETAILS adetails INNER JOIN ts2_employee_details edetails ON adetails.emp_id = edetails.emp_id " +
                                           "WHERE ATTENDANCE_TYPE_ID =2 and lower(edetails.cuid) not like ('x%') and " +
                                           "START_DATE >= To_date('" + startDate.ToShortDateString() + "','MM/DD/YYYY') and END_DATE <= To_date('" + endDate.ToShortDateString() + "','MM/DD/YYYY') " +
                                           "order by EMP_ID,START_DATE";
                    command.CommandType = CommandType.Text;

                    adapter.Fill(dataSet);
                }

                // Check if the dataset is empty
                if (dataSet.Tables.Count == 0 || dataSet.Tables[0].Rows.Count == 0)
                {
                    return StatusCode(200, "No data available for the given date range.");
                }

                string json = JsonConvert.SerializeObject(dataSet.Tables[0]);

                // Convert JSON to DataTable
                DataTable dataTable = JsonConvert.DeserializeObject<DataTable>(json);

                // Generate Excel file with current date and time in the filename
                string currentDateTime = DateTime.Now.ToString("yyyyMMdd_HHmmss");
                string filePath = $"WFHReport_{currentDateTime}.xlsx";
                ExcelHelper.CreateExcelFromDataTable(dataTable, filePath);

                byte[] content = System.IO.File.ReadAllBytes(filePath);
                return File(content, "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", filePath);
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Error retrieving WFH report.");
                return StatusCode(500, "Internal server error");
            }
        }



        [HttpGet("WFHReport2")]
        public async Task<IActionResult> GetWFHReport2([FromQuery] DateTime startDate, [FromQuery] DateTime endDate)
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

                // Check if the dataset is empty
                if (dataSet.Tables.Count == 0 || dataSet.Tables[0].Rows.Count == 0)
                {
                    return StatusCode(200, "No data available for the given date range.");
                }

                // Convert DataTable to a list of objects with all required fields
                var dataTable = dataSet.Tables[0];
                var data = dataTable.AsEnumerable().Select(row => new
                {

                    EMP_ID = row["emp_id"] != DBNull.Value ? Convert.ToInt32(row["emp_id"]) : (int?)null,
                    EMP_NAME = row["EMP_NAME"] != DBNull.Value ? row["EMP_NAME"].ToString() : null,
                    START_DATE = row["start_date"] != DBNull.Value ? Convert.ToDateTime(row["start_date"]).ToString("yyyy-MM-dd") : null,
                    END_DATE = row["end_date"] != DBNull.Value ? Convert.ToDateTime(row["end_date"]).ToString("yyyy-MM-dd") : null,
                    STATUS = row["status"] != DBNull.Value ? row["status"].ToString() : null,
                    COMMENTS = row["Comments"] != DBNull.Value ? row["Comments"].ToString() : null,
                    APPLIED_DATE = row["applied_date"] != DBNull.Value ? Convert.ToDateTime(row["applied_date"]).ToString("yyyy-MM-dd") : null,
                    PROCESSED_DATE = row["processed_date"] != DBNull.Value ? Convert.ToDateTime(row["processed_date"]).ToString("yyyy-MM-dd") : null,
                    APPROVER_ID = row["approver_id"] != DBNull.Value ? Convert.ToInt32(row["approver_id"]) : (int?)null,
                    APPROVERNAME = row["ApproverName"] != DBNull.Value ? row["ApproverName"].ToString() : null,
                    APPROVAL_REJECTION_REASON = row["APPROVAL_REJECTION_REASON"] != DBNull.Value ? row["APPROVAL_REJECTION_REASON"].ToString() : null

                }).ToList();

                return Ok(new
                {
                    status = "success",
                    data = data
                });
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Error retrieving WFH report.");
                return StatusCode(500, new { message = "Internal server error" });
            }
        }



        [HttpGet("EscalatedEmailLog")]
        public async Task<IActionResult> GetEscalatedEmailLog(
         [FromQuery] string escalationEmailTemplate,
         [FromQuery] DateTime startDate,
         [FromQuery] DateTime endDate)
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
                    command.Parameters.Add(new OracleParameter("templateName", escalationEmailTemplate));
                    command.Parameters.Add(new OracleParameter("startDate", startDate));
                    command.Parameters.Add(new OracleParameter("endDate", endDate));
                    command.CommandType = CommandType.Text;

                    adapter.Fill(dataSet);
                }

                if (dataSet.Tables.Count == 0 || dataSet.Tables[0].Rows.Count == 0)
                {
                    return StatusCode(200, "No data available for the given date range.");
                }

                string json = JsonConvert.SerializeObject(dataSet.Tables[0]);
                DataTable dataTable = JsonConvert.DeserializeObject<DataTable>(json);

                string currentDateTime = DateTime.Now.ToString("yyyyMMdd_HHmmss");
                string filePath = $"EscalatedEmailLog_{currentDateTime}.xlsx";
                ExcelHelper.CreateExcelFromDataTable(dataTable, filePath);

                byte[] content = System.IO.File.ReadAllBytes(filePath);
                return File(content, "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", filePath);
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Error retrieving escalated email log.");
                return StatusCode(500, "Internal server error");
            }
        }


















    }


}
