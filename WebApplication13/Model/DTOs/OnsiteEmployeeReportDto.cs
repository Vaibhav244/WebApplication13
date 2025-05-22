using System;

namespace YourCompany.ResourceManagement.Models
{
    public class OnsiteEmployeeReportDto
    {
        public int EmpId { get; set; }
        public string EmpName { get; set; }
        public string JobGrade { get; set; }
        public string JobTitle { get; set; }
        public DateTime OnsiteStartDate { get; set; }
        public DateTime OnsiteEndDate { get; set; }
        public string Location { get; set; }
        public string Status { get; set; }
        public int? AttendanceTypeId { get; set; }
        public decimal? FeedbackScore { get; set; }
        public int? ManagerId { get; set; }
        public string ManagerName { get; set; }
        public string EmailId { get; set; }
        public string Gender { get; set; }
    }
}