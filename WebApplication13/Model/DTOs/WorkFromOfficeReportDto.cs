using System;

namespace WebApplication13.Models.DTOs
{
    public class WorkFromOfficeReportDto
    {
        public string Month { get; set; }
        public string AccessTime { get; set; }
        public string FirstLogin { get; set; }
        public string LastLogout { get; set; }
        public string TotalDuration { get; set; }
        public int EmpId { get; set; }
        public string EmpName { get; set; }
        public string EmpStatus { get; set; }
        public string JobGrade { get; set; }
        public int ReportingToEmpId { get; set; }
        public string ReportingTo { get; set; }
        public int ManagerId { get; set; }
        public string Manager { get; set; }
        public int DirectorId { get; set; }
        public string Director { get; set; }
        public int DirectReportId { get; set; }
        public string DirectReport { get; set; }
        public int VpId { get; set; }
        public string Vp { get; set; }
        public string Gender { get; set; }
        public string Location { get; set; }
        public string EmployeeEmail { get; set; }
        public string EmployeeType { get; set; }
        public string JobTitle { get; set; }
    }
}