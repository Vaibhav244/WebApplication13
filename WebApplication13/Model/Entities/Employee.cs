using System;

namespace WebApplication13.Models.Entities
{
    public class Employee
    {
        public int EmpId { get; set; }
        public string Cuid { get; set; }
        public string EmpName { get; set; }
        public string EmpStatus { get; set; }
        public string JobGrade { get; set; }
        public int? EmpLevel { get; set; }
        public string JobTitleFull { get; set; }
        public string City { get; set; }
        public int? ManagerId { get; set; }
        public DateTime? TerminationDate { get; set; }
        public string Gender { get; set; }
        public string EmailIdOff { get; set; }
        public string PsEmpType { get; set; }
    }
}