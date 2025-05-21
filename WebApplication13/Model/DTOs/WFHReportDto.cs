using System;

namespace WebApplication13.Models.DTOs
{
    public class WFHReportDto
    {
        public int? EmpId { get; set; }
        public string EmpName { get; set; }
        public string StartDate { get; set; }
        public string EndDate { get; set; }
        public string Status { get; set; }
        public string Comments { get; set; }
        public string AppliedDate { get; set; }
        public string ProcessedDate { get; set; }
        public int? ApproverId { get; set; }
        public string ApproverName { get; set; }
        public string ApprovalRejectionReason { get; set; }
    }
}