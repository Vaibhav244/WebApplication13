using System;

namespace WebApplication13.Models.Entities
{
    public class AttendanceDetail
    {
        public int EmpId { get; set; }
        public DateTime StartDate { get; set; }
        public DateTime EndDate { get; set; }
        public string Status { get; set; }
        public string Comments { get; set; }
        public DateTime AppliedDate { get; set; }
        public DateTime? ProcessedDate { get; set; }
        public int? ApproverId { get; set; }
        public string ApproverName { get; set; }
        public string ApprovalRejectionReason { get; set; }
        public int AttendanceTypeId { get; set; }
    }
}