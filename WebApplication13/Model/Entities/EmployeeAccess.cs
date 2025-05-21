using System;

namespace WebApplication13.Models.Entities
{
    public class EmployeeAccess
    {
        public string BatchId { get; set; }
        public int EmpId { get; set; }
        public DateTime AccessTime { get; set; }
        public string BatchReaderLoc { get; set; }
        public string SwipeMonth { get; set; }
    }
}