using System;

namespace WebApplication13.Models.Entities
{
    public class MailLog
    {
        public int EmpId { get; set; }
        public string MailTo { get; set; }
        public string MailCc { get; set; }
        public DateTime UpdatedOn { get; set; }
        public string TemplateName { get; set; }
    }
}