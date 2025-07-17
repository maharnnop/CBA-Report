namespace BestPolicyReport.Models.DailyPolicyReport
{
    public class InvoiceReportInput
    {
        public string? StartPolicyDate { get; set; }
        public string? EndPolicyDate { get; set; }
        public string? CreateUserCode { get; set; }
        public string? ContactPersonId1 { get; set; }
        public string? ContactPersonId2 { get; set; }
        public string? AgentCode1 { get; set; }
        public string? AgentCode2 { get; set; }
        public string? InsurerCode { get; set; }
        public string? Status { get; set; }
        public string? Class { get; set; }
        public string? SubClass { get; set; }

        public string? InstallmentType { get; set; }
        public string? OrderBy { get; set; }
        public int? Limit { get; set; }
        public int? Pagecount { get; set; }
    }
}
