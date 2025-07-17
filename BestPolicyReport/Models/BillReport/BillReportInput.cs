namespace BestPolicyReport.Models.BillReport
{
    public class BillReportInput
    {
        public string? InsurerCode { get; set; }
        public string? AgentCode1 { get; set; }
        public string? AgentCode2 { get; set; }
        public string? StartBillAdvisorNo { get; set; }
        public string? EndBillAdvisorNo { get; set; }
        public string? StartBillDate { get; set; }
        public string? EndBillDate { get; set; }
        public int? Limit { get; set; }
        public int? Pagecount { get; set; }
    }
}
