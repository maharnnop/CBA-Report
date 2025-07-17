namespace BestPolicyReport.Models.WhtReport
{
    public class WhtReportInput
    {
        public string? AgentCode { get; set; }
        public string? StartRpRefDate { get; set; }
        public string? EndRpRefDate { get; set; }
        public int? Limit { get; set; }
        public int? Pagecount { get; set; }
    }
}
