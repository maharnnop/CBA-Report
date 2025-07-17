namespace BestPolicyReport.Models.BillReport
{
    public class InvoiceReportInput
    {
        public string? AgentCode { get; set; }
        public string? CreatedDateStart { get; set; }
        public string? CreatedDateEnd { get; set; }
        public string? DueDateStatus { get; set; }
        public string? EndInvoiceNo { get; set; }

        public string? EndorseNo { get; set; }
        public string? Insuree_fn { get; set; }
        public string? Insuree_ln { get; set; }
        public string? Insuree_ogn { get; set; }
        public string? InsurerCode { get; set; }
        public string? LicenseNo { get; set; }
        public int? Limit { get; set; }
        public int? Pagecount { get; set; }

        public string? PersonType { get; set; }
        public string? PolicyNo { get; set; }
        public string? StartInvoiceNo { get; set; }

        public string? Status { get; set; }

    }
}
