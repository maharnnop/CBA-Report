using System.ComponentModel.DataAnnotations.Schema;

namespace BestPolicyReport.Models.ArApReport
{
    public class ArApReportInput
    {
        public string? StartPolicyIssueDate { get; set; }
        public string? EndPolicyIssueDate { get; set; }
        public string? AsAtDate { get; set; }
        public string? CreateUserCode { get; set; }
        public string? MainAccountContactPersonId { get; set; }
        public string? AgentCode { get; set; }
        public string? MainAccountCode { get; set; }
        public string? InsurerCode { get; set; }
        public string? Insurancestatus { get; set; }
        public string? Class { get; set; }
        public string? SubClass { get; set; }
        public string? TransactionType { get; set; }

        public string? PreminRprefdatestart { get; set; }
        public string? PreminRprefdateend { get; set; }
        public string? PremoutRprefdatestart { get; set; }
        public string? PremoutRprefdateend { get; set; }
        public int? Limit { get; set; }
        public int? Pagecount { get; set;}


    }
    public class ArApOutPut
    {
        public string? PolicyNo { get; set; }
        public string? EndorseNo { get; set; }
        public string? InvoiceNo { get; set; }
        public string? TransType { get; set; }
        public int? SeqNo { get; set; }
        public string? CashierReceiveNo { get; set; }
        public DateTime? CashierDate { get; set; }
        public double? CashierAmt { get; set; }
        public string? CashierReceiveType { get; set; }
        public string? CashierRefNo { get; set; }
        [NotMapped]
        public DateTime? CashierRefDate { get; set; }
        public string? PremInDfRpReferNo { get; set; }
        public DateTime? PremInRpRefDate { get; set; }
        public string? DfRpReferNo { get; set; }
        public DateTime? RpRefDate { get; set; }
        public double? GrossPrem { get; set; }
        public double? SpecDiscRate { get; set; }
        public double? SpecDiscAmt { get; set; }
        public double? NetGrossPrem { get; set; }
        public double? Duty { get; set; }
        public double? Tax { get; set; }
        public double? TotalPrem { get; set; }
        public string? NetFlag { get; set; }
        public DateTime? ActDate { get; set; }
        public DateTime? ExpDate { get; set; }
        public string? MainAccountCode { get; set; }
        public string? MainAccountName { get; set; }
        public string? InsureeCode { get; set; }
        public string? InsureeName { get; set; }
        public string? Class { get; set; }
        public string? SubClass { get; set; }
        public double? CommOutRate { get; set; }
        public double? CommOutAmt { get; set; }
        public double? OvOutRate { get; set; }
        public double? OvOutAmt { get; set; }
       
        public DateTime? IssueDate { get; set; }
        public string? PolicyCreateUserCode { get; set; }
        public int? MainAccountContactPersonId { get; set; }
        public string? InsurerCode { get; set; }
        public string? Insurancestatus { get; set; }
        public string? TransactionType { get; set; }
    }

    public class DatasCount
    {
        public int Count { get; set; }
    }
}
