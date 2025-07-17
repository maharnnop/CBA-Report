namespace BestPolicyReport.Models.DailyPolicyReport
{
    public class InvoiceReportResultXX
    {
        public string? ApplicationNo { get; set; }
        public string? PolicyNo { get; set; }
        public string? PolicyType { get; set; }
        public string? InsurerCode { get; set; }
        public string? InsurerName { get; set; }
        public string? TaxInvoiceNo { get; set; }
        public string? InvoiceNo { get; set; }
        public string? SeqNo { get; set; }
        public DateTime? GenerateDate { get; set; }
        public DateTime? DueDate { get; set; }
        public DateTime? PolicyDate { get; set; }
        public DateTime? ActDate { get; set; }
        public DateTime? ExpDate { get; set; }
        public DateTime? IssueDate { get; set; }
        public string? CreateUserCode { get; set; }
        public string? Username { get; set; }
        public int? ContactPersonId1 { get; set; }
        public string? ContactPersonName1 { get; set; }
        //public int? ContactPersonId2 { get; set; }
        //public string? ContactPersonName2 { get; set; }
        public string? AgentCode1 { get; set; }
        public string? AgentName1 { get; set; }
        public string? AgentCode2 { get; set; }
        public string? AgentName2 { get; set; }
        public string? InsureeCode { get; set; }
        public string? InsureeName { get; set; }
        public string? Class { get; set; }
        public string? SubClass { get; set; }
        public string? LicenseNo { get; set; }
        public string? Province { get; set; }
        public string? ChassisNo { get; set; }
        public double? POLGrossPrem { get; set; }
        public double? SpecDiscRate { get; set; }
        public double? POLSpecDiscAmt { get; set; }
        public double? POLNetGrossPrem { get; set; }
        public double? POLDuty { get; set; }
        public double? POLTax { get; set; }
        public double? POLTotalPrem { get; set; }
        public double? INVGrossPrem { get; set; }
        public double? INVSpecDiscAmt { get; set; }
        public double? INVNetGrossPrem { get; set; }
        public double? INVDuty { get; set; }
        public double? INVTax { get; set; }
        public double? INVTotalPrem { get; set; }
        public double? CommInRate { get; set; }
        public double? INVCommInAmt { get; set; }
        public double? INVCommInTaxAmt { get; set; }
        public double? OvInRate { get; set; }
        public double? INVOvInAmt { get; set; }
        public double? INVOvInTaxAmt { get; set; }
        public double? CommOutRate { get; set; }
        public double? INVCommOutAmt { get; set; }
        public double? OvOutRate { get; set; }
        public double? INVOvOutAmt { get; set; }
       
    }
}
