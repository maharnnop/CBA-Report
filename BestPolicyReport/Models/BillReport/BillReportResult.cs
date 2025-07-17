namespace BestPolicyReport.Models.BillReport
{
    public class BillReportResult
    {
        public BillReportResult() { }
        public BillReportResult(List<BillData> datas, int count ) {
            this.Datas = datas;
            this.Count = count;
        }
        public List<BillData>? Datas { get; set; }
        public int Count { get; set; }
    }

    public class BillData {

        public string? PolicyNo { get; set; }
        public string? EndorseNo { get; set; }
        public DateTime? IssueDate { get; set; }
        public DateTime? ActDate { get; set; }
        public DateTime? ExpDate { get; set; }
        public string? InvoiceNo { get; set; }
        public string? TaxInvoiceNo { get; set; }
        public string? InvoiceName { get; set; }
        public string? InsurerCode { get; set; }
        public string? Class { get; set; }
        public string? SubClass { get; set; }
        public string? AgentCode1 { get; set; }
        public string? AgentCode2 { get; set; }
        public int? SeqNo { get; set; }
        public string? InsureeCode { get; set; }
        public string? InsureeName { get; set; }

        public string? VoluntaryCode { get; set; }
        public string? CompulsoryCode { get; set; }
        public string? Unregisterflag { get; set; }
        public string? LicenseNo { get; set; }
        public string? Province { get; set; }
        public string? ChassisNo { get; set; }
        public string? Addition_access { get; set; }
        public string? EngineNo { get; set; }
        public int? ModelYear { get; set; }
        public string? Brand { get; set; }
        public string? Model { get; set; }
        public string? Specname { get; set; }
        public int? Cc { get; set; }
        public int? Seat { get; set; }
        public int? Gvw { get; set; }

        public DateTime? DueDate { get; set; }
        public double? GrossPrem { get; set; }
        public double? SpecDiscRate { get; set; }
        public double? SpecDiscAmt { get; set; }
        public double? NetGrossPrem { get; set; }
        public double? Duty { get; set; }
        public double? Tax { get; set; }
        public double? TotalPrem { get; set; }
        public double? CommOutRate1 { get; set; }
        public double? CommOutAmt1 { get; set; }
        public double? OvOutRate1 { get; set; }
        public double? OvOutAmt1 { get; set; }
        public double? CommOutRate2 { get; set; }
        public double? CommOutAmt2 { get; set; }
        public double? OvOutRate2 { get; set; }
        public double? OvOutAmt2 { get; set; }
        public double? CommOutRate { get; set; }
        public double? CommOutAmt { get; set; }
        public double? OvOutRate { get; set; }
        public double? OvOutAmt { get; set; }
        public string? NetFlag { get; set; }
        public double? BillPremium { get; set; }
        public string? BillAdvisorNo { get; set; }
        public DateTime? BillDate { get; set; }

        public string? Beneficiary { get; set; }
    }
}
