namespace BestPolicyReport.Models.ArApReport
{
    public class PremOutCommInOvInReportResult
    {
        public PremOutCommInOvInReportResult()
        {
            this.Count = 0;
        }
        public PremOutCommInOvInReportResult(List<PremOutCommInOvInData> data, int count)
        {
            this.Count = count;
            this.Datas = data;
        }
        public List<PremOutCommInOvInData> Datas { get; set; }
        public int Count { get; set; }

    }
    
    public class PremOutCommInOvInData: ArApOutPut
    {
        public string? AgentCode { get; set; }
        public string? TaxInvoiceNo { get; set; }
        public string? PremOutDfRpReferNo { get; set; }
        public DateTime? PremOutRpRefDate { get; set; }
        public string? InsurerName { get; set; }
        public double? CommInRate { get; set; }
        public double? CommInAmt { get; set; }
        public double? CommOvInRate { get; set; }
        public double? CommOvInAmt { get; set; }
        public string? CommInDfRpReferNo { get; set; }
        public DateTime? CommInRpRefDate { get; set; }
        public double? CommInPaidAmt { get; set; }
        public double? CommInDiffAmt { get; set; }
        public double? OvInRate { get; set; }
        public double? OvInAmt { get; set; }
        public string? OvInDfRpReferNo { get; set; }
        public DateTime? OvInRpRefDate { get; set; }
        public double? OvInPaidAmt { get; set; }
        public double? OvInDiffAmt { get; set; }
    }
}
