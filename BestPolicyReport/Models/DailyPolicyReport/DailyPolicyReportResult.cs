namespace BestPolicyReport.Models.DailyPolicyReport
{
    public class DailyPolicyReportResult
    {
       public DailyPolicyReportResult() { }
        public DailyPolicyReportResult(List<DailyPolicyData>? datas, int count) {
            this.Datas = datas; 
            this.Count = count; 
        }
       
       public List<DailyPolicyData>? Datas { get; set; }
        public int Count { get; set; }  
    }
    public class DailyPolicyData
    {
        public string? ApplicationNo { get; set; }
        public string? PolicyNo { get; set; }
        public DateTime? PolicyDate { get; set; }
        public DateTime? ActDate { get; set; }
        public DateTime? ExpDate { get; set; }
        public DateTime? IssueDate { get; set; }
        public string? CreateUserCode { get; set; }
        public string? Username { get; set; }
        public int? ContactPersonId1 { get; set; }
        public string? ContactPersonName1 { get; set; }
        public int? ContactPersonId2 { get; set; }
        public string? ContactPersonName2 { get; set; }
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
        public double? GrossPrem { get; set; }
        public double? SpecDiscRate { get; set; }
        public double? SpecDiscAmt { get; set; }
        public double? NetGrossPrem { get; set; }
        public double? Duty { get; set; }
        public double? Tax { get; set; }
        public double? TotalPrem { get; set; }
        public double? Withheld { get; set; }
        public double? TotalAmt { get; set; }
        public double? CommInRate { get; set; }
        public double? CommInAmt { get; set; }
        public double? CommInTaxAmt { get; set; }
        public double? OvInRate { get; set; }
        public double? OvInAmt { get; set; }
        public double? OvInTaxAmt { get; set; }
        public double? CommOutRate { get; set; }
        public double? CommOutAmt { get; set; }
        public double? OvOutRate { get; set; }
        public double? OvOutAmt { get; set; }
        public string? InsurerCode { get; set; }
    }
}
