namespace BestPolicyReport.Models.DailyPolicyReport
{
    public class EndorseReportInput
    {
        public string? StartedDate { get; set; }
        public string? EndedDate { get; set; }
        public string? StartpolicyNo { get; set; }
        public string? EndpolicyNo { get; set; }
        public string? StartendorseNo { get; set; }
        public string? EndendorseNo { get; set; }
        public string? Edtypecode { get; set; }
        public string? Class { get; set; }
        public string? SubClass { get; set; }
        public int? Limit { get; set; }
        public int? Pagecount { get; set; }
    }
}
