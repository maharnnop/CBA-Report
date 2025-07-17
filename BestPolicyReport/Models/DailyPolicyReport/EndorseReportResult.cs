namespace BestPolicyReport.Models.DailyPolicyReport
{
    public class EndorseReportResult
    {
       public EndorseReportResult() { }
        public EndorseReportResult(List<EndorseData>? datas, int count) {
            this.Datas = datas;
            this.Count = count; 
        }
        public List<EndorseData>? Datas { get; set; }
        public int Count { get; set; }

        
    }

    public class EndorseData
    {
     public string? ApplicationNo { get; set; }
    public string? PolicyNo { get; set; }
        public string? EndorseType { get; set; }
        public DateTime? EndorseDate { get; set; }
        public string? EndorseNo { get; set; }
    public int? Endorseseries { get; set; }

    public string? Edtypecode { get; set; }
    public string? Edtitle { get; set; }
    public string? Eddetail { get; set; }

    public double? Ednetgrossprem { get; set; }
    public double? Edduty { get; set; }
    public double? Edtax { get; set; }
    public double? Edtotalprem { get; set; }
}
}
