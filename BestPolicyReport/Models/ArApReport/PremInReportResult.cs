using System.ComponentModel.DataAnnotations.Schema;

namespace BestPolicyReport.Models.ArApReport
{
    public class PremInReportResult
    {
        public PremInReportResult()
        {
            this.Count = 0;
        }
        public PremInReportResult(List<PremInData> data, int count)
        {
            this.Count = count;
            this.Datas = data;
        }
        public List<PremInData> Datas { get; set; }
        public int Count { get; set; }

    }

    public class PremInData : ArApOutPut
    {
        public string? Billadvisorno { get; set; }
        public DateTime? Billdate { get; set; }
        public double? PremInPaidAmt { get; set; }
        public double? PremInDiffAmt { get; set; }
        public double? CommOutRate1 { get; set; }
        public double? CommOutAmt1 { get; set; }
        public double? OvOutRate1 { get; set; }
        public double? OvOutAmt1 { get; set; }
        public double? CommOutRate2 { get; set; }
        public double? CommOutAmt2 { get; set; }
       
    }
}
