using System.ComponentModel.DataAnnotations.Schema;

namespace BestPolicyReport.Models.ArApReport
{
    public class CommOutOvOutReportResult
    {
        public CommOutOvOutReportResult()
        {
            this.Count = 0;
        }
        public CommOutOvOutReportResult(List<CommOutOvOutData> data, int count)
        {
            this.Count = count;
            this.Datas = data;
        }
        public List<CommOutOvOutData> Datas { get; set; }
        public int Count { get; set; }

    }
    public class CommOutOvOutData : ArApOutPut
        {
        public string? LicenseNo { get; set; }
        public string? Province { get; set; }
        public string? ChassisNo { get; set; }
        public double? CommOvOutRate { get; set; }
        public double? CommOvOutAmt { get; set; }
        public string? CommOutDfRpReferNo { get; set; }
        public DateTime? CommOutRpRefDate { get; set; }
        public double? CommOutPaidAmt { get; set; }
        public double? CommOutDiffAmt { get; set; }
        public double? OvOutPaidAmt { get; set; }
        public double? OvOutDiffAmt { get; set; }
       
    }

    
}
