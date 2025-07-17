namespace BestPolicyReport.Models.BillReport
{
    public class InvoiceReportResult
    {
        public InvoiceReportResult() { }

        public InvoiceReportResult( List<InvoiceData?> datas , int count) 
            { 
                this.Count = count;
                this.Datas = datas;
            }
        public List<InvoiceData?> Datas { get; set; }
        public int Count { get; set; }
    }
    public class InvoiceData {


        public string? AgentCode { get; set; }
    public string? InsurerCode { get; set; }

        public string? InsureeName { get; set; } // 09-04-2025
        public string? ApplicationNo { get; set; }
    public string? PolicyNo { get; set; }
    public string? EndorseNo { get; set; }
    public string? InvoiceNo { get; set; }
    public int? SeqNo { get; set; }
    public int? Customerid { get; set; }
    public string? InsureeCode { get; set; }
    public string? Agentname { get; set; }
    public int? Polid { get; set; }
    public double? Grossprem { get; set; }
    public double? Specdiscrate { get; set; }
    public double? Specdiscamt { get; set; }
    public double? Netgrossprem { get; set; }
    public double? Duty { get; set; }
    public double? Tax { get; set; }
    public double? Totalprem { get; set; }
    public double? Withheld { get; set; }
    public double? Totalamt { get; set; }
    public double? Commout_rate { get; set; }
    public double? Commout_amt { get; set; }
    public double? Ovout_rate { get; set; }
    public double? Ovout_amt { get; set; }
    public double? Commout1_rate { get; set; }
    public double? Commout1_amt { get; set; }
    public double? Ovout1_rate { get; set; }
    public double? Ovout1_amt { get; set; }
    public DateTime? DueDate { get; set; }

    public string? Dfrpreferno { get; set; }

    public DateTime? Rprefdate { get; set; }




}
}
