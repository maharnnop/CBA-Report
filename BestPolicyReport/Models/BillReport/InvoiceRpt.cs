namespace BestPolicyReport.Models.BillReport
{
    public class InvoiceRpt
    {
        public int Id { get; set; }
        public string? invoiceNo { get; set; }
        public DateOnly dueDate { get; set; }
        public string? insureeName { get; set; }
        public string? insureeCode { get; set; }
        public string? insureeLocation { get; set; }

        public string? insurerName { get; set; }
        public string? insureName { get; set; }
        public string? policyNo { get; set; }
        public DateOnly actDate { get; set; }
        public DateOnly expDate { get; set; }
        public string? endorseNo { get; set; }
        public double? cover_amt { get; set; }
        public double? netgrossprem { get; set; }

        public double? duty { get; set; }
        public double? tax { get; set; }
        public double? totalprem { get; set; }

        public double? specdiscamt { get; set; }
        public double? totalamt { get; set; }
        public double? seqamt { get; set; }

    }
}
