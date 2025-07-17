using BestPolicyReport.Models.BillReport;
using Microsoft.AspNetCore.Mvc;

namespace BestPolicyReport.Services.BillService
{
    public interface IBillService
    {
        Task<BillReportResult?> GetBillReportJson(BillReportInput data);
        Task<List<PolicyGroupBillReportResult>?> GetPolicyGroupBillReportJson(PolicyGroupBillReportInput data);

        Task<InvoiceReportResult>? GetInvoiceJson(InvoiceReportInput data);
        Task<InvoiceRpt>? GetInvoicerpt(string invoiceNo, string usercode);

    }
}
