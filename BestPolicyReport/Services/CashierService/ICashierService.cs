using BestPolicyReport.Models.CashierReport;

namespace BestPolicyReport.Services.CashierService
{
    public interface ICashierService
    {
        Task<CashierReportResult?> GetCashierReportJson(CashierReportInput data);
    }
}
