using BestPolicyReport.Models.PremInDirectReport;

namespace BestPolicyReport.Services.PremInDirectService
{
    public interface IPremInDirectService
    {
        Task<PremInDirectReportResult?> GetPremInDirectReportJson(PremInDirectReportInput data);
    }
}
