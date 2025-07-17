using BestPolicyReport.Models.DailyPolicyReport;

namespace BestPolicyReport.Services.DailyPolicyService
{
    public interface IOICService
    {
        Task<DailyPolicyReportResult?> GetDailyPolicyReportJson(DailyPolicyReportInput data);
        Task<EndorseReportResult?> GetEndorseReportJson(EndorseReportInput data);

    }

    
}
