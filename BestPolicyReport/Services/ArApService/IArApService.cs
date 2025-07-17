using BestPolicyReport.Models.ArApReport;

namespace BestPolicyReport.Services.ArApService
{
    public interface IArApService
    {
        Task<PremInReportResult?> GetPremInOpenItemReportJson(ArApReportInput data);
        Task<PremInReportResult?> GetPremInClearingReportJson(ArApReportInput data);
        Task<PremInReportResult?> GetPremInOutstandingReportJson(ArApReportInput data);
        Task<CommOutOvOutReportResult?> GetCommOutOvOutOpenItemReportJson(ArApReportInput data);
        Task<CommOutOvOutReportResult?> GetCommOutOvOutClearingReportJson(ArApReportInput data);
        Task<CommOutOvOutReportResult?> GetCommOutOvOutOutstandingReportJson(ArApReportInput data);
        Task<PremOutCommInOvInReportResult?> GetPremOutOpenItemReportJson(ArApReportInput data);
        Task<PremOutCommInOvInReportResult?> GetPremOutClearingReportJson(ArApReportInput data);
        Task<PremOutCommInOvInReportResult?> GetPremOutOutstandingReportJson(ArApReportInput data);
        Task<PremOutCommInOvInReportResult?> GetCommInOvInOpenItemReportJson(ArApReportInput data);
        Task<PremOutCommInOvInReportResult?> GetCommInOvInClearingReportJson(ArApReportInput data);
        Task<PremOutCommInOvInReportResult?> GetCommInOvInOutstandingReportJson(ArApReportInput data);
    }
}
