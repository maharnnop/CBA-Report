using Microsoft.EntityFrameworkCore;
using BestPolicyReport.Models.DailyPolicyReport;
using BestPolicyReport.Models.BillReport;
using BestPolicyReport.Models.CashierReport;
using BestPolicyReport.Models.OutputVatReport;
using BestPolicyReport.Models.ArApReport;
using BestPolicyReport.Models.PremInDirectReport;
using BestPolicyReport.Models.WhtReport;

namespace BestPolicyReport.Data
{
    public class DataContext : DbContext
    {
        public DataContext(DbContextOptions<DataContext> options) : base(options)
        {

        }

        protected override void OnModelCreating(ModelBuilder modelBuilder)
        {
            modelBuilder.Entity<DailyPolicyData>().HasNoKey();
            modelBuilder.Entity<EndorseData>().HasNoKey();
            modelBuilder.Entity<InvoiceData>().HasNoKey();
            modelBuilder.Entity<BillData>().HasNoKey();
            modelBuilder.Entity<CashierData>().HasNoKey();
            modelBuilder.Entity<CommOutOvOutData>().HasNoKey();
            modelBuilder.Entity<PremOutCommInOvInData>().HasNoKey();
            modelBuilder.Entity<PremInData>().HasNoKey();
            modelBuilder.Entity<PremInDirectData>().HasNoKey();
            modelBuilder.Entity<DatasCount>().HasNoKey();

            modelBuilder.Entity<PolicyGroupBillReportResult>().HasNoKey();
            modelBuilder.Entity<OutputVatCommInReportResult>().HasNoKey();
            modelBuilder.Entity<OutputVatOvInReportResult>().HasNoKey();
            modelBuilder.Entity<WhtCommOutOvOutReportResult>().HasNoKey();
            


        }

        public DbSet<DailyPolicyData> DailyPolicyDatas { get; set; }
        public DbSet<EndorseData> EndorseDatas { get; set; }
        public DbSet<InvoiceData> InvoiceDatas { get; set; }

        public DbSet<InvoiceRpt> InvoiceRptDatas { get; set; }
        public DbSet<BillData> BillDatas { get; set; }
        public DbSet<CashierData> CashierDatas { get; set; }
        public DbSet<OutputVatCommInReportResult> OutputVatCommInReportResults { get; set; }
        public DbSet<OutputVatOvInReportResult> OutputVatOvInReportResults { get; set; }
        public DbSet<CommOutOvOutData> CommOutOvOutDatas { get; set; }
        public DbSet<PremOutCommInOvInData> PremOutCommInOvInDatas { get; set; }
        public DbSet<PremInData> PremInDatas { get; set; }
        public DbSet<PolicyGroupBillReportResult> PolicyGroupBillReportResults { get; set; }
        public DbSet<PremInDirectData> PremInDirectDatas { get; set; }
        public DbSet<WhtCommOutOvOutReportResult> WhtCommOutOvOutReportResults { get; set; }

        public DbSet<DatasCount> DatasCount { get; set; }
    }
}
