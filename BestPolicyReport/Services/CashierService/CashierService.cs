using BestPolicyReport.Data;
using BestPolicyReport.Models.BillReport;
using BestPolicyReport.Models.CashierReport;
using Microsoft.EntityFrameworkCore;

namespace BestPolicyReport.Services.CashierService
{
    public class CashierService : ICashierService
    {
        private readonly DataContext _dataContext;

        public CashierService(DataContext dataContext)
        {
            _dataContext = dataContext;
        }

        public async Task<CashierReportResult?> GetCashierReportJson(CashierReportInput data)
        {
            var sql = $@"(select t.""policyNo"" ,getname(i.""entityID"") as insureename ,c.cashierreceiveno as ""cashierReceiveNo"", c.cashierdate as ""cashierDate"", c.billadvisorno as ""billAdvisorNo"", a.billdate as ""billDate"",
                         c.receivefrom as ""receiveFrom"", c.receivename as ""receiveName"", c.receivetype as ""receiveType"", c.refno as ""refNo"",
                         --c.refdate as ""refDate"", 
                         c.transactiontype as ""transactionType"", c.amt as ""cashierAmt"", c.dfrpreferno as ""dfRpReferNo"", r.rprefdate as ""rpRefDate"", 
                         r.actualvalue as ""actualValue"", r.diffamt as ""diffAmt"", r.status, c.createusercode as ""createusercode""
                         from static_data.b_jacashiers c
                               left join static_data.b_jabilladvisors a on  a.billadvisorno = c.billadvisorno and a.active ='Y'
                               left join  static_data.b_jaaraps r on c.dfrpreferno = r.dfrpreferno and c.cashierreceiveno = r.cashierreceiveno 
                            left join (select distinct ""policyNo"",dfrpreferno,""transType"" from static_data.""Transactions""  where ""transType"" = 'PREM-IN' ) t on t.dfrpreferno =c.dfrpreferno  and c.billadvisorno ='-'
                           left join static_data.""Policies"" p on t.""policyNo""  = p.""policyNo"" and p.""lastVersion"" ='Y'
                           left join static_data.""Insurees"" i on i.""insureeCode"" =p.""insureeCode""  and i.lastversion ='Y'
                        order by c.cashierreceiveno ) as query where true ";
            string currentDate = DateTime.Now.ToString("yyyy-MM-dd", new System.Globalization.CultureInfo("en-US"));
            if (!string.IsNullOrEmpty(data.StartCashierDate?.ToString()))
            {
                if (!string.IsNullOrEmpty(data.EndCashierDate?.ToString()))
                {
                    sql += $@"and ""cashierDate"" between '{data.StartCashierDate}' and '{data.EndCashierDate}' ";
                }
                else
                {
                    sql += $@"and ""cashierDate"" between '{data.StartCashierDate}' and '{currentDate}' ";
                }
            }
            if (!string.IsNullOrEmpty(data.StartCashierReceiveNo) && !string.IsNullOrEmpty(data.EndCashierReceiveNo))
            {
                sql += $@"and ""cashierReceiveNo"" between '{data.StartCashierReceiveNo}' and '{data.EndCashierReceiveNo}' ";
            }
            if (!string.IsNullOrEmpty(data.TransactionType))
            {
                sql += $@"and ""transactionType"" = '{data.TransactionType}' ";
            }
            string sql1 = "select * from " + sql;

            if ((data.Limit).HasValue && (data.Pagecount).HasValue)
            {
                sql1 += $@" LIMIT {data.Limit}  OFFSET ({data.Pagecount} - 1) * {data.Limit} ;";
            }
            else { sql1 += " ;"; }

            string sql2 = "select count(*) as count from " + sql + " ;";
            
            var json = await _dataContext.CashierDatas.FromSqlRaw(sql1).ToListAsync();

            var count = await _dataContext.DatasCount.FromSqlRaw(sql2).ToListAsync();
            return new CashierReportResult(json, count[0].Count);
        }
    }
}
