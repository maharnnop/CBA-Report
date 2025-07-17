using BestPolicyReport.Data;
using BestPolicyReport.Models.ArApReport;
using DocumentFormat.OpenXml.Wordprocessing;
using Microsoft.EntityFrameworkCore;

namespace BestPolicyReport.Services.ArApService
{
    public class ArApService : IArApService
    {
        private readonly DataContext _dataContext;

        public ArApService(DataContext dataContext)
        {
            _dataContext = dataContext;
        }

        private static Task<string> GetWhereSql(ArApReportInput data, string sql, string type, Boolean count = false)
        {
            string currentDate = DateTime.Now.ToString("yyyy-MM-dd", new System.Globalization.CultureInfo("en-US"));
            if (!string.IsNullOrEmpty(data.StartPolicyIssueDate?.ToString()))
            {
                if (!string.IsNullOrEmpty(data.EndPolicyIssueDate?.ToString()))
                {
                    sql += $@"and p.""issueDate"" between '{data.StartPolicyIssueDate}' and '{data.EndPolicyIssueDate}' ";
                }
                else
                {
                    sql += $@"and p.""issueDate"" between '{data.StartPolicyIssueDate}' and '{currentDate}' ";
                }
            }
            if (!string.IsNullOrEmpty(data.CreateUserCode))
            {
                sql += $@"and p.createusercode = '{data.CreateUserCode}' ";
            }
            if (!string.IsNullOrEmpty(data.MainAccountContactPersonId))
            {
                sql += $@"and a.""contactPersonID"" = {data.MainAccountContactPersonId} ";
            }
            if (!string.IsNullOrEmpty(data.MainAccountCode))
            {
                sql += $@"and t.mainaccountcode = '{data.MainAccountCode}' ";
            }
            if (!string.IsNullOrEmpty(data.AgentCode))
            {
                sql += $@"and t.""agentCode"" = '{data.AgentCode}' ";
            }
            if (!string.IsNullOrEmpty(data.InsurerCode))
            {
                sql += $@"and t.""insurerCode"" = '{data.InsurerCode}' ";
            }
            if (!string.IsNullOrEmpty(data.Insurancestatus))
            {
                sql += $@"and p.insurancestatus = '{data.Insurancestatus}' ";
            }
            if (!string.IsNullOrEmpty(data.Class))
            {
                sql += $@"and it.""class"" = '{data.Class}' ";
            }
            if (!string.IsNullOrEmpty(data.SubClass))
            {
                sql += $@"and it.""subClass"" = '{data.SubClass}' ";
            }
            if (!string.IsNullOrEmpty(data.PreminRprefdatestart))
            {
                sql += $@"and t.""premin-rprefdate"" >= '{data.PreminRprefdatestart}' ";
            }
            if (!string.IsNullOrEmpty(data.PreminRprefdateend))
            {
                sql += $@"and t.""premin-rprefdate"" <= '{data.PreminRprefdateend}' ";
            }
            if (!string.IsNullOrEmpty(data.PremoutRprefdatestart))
            {
                sql += $@"and t.""premout-rprefdate"" >= '{data.PremoutRprefdatestart}' ";
            }
            if (!string.IsNullOrEmpty(data.PremoutRprefdateend))
            {
                sql += $@"and t.""premout-rprefdate"" <= '{data.PremoutRprefdateend}' ";
            }
            if (type == "CommOutOvOut")
            {
               /*
                if (!string.IsNullOrEmpty(data.TransactionType))
                {
                    sql += $@"and t.""transType"" = '{data.TransactionType}' ";
                }
               */
            }
            if (!count)
            {
                sql += $@"order by t.""policyNo"" asc,  t.""endorseNo"" asc, t.""transType"" asc, t.""seqNo"" asc ";
                if ((data.Limit).HasValue && (data.Pagecount).HasValue)
                {
                    sql += $@"LIMIT {data.Limit}  OFFSET ({data.Pagecount} - 1) * {data.Limit} ;";
                }
                else { sql += ";"; }
            }
            else
            {
                sql +=  ";";
            }
            return Task.FromResult(sql);
        }
        public async Task<PremInReportResult?> GetPremInOpenItemReportJson(ArApReportInput data)
        {
            var sql = $@"select
                         p.grossprem as ""grossPrem"",
                         p.specdiscrate as ""specDiscRate"",
                         p.specdiscamt as ""specDiscAmt"",
                         t.""transType"",
                         t.""policyNo"",
                         t.""endorseNo"",
                         jupgr.""invoiceNo"",
                         t.""seqNo"",
                         t.billadvisorno, (select billdate from static_data.b_jabilladvisors where billadvisorno = t.billadvisorno and active = 'Y'),
                         ar.cashierreceiveno as ""cashierReceiveNo"",
                         c.cashierdate as ""cashierDate"",
                         ar.cashieramt as ""cashierAmt"",
                         c.receivetype as ""cashierReceiveType"",
                         c.refno as ""cashierRefNo"",
                         --c.refdate,
                         t.""dfrpreferno"" as ""dfRpReferNo"",
                         t.""premin-dfrpreferno"" as ""premInDfRpReferNo"",
                         t.""premin-rprefdate"" as ""premInRpRefDate"",
                         t.rprefdate as ""rpRefDate"",
                         t.netflag as ""netFlag"",
                         t.paidamt as ""premInPaidAmt"", 
                         t.remainamt as ""premInDiffAmt"",
                         p.insurancestatus as ""insurancestatus"",
                         p.""actDate"",
                         p.""expDate"",
                         t.mainaccountcode as ""mainAccountCode"",
                         case
                             when e.""personType"" = 'O' then concat(tt.""TITLETHAIBEGIN"", ' ', e.""t_ogName"", ' ', tt.""TITLETHAIEND"")
                             when e.""personType"" = 'P' then concat(tt.""TITLETHAIBEGIN"", ' ', e.""t_firstName"", ' ', e.""t_lastName"", ' ', tt.""TITLETHAIEND"")
                             else null
                         end as ""MainAccountName"",
                         p.""insureeCode"",
                         case
                             when e_i.""personType"" = 'O' then concat(tt_i.""TITLETHAIBEGIN"", ' ', e_i.""t_ogName"", ' ', tt_i.""TITLETHAIEND"")
                             when e_i.""personType"" = 'P' then concat(tt_i.""TITLETHAIBEGIN"", ' ', e_i.""t_firstName"", ' ', e_i.""t_lastName"", ' ', tt_i.""TITLETHAIEND"")
                             else null
                         end as ""insureeName"",
                         it.""class"",
                         it.""subClass"",
                         (case when t.""transType"" ='DISC-IN' or t.netgrossprem is null then 0 else ((t.""subType"" *2) -1) * t.netgrossprem end ) as ""netgrossprem"",
                         (case when t.""transType"" ='DISC-IN' or t.duty is null then 0 else ((t.""subType"" *2) -1) * t.duty end )  as ""duty"",
                         (case when t.""transType"" ='DISC-IN' or t.tax is null then 0 else  ((t.""subType"" *2) -1) * t.tax end )  as ""tax"",
                         (case when t.""transType"" ='DISC-IN' or t.totalprem is null then 0 else ((t.""subType"" *2) -1) * t.totalprem end )  as ""totalPrem"",
                         (case when t.""transType"" ='DISC-IN' or t.withheld is null then 0 else ((t.""subType"" *2) -1) * t.withheld end )  as ""withheld"",
                         (((t.""subType"" *2) -1) * t.totalamt) as totalamt,
                         p.commout_rate as ""commOutRate"",
                         ((t.""subType"" *2) -1) * p.commout_amt as ""commOutAmt"",
                         p.ovout_rate as ""ovOutRate"",
                         ((t.""subType"" *2) -1) * p.ovout_amt as ""ovOutAmt"",
                         p.commout1_rate as ""commOutRate1"",
                         ((t.""subType"" *2) -1) * p.commout1_amt as ""commOutAmt1"",
                         p.ovout1_rate as ""ovOutRate1"",
                         ((t.""subType"" *2) -1) * p.ovout1_amt as ""ovOutAmt1"",
                         p.commout2_rate as ""commOutRate2"",
                         ((t.""subType"" *2) -1) * p.commout2_amt as ""commOutAmt2"",
                         p.""issueDate"",
                         p.createusercode as ""policyCreateUserCode"",
                         a.""contactPersonID"" as ""mainAccountContactPersonId"",
                         t.""insurerCode"",
                         t.""transType"" as ""transactionType""
                         from
                         static_data.""Transactions"" t 
                         left join static_data.""b_jupgrs"" jupgr on jupgr.""policyNo"" = t.""policyNo"" and t.dftxno = jupgr.dftxno and jupgr.""seqNo"" = t.""seqNo"" 
                         left join static_data.""Policies"" p on jupgr.polid = p.id
                         left join static_data.""Insurees"" i on p.""insureeCode"" = i.""insureeCode""
                         left join static_data.""Entities"" e_i on i.""entityID"" = e_i.id
                         left join static_data.""Titles"" tt_i on tt_i.""TITLEID"" = e_i.""titleID""
                         left join static_data.""InsureTypes"" it on p.""insureID"" = it.id
                         left join static_data.b_jacashiers c on t.""premin-dfrpreferno"" = c.dfrpreferno
                         left join static_data.b_jaaraps ar on ar.cashierreceiveno = c.cashierreceiveno
                         left join static_data.b_jaarapds ard on ard.keyidm = ar.id and ard.polid = p.id
                         left join static_data.""Agents"" a on a.""agentCode"" = t.mainaccountcode  and a.""lastversion"" = 'Y'
                         left join static_data.""Entities"" e on a.""entityID"" = e.id
                         left join static_data.""Titles"" tt on tt.""TITLEID"" = e.""titleID""
                         where  t.txtype2 in ('1', '2', '3', '4', '5')
                         and t.""transType"" in ('PREM-IN', 'DISC-IN')
                         --( ( t.""transType"" in ('PREM-IN', 'DISC-IN') and t.txtype2 in ('1', '2' ) ) or ( t.""transType"" in ('PREM-OUT', 'DISC-OUT') and t.txtype2 in ('3', '4', '5' ) ) )
                         -- and t.txtype2 in ('1', '2', '3', '4', '5')
                         and t.status = 'N' 
                         and p.""lastVersion"" = 'Y'
                         and jupgr.""installmenttype"" = 'A' ";
            var sqlcount = $@"select
                         count(*) as Count
                         from
                         static_data.""Transactions"" t 
                         left join static_data.""b_jupgrs"" jupgr on jupgr.""policyNo"" = t.""policyNo"" and t.dftxno = jupgr.dftxno and jupgr.""seqNo"" = t.""seqNo"" 
                         left join static_data.""Policies"" p on jupgr.polid = p.id
                         left join static_data.""Insurees"" i on p.""insureeCode"" = i.""insureeCode""
                         left join static_data.""Entities"" e_i on i.""entityID"" = e_i.id
                         left join static_data.""Titles"" tt_i on tt_i.""TITLEID"" = e_i.""titleID""
                         left join static_data.""InsureTypes"" it on p.""insureID"" = it.id
                         left join static_data.b_jacashiers c on t.""premin-dfrpreferno"" = c.dfrpreferno
                         left join static_data.b_jaaraps ar on ar.cashierreceiveno = c.cashierreceiveno
                         left join static_data.b_jaarapds ard on ard.keyidm = ar.id and ard.polid = p.id
                         left join static_data.""Agents"" a on a.""agentCode"" = t.mainaccountcode  and a.""lastversion"" = 'Y'
                         left join static_data.""Entities"" e on a.""entityID"" = e.id
                         left join static_data.""Titles"" tt on tt.""TITLEID"" = e.""titleID""
                         where  t.txtype2 in ('1', '2', '3', '4', '5')
                         and t.""transType"" in ('PREM-IN', 'DISC-IN')
                         --( ( t.""transType"" in ('PREM-IN', 'DISC-IN') and t.txtype2 in ('1', '2' ) ) or ( t.""transType"" in ('PREM-OUT', 'DISC-OUT') and t.txtype2 in ('3', '4', '5' ) ) )
                         -- and t.txtype2 in ('1', '2', '3', '4', '5')
                         and t.status = 'N' 
                         and p.""lastVersion"" = 'Y'
                         and jupgr.""installmenttype"" = 'A' ";
            sql = await GetWhereSql(data, sql, "PremIn");
            sqlcount = await GetWhereSql(data, sqlcount, "PremIn", true);
            var json = await _dataContext.PremInDatas.FromSqlRaw(sql).ToListAsync();
            var count = await _dataContext.DatasCount.FromSqlRaw(sqlcount).ToListAsync();
            return new PremInReportResult(json, count[0].Count); 
        }

        public async Task<PremInReportResult?> GetPremInClearingReportJson(ArApReportInput data)
        {
            var sql = $@"select
                         t.""policyNo"",
                         p.specdiscrate as ""specDiscRate"",
                         p.specdiscamt as ""specDiscAmt"",
                         p.grossprem as ""grossPrem"",
                         t.""transType"",
                         t.""endorseNo"",
                         jupgr.""invoiceNo"",
                         t.billadvisorno, (select billdate from static_data.b_jabilladvisors where billadvisorno = t.billadvisorno and active = 'Y'),
                         t.""seqNo"",
                         ar.cashierreceiveno as ""cashierReceiveNo"",
                         c.cashierdate as ""cashierDate"",
                         ar.cashieramt as ""cashierAmt"",
                         c.receivetype as ""cashierReceiveType"",
                         c.refno as ""cashierRefNo"",
                         --c.refdate,
                         t.""dfrpreferno"" as ""dfRpReferNo"",
                         t.""premin-dfrpreferno"" as ""premInDfRpReferNo"",
                         t.""premin-rprefdate"" as ""premInRpRefDate"",
                         t.rprefdate as ""rpRefDate"",
                         t.netflag as ""netFlag"",
                         t.paidamt as ""premInPaidAmt"", 
                         t.remainamt as ""premInDiffAmt"",
                         p.insurancestatus as ""insurancestatus"",
                         p.""actDate"",
                         p.""expDate"",
                         t.mainaccountcode as ""mainAccountCode"",
                         case
                             when e.""personType"" = 'O' then concat(tt.""TITLETHAIBEGIN"", ' ', e.""t_ogName"", ' ', tt.""TITLETHAIEND"")
                             when e.""personType"" = 'P' then concat(tt.""TITLETHAIBEGIN"", ' ', e.""t_firstName"", ' ', e.""t_lastName"", ' ', tt.""TITLETHAIEND"")
                             else null
                         end as ""MainAccountName"",
                         p.""insureeCode"",
                         case
                             when e_i.""personType"" = 'O' then concat(tt_i.""TITLETHAIBEGIN"", ' ', e_i.""t_ogName"", ' ', tt_i.""TITLETHAIEND"")
                             when e_i.""personType"" = 'P' then concat(tt_i.""TITLETHAIBEGIN"", ' ', e_i.""t_firstName"", ' ', e_i.""t_lastName"", ' ', tt_i.""TITLETHAIEND"")
                             else null
                         end as ""insureeName"",
                         it.""class"",
                         it.""subClass"",
                        (case when t.""transType"" ='DISC-IN' or t.netgrossprem is null then 0 else  ((t.""subType"" *2) -1) * t.netgrossprem end ) as ""netgrossprem"",
                         (case when t.""transType"" ='DISC-IN' or t.duty is null then 0 else  ((t.""subType"" *2) -1) *  t.duty end )  as ""duty"",
                         (case when t.""transType"" ='DISC-IN' or t.tax is null then 0 else  ((t.""subType"" *2) -1) * t.tax end )  as ""tax"",
                         (case when t.""transType"" ='DISC-IN' or t.totalprem is null then 0 else  ((t.""subType"" *2) -1) * t.totalprem end )  as ""totalPrem"",
                         (case when t.""transType"" ='DISC-IN' or t.withheld is null then 0 else  ((t.""subType"" *2) -1) * t.withheld end )  as ""withheld"",
                         ( ((t.""subType"" *2) -1) * t.totalamt) as totalamt,
                         p.commout_rate as ""commOutRate"",
                          ((t.""subType"" *2) -1) * p.commout_amt as ""commOutAmt"",
                         p.ovout_rate as ""ovOutRate"",
                          ((t.""subType"" *2) -1) * p.ovout_amt as ""ovOutAmt"",
                         p.commout1_rate as ""commOutRate1"",
                          ((t.""subType"" *2) -1) * p.commout1_amt as ""commOutAmt1"",
                         p.ovout1_rate as ""ovOutRate1"",
                         ((t.""subType"" *2) -1) * p.ovout1_amt as ""ovOutAmt1"",
                         p.commout2_rate as ""commOutRate2"",
                          ((t.""subType"" *2) -1) * p.commout2_amt as ""commOutAmt2"",
                         p.""issueDate"",
                         p.createusercode as ""policyCreateUserCode"",
                         a.""contactPersonID"" as ""mainAccountContactPersonId"",
                         t.""insurerCode"",
                         t.""transType"" as ""transactionType""
                         from
                         static_data.""Transactions"" t 
                         left join static_data.""b_jupgrs"" jupgr on jupgr.""policyNo"" = t.""policyNo"" and  t.dftxno = jupgr.dftxno and jupgr.""seqNo"" = t.""seqNo"" 
                         left join static_data.""Policies"" p on jupgr.polid = p.id
                         left join static_data.""Insurees"" i on p.""insureeCode"" = i.""insureeCode""
                         left join static_data.""Entities"" e_i on i.""entityID"" = e_i.id
                         left join static_data.""Titles"" tt_i on tt_i.""TITLEID"" = e_i.""titleID""
                         left join static_data.""InsureTypes"" it on p.""insureID"" = it.id
                         left join static_data.b_jacashiers c on t.""premin-dfrpreferno"" = c.dfrpreferno
                         left join static_data.b_jaaraps ar on (ar.cashierreceiveno = c.cashierreceiveno)
                         left join static_data.b_jaarapds ard on ard.keyidm = ar.id and ard.polid = p.id
                         left join static_data.""Agents"" a on a.""agentCode"" = t.mainaccountcode  and a.""lastversion"" = 'Y'
                         left join static_data.""Entities"" e on a.""entityID"" = e.id
                         left join static_data.""Titles"" tt on tt.""TITLEID"" = e.""titleID""
                         where t.""transType"" in ('PREM-IN', 'DISC-IN')
                         and t.txtype2 in ('1', '2', '3', '4', '5')
                         -- ( ( t.""transType"" in ('PREM-IN', 'DISC-IN') and t.txtype2 in ('1', '2' ) ) or ( t.""transType"" in ('PREM-OUT', 'DISC-OUT') and t.txtype2 in ('3', '4', '5' ) ) )
                         -- and t.txtype2 in ('1', '2', '3', '4', '5')
                         and t.status = 'N'
                         and t.""premin-dfrpreferno"" is not null
                         and t.""premin-rprefdate"" is not null
                         and t.dfrpreferno is not null
                         and t.rprefdate is not null 
                         and p.""lastVersion"" = 'Y'
                         and jupgr.""installmenttype"" = 'A' ";
            var sqlcount = $@"select
                         count(*) as count
                         from
                         static_data.""Transactions"" t 
                         left join static_data.""b_jupgrs"" jupgr on jupgr.""policyNo"" = t.""policyNo"" and  t.dftxno = jupgr.dftxno and jupgr.""seqNo"" = t.""seqNo"" 
                         left join static_data.""Policies"" p on jupgr.polid = p.id
                         left join static_data.""Insurees"" i on p.""insureeCode"" = i.""insureeCode""
                         left join static_data.""Entities"" e_i on i.""entityID"" = e_i.id
                         left join static_data.""Titles"" tt_i on tt_i.""TITLEID"" = e_i.""titleID""
                         left join static_data.""InsureTypes"" it on p.""insureID"" = it.id
                         left join static_data.b_jacashiers c on t.""premin-dfrpreferno"" = c.dfrpreferno
                         left join static_data.b_jaaraps ar on (ar.cashierreceiveno = c.cashierreceiveno)
                         left join static_data.b_jaarapds ard on ard.keyidm = ar.id and ard.polid = p.id
                         left join static_data.""Agents"" a on a.""agentCode"" = t.mainaccountcode  and a.""lastversion"" = 'Y'
                         left join static_data.""Entities"" e on a.""entityID"" = e.id
                         left join static_data.""Titles"" tt on tt.""TITLEID"" = e.""titleID""
                         where t.""transType"" in ('PREM-IN', 'DISC-IN')
                         and t.txtype2 in ('1', '2', '3', '4', '5')
                         -- ( ( t.""transType"" in ('PREM-IN', 'DISC-IN') and t.txtype2 in ('1', '2' ) ) or ( t.""transType"" in ('PREM-OUT', 'DISC-OUT') and t.txtype2 in ('3', '4', '5' ) ) )
                         -- and t.txtype2 in ('1', '2', '3', '4', '5')
                         and t.status = 'N'
                         and t.""premin-dfrpreferno"" is not null
                         and t.""premin-rprefdate"" is not null
                         and t.dfrpreferno is not null
                         and t.rprefdate is not null 
                         and p.""lastVersion"" = 'Y'
                         and jupgr.""installmenttype"" = 'A' ";
            sql = await GetWhereSql(data, sql, "PremIn");
            sqlcount = await GetWhereSql(data, sqlcount, "PremIn", true);
            var json = await _dataContext.PremInDatas.FromSqlRaw(sql).ToListAsync();
            var count = await _dataContext.DatasCount.FromSqlRaw(sqlcount).ToListAsync();
            return new PremInReportResult(json, count[0].Count);
        }

        public async Task<PremInReportResult?> GetPremInOutstandingReportJson(ArApReportInput data)
        {
            var sql = $@"select
                         p.grossprem as ""grossPrem"",
                         p.specdiscrate as ""specDiscRate"",
                         p.specdiscamt as ""specDiscAmt"",
                         t.""transType"",
                         t.""policyNo"",
                         t.""endorseNo"",
                         jupgr.""invoiceNo"",
                         t.billadvisorno, (select billdate from static_data.b_jabilladvisors where billadvisorno = t.billadvisorno and active = 'Y'),
                         t.""seqNo"",
                         ar.cashierreceiveno as ""cashierReceiveNo"",
                         c.cashierdate as ""cashierDate"",
                         ar.cashieramt as ""cashierAmt"",
                         c.receivetype as ""cashierReceiveType"",
                         c.refno as ""cashierRefNo"",
                         --c.refdate,
                         t.""dfrpreferno"" as ""dfRpReferNo"",
                         t.""premin-dfrpreferno"" as ""premInDfRpReferNo"",
                         t.""premin-rprefdate"" as ""premInRpRefDate"",
                         t.rprefdate as ""rpRefDate"",
                         t.netflag as ""netFlag"",
                         t.paidamt as ""premInPaidAmt"", 
                         t.remainamt as ""premInDiffAmt"",
                         p.insurancestatus as ""insurancestatus"",
                         p.""actDate"",
                         p.""expDate"",
                         t.mainaccountcode as ""mainAccountCode"",
                         case
                             when e.""personType"" = 'O' then concat(tt.""TITLETHAIBEGIN"", ' ', e.""t_ogName"", ' ', tt.""TITLETHAIEND"")
                             when e.""personType"" = 'P' then concat(tt.""TITLETHAIBEGIN"", ' ', e.""t_firstName"", ' ', e.""t_lastName"", ' ', tt.""TITLETHAIEND"")
                             else null
                         end as ""MainAccountName"",
                         p.""insureeCode"",
                         case
                             when e_i.""personType"" = 'O' then concat(tt_i.""TITLETHAIBEGIN"", ' ', e_i.""t_ogName"", ' ', tt_i.""TITLETHAIEND"")
                             when e_i.""personType"" = 'P' then concat(tt_i.""TITLETHAIBEGIN"", ' ', e_i.""t_firstName"", ' ', e_i.""t_lastName"", ' ', tt_i.""TITLETHAIEND"")
                             else null
                         end as ""insureeName"",
                         it.""class"",
                         it.""subClass"",
                         (case when t.""transType"" ='DISC-IN' or t.netgrossprem is null then 0 else ((t.""subType"" *2) -1) * t.netgrossprem end ) as ""netgrossprem"",
                         (case when t.""transType"" ='DISC-IN' or t.duty is null then 0 else ((t.""subType"" *2) -1) * t.duty end )  as ""duty"",
                         (case when t.""transType"" ='DISC-IN' or t.tax is null then 0 else ((t.""subType"" *2) -1) * t.tax end )  as ""tax"",
                         (case when t.""transType"" ='DISC-IN' or t.totalprem is null then 0 else ((t.""subType"" *2) -1) * t.totalprem end )  as ""totalPrem"",
                         (case when t.""transType"" ='DISC-IN' or t.withheld is null then 0 else ((t.""subType"" *2) -1) * t.withheld end )  as ""withheld"",
                         (((t.""subType"" *2) -1) * t.totalamt) as totalamt,
                         p.commout_rate as ""commOutRate"",
                         ((t.""subType"" *2) -1) * p.commout_amt as ""commOutAmt"",
                         p.ovout_rate as ""ovOutRate"",
                         ((t.""subType"" *2) -1) * p.ovout_amt as ""ovOutAmt"",
                         p.commout1_rate as ""commOutRate1"",
                       ((t.""subType"" *2) -1) *  p.commout1_amt as ""commOutAmt1"",
                         p.ovout1_rate as ""ovOutRate1"",
                        ((t.""subType"" *2) -1) *  p.ovout1_amt as ""ovOutAmt1"",
                         p.commout2_rate as ""commOutRate2"",
                       ((t.""subType"" *2) -1) *  p.commout2_amt as ""commOutAmt2"",
                         p.""issueDate"",
                         p.createusercode as ""policyCreateUserCode"",
                         a.""contactPersonID"" as ""mainAccountContactPersonId"",
                         t.""insurerCode"",
                         t.""transType"" as ""transactionType""
                         from
                         static_data.""Transactions"" t 
                          left join static_data.""b_jupgrs"" jupgr on jupgr.""policyNo"" = t.""policyNo"" and  t.dftxno = jupgr.dftxno and jupgr.""seqNo"" = t.""seqNo"" 
                         left join static_data.""Policies"" p on jupgr.polid = p.id
                         left join static_data.""Insurees"" i on p.""insureeCode"" = i.""insureeCode""
                         left join static_data.""Entities"" e_i on i.""entityID"" = e_i.id
                         left join static_data.""Titles"" tt_i on tt_i.""TITLEID"" = e_i.""titleID""
                         left join static_data.""InsureTypes"" it on p.""insureID"" = it.id
                         left join static_data.b_jacashiers c on t.""premin-dfrpreferno"" = c.dfrpreferno
                         left join static_data.b_jaaraps ar on (ar.cashierreceiveno = c.cashierreceiveno)
                         left join static_data.b_jaarapds ard on ard.keyidm = ar.id and ard.polid = p.id
                         left join static_data.""Agents"" a on a.""agentCode"" = t.mainaccountcode  and a.""lastversion"" = 'Y'
                         left join static_data.""Entities"" e on a.""entityID"" = e.id
                         left join static_data.""Titles"" tt on tt.""TITLEID"" = e.""titleID""
                         where t.""transType"" in ('PREM-IN', 'DISC-IN')
                         and t.txtype2 in ('1', '2', '3', '4', '5')
                         -- ( ( t.""transType"" in ('PREM-IN', 'DISC-IN') and t.txtype2 in ('1', '2' ) ) or ( t.""transType"" in ('PREM-OUT', 'DISC-OUT') and t.txtype2 in ('3', '4', '5' ) ) )
                         -- and t.txtype2 in ('1', '2', '3', '4', '5')
                         and t.status = 'N'
                         and p.""lastVersion"" = 'Y'
                         and t.""premin-dfrpreferno"" is null
                         and t.""premin-rprefdate"" is null
                         and t.dfrpreferno is null 
                         and t.rprefdate is null
                         and jupgr.""installmenttype"" = 'A' ";
            //if (!string.IsNullOrEmpty(data.AsAtDate))
            //{
            //    sql += $@"or (t.dfrpreferno is not null and t.rprefdate is not null and t.rprefdate <= '{data.AsAtDate}') ";
            //}
            //sql += $@")";
            var sqlcount = $@"select
                         count(*) as count
                         from
                         static_data.""Transactions"" t 
                          left join static_data.""b_jupgrs"" jupgr on jupgr.""policyNo"" = t.""policyNo"" and  t.dftxno = jupgr.dftxno and jupgr.""seqNo"" = t.""seqNo"" 
                         left join static_data.""Policies"" p on jupgr.polid = p.id
                         left join static_data.""Insurees"" i on p.""insureeCode"" = i.""insureeCode""
                         left join static_data.""Entities"" e_i on i.""entityID"" = e_i.id
                         left join static_data.""Titles"" tt_i on tt_i.""TITLEID"" = e_i.""titleID""
                         left join static_data.""InsureTypes"" it on p.""insureID"" = it.id
                         left join static_data.b_jacashiers c on t.""premin-dfrpreferno"" = c.dfrpreferno
                         left join static_data.b_jaaraps ar on (ar.cashierreceiveno = c.cashierreceiveno)
                         left join static_data.b_jaarapds ard on ard.keyidm = ar.id and ard.polid = p.id
                         left join static_data.""Agents"" a on a.""agentCode"" = t.mainaccountcode  and a.""lastversion"" = 'Y'
                         left join static_data.""Entities"" e on a.""entityID"" = e.id
                         left join static_data.""Titles"" tt on tt.""TITLEID"" = e.""titleID""
                         where t.""transType"" in ('PREM-IN', 'DISC-IN')
                         and t.txtype2 in ('1', '2', '3', '4', '5')
                         -- ( ( t.""transType"" in ('PREM-IN', 'DISC-IN') and t.txtype2 in ('1', '2' ) ) or ( t.""transType"" in ('PREM-OUT', 'DISC-OUT') and t.txtype2 in ('3', '4', '5' ) ) )
                         -- and t.txtype2 in ('1', '2', '3', '4', '5')
                         and t.status = 'N'
                         and p.""lastVersion"" = 'Y'
                         and t.""premin-dfrpreferno"" is null
                         and t.""premin-rprefdate"" is null
                         and t.dfrpreferno is null 
                         and t.rprefdate is null
                         and jupgr.""installmenttype"" = 'A' ";
            sql = await GetWhereSql(data, sql, "PremIn");
            sqlcount = await GetWhereSql(data, sqlcount, "PremIn",true);
            var json = await _dataContext.PremInDatas.FromSqlRaw(sql).ToListAsync();
            var count = await _dataContext.DatasCount.FromSqlRaw(sqlcount).ToListAsync();
            return new PremInReportResult(json, count[0].Count);
        }

        public async Task<CommOutOvOutReportResult?> GetCommOutOvOutOpenItemReportJson(ArApReportInput data)
        {
            var sql = $@"select
                         -1 * ((t.""subType"" *2) -1) * t.withheld as ""withheld"",
                         -1 * ((t.""subType"" *2) -1) * t.totalamt as ""totalamt"",
                         t.""policyNo"",
                         t.""endorseNo"",
                         jupgr.""invoiceNo"",
                         t.""transType"",
                         t.""seqNo"",
                         ar.cashierreceiveno as ""cashierReceiveNo"",
                         c.cashierdate as ""cashierDate"",
                         ar.cashieramt as ""cashierAmt"",
                         c.receivetype as ""cashierReceiveType"",
                         c.refno as ""cashierRefNo"",
                         --c.refdate,
                         t.""premin-dfrpreferno"" as ""premInDfRpReferNo"",
                         t.""premin-rprefdate"" as ""premInRpRefDate"",
                         t.""dfrpreferno"" as ""dfRpReferNo"",
                         t.""rprefdate"" as ""rpRefDate"",
                         p.grossprem as ""grossPrem"",
                         p.specdiscrate as ""specDiscRate"",
                         p.specdiscamt as ""specDiscAmt"",
                        (  -1 * ((t.""subType"" *2) -1) * t.netgrossprem) as ""netGrossPrem"",
                        ( -1 * ((t.""subType"" *2) -1) * t.duty) as ""duty"",
                         (-1 * ((t.""subType"" *2) -1) * t.tax) as ""tax"",
                         (-1 * ((t.""subType"" *2) -1) * t.totalprem) as ""totalPrem"",
                         t.netflag as ""netFlag"",
                         p.""actDate"",
                         p.""expDate"",
                         t.mainaccountcode as ""mainAccountCode"",
                         case
                             when e.""personType"" = 'O' then concat(tt.""TITLETHAIBEGIN"", ' ', e.""t_ogName"", ' ', tt.""TITLETHAIEND"")
                             when e.""personType"" = 'P' then concat(tt.""TITLETHAIBEGIN"", ' ', e.""t_firstName"", ' ', e.""t_lastName"", ' ', tt.""TITLETHAIEND"")
                             else null
                         end as ""MainAccountName"",
                         p.""insureeCode"",
                         case
                             when e_i.""personType"" = 'O' then concat(tt_i.""TITLETHAIBEGIN"", ' ', e_i.""t_ogName"", ' ', tt_i.""TITLETHAIEND"")
                             when e_i.""personType"" = 'P' then concat(tt_i.""TITLETHAIBEGIN"", ' ', e_i.""t_firstName"", ' ', e_i.""t_lastName"", ' ', tt_i.""TITLETHAIEND"")
                             else null
                         end as ""insureeName"",
                         it.""class"",
                         it.""subClass"",
                         m.""licenseNo"",
                         pv.t_provincename as ""province"",
                         m.""chassisNo"",
                         p.commout_rate as ""commOutRate"",
                         p.ovout_rate as ""ovOutRate"",
                        ( -1 * ((t.""subType"" *2) -1) * t.ovamt) as ""ovOutAmt"",
                        ( -1 * ((t.""subType"" *2) -1) * t.commamt) as ""commOutAmt"",
                         (case when t.""transType"" = 'DISC-OUT' then null when t.""transType"" = 'COMM-OUT' then (case when t.""agentCode2"" is null then p.commout1_rate else p.commout2_rate end ) else (case when t.""agentCode2"" is null then p.ovout1_rate else p.ovout2_rate end ) end ) as ""commOvOutRate"",
                         -- (case when t.""transType"" = 'COMM-OUT' then t.commamt else t.ovamt end ) as ""commOvOutAmt"",
                        ( -1 * ((t.""subType"" *2) -1)  * t.totalamt  ) as ""commOvOutAmt"",
                         (select t2.dfrpreferno from static_data.""Transactions"" t2 where ""transType"" = 'COMM-OUT' and t2.id = t.id) as ""commOutDfRpReferNo"",
                         (select t2.rprefdate from static_data.""Transactions"" t2 where ""transType"" = 'COMM-OUT' and t2.id = t.id) as ""commOutRpRefDate"",
                         (select t2.paidamt from static_data.""Transactions"" t2 where ""transType"" = 'COMM-OUT' and t2.id = t.id) as ""commOutPaidAmt"",
                         (select t2.remainamt from static_data.""Transactions"" t2 where ""transType"" = 'COMM-OUT' and t2.id = t.id) as ""commOutDiffAmt"",
                         (select t2.paidamt from static_data.""Transactions"" t2 where ""transType"" = 'OV-OUT' and t2.id = t.id) as ""ovOutPaidAmt"",
                         (select t2.remainamt from static_data.""Transactions"" t2 where ""transType"" = 'OV-OUT' and t2.id = t.id) as ""ovOutDiffAmt"",
                         p.""issueDate"",
                         p.createusercode as ""policyCreateUserCode"",
                         a.""contactPersonID"" as ""mainAccountContactPersonId"",
                         t.""insurerCode"",
                         p.insurancestatus as ""insurancestatus"",
                         t.""transType"" as ""transactionType""
                         from
                         static_data.""Transactions"" t 
                         left join static_data.""b_jupgrs"" jupgr on jupgr.""policyNo"" = t.""policyNo"" and  t.dftxno = jupgr.dftxno and jupgr.""seqNo"" = t.""seqNo"" 
                         left join static_data.""Policies"" p on jupgr.polid = p.id
                         left join static_data.""Insurees"" i on p.""insureeCode"" = i.""insureeCode""
                         left join static_data.""Entities"" e_i on i.""entityID"" = e_i.id
                         left join static_data.""Titles"" tt_i on tt_i.""TITLEID"" = e_i.""titleID""
                         left join static_data.""InsureTypes"" it on p.""insureID"" = it.id
                         left join static_data.""Motors"" m on p.""itemList"" = m.id
                         left join static_data.provinces pv on m.""motorprovinceID"" = pv.provinceid
                         left join static_data.b_jacashiers c on t.""premin-dfrpreferno"" = c.dfrpreferno
                         left join static_data.b_jaaraps ar on (ar.cashierreceiveno = c.cashierreceiveno)
                         left join static_data.b_jaarapds ard on ard.keyidm = ar.id and ard.polid = p.id
                         left join static_data.""Agents"" a on a.""agentCode"" = t.mainaccountcode  and a.""lastversion"" = 'Y'
                         left join static_data.""Entities"" e on a.""entityID"" = e.id
                         left join static_data.""Titles"" tt on tt.""TITLEID"" = e.""titleID""
                         where t.""transType"" in ('COMM-OUT', 'OV-OUT','DISC-OUT')
                         and t.txtype2 in ('1', '2', '3', '4', '5')
                         -- ( ( t.""transType"" in ('COMM-OUT', 'OV-OUT','DISC-OUT') and t.txtype2 in ('1', '2' ) ) or ( t.""transType"" in ('COMM-IN', 'OV-IN','DISC-IN') and t.txtype2 in ('3', '4', '5' ) ) )
                         -- and t.txtype2 in ('1', '2', '3', '4', '5')
                         and t.status = 'N'
                         and p.""lastVersion"" = 'Y'
                         and t.""premin-dfrpreferno"" is not null
                         and t.""premin-rprefdate"" is not null 
                         and jupgr.""installmenttype"" = 'A' ";
            var sqlcount = $@"select
                         count(*) as count
                         from
                         static_data.""Transactions"" t 
                         left join static_data.""b_jupgrs"" jupgr on jupgr.""policyNo"" = t.""policyNo"" and  t.dftxno = jupgr.dftxno and jupgr.""seqNo"" = t.""seqNo"" 
                         left join static_data.""Policies"" p on jupgr.polid = p.id
                         left join static_data.""Insurees"" i on p.""insureeCode"" = i.""insureeCode""
                         left join static_data.""Entities"" e_i on i.""entityID"" = e_i.id
                         left join static_data.""Titles"" tt_i on tt_i.""TITLEID"" = e_i.""titleID""
                         left join static_data.""InsureTypes"" it on p.""insureID"" = it.id
                         left join static_data.""Motors"" m on p.""itemList"" = m.id
                         left join static_data.provinces pv on m.""motorprovinceID"" = pv.provinceid
                         left join static_data.b_jacashiers c on t.""premin-dfrpreferno"" = c.dfrpreferno
                         left join static_data.b_jaaraps ar on (ar.cashierreceiveno = c.cashierreceiveno)
                         left join static_data.b_jaarapds ard on ard.keyidm = ar.id and ard.polid = p.id
                         left join static_data.""Agents"" a on a.""agentCode"" = t.mainaccountcode  and a.""lastversion"" = 'Y'
                         left join static_data.""Entities"" e on a.""entityID"" = e.id
                         left join static_data.""Titles"" tt on tt.""TITLEID"" = e.""titleID""
                         where t.""transType"" in ('COMM-OUT', 'OV-OUT','DISC-OUT')
                         and t.txtype2 in ('1', '2', '3', '4', '5')
                         -- ( ( t.""transType"" in ('COMM-OUT', 'OV-OUT','DISC-OUT') and t.txtype2 in ('1', '2' ) ) or ( t.""transType"" in ('COMM-IN', 'OV-IN','DISC-IN') and t.txtype2 in ('3', '4', '5' ) ) )
                         -- and t.txtype2 in ('1', '2', '3', '4', '5')
                         and t.status = 'N'
                         and p.""lastVersion"" = 'Y'
                         and t.""premin-dfrpreferno"" is not null
                         and t.""premin-rprefdate"" is not null 
                         and jupgr.""installmenttype"" = 'A' ";
            sql = await GetWhereSql(data, sql, "CommOutOvOut");
            sqlcount = await GetWhereSql(data, sqlcount, "CommOutOvOut", true);
            var json = await _dataContext.CommOutOvOutDatas.FromSqlRaw(sql).ToListAsync();
            var count = await _dataContext.DatasCount.FromSqlRaw(sqlcount).ToListAsync();
            return new CommOutOvOutReportResult(json, count[0].Count);
        }

        public async Task<CommOutOvOutReportResult?> GetCommOutOvOutClearingReportJson(ArApReportInput data)
        {
            var sql = $@"select
                         -1 * ((t.""subType"" *2) -1) * t.withheld as ""withheld"",
                         -1 * ((t.""subType"" *2) -1) * t.totalamt as ""totalamt"",
                         t.""policyNo"",
                         t.""endorseNo"",
                         jupgr.""invoiceNo"",
                         t.""seqNo"",
                         t.""transType"",
                         ar.cashierreceiveno as ""cashierReceiveNo"",
                         c.cashierdate as ""cashierDate"",
                         ar.cashieramt as ""cashierAmt"",
                         c.receivetype as ""cashierReceiveType"",
                         c.refno as ""cashierRefNo"",
                         --c.refdate,
                         t.""premin-dfrpreferno"" as ""premInDfRpReferNo"",
                         t.""premin-rprefdate"" as ""premInRpRefDate"",
                         t.""dfrpreferno"" as ""dfRpReferNo"",
                         t.""rprefdate"" as ""rpRefDate"",
                         p.grossprem as ""grossPrem"",
                         p.specdiscrate as ""specDiscRate"",
                         p.specdiscamt as ""specDiscAmt"",
                         -1 * ((t.""subType"" *2) -1) * t.netgrossprem as ""netGrossPrem"",
                         -1 * ((t.""subType"" *2) -1) * t.duty as ""duty"",
                         -1 * ((t.""subType"" *2) -1) * t.tax as ""tax"",
                         -1 * ((t.""subType"" *2) -1) *  t.totalprem as ""totalPrem"",
                         t.netflag as ""netFlag"",
                         p.""actDate"",
                         p.""expDate"",
                         t.mainaccountcode as ""mainAccountCode"",
                         case
                             when e.""personType"" = 'O' then concat(tt.""TITLETHAIBEGIN"", ' ', e.""t_ogName"", ' ', tt.""TITLETHAIEND"")
                             when e.""personType"" = 'P' then concat(tt.""TITLETHAIBEGIN"", ' ', e.""t_firstName"", ' ', e.""t_lastName"", ' ', tt.""TITLETHAIEND"")
                             else null
                         end as ""MainAccountName"",
                         p.""insureeCode"",
                         case
                             when e_i.""personType"" = 'O' then concat(tt_i.""TITLETHAIBEGIN"", ' ', e_i.""t_ogName"", ' ', tt_i.""TITLETHAIEND"")
                             when e_i.""personType"" = 'P' then concat(tt_i.""TITLETHAIBEGIN"", ' ', e_i.""t_firstName"", ' ', e_i.""t_lastName"", ' ', tt_i.""TITLETHAIEND"")
                             else null
                         end as ""insureeName"",
                         it.""class"",
                         it.""subClass"",
                         m.""licenseNo"",
                         pv.t_provincename as ""province"",
                         m.""chassisNo"",
                         p.commout_rate as ""commOutRate"",
                         p.ovout_rate as ""ovOutRate"",
                         -1 * ((t.""subType"" *2) -1) *  t.ovamt as ""ovOutAmt"",
                         -1 * ((t.""subType"" *2) -1) *  t.commamt as ""commOutAmt"",
                         (case when t.""transType"" = 'DISC-OUT' then null when t.""transType"" = 'COMM-OUT' then (case when t.""agentCode2"" is null then p.commout1_rate else p.commout2_rate end ) else (case when t.""agentCode2"" is null then p.ovout1_rate else p.ovout2_rate end ) end ) as ""commOvOutRate"",
                         -- (case when t.""transType"" = 'COMM-OUT' then t.commamt else t.ovamt end ) as ""commOvOutAmt"",
                         ( -1 * ((t.""subType"" *2) -1) *  t.totalamt  ) as ""commOvOutAmt"",
                         (select t2.dfrpreferno from static_data.""Transactions"" t2 where ""transType"" = 'COMM-OUT' and t2.id = t.id) as ""commOutDfRpReferNo"",
                         (select t2.rprefdate from static_data.""Transactions"" t2 where ""transType"" = 'COMM-OUT' and t2.id = t.id) as ""commOutRpRefDate"",
                         (select t2.paidamt from static_data.""Transactions"" t2 where ""transType"" = 'COMM-OUT' and t2.id = t.id) as ""commOutPaidAmt"",
                         (select t2.remainamt from static_data.""Transactions"" t2 where ""transType"" = 'COMM-OUT' and t2.id = t.id) as ""commOutDiffAmt"",
                         (select t2.paidamt from static_data.""Transactions"" t2 where ""transType"" = 'OV-OUT' and t2.id = t.id) as ""ovOutPaidAmt"",
                         (select t2.remainamt from static_data.""Transactions"" t2 where ""transType"" = 'OV-OUT' and t2.id = t.id) as ""ovOutDiffAmt"",
                         p.""issueDate"",
                         p.createusercode as ""policyCreateUserCode"",
                         a.""contactPersonID"" as ""mainAccountContactPersonId"",
                         t.""insurerCode"",
                         p.insurancestatus as ""insurancestatus"",
                         t.""transType"" as ""transactionType""
                         from
                         static_data.""Transactions"" t 
                         left join static_data.""b_jupgrs"" jupgr on jupgr.""policyNo"" = t.""policyNo"" and  t.dftxno = jupgr.dftxno and jupgr.""seqNo"" = t.""seqNo"" 
                         left join static_data.""Policies"" p on jupgr.polid = p.id
                         left join static_data.""Insurees"" i on p.""insureeCode"" = i.""insureeCode""
                         left join static_data.""Entities"" e_i on i.""entityID"" = e_i.id
                         left join static_data.""Titles"" tt_i on tt_i.""TITLEID"" = e_i.""titleID""
                         left join static_data.""InsureTypes"" it on p.""insureID"" = it.id
                         left join static_data.""Motors"" m on p.""itemList"" = m.id
                         left join static_data.provinces pv on m.""motorprovinceID"" = pv.provinceid
                         left join static_data.b_jacashiers c on t.""premin-dfrpreferno"" = c.dfrpreferno
                         left join static_data.b_jaaraps ar on (ar.cashierreceiveno = c.cashierreceiveno)
                         left join static_data.b_jaarapds ard on ard.keyidm = ar.id and ard.polid = p.id
                         left join static_data.""Agents"" a on a.""agentCode"" = t.mainaccountcode  and a.""lastversion"" = 'Y'
                         left join static_data.""Entities"" e on a.""entityID"" = e.id
                         left join static_data.""Titles"" tt on tt.""TITLEID"" = e.""titleID""
                         where t.""transType"" in ('COMM-OUT', 'OV-OUT','DISC-OUT')
                         and t.txtype2 in ('1', '2', '3', '4', '5')
                         -- ( ( t.""transType"" in ('COMM-OUT', 'OV-OUT','DISC-OUT') and t.txtype2 in ('1', '2' ) ) or ( t.""transType"" in ('COMM-IN', 'OV-IN','DISC-IN') and t.txtype2 in ('3', '4', '5' ) ) )
                         -- and t.txtype2 in ('1', '2', '3', '4', '5')
                         and t.status = 'N' 
                         and p.""lastVersion"" = 'Y'
                         and t.""premin-dfrpreferno"" is not null
                         and t.""premin-rprefdate"" is not null
                         and t.dfrpreferno is not null
                         and t.rprefdate is not null 
                         and jupgr.""installmenttype"" = 'A' ";
            var sqlcount = $@"select
                         count (*) as count
                         from
                         static_data.""Transactions"" t 
                         left join static_data.""b_jupgrs"" jupgr on jupgr.""policyNo"" = t.""policyNo"" and  t.dftxno = jupgr.dftxno and jupgr.""seqNo"" = t.""seqNo"" 
                         left join static_data.""Policies"" p on jupgr.polid = p.id
                         left join static_data.""Insurees"" i on p.""insureeCode"" = i.""insureeCode""
                         left join static_data.""Entities"" e_i on i.""entityID"" = e_i.id
                         left join static_data.""Titles"" tt_i on tt_i.""TITLEID"" = e_i.""titleID""
                         left join static_data.""InsureTypes"" it on p.""insureID"" = it.id
                         left join static_data.""Motors"" m on p.""itemList"" = m.id
                         left join static_data.provinces pv on m.""motorprovinceID"" = pv.provinceid
                         left join static_data.b_jacashiers c on t.""premin-dfrpreferno"" = c.dfrpreferno
                         left join static_data.b_jaaraps ar on (ar.cashierreceiveno = c.cashierreceiveno)
                         left join static_data.b_jaarapds ard on ard.keyidm = ar.id and ard.polid = p.id
                         left join static_data.""Agents"" a on a.""agentCode"" = t.mainaccountcode  and a.""lastversion"" = 'Y'
                         left join static_data.""Entities"" e on a.""entityID"" = e.id
                         left join static_data.""Titles"" tt on tt.""TITLEID"" = e.""titleID""
                         where t.""transType"" in ('COMM-OUT', 'OV-OUT','DISC-OUT')
                         and t.txtype2 in ('1', '2', '3', '4', '5')
                         -- ( ( t.""transType"" in ('COMM-OUT', 'OV-OUT','DISC-OUT') and t.txtype2 in ('1', '2' ) ) or ( t.""transType"" in ('COMM-IN', 'OV-IN','DISC-IN') and t.txtype2 in ('3', '4', '5' ) ) )
                         -- and t.txtype2 in ('1', '2', '3', '4', '5')
                         and t.status = 'N' 
                         and p.""lastVersion"" = 'Y'
                         and t.""premin-dfrpreferno"" is not null
                         and t.""premin-rprefdate"" is not null
                         and t.dfrpreferno is not null
                         and t.rprefdate is not null 
                         and jupgr.""installmenttype"" = 'A' ";
            sql = await GetWhereSql(data, sql, "CommOutOvOut");
            sqlcount = await GetWhereSql(data, sqlcount, "CommOutOvOut", true);
            var json = await _dataContext.CommOutOvOutDatas.FromSqlRaw(sql).ToListAsync();
            var count = await _dataContext.DatasCount.FromSqlRaw(sqlcount).ToListAsync();
            return new CommOutOvOutReportResult(json, count[0].Count);
        }

        public async Task<CommOutOvOutReportResult?> GetCommOutOvOutOutstandingReportJson(ArApReportInput data)
        {
            var sql = $@"select
                         -1 * ((t.""subType"" *2) -1) * t.withheld as ""withheld"",
                         -1 * ((t.""subType"" *2) -1) * t.totalamt as ""totalamt"",
                         t.""policyNo"",
                         t.""endorseNo"",
                         jupgr.""invoiceNo"",
                         t.""seqNo"",
                         t.""transType"",
                         ar.cashierreceiveno as ""cashierReceiveNo"",
                         c.cashierdate as ""cashierDate"",
                         ar.cashieramt as ""cashierAmt"",
                         c.receivetype as ""cashierReceiveType"",
                         c.refno as ""cashierRefNo"",
                         --c.refdate,
                         t.""premin-dfrpreferno"" as ""premInDfRpReferNo"",
                         t.""premin-rprefdate"" as ""premInRpRefDate"",
                         t.""dfrpreferno"" as ""dfRpReferNo"",
                         t.""rprefdate"" as ""rpRefDate"",
                         p.grossprem as ""grossPrem"",
                         p.specdiscrate as ""specDiscRate"",
                         p.specdiscamt as ""specDiscAmt"",
                         -1 * ((t.""subType"" *2) -1) * t.netgrossprem as ""netGrossPrem"",
                         -1 * ((t.""subType"" *2) -1) * t.duty as ""duty"",
                         -1 * ((t.""subType"" *2) -1) * t.tax as ""tax"",
                         -1 * ((t.""subType"" *2) -1) * t.totalprem as ""totalPrem"",
                         t.netflag as ""netFlag"",
                         p.""actDate"",
                         p.""expDate"",
                         t.mainaccountcode as ""mainAccountCode"",
                         case
                             when e.""personType"" = 'O' then concat(tt.""TITLETHAIBEGIN"", ' ', e.""t_ogName"", ' ', tt.""TITLETHAIEND"")
                             when e.""personType"" = 'P' then concat(tt.""TITLETHAIBEGIN"", ' ', e.""t_firstName"", ' ', e.""t_lastName"", ' ', tt.""TITLETHAIEND"")
                             else null
                         end as ""MainAccountName"",
                         p.""insureeCode"",
                         case
                             when e_i.""personType"" = 'O' then concat(tt_i.""TITLETHAIBEGIN"", ' ', e_i.""t_ogName"", ' ', tt_i.""TITLETHAIEND"")
                             when e_i.""personType"" = 'P' then concat(tt_i.""TITLETHAIBEGIN"", ' ', e_i.""t_firstName"", ' ', e_i.""t_lastName"", ' ', tt_i.""TITLETHAIEND"")
                             else null
                         end as ""insureeName"",
                         it.""class"",
                         it.""subClass"",
                         m.""licenseNo"",
                         pv.t_provincename as ""province"",
                         m.""chassisNo"",
                         p.commout_rate as ""commOutRate"",
                         p.ovout_rate as ""ovOutRate"",
                         -1 * ((t.""subType"" *2) -1) * t.ovamt as ""ovOutAmt"",
                         -1 * ((t.""subType"" *2) -1) * t.commamt as ""commOutAmt"",
                         (case when t.""transType"" = 'DISC-OUT' then null when t.""transType"" = 'COMM-OUT' then (case when t.""agentCode2"" is null then p.commout1_rate else p.commout2_rate end ) else (case when t.""agentCode2"" is null then p.ovout1_rate else p.ovout2_rate end ) end ) as ""commOvOutRate"",
                         -- (case when t.""transType"" = 'COMM-OUT' then t.commamt else t.ovamt end ) as ""commOvOutAmt"",
                          ( -1 * ((t.""subType"" *2) -1) * t.totalamt  ) as ""commOvOutAmt"",
                         (select t2.dfrpreferno from static_data.""Transactions"" t2 where ""transType"" = 'COMM-OUT' and t2.id = t.id) as ""commOutDfRpReferNo"",
                         (select t2.rprefdate from static_data.""Transactions"" t2 where ""transType"" = 'COMM-OUT' and t2.id = t.id) as ""commOutRpRefDate"",
                         (select t2.paidamt from static_data.""Transactions"" t2 where ""transType"" = 'COMM-OUT' and t2.id = t.id) as ""commOutPaidAmt"",
                         (select t2.remainamt from static_data.""Transactions"" t2 where ""transType"" = 'COMM-OUT' and t2.id = t.id) as ""commOutDiffAmt"",
                         (select t2.paidamt from static_data.""Transactions"" t2 where ""transType"" = 'OV-OUT' and t2.id = t.id) as ""ovOutPaidAmt"",
                         (select t2.remainamt from static_data.""Transactions"" t2 where ""transType"" = 'OV-OUT' and t2.id = t.id) as ""ovOutDiffAmt"",
                         p.""issueDate"",
                         p.createusercode as ""policyCreateUserCode"",
                         a.""contactPersonID"" as ""mainAccountContactPersonId"",
                         t.""insurerCode"",
                         p.insurancestatus as ""insurancestatus"",
                         t.""transType"" as ""transactionType""
                         from
                         static_data.""Transactions"" t 
                         left join static_data.""b_jupgrs"" jupgr on jupgr.""policyNo"" = t.""policyNo"" and  t.dftxno = jupgr.dftxno and jupgr.""seqNo"" = t.""seqNo"" 
                         left join static_data.""Policies"" p on jupgr.polid = p.id
                         left join static_data.""Insurees"" i on p.""insureeCode"" = i.""insureeCode""
                         left join static_data.""Entities"" e_i on i.""entityID"" = e_i.id
                         left join static_data.""Titles"" tt_i on tt_i.""TITLEID"" = e_i.""titleID""
                         left join static_data.""InsureTypes"" it on p.""insureID"" = it.id
                         left join static_data.""Motors"" m on p.""itemList"" = m.id
                         left join static_data.provinces pv on m.""motorprovinceID"" = pv.provinceid
                         left join static_data.b_jacashiers c on t.""premin-dfrpreferno"" = c.dfrpreferno
                         left join static_data.b_jaaraps ar on (ar.cashierreceiveno = c.cashierreceiveno)
                         left join static_data.b_jaarapds ard on ard.keyidm = ar.id and ard.polid = p.id
                         left join static_data.""Agents"" a on a.""agentCode"" = t.mainaccountcode  and a.""lastversion"" = 'Y'
                         left join static_data.""Entities"" e on a.""entityID"" = e.id
                         left join static_data.""Titles"" tt on tt.""TITLEID"" = e.""titleID""
                         where t.""transType"" in ('COMM-OUT', 'OV-OUT','DISC-OUT')
                         -- ( ( t.""transType"" in ('COMM-OUT', 'OV-OUT','DISC-OUT') and t.txtype2 in ('1', '2' ) ) or ( t.""transType"" in ('COMM-IN', 'OV-IN','DISC-IN') and t.txtype2 in ('3', '4', '5' ) ) )
                         and t.txtype2 in ('1', '2', '3', '4', '5')
                         and t.status = 'N' 
                         and p.""lastVersion"" = 'Y'
                         and t.""premin-dfrpreferno"" is not null
                         and t.""premin-rprefdate"" is not null
                         and t.dfrpreferno is null 
                         and t.rprefdate is null
                         and jupgr.""installmenttype"" = 'A' ";
            //if (!string.IsNullOrEmpty(data.AsAtDate))
            //{
            //    sql += $@"or (t.dfrpreferno is not null and t.rprefdate is not null and t.rprefdate <= '{data.AsAtDate}') ";
            //}
            //sql += $@")";
            var sqlcount = $@"select
                         count (*) as count
                         from
                         static_data.""Transactions"" t 
                         left join static_data.""b_jupgrs"" jupgr on jupgr.""policyNo"" = t.""policyNo"" and  t.dftxno = jupgr.dftxno and jupgr.""seqNo"" = t.""seqNo"" 
                         left join static_data.""Policies"" p on jupgr.polid = p.id
                         left join static_data.""Insurees"" i on p.""insureeCode"" = i.""insureeCode""
                         left join static_data.""Entities"" e_i on i.""entityID"" = e_i.id
                         left join static_data.""Titles"" tt_i on tt_i.""TITLEID"" = e_i.""titleID""
                         left join static_data.""InsureTypes"" it on p.""insureID"" = it.id
                         left join static_data.""Motors"" m on p.""itemList"" = m.id
                         left join static_data.provinces pv on m.""motorprovinceID"" = pv.provinceid
                         left join static_data.b_jacashiers c on t.""premin-dfrpreferno"" = c.dfrpreferno
                         left join static_data.b_jaaraps ar on (ar.cashierreceiveno = c.cashierreceiveno)
                         left join static_data.b_jaarapds ard on ard.keyidm = ar.id and ard.polid = p.id
                         left join static_data.""Agents"" a on a.""agentCode"" = t.mainaccountcode  and a.""lastversion"" = 'Y'
                         left join static_data.""Entities"" e on a.""entityID"" = e.id
                         left join static_data.""Titles"" tt on tt.""TITLEID"" = e.""titleID""
                         where t.""transType"" in ('COMM-OUT', 'OV-OUT','DISC-OUT')
                         -- ( ( t.""transType"" in ('COMM-OUT', 'OV-OUT','DISC-OUT') and t.txtype2 in ('1', '2' ) ) or ( t.""transType"" in ('COMM-IN', 'OV-IN','DISC-IN') and t.txtype2 in ('3', '4', '5' ) ) )
                         and t.txtype2 in ('1', '2', '3', '4', '5')
                         and t.status = 'N' 
                         and p.""lastVersion"" = 'Y'
                         and t.""premin-dfrpreferno"" is not null
                         and t.""premin-rprefdate"" is not null
                         and t.dfrpreferno is null 
                         and t.rprefdate is null
                         and jupgr.""installmenttype"" = 'A' ";
            sql = await GetWhereSql(data, sql, "CommOutOvOut");
            sqlcount = await GetWhereSql(data, sqlcount, "CommOutOvOut", true);
            var json = await _dataContext.CommOutOvOutDatas.FromSqlRaw(sql).ToListAsync();
            var count = await _dataContext.DatasCount.FromSqlRaw(sqlcount).ToListAsync();
            return new CommOutOvOutReportResult(json, count[0].Count);

           
        }

        public async Task<PremOutCommInOvInReportResult?> GetPremOutOpenItemReportJson(ArApReportInput data)
        {
            var sql = $@"select
                        null as ""cashierReceiveNo"",
                        null as ""CashierDate"",
                        null as ""CashierAmt"",
                        null as ""CashierReceiveType"",
                        null as ""CashierRefNo"",
                        null as ""CashierRefDate"",
                        null as ""MainAccountName"",
                        null as ""CommOutRate"",
                        null as ""CommOutAmt"",
                        null as ""OvOutRate"",
                        null as ""OvOutAmt"",
                         t.""policyNo"",
                         t.""endorseNo"",
                         t.""agentCode"",
                         jupgr.""invoiceNo"",
                         jupgr.""taxInvoiceNo"",
                         null as ""DfRpReferNo"",
                         null as ""transType"",
                         null as ""RpRefDate"",
                         null as ""CommOvInAmt"",
                         null as ""CommOvInRate"",
                         t.""seqNo"",
                         t.""premin-dfrpreferno"" as ""premInDfRpReferNo"",
                         t.""premin-rprefdate"" as ""premInRpRefDate"",
                         t.""premout-dfrpreferno"" as ""premOutDfRpReferNo"",
                         t.""premout-rprefdate"" as ""premOutRpRefDate"",
                         p.""actDate"",
                         p.""expDate"",
                         t.""insurerCode"",
                         case
                             when e_ier.""personType"" = 'O' then concat(tt_ier.""TITLETHAIBEGIN"", ' ', e_ier.""t_ogName"", ' ', tt_ier.""TITLETHAIEND"")
                             when e_ier.""personType"" = 'P' then concat(tt_ier.""TITLETHAIBEGIN"", ' ', e_ier.""t_firstName"", ' ', e_ier.""t_lastName"", ' ', tt_ier.""TITLETHAIEND"")
                             else null
                         end as ""insurerName"",
                         p.""insureeCode"",
                         case
                             when e_iee.""personType"" = 'O' then concat(tt_iee.""TITLETHAIBEGIN"", ' ', e_iee.""t_ogName"", ' ', tt_iee.""TITLETHAIEND"")
                             when e_iee.""personType"" = 'P' then concat(tt_iee.""TITLETHAIBEGIN"", ' ', e_iee.""t_firstName"", ' ', e_iee.""t_lastName"", ' ', tt_iee.""TITLETHAIEND"")
                             else null
                         end as ""insureeName"",
                         it.""class"",
                         it.""subClass"",
                         p.grossprem as ""grossPrem"",
                         p.specdiscrate as ""specDiscRate"",
                         p.specdiscamt as ""specDiscAmt"",
                         -1 * ((t.""subType"" *2) -1) * t.netgrossprem as ""netGrossPrem"",
                         -1 * ((t.""subType"" *2) -1) * t.duty as duty,
                         -1 * ((t.""subType"" *2) -1) * t.tax as tax,
                         -1 * ((t.""subType"" *2) -1) * t.totalprem as ""totalPrem"",
                         -1 * ((t.""subType"" *2) -1) * t.withheld as ""withheld"",
                         -1 * ((t.""subType"" *2) -1) * t.totalamt as ""totalamt"",
                         t.netflag as ""netFlag"",
                         p.commin_rate as ""commInRate"",
                         p.commin_amt as ""commInAmt"",
                         (select t2.dfrpreferno from static_data.""Transactions"" t2 where ""transType"" = 'COMM-IN' and t2.id = t.id) as ""commInDfRpReferNo"",
                         (select t2.rprefdate from static_data.""Transactions"" t2 where ""transType"" = 'COMM-IN' and t2.id = t.id) as ""commInRpRefDate"",
                         (select t2.paidamt from static_data.""Transactions"" t2 where ""transType"" = 'COMM-IN' and t2.id = t.id) as ""commInPaidAmt"",
                         (select t2.remainamt from static_data.""Transactions"" t2 where ""transType"" = 'COMM-IN' and t2.id = t.id) as ""commInDiffAmt"",
                         p.ovin_rate as ""ovInRate"",
                         p.ovin_amt as ""ovInAmt"",
                         (select t2.dfrpreferno from static_data.""Transactions"" t2 where ""transType"" = 'OV-IN' and t2.id = t.id) as ""ovInDfRpReferNo"",
                         (select t2.rprefdate from static_data.""Transactions"" t2 where ""transType"" = 'OV-IN' and t2.id = t.id) as ""ovInRpRefDate"",
                         (select t2.paidamt from static_data.""Transactions"" t2 where ""transType"" = 'OV-IN' and t2.id = t.id) as ""ovInPaidAmt"",
                         (select t2.remainamt from static_data.""Transactions"" t2 where ""transType"" = 'OV-IN' and t2.id = t.id) as ""ovInDiffAmt"",
                         p.""issueDate"",
                         p.createusercode as ""policyCreateUserCode"",
                         a.""contactPersonID"" as ""mainAccountContactPersonId"",
                         t.mainaccountcode as ""mainAccountCode"",
                         p.insurancestatus as ""insurancestatus"",
                         t.""transType"" as ""transactionType""
                         from
                         static_data.""Transactions"" t 
                         left join static_data.""b_jupgrs"" jupgr on jupgr.""policyNo"" = t.""policyNo"" and  t.dftxno = jupgr.dftxno and jupgr.""seqNo"" = t.""seqNo"" 
                         left join static_data.""Policies"" p on jupgr.polid = p.id
                         left join static_data.""Insurers"" ier on t.""insurerCode"" = ier.""insurerCode""  and ier.""lastversion"" = 'Y'
                         left join static_data.""Entities"" e_ier on ier.""entityID"" = e_ier.id
                         left join static_data.""Titles"" tt_ier on tt_ier.""TITLEID"" = e_ier.""titleID""
                         left join static_data.""Insurees"" iee on p.""insureeCode"" = iee.""insureeCode""
                         left join static_data.""Entities"" e_iee on iee.""entityID"" = e_iee.id
                         left join static_data.""Titles"" tt_iee on tt_iee.""TITLEID"" = e_iee.""titleID""
                         left join static_data.""InsureTypes"" it on p.""insureID"" = it.id
                         left join static_data.b_jaaraps ar on (ar.billadvisorno = t.billadvisorno)
                         left join static_data.b_jaarapds ard on ard.keyidm = ar.id and ard.polid = p.id
                         left join static_data.""Agents"" a on a.""agentCode"" = t.mainaccountcode  and a.""lastversion"" = 'Y'
                         left join static_data.""Entities"" e on a.""entityID"" = e.id
                         left join static_data.""Titles"" tt on tt.""TITLEID"" = e.""titleID""
                         where t.""transType"" = 'PREM-OUT'
                         and t.txtype2 in ('1', '2', '3', '4', '5')
                         and t.status = 'N' 
                         and p.""lastVersion"" = 'Y'
                         and t.""premin-dfrpreferno"" is not null
                         and t.""premin-rprefdate"" is not null 
                         and jupgr.""installmenttype"" = 'I' ";
            var sqlcount = $@"select
                         count (*) as count
                         from
                         static_data.""Transactions"" t 
                         left join static_data.""b_jupgrs"" jupgr on jupgr.""policyNo"" = t.""policyNo"" and  t.dftxno = jupgr.dftxno and jupgr.""seqNo"" = t.""seqNo"" 
                         left join static_data.""Policies"" p on jupgr.polid = p.id
                         left join static_data.""Insurers"" ier on t.""insurerCode"" = ier.""insurerCode""  and ier.""lastversion"" = 'Y'
                         left join static_data.""Entities"" e_ier on ier.""entityID"" = e_ier.id
                         left join static_data.""Titles"" tt_ier on tt_ier.""TITLEID"" = e_ier.""titleID""
                         left join static_data.""Insurees"" iee on p.""insureeCode"" = iee.""insureeCode""
                         left join static_data.""Entities"" e_iee on iee.""entityID"" = e_iee.id
                         left join static_data.""Titles"" tt_iee on tt_iee.""TITLEID"" = e_iee.""titleID""
                         left join static_data.""InsureTypes"" it on p.""insureID"" = it.id
                         left join static_data.b_jaaraps ar on (ar.billadvisorno = t.billadvisorno)
                         left join static_data.b_jaarapds ard on ard.keyidm = ar.id and ard.polid = p.id
                         left join static_data.""Agents"" a on a.""agentCode"" = t.mainaccountcode  and a.""lastversion"" = 'Y'
                         left join static_data.""Entities"" e on a.""entityID"" = e.id
                         left join static_data.""Titles"" tt on tt.""TITLEID"" = e.""titleID""
                         where t.""transType"" = 'PREM-OUT'
                         and t.txtype2 in ('1', '2', '3', '4', '5')
                         and t.status = 'N' 
                         and p.""lastVersion"" = 'Y'
                         and t.""premin-dfrpreferno"" is not null
                         and t.""premin-rprefdate"" is not null 
                         and jupgr.""installmenttype"" = 'I' ";
            sql = await GetWhereSql(data, sql, "PremOut");
            sqlcount = await GetWhereSql(data, sqlcount, "PremOut", true);
            var json = await _dataContext.PremOutCommInOvInDatas.FromSqlRaw(sql).ToListAsync();
            var count = await _dataContext.DatasCount.FromSqlRaw(sqlcount).ToListAsync();
            return new PremOutCommInOvInReportResult(json, count[0].Count);
        }

        public async Task<PremOutCommInOvInReportResult?> GetPremOutClearingReportJson(ArApReportInput data)
        {
            var sql = $@"select
                        null as ""cashierReceiveNo"",
                        null as ""CashierDate"",
                        null as ""CashierAmt"",
                        null as ""CashierReceiveType"",
                        null as ""CashierRefNo"",
                        null as ""CashierRefDate"",
                        null as ""MainAccountName"",
                        null as ""CommOutRate"",
                        null as ""CommOutAmt"",
                        null as ""OvOutRate"",
                        null as ""OvOutAmt"",
                         t.""policyNo"",
                         t.""endorseNo"",
                         t.""agentCode"",
                         jupgr.""invoiceNo"",
                         jupgr.""taxInvoiceNo"",
                         null as ""DfRpReferNo"",
                         null as ""transType"",
                         null as ""RpRefDate"",
                         null as ""CommOvInAmt"",
                         null as ""CommOvInRate"",
                         t.""seqNo"",
                         t.""premin-dfrpreferno"" as ""premInDfRpReferNo"",
                         t.""premin-rprefdate"" as ""premInRpRefDate"",
                         t.""premout-dfrpreferno"" as ""premOutDfRpReferNo"",
                         t.""premout-rprefdate"" as ""premOutRpRefDate"",
                         p.""actDate"",
                         p.""expDate"",
                         t.""insurerCode"",
                         case
                             when e_ier.""personType"" = 'O' then concat(tt_ier.""TITLETHAIBEGIN"", ' ', e_ier.""t_ogName"", ' ', tt_ier.""TITLETHAIEND"")
                             when e_ier.""personType"" = 'P' then concat(tt_ier.""TITLETHAIBEGIN"", ' ', e_ier.""t_firstName"", ' ', e_ier.""t_lastName"", ' ', tt_ier.""TITLETHAIEND"")
                             else null
                         end as ""insurerName"",
                         p.""insureeCode"",
                         case
                             when e_iee.""personType"" = 'O' then concat(tt_iee.""TITLETHAIBEGIN"", ' ', e_iee.""t_ogName"", ' ', tt_iee.""TITLETHAIEND"")
                             when e_iee.""personType"" = 'P' then concat(tt_iee.""TITLETHAIBEGIN"", ' ', e_iee.""t_firstName"", ' ', e_iee.""t_lastName"", ' ', tt_iee.""TITLETHAIEND"")
                             else null
                         end as ""insureeName"",
                         it.""class"",
                         it.""subClass"",
                         p.grossprem as ""grossPrem"",
                         p.specdiscrate as ""specDiscRate"",
                         p.specdiscamt as ""specDiscAmt"",
                         -1 * ((t.""subType"" *2) -1) * t.netgrossprem as ""netGrossPrem"",
                         -1 * ((t.""subType"" *2) -1) * t.duty as duty,
                         -1 * ((t.""subType"" *2) -1) * t.tax as tax,
                         -1 * ((t.""subType"" *2) -1) * t.totalprem as ""totalPrem"",
                         -1 * ((t.""subType"" *2) -1) * t.withheld as ""withheld"",
                         -1 * ((t.""subType"" *2) -1) * t.totalamt as ""totalamt"",
                         t.netflag as ""netFlag"",
                         p.commin_rate as ""commInRate"",
                          p.commin_amt as ""commInAmt"",
                         (select t2.dfrpreferno from static_data.""Transactions"" t2 where ""transType"" = 'COMM-IN' and t2.id = t.id) as ""commInDfRpReferNo"",
                         (select t2.rprefdate from static_data.""Transactions"" t2 where ""transType"" = 'COMM-IN' and t2.id = t.id) as ""commInRpRefDate"",
                         (select t2.paidamt from static_data.""Transactions"" t2 where ""transType"" = 'COMM-IN' and t2.id = t.id) as ""commInPaidAmt"",
                         (select t2.remainamt from static_data.""Transactions"" t2 where ""transType"" = 'COMM-IN' and t2.id = t.id) as ""commInDiffAmt"",
                         p.ovin_rate as ""ovInRate"",
                         p.ovin_amt as ""ovInAmt"",
                         (select t2.dfrpreferno from static_data.""Transactions"" t2 where ""transType"" = 'OV-IN' and t2.id = t.id) as ""ovInDfRpReferNo"",
                         (select t2.rprefdate from static_data.""Transactions"" t2 where ""transType"" = 'OV-IN' and t2.id = t.id) as ""ovInRpRefDate"",
                         (select t2.paidamt from static_data.""Transactions"" t2 where ""transType"" = 'OV-IN' and t2.id = t.id) as ""ovInPaidAmt"",
                         (select t2.remainamt from static_data.""Transactions"" t2 where ""transType"" = 'OV-IN' and t2.id = t.id) as ""ovInDiffAmt"",
                         p.""issueDate"",
                         p.createusercode as ""policyCreateUserCode"",
                         a.""contactPersonID"" as ""mainAccountContactPersonId"",
                         t.mainaccountcode as ""mainAccountCode"",
                         p.insurancestatus as ""insurancestatus"",
                         t.""transType"" as ""transactionType""
                         from
                         static_data.""Transactions"" t 
                         left join static_data.""b_jupgrs"" jupgr on jupgr.""policyNo"" = t.""policyNo"" and  t.dftxno = jupgr.dftxno and jupgr.""seqNo"" = t.""seqNo"" 
                         left join static_data.""Policies"" p on jupgr.polid = p.id
                         left join static_data.""Insurers"" ier on t.""insurerCode"" = ier.""insurerCode""  and ier.""lastversion"" = 'Y'
                         left join static_data.""Entities"" e_ier on ier.""entityID"" = e_ier.id
                         left join static_data.""Titles"" tt_ier on tt_ier.""TITLEID"" = e_ier.""titleID""
                         left join static_data.""Insurees"" iee on p.""insureeCode"" = iee.""insureeCode""
                         left join static_data.""Entities"" e_iee on iee.""entityID"" = e_iee.id
                         left join static_data.""Titles"" tt_iee on tt_iee.""TITLEID"" = e_iee.""titleID""
                         left join static_data.""InsureTypes"" it on p.""insureID"" = it.id
                         left join static_data.b_jaaraps ar on (ar.billadvisorno = t.billadvisorno)
                         left join static_data.b_jaarapds ard on ard.keyidm = ar.id and ard.polid = p.id
                         left join static_data.""Agents"" a on a.""agentCode"" = t.mainaccountcode   and a.""lastversion"" = 'Y'
                         left join static_data.""Entities"" e on a.""entityID"" = e.id
                         left join static_data.""Titles"" tt on tt.""TITLEID"" = e.""titleID""
                         where t.""transType"" = 'PREM-OUT'
                         and t.txtype2 in ('1', '2', '3', '4', '5')
                         and t.status = 'N' 
                         and p.""lastVersion"" = 'Y'
                         and t.""premin-dfrpreferno"" is not null
                         and t.""premin-rprefdate"" is not null
                         and t.""premout-dfrpreferno"" is not null
                         and t.""premout-rprefdate"" is not null 
                         and jupgr.""installmenttype"" = 'I' ";
            var sqlcount = $@"select
                         count (*) as count
                         from
                         static_data.""Transactions"" t 
                         left join static_data.""b_jupgrs"" jupgr on jupgr.""policyNo"" = t.""policyNo"" and  t.dftxno = jupgr.dftxno and jupgr.""seqNo"" = t.""seqNo"" 
                         left join static_data.""Policies"" p on jupgr.polid = p.id
                         left join static_data.""Insurers"" ier on t.""insurerCode"" = ier.""insurerCode""  and ier.""lastversion"" = 'Y'
                         left join static_data.""Entities"" e_ier on ier.""entityID"" = e_ier.id
                         left join static_data.""Titles"" tt_ier on tt_ier.""TITLEID"" = e_ier.""titleID""
                         left join static_data.""Insurees"" iee on p.""insureeCode"" = iee.""insureeCode""
                         left join static_data.""Entities"" e_iee on iee.""entityID"" = e_iee.id
                         left join static_data.""Titles"" tt_iee on tt_iee.""TITLEID"" = e_iee.""titleID""
                         left join static_data.""InsureTypes"" it on p.""insureID"" = it.id
                         left join static_data.b_jaaraps ar on (ar.billadvisorno = t.billadvisorno)
                         left join static_data.b_jaarapds ard on ard.keyidm = ar.id and ard.polid = p.id
                         left join static_data.""Agents"" a on a.""agentCode"" = t.mainaccountcode   and a.""lastversion"" = 'Y'
                         left join static_data.""Entities"" e on a.""entityID"" = e.id
                         left join static_data.""Titles"" tt on tt.""TITLEID"" = e.""titleID""
                         where t.""transType"" = 'PREM-OUT'
                         and t.txtype2 in ('1', '2', '3', '4', '5')
                         and t.status = 'N' 
                         and p.""lastVersion"" = 'Y'
                         and t.""premin-dfrpreferno"" is not null
                         and t.""premin-rprefdate"" is not null
                         and t.""premout-dfrpreferno"" is not null
                         and t.""premout-rprefdate"" is not null 
                         and jupgr.""installmenttype"" = 'I' ";
            sql = await GetWhereSql(data, sql, "PremOut");
            sqlcount = await GetWhereSql(data, sqlcount, "PremOut", true);
            var json = await _dataContext.PremOutCommInOvInDatas.FromSqlRaw(sql).ToListAsync();
            var count = await _dataContext.DatasCount.FromSqlRaw(sqlcount).ToListAsync();
            return new PremOutCommInOvInReportResult(json, count[0].Count);
        }

        public async Task<PremOutCommInOvInReportResult?> GetPremOutOutstandingReportJson(ArApReportInput data)
        {
            var sql = $@"select
                        null as ""cashierReceiveNo"",
                        null as ""CashierDate"",
                        null as ""CashierAmt"",
                        null as ""CashierReceiveType"",
                        null as ""CashierRefNo"",
                        null as ""CashierRefDate"",
                        null as ""MainAccountName"",
                        null as ""CommOutRate"",
                        null as ""CommOutAmt"",
                        null as ""OvOutRate"",
                        null as ""OvOutAmt"",
                         t.""policyNo"",
                         t.""endorseNo"",
                         t.""agentCode"",
                         null as ""DfRpReferNo"",
                         null as ""transType"",
                         null as ""RpRefDate"",
                         null as ""CommOvInAmt"",
                         null as ""CommOvInRate"",
                         jupgr.""invoiceNo"",
                         jupgr.""taxInvoiceNo"",
                         t.""seqNo"",
                         t.""premin-dfrpreferno"" as ""premInDfRpReferNo"",
                         t.""premin-rprefdate"" as ""premInRpRefDate"",
                         t.""premout-dfrpreferno"" as ""premOutDfRpReferNo"",
                         t.""premout-rprefdate"" as ""premOutRpRefDate"",
                         p.""actDate"",
                         p.""expDate"",
                         t.""insurerCode"",
                         case
                             when e_ier.""personType"" = 'O' then concat(tt_ier.""TITLETHAIBEGIN"", ' ', e_ier.""t_ogName"", ' ', tt_ier.""TITLETHAIEND"")
                             when e_ier.""personType"" = 'P' then concat(tt_ier.""TITLETHAIBEGIN"", ' ', e_ier.""t_firstName"", ' ', e_ier.""t_lastName"", ' ', tt_ier.""TITLETHAIEND"")
                             else null
                         end as ""insurerName"",
                         p.""insureeCode"",
                         case
                             when e_iee.""personType"" = 'O' then concat(tt_iee.""TITLETHAIBEGIN"", ' ', e_iee.""t_ogName"", ' ', tt_iee.""TITLETHAIEND"")
                             when e_iee.""personType"" = 'P' then concat(tt_iee.""TITLETHAIBEGIN"", ' ', e_iee.""t_firstName"", ' ', e_iee.""t_lastName"", ' ', tt_iee.""TITLETHAIEND"")
                             else null
                         end as ""insureeName"",
                         it.""class"",
                         it.""subClass"",
                         p.grossprem as ""grossPrem"",
                         p.specdiscrate as ""specDiscRate"",
                         p.specdiscamt as ""specDiscAmt"",
                         -1 * ((t.""subType"" *2) -1) * t.netgrossprem as ""netGrossPrem"",
                         -1 * ((t.""subType"" *2) -1) * t.duty as duty,
                         -1 * ((t.""subType"" *2) -1) * t.tax as tax,
                         -1 * ((t.""subType"" *2) -1) * t.totalprem as ""totalPrem"",
                         -1 * ((t.""subType"" *2) -1) * t.withheld as ""withheld"",
                         -1 * ((t.""subType"" *2) -1) * t.totalamt as ""totalamt"",
                         t.netflag as ""netFlag"",
                         p.commin_rate as ""commInRate"",
                         p.commin_amt as ""commInAmt"",
                         (select t2.dfrpreferno from static_data.""Transactions"" t2 where ""transType"" = 'COMM-IN' and t2.id = t.id) as ""commInDfRpReferNo"",
                         (select t2.rprefdate from static_data.""Transactions"" t2 where ""transType"" = 'COMM-IN' and t2.id = t.id) as ""commInRpRefDate"",
                         (select t2.paidamt from static_data.""Transactions"" t2 where ""transType"" = 'COMM-IN' and t2.id = t.id) as ""commInPaidAmt"",
                         (select t2.remainamt from static_data.""Transactions"" t2 where ""transType"" = 'COMM-IN' and t2.id = t.id) as ""commInDiffAmt"",
                         p.ovin_rate as ""ovInRate"",
                         p.ovin_amt as ""ovInAmt"",
                         (select t2.dfrpreferno from static_data.""Transactions"" t2 where ""transType"" = 'OV-IN' and t2.id = t.id) as ""ovInDfRpReferNo"",
                         (select t2.rprefdate from static_data.""Transactions"" t2 where ""transType"" = 'OV-IN' and t2.id = t.id) as ""ovInRpRefDate"",
                         (select t2.paidamt from static_data.""Transactions"" t2 where ""transType"" = 'OV-IN' and t2.id = t.id) as ""ovInPaidAmt"",
                         (select t2.remainamt from static_data.""Transactions"" t2 where ""transType"" = 'OV-IN' and t2.id = t.id) as ""ovInDiffAmt"",
                         p.""issueDate"",
                         p.createusercode as ""policyCreateUserCode"",
                         a.""contactPersonID"" as ""mainAccountContactPersonId"",
                         t.mainaccountcode as ""mainAccountCode"",
                         p.insurancestatus as ""insurancestatus"",
                         t.""transType"" as ""transactionType""
                         from
                         static_data.""Transactions"" t 
                         left join static_data.""b_jupgrs"" jupgr on jupgr.""policyNo"" = t.""policyNo"" and  t.dftxno = jupgr.dftxno and jupgr.""seqNo"" = t.""seqNo"" 
                         left join static_data.""Policies"" p on jupgr.polid = p.id
                         left join static_data.""Insurers"" ier on t.""insurerCode"" = ier.""insurerCode""  and ier.""lastversion"" = 'Y'
                         left join static_data.""Entities"" e_ier on ier.""entityID"" = e_ier.id
                         left join static_data.""Titles"" tt_ier on tt_ier.""TITLEID"" = e_ier.""titleID""
                         left join static_data.""Insurees"" iee on p.""insureeCode"" = iee.""insureeCode""
                         left join static_data.""Entities"" e_iee on iee.""entityID"" = e_iee.id
                         left join static_data.""Titles"" tt_iee on tt_iee.""TITLEID"" = e_iee.""titleID""
                         left join static_data.""InsureTypes"" it on p.""insureID"" = it.id
                         left join static_data.b_jaaraps ar on (ar.billadvisorno = t.billadvisorno)
                         left join static_data.b_jaarapds ard on ard.keyidm = ar.id and ard.polid = p.id
                         left join static_data.""Agents"" a on a.""agentCode"" = t.mainaccountcode  and a.""lastversion"" = 'Y'
                         left join static_data.""Entities"" e on a.""entityID"" = e.id
                         left join static_data.""Titles"" tt on tt.""TITLEID"" = e.""titleID""
                         where t.""transType"" = 'PREM-OUT'
                         and t.txtype2 in ('1', '2', '3', '4', '5')
                         and t.status = 'N' 
                         and p.""lastVersion"" = 'Y'
                         and t.""premin-dfrpreferno"" is not null
                         and t.""premin-rprefdate"" is not null
                         and t.""premout-dfrpreferno"" is null 
                         and t.""premout-rprefdate"" is null
                         and jupgr.""installmenttype"" = 'I' ";
            //if (!string.IsNullOrEmpty(data.AsAtDate))
            //{
            //    sql += $@"or (t.""premout-dfrpreferno"" is not null and t.""premout-rprefdate"" is not null and t.""premout-rprefdate"" <= '{data.AsAtDate}')";
            //}
            //sql += $@")";
            var sqlcount = $@"select
                         count(*) as count
                         from
                         static_data.""Transactions"" t 
                         left join static_data.""b_jupgrs"" jupgr on jupgr.""policyNo"" = t.""policyNo"" and  t.dftxno = jupgr.dftxno and jupgr.""seqNo"" = t.""seqNo"" 
                         left join static_data.""Policies"" p on jupgr.polid = p.id
                         left join static_data.""Insurers"" ier on t.""insurerCode"" = ier.""insurerCode""  and ier.""lastversion"" = 'Y'
                         left join static_data.""Entities"" e_ier on ier.""entityID"" = e_ier.id
                         left join static_data.""Titles"" tt_ier on tt_ier.""TITLEID"" = e_ier.""titleID""
                         left join static_data.""Insurees"" iee on p.""insureeCode"" = iee.""insureeCode""
                         left join static_data.""Entities"" e_iee on iee.""entityID"" = e_iee.id
                         left join static_data.""Titles"" tt_iee on tt_iee.""TITLEID"" = e_iee.""titleID""
                         left join static_data.""InsureTypes"" it on p.""insureID"" = it.id
                         left join static_data.b_jaaraps ar on (ar.billadvisorno = t.billadvisorno)
                         left join static_data.b_jaarapds ard on ard.keyidm = ar.id and ard.polid = p.id
                         left join static_data.""Agents"" a on a.""agentCode"" = t.mainaccountcode  and a.""lastversion"" = 'Y'
                         left join static_data.""Entities"" e on a.""entityID"" = e.id
                         left join static_data.""Titles"" tt on tt.""TITLEID"" = e.""titleID""
                         where t.""transType"" = 'PREM-OUT'
                         and t.txtype2 in ('1', '2', '3', '4', '5')
                         and t.status = 'N' 
                         and p.""lastVersion"" = 'Y'
                         and t.""premin-dfrpreferno"" is not null
                         and t.""premin-rprefdate"" is not null
                         and t.""premout-dfrpreferno"" is null 
                         and t.""premout-rprefdate"" is null
                         and jupgr.""installmenttype"" = 'I' ";
            sql = await GetWhereSql(data, sql, "PremOut");
            sqlcount = await GetWhereSql(data, sqlcount, "PremOut", true);
            var json = await _dataContext.PremOutCommInOvInDatas.FromSqlRaw(sql).ToListAsync();
            var count = await _dataContext.DatasCount.FromSqlRaw(sqlcount).ToListAsync();
            return new PremOutCommInOvInReportResult(json, count[0].Count);
        }

        public async Task<PremOutCommInOvInReportResult?> GetCommInOvInOpenItemReportJson(ArApReportInput data)
        {
            var sql = $@"select
                        null as ""cashierReceiveNo"",
                        null as ""CashierDate"",
                        null as ""CashierAmt"",
                        null as ""CashierReceiveType"",
                        null as ""CashierRefNo"",
                        null as ""CashierRefDate"",
                        null as ""MainAccountName"",
                        null as ""CommOutRate"",
                        null as ""CommOutAmt"",
                        null as ""OvOutRate"",
                        null as ""OvOutAmt"",
                         t.""policyNo"",
                         t.""endorseNo"",
                         t.""agentCode"",
                         jupgr.""invoiceNo"",
                         jupgr.""taxInvoiceNo"",
                         t.""transType"",
                         t.""seqNo"",
                         t.""dfrpreferno"" as ""dfRpReferNo"",
                         t.""rprefdate"" as ""rpRefDate"",
                         t.""premin-dfrpreferno"" as ""premInDfRpReferNo"",
                         t.""premin-rprefdate"" as ""premInRpRefDate"",
                         t.""premout-dfrpreferno"" as ""premOutDfRpReferNo"",
                         t.""premout-rprefdate"" as ""premOutRpRefDate"",
                         p.""actDate"",
                         p.""expDate"",
                         t.""insurerCode"",
                         case
                             when e_ier.""personType"" = 'O' then concat(tt_ier.""TITLETHAIBEGIN"", ' ', e_ier.""t_ogName"", ' ', tt_ier.""TITLETHAIEND"")
                             when e_ier.""personType"" = 'P' then concat(tt_ier.""TITLETHAIBEGIN"", ' ', e_ier.""t_firstName"", ' ', e_ier.""t_lastName"", ' ', tt_ier.""TITLETHAIEND"")
                             else null
                         end as ""insurerName"",
                         p.""insureeCode"",
                         case
                             when e_iee.""personType"" = 'O' then concat(tt_iee.""TITLETHAIBEGIN"", ' ', e_iee.""t_ogName"", ' ', tt_iee.""TITLETHAIEND"")
                             when e_iee.""personType"" = 'P' then concat(tt_iee.""TITLETHAIBEGIN"", ' ', e_iee.""t_firstName"", ' ', e_iee.""t_lastName"", ' ', tt_iee.""TITLETHAIEND"")
                             else null
                         end as ""insureeName"",
                         it.""class"",
                         it.""subClass"",
                         p.grossprem as ""grossPrem"",
                         p.specdiscrate as ""specDiscRate"",
                         p.specdiscamt as ""specDiscAmt"",
                         ((t.""subType"" *2) -1) * t.netgrossprem as ""netGrossPrem"",
                         ((t.""subType"" *2) -1) * t.duty as duty,
                         ((t.""subType"" *2) -1) * t.tax as tax,
                         ((t.""subType"" *2) -1) * t.totalprem as ""totalPrem"",
                         ((t.""subType"" *2) -1) * t.withheld as withheld,
                        ((t.""subType"" *2) -1) * t.totalamt as ""totalamt"",
                         t.netflag as ""netFlag"",
                         p.commin_rate as ""commInRate"",
                         p.commin_amt as ""commInAmt"",
                         (case when t.""transType"" = 'COMM-IN' then p.commin_rate else p.ovin_rate end ) as ""commOvInRate"",
                         (case when t.""transType"" = 'COMM-IN' then t.commamt else t.ovamt end ) as ""commOvInAmt"",
                         (select t2.dfrpreferno from static_data.""Transactions"" t2 where ""transType"" = 'COMM-IN' and t2.id = t.id) as ""commInDfRpReferNo"",
                         (select t2.rprefdate from static_data.""Transactions"" t2 where ""transType"" = 'COMM-IN' and t2.id = t.id) as ""commInRpRefDate"",
                         (select t2.paidamt from static_data.""Transactions"" t2 where ""transType"" = 'COMM-IN' and t2.id = t.id) as ""commInPaidAmt"",
                         (select t2.remainamt from static_data.""Transactions"" t2 where ""transType"" = 'COMM-IN' and t2.id = t.id) as ""commInDiffAmt"",
                         p.ovin_rate as ""ovInRate"",
                         p.ovin_amt as ""ovInAmt"",
                         (select t2.dfrpreferno from static_data.""Transactions"" t2 where ""transType"" = 'OV-IN' and t2.id = t.id) as ""ovInDfRpReferNo"",
                         (select t2.rprefdate from static_data.""Transactions"" t2 where ""transType"" = 'OV-IN' and t2.id = t.id) as ""ovInRpRefDate"",
                         (select t2.paidamt from static_data.""Transactions"" t2 where ""transType"" = 'OV-IN' and t2.id = t.id) as ""ovInPaidAmt"",
                         (select t2.remainamt from static_data.""Transactions"" t2 where ""transType"" = 'OV-IN' and t2.id = t.id) as ""ovInDiffAmt"",
                         p.""issueDate"",
                         p.createusercode as ""policyCreateUserCode"",
                         a.""contactPersonID"" as ""mainAccountContactPersonId"",
                         t.mainaccountcode as ""mainAccountCode"",
                         p.insurancestatus as ""insurancestatus"",
                         t.""transType"" as ""transactionType""
                         from
                         static_data.""Transactions"" t 
                         left join static_data.""b_jupgrs"" jupgr on jupgr.""policyNo"" = t.""policyNo"" and  t.dftxno = jupgr.dftxno and jupgr.""seqNo"" = t.""seqNo"" 
                         left join static_data.""Policies"" p on jupgr.polid = p.id
                         left join static_data.""Insurers"" ier on t.""insurerCode"" = ier.""insurerCode"" and ier.""lastversion"" = 'Y'
                         left join static_data.""Entities"" e_ier on ier.""entityID"" = e_ier.id
                         left join static_data.""Titles"" tt_ier on tt_ier.""TITLEID"" = e_ier.""titleID""
                         left join static_data.""Insurees"" iee on p.""insureeCode"" = iee.""insureeCode""
                         left join static_data.""Entities"" e_iee on iee.""entityID"" = e_iee.id
                         left join static_data.""Titles"" tt_iee on tt_iee.""TITLEID"" = e_iee.""titleID""
                         left join static_data.""InsureTypes"" it on p.""insureID"" = it.id
                         left join static_data.b_jaaraps ar on (ar.billadvisorno = t.billadvisorno)
                         left join static_data.b_jaarapds ard on ard.keyidm = ar.id and ard.polid = p.id
                         left join static_data.""Agents"" a on a.""agentCode"" = t.mainaccountcode and a.""lastversion"" = 'Y'
                         left join static_data.""Entities"" e on a.""entityID"" = e.id
                         left join static_data.""Titles"" tt on tt.""TITLEID"" = e.""titleID""
                         where t.""transType"" in ('COMM-IN', 'OV-IN')
                         and t.txtype2 in ('1', '2', '3', '4', '5')
                         and t.status = 'N' 
                         and p.""lastVersion"" = 'Y'
                         and t.""premout-dfrpreferno"" is not null
                         and t.""premout-rprefdate"" is not null 
                         and jupgr.""installmenttype"" = 'I' ";
            var sqlcount = $@"select
                         count (*) as count
                         from
                         static_data.""Transactions"" t 
                         left join static_data.""b_jupgrs"" jupgr on jupgr.""policyNo"" = t.""policyNo"" and  t.dftxno = jupgr.dftxno and jupgr.""seqNo"" = t.""seqNo"" 
                         left join static_data.""Policies"" p on jupgr.polid = p.id
                         left join static_data.""Insurers"" ier on t.""insurerCode"" = ier.""insurerCode"" and ier.""lastversion"" = 'Y'
                         left join static_data.""Entities"" e_ier on ier.""entityID"" = e_ier.id
                         left join static_data.""Titles"" tt_ier on tt_ier.""TITLEID"" = e_ier.""titleID""
                         left join static_data.""Insurees"" iee on p.""insureeCode"" = iee.""insureeCode""
                         left join static_data.""Entities"" e_iee on iee.""entityID"" = e_iee.id
                         left join static_data.""Titles"" tt_iee on tt_iee.""TITLEID"" = e_iee.""titleID""
                         left join static_data.""InsureTypes"" it on p.""insureID"" = it.id
                         left join static_data.b_jaaraps ar on (ar.billadvisorno = t.billadvisorno)
                         left join static_data.b_jaarapds ard on ard.keyidm = ar.id and ard.polid = p.id
                         left join static_data.""Agents"" a on a.""agentCode"" = t.mainaccountcode and a.""lastversion"" = 'Y'
                         left join static_data.""Entities"" e on a.""entityID"" = e.id
                         left join static_data.""Titles"" tt on tt.""TITLEID"" = e.""titleID""
                         where t.""transType"" in ('COMM-IN', 'OV-IN')
                         and t.txtype2 in ('1', '2', '3', '4', '5')
                         and t.status = 'N' 
                         and p.""lastVersion"" = 'Y'
                         and t.""premout-dfrpreferno"" is not null
                         and t.""premout-rprefdate"" is not null 
                         and jupgr.""installmenttype"" = 'I' ";
            sql = await GetWhereSql(data, sql, "CommInOvIn");
            sqlcount = await GetWhereSql(data, sqlcount, "CommInOvIn", true);
            var json = await _dataContext.PremOutCommInOvInDatas.FromSqlRaw(sql).ToListAsync();
            var count = await _dataContext.DatasCount.FromSqlRaw(sqlcount).ToListAsync();
            return new PremOutCommInOvInReportResult(json, count[0].Count);
        }

        public async Task<PremOutCommInOvInReportResult?> GetCommInOvInClearingReportJson(ArApReportInput data)
        {
            var sql = $@"select
                        null as ""cashierReceiveNo"",
                        null as ""CashierDate"",
                        null as ""CashierAmt"",
                        null as ""CashierReceiveType"",
                        null as ""CashierRefNo"",
                        null as ""CashierRefDate"",
                        null as ""MainAccountName"",
                        null as ""CommOutRate"",
                        null as ""CommOutAmt"",
                        null as ""OvOutRate"",
                        null as ""OvOutAmt"",
                         t.""policyNo"",
                         t.""endorseNo"",
                         t.""agentCode"",
                         jupgr.""invoiceNo"",
                         jupgr.""taxInvoiceNo"",
                         t.""transType"",
                         t.""seqNo"",
                         t.""dfrpreferno"" as ""dfRpReferNo"",
                         t.""rprefdate"" as ""rpRefDate"",
                         t.""premin-dfrpreferno"" as ""premInDfRpReferNo"",
                         t.""premin-rprefdate"" as ""premInRpRefDate"",
                         t.""premout-dfrpreferno"" as ""premOutDfRpReferNo"",
                         t.""premout-rprefdate"" as ""premOutRpRefDate"",
                         p.""actDate"",
                         p.""expDate"",
                         t.""insurerCode"",
                         case
                             when e_ier.""personType"" = 'O' then concat(tt_ier.""TITLETHAIBEGIN"", ' ', e_ier.""t_ogName"", ' ', tt_ier.""TITLETHAIEND"")
                             when e_ier.""personType"" = 'P' then concat(tt_ier.""TITLETHAIBEGIN"", ' ', e_ier.""t_firstName"", ' ', e_ier.""t_lastName"", ' ', tt_ier.""TITLETHAIEND"")
                             else null
                         end as ""insurerName"",
                         p.""insureeCode"",
                         case
                             when e_iee.""personType"" = 'O' then concat(tt_iee.""TITLETHAIBEGIN"", ' ', e_iee.""t_ogName"", ' ', tt_iee.""TITLETHAIEND"")
                             when e_iee.""personType"" = 'P' then concat(tt_iee.""TITLETHAIBEGIN"", ' ', e_iee.""t_firstName"", ' ', e_iee.""t_lastName"", ' ', tt_iee.""TITLETHAIEND"")
                             else null
                         end as ""insureeName"",
                         it.""class"",
                         it.""subClass"",
                         p.grossprem as ""grossPrem"",
                         p.specdiscrate as ""specDiscRate"",
                         p.specdiscamt as ""specDiscAmt"",
                          ((t.""subType"" *2) -1) * t.netgrossprem as ""netGrossPrem"",
                          ((t.""subType"" *2) -1) * t.duty as duty,
                          ((t.""subType"" *2) -1) * t.tax as tax,
                          ((t.""subType"" *2) -1) * t.totalprem as ""totalPrem"",
                          ((t.""subType"" *2) -1) * t.withheld as withheld,
                         ((t.""subType"" *2) -1) * t.totalamt as ""totalamt"",
                         t.netflag as ""netFlag"",
                         (case when t.""transType"" = 'COMM-IN' then p.commin_rate else p.ovin_rate end ) as ""commOvInRate"",
                         (case when t.""transType"" = 'COMM-IN' then t.commamt else t.ovamt end ) as ""commOvInAmt"",
                         p.commin_rate as ""commInRate"",
                         p.commin_amt as ""commInAmt"",
                         (select t2.dfrpreferno from static_data.""Transactions"" t2 where ""transType"" = 'COMM-IN' and t2.id = t.id) as ""commInDfRpReferNo"",
                         (select t2.rprefdate from static_data.""Transactions"" t2 where ""transType"" = 'COMM-IN' and t2.id = t.id) as ""commInRpRefDate"",
                         (select t2.paidamt from static_data.""Transactions"" t2 where ""transType"" = 'COMM-IN' and t2.id = t.id) as ""commInPaidAmt"",
                         (select t2.remainamt from static_data.""Transactions"" t2 where ""transType"" = 'COMM-IN' and t2.id = t.id) as ""commInDiffAmt"",
                         p.ovin_rate as ""ovInRate"",
                         p.ovin_amt as ""ovInAmt"",
                         (select t2.dfrpreferno from static_data.""Transactions"" t2 where ""transType"" = 'OV-IN' and t2.id = t.id) as ""ovInDfRpReferNo"",
                         (select t2.rprefdate from static_data.""Transactions"" t2 where ""transType"" = 'OV-IN' and t2.id = t.id) as ""ovInRpRefDate"",
                         (select t2.paidamt from static_data.""Transactions"" t2 where ""transType"" = 'OV-IN' and t2.id = t.id) as ""ovInPaidAmt"",
                         (select t2.remainamt from static_data.""Transactions"" t2 where ""transType"" = 'OV-IN' and t2.id = t.id) as ""ovInDiffAmt"",
                         p.""issueDate"",
                         p.createusercode as ""policyCreateUserCode"",
                         a.""contactPersonID"" as ""mainAccountContactPersonId"",
                         t.mainaccountcode as ""mainAccountCode"",
                         p.insurancestatus as ""insurancestatus"",
                         t.""transType"" as ""transactionType""
                         from
                         static_data.""Transactions"" t 
                         left join static_data.""b_jupgrs"" jupgr on jupgr.""policyNo"" = t.""policyNo"" and  t.dftxno = jupgr.dftxno and jupgr.""seqNo"" = t.""seqNo"" 
                         left join static_data.""Policies"" p on jupgr.polid = p.id
                         left join static_data.""Insurers"" ier on t.""insurerCode"" = ier.""insurerCode"" and ier.""lastversion"" = 'Y'
                         left join static_data.""Entities"" e_ier on ier.""entityID"" = e_ier.id
                         left join static_data.""Titles"" tt_ier on tt_ier.""TITLEID"" = e_ier.""titleID""
                         left join static_data.""Insurees"" iee on p.""insureeCode"" = iee.""insureeCode""
                         left join static_data.""Entities"" e_iee on iee.""entityID"" = e_iee.id
                         left join static_data.""Titles"" tt_iee on tt_iee.""TITLEID"" = e_iee.""titleID""
                         left join static_data.""InsureTypes"" it on p.""insureID"" = it.id
                         left join static_data.b_jaaraps ar on (ar.billadvisorno = t.billadvisorno)
                         left join static_data.b_jaarapds ard on ard.keyidm = ar.id and ard.polid = p.id
                         left join static_data.""Agents"" a on a.""agentCode"" = t.mainaccountcode   and a.""lastversion"" = 'Y'
                         left join static_data.""Entities"" e on a.""entityID"" = e.id
                         left join static_data.""Titles"" tt on tt.""TITLEID"" = e.""titleID""
                         where t.""transType"" in ('COMM-IN', 'OV-IN')
                         and t.txtype2 in ('1', '2', '3', '4', '5')
                         and t.status = 'N' 
                         and p.""lastVersion"" = 'Y'
                         and t.""premout-dfrpreferno"" is not null
                         and t.""premout-rprefdate"" is not null
                         and t.dfrpreferno is not null
                         and t.rprefdate is not null 
                         and jupgr.""installmenttype"" = 'I' ";
            var sqlcount = $@"select
                         count(*) as count
                         from
                         static_data.""Transactions"" t 
                         left join static_data.""b_jupgrs"" jupgr on jupgr.""policyNo"" = t.""policyNo"" and  t.dftxno = jupgr.dftxno and jupgr.""seqNo"" = t.""seqNo"" 
                         left join static_data.""Policies"" p on jupgr.polid = p.id
                         left join static_data.""Insurers"" ier on t.""insurerCode"" = ier.""insurerCode"" and ier.""lastversion"" = 'Y'
                         left join static_data.""Entities"" e_ier on ier.""entityID"" = e_ier.id
                         left join static_data.""Titles"" tt_ier on tt_ier.""TITLEID"" = e_ier.""titleID""
                         left join static_data.""Insurees"" iee on p.""insureeCode"" = iee.""insureeCode""
                         left join static_data.""Entities"" e_iee on iee.""entityID"" = e_iee.id
                         left join static_data.""Titles"" tt_iee on tt_iee.""TITLEID"" = e_iee.""titleID""
                         left join static_data.""InsureTypes"" it on p.""insureID"" = it.id
                         left join static_data.b_jaaraps ar on (ar.billadvisorno = t.billadvisorno)
                         left join static_data.b_jaarapds ard on ard.keyidm = ar.id and ard.polid = p.id
                         left join static_data.""Agents"" a on a.""agentCode"" = t.mainaccountcode   and a.""lastversion"" = 'Y'
                         left join static_data.""Entities"" e on a.""entityID"" = e.id
                         left join static_data.""Titles"" tt on tt.""TITLEID"" = e.""titleID""
                         where t.""transType"" in ('COMM-IN', 'OV-IN')
                         and t.txtype2 in ('1', '2', '3', '4', '5')
                         and t.status = 'N' 
                         and p.""lastVersion"" = 'Y'
                         and t.""premout-dfrpreferno"" is not null
                         and t.""premout-rprefdate"" is not null
                         and t.dfrpreferno is not null
                         and t.rprefdate is not null 
                         and jupgr.""installmenttype"" = 'I' ";
            sql = await GetWhereSql(data, sql, "CommInOvIn");
            sqlcount = await GetWhereSql(data, sqlcount, "CommInOvIn", true);
            var json = await _dataContext.PremOutCommInOvInDatas.FromSqlRaw(sql).ToListAsync();
            var count = await _dataContext.DatasCount.FromSqlRaw(sqlcount).ToListAsync();
            return new PremOutCommInOvInReportResult(json, count[0].Count);
        }

        public async Task<PremOutCommInOvInReportResult?> GetCommInOvInOutstandingReportJson(ArApReportInput data)
        {
            var sql = $@"select
                        null as ""cashierReceiveNo"",
                        null as ""CashierDate"",
                        null as ""CashierAmt"",
                        null as ""CashierReceiveType"",
                        null as ""CashierRefNo"",
                        null as ""CashierRefDate"",
                        null as ""MainAccountName"",
                        null as ""CommOutRate"",
                        null as ""CommOutAmt"",
                        null as ""OvOutRate"",
                        null as ""OvOutAmt"",
                         t.""policyNo"",
                         t.""endorseNo"",
                         t.""agentCode"",
                         jupgr.""invoiceNo"",
                         jupgr.""taxInvoiceNo"",
                         t.""transType"",
                         t.""seqNo"",
                         t.""dfrpreferno"" as ""dfRpReferNo"",
                         t.""rprefdate"" as ""rpRefDate"",
                         t.""premin-dfrpreferno"" as ""premInDfRpReferNo"",
                         t.""premin-rprefdate"" as ""premInRpRefDate"",
                         t.""premout-dfrpreferno"" as ""premOutDfRpReferNo"",
                         t.""premout-rprefdate"" as ""premOutRpRefDate"",
                         p.""actDate"",
                         p.""expDate"",
                         t.""insurerCode"",
                         case
                             when e_ier.""personType"" = 'O' then concat(tt_ier.""TITLETHAIBEGIN"", ' ', e_ier.""t_ogName"", ' ', tt_ier.""TITLETHAIEND"")
                             when e_ier.""personType"" = 'P' then concat(tt_ier.""TITLETHAIBEGIN"", ' ', e_ier.""t_firstName"", ' ', e_ier.""t_lastName"", ' ', tt_ier.""TITLETHAIEND"")
                             else null
                         end as ""insurerName"",
                         p.""insureeCode"",
                         case
                             when e_iee.""personType"" = 'O' then concat(tt_iee.""TITLETHAIBEGIN"", ' ', e_iee.""t_ogName"", ' ', tt_iee.""TITLETHAIEND"")
                             when e_iee.""personType"" = 'P' then concat(tt_iee.""TITLETHAIBEGIN"", ' ', e_iee.""t_firstName"", ' ', e_iee.""t_lastName"", ' ', tt_iee.""TITLETHAIEND"")
                             else null
                         end as ""insureeName"",
                         it.""class"",
                         it.""subClass"",
                         p.grossprem as ""grossPrem"",
                         p.specdiscrate as ""specDiscRate"",
                         p.specdiscamt as ""specDiscAmt"",
                         ((t.""subType"" *2) -1) * t.netgrossprem as ""netGrossPrem"",
                         ((t.""subType"" *2) -1) * t.duty as duty,
                         ((t.""subType"" *2) -1) * t.tax as tax,
                         ((t.""subType"" *2) -1) * t.totalprem as ""totalPrem"",
                         ((t.""subType"" *2) -1) * t.withheld as withheld,
                         ((t.""subType"" *2) -1) * t.totalamt as ""totalamt"",
                         t.netflag as ""netFlag"",
                         (case when t.""transType"" = 'COMM-IN' then p.commin_rate else p.ovin_rate end ) as ""commOvInRate"",
                         (case when t.""transType"" = 'COMM-IN' then t.commamt else t.ovamt end ) as ""commOvInAmt"",
                         p.commin_rate as ""commInRate"",
                         p.commin_amt as ""commInAmt"",
                         (select t2.dfrpreferno from static_data.""Transactions"" t2 where ""transType"" = 'COMM-IN' and t2.id = t.id) as ""commInDfRpReferNo"",
                         (select t2.rprefdate from static_data.""Transactions"" t2 where ""transType"" = 'COMM-IN' and t2.id = t.id) as ""commInRpRefDate"",
                         (select t2.paidamt from static_data.""Transactions"" t2 where ""transType"" = 'COMM-IN' and t2.id = t.id) as ""commInPaidAmt"",
                         (select t2.remainamt from static_data.""Transactions"" t2 where ""transType"" = 'COMM-IN' and t2.id = t.id) as ""commInDiffAmt"",
                         p.ovin_rate as ""ovInRate"",
                         p.ovin_amt as ""ovInAmt"",
                         (select t2.dfrpreferno from static_data.""Transactions"" t2 where ""transType"" = 'OV-IN' and t2.id = t.id) as ""ovInDfRpReferNo"",
                         (select t2.rprefdate from static_data.""Transactions"" t2 where ""transType"" = 'OV-IN' and t2.id = t.id) as ""ovInRpRefDate"",
                         (select t2.paidamt from static_data.""Transactions"" t2 where ""transType"" = 'OV-IN' and t2.id = t.id) as ""ovInPaidAmt"",
                         (select t2.remainamt from static_data.""Transactions"" t2 where ""transType"" = 'OV-IN' and t2.id = t.id) as ""ovInDiffAmt"",
                         p.""issueDate"",
                         p.createusercode as ""policyCreateUserCode"",
                         a.""contactPersonID"" as ""mainAccountContactPersonId"",
                         t.mainaccountcode as ""mainAccountCode"",
                         p.insurancestatus as ""insurancestatus"",
                         t.""transType"" as ""transactionType""
                         from
                         static_data.""Transactions"" t 
                         left join static_data.""b_jupgrs"" jupgr on jupgr.""policyNo"" = t.""policyNo"" and  t.dftxno = jupgr.dftxno and jupgr.""seqNo"" = t.""seqNo"" 
                        left join static_data.""Policies"" p on jupgr.polid = p.id
                         left join static_data.""Insurers"" ier on t.""insurerCode"" = ier.""insurerCode"" and ier.""lastversion"" = 'Y'
                         left join static_data.""Entities"" e_ier on ier.""entityID"" = e_ier.id
                         left join static_data.""Titles"" tt_ier on tt_ier.""TITLEID"" = e_ier.""titleID""
                         left join static_data.""Insurees"" iee on p.""insureeCode"" = iee.""insureeCode""
                         left join static_data.""Entities"" e_iee on iee.""entityID"" = e_iee.id
                         left join static_data.""Titles"" tt_iee on tt_iee.""TITLEID"" = e_iee.""titleID""
                         left join static_data.""InsureTypes"" it on p.""insureID"" = it.id
                         left join static_data.b_jaaraps ar on (ar.billadvisorno = t.billadvisorno)
                         left join static_data.b_jaarapds ard on ard.keyidm = ar.id and ard.polid = p.id
                         left join static_data.""Agents"" a on a.""agentCode"" = t.mainaccountcode   and a.""lastversion"" = 'Y'
                         left join static_data.""Entities"" e on a.""entityID"" = e.id
                         left join static_data.""Titles"" tt on tt.""TITLEID"" = e.""titleID""
                         where t.""transType"" in ('COMM-IN', 'OV-IN')
                         and t.txtype2 in ('1', '2', '3', '4', '5')
                         and t.status = 'N' 
                         and p.""lastVersion"" = 'Y'
                         and t.""premout-dfrpreferno"" is not null
                         and t.""premout-rprefdate"" is not null
                         and t.dfrpreferno is null 
                         and t.rprefdate is null
                         and jupgr.""installmenttype"" = 'I' ";
            //if (!string.IsNullOrEmpty(data.AsAtDate))
            //{
            //    sql += $@"or (t.dfrpreferno is not null and t.rprefdate is not null and t.rprefdate <= '{data.AsAtDate}')";
            //}
            //sql += $@")";
            var sqlcount = $@"select
                         count (*) as count
                         from
                         static_data.""Transactions"" t 
                         left join static_data.""b_jupgrs"" jupgr on jupgr.""policyNo"" = t.""policyNo"" and  t.dftxno = jupgr.dftxno and jupgr.""seqNo"" = t.""seqNo"" 
                        left join static_data.""Policies"" p on jupgr.polid = p.id
                         left join static_data.""Insurers"" ier on t.""insurerCode"" = ier.""insurerCode"" and ier.""lastversion"" = 'Y'
                         left join static_data.""Entities"" e_ier on ier.""entityID"" = e_ier.id
                         left join static_data.""Titles"" tt_ier on tt_ier.""TITLEID"" = e_ier.""titleID""
                         left join static_data.""Insurees"" iee on p.""insureeCode"" = iee.""insureeCode""
                         left join static_data.""Entities"" e_iee on iee.""entityID"" = e_iee.id
                         left join static_data.""Titles"" tt_iee on tt_iee.""TITLEID"" = e_iee.""titleID""
                         left join static_data.""InsureTypes"" it on p.""insureID"" = it.id
                         left join static_data.b_jaaraps ar on (ar.billadvisorno = t.billadvisorno)
                         left join static_data.b_jaarapds ard on ard.keyidm = ar.id and ard.polid = p.id
                         left join static_data.""Agents"" a on a.""agentCode"" = t.mainaccountcode   and a.""lastversion"" = 'Y'
                         left join static_data.""Entities"" e on a.""entityID"" = e.id
                         left join static_data.""Titles"" tt on tt.""TITLEID"" = e.""titleID""
                         where t.""transType"" in ('COMM-IN', 'OV-IN')
                         and t.txtype2 in ('1', '2', '3', '4', '5')
                         and t.status = 'N' 
                         and p.""lastVersion"" = 'Y'
                         and t.""premout-dfrpreferno"" is not null
                         and t.""premout-rprefdate"" is not null
                         and t.dfrpreferno is null 
                         and t.rprefdate is null
                         and jupgr.""installmenttype"" = 'I' ";

            sql = await GetWhereSql(data, sql, "CommInOvIn");
            sqlcount = await GetWhereSql(data, sqlcount, "CommInOvIn", true);
            var json = await _dataContext.PremOutCommInOvInDatas.FromSqlRaw(sql).ToListAsync();
            var count = await _dataContext.DatasCount.FromSqlRaw(sqlcount).ToListAsync();
            return new PremOutCommInOvInReportResult(json, count[0].Count);
        }
    }
}
