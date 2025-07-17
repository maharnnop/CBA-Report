using AspNetCore.Reporting;
using BestPolicyReport.Data;
using BestPolicyReport.Models.ArApReport;
using BestPolicyReport.Models.BillReport;
using Microsoft.AspNetCore.Mvc;
using Microsoft.EntityFrameworkCore;
using System.Reflection;
using static Microsoft.EntityFrameworkCore.DbLoggerCategory;

namespace BestPolicyReport.Services.BillService
{
    public class BillService : IBillService
    {
        private readonly DataContext _dataContext;

        public BillService(DataContext dataContext)
        {
            _dataContext = dataContext;
        }
        //ใบแจ้งหนี้
        public async Task<BillReportResult?> GetBillReportJson(BillReportInput data)
        {
            var subString = "BILL";
            // mai's query 
            var sql = $@"(select p.""policyNo"",   t.""endorseNo"" , p.""issueDate"", p.""actDate"", p.""expDate"",
                         jupgr.""invoiceNo"", jupgr.""taxInvoiceNo"", p.""invoiceName"",
                         p.""insurerCode"", it.class, it.""subClass"",  p.""agentCode"" as ""agentCode1"", p.""agentCode2"", bjd.seqno as ""seqNo"", i.""insureeCode"",
                         case 
                         	when e_i.""personType"" = 'O' then concat(t_i.""TITLETHAIBEGIN"", ' ', e_i.""t_ogName"", ' ', t_i.""TITLETHAIEND"") 
                         	when e_i.""personType"" = 'P' then concat(t_i.""TITLETHAIBEGIN"", ' ', e_i.""t_firstName"", ' ', e_i.""t_lastName"", ' ', t_i.""TITLETHAIEND"") 
                         	else ''
                         end	as ""insureeName"",
                         -- case 
                         	-- when p.""itemList"" is not null then m.""licenseNo""
                         	-- else null
                         -- end	as ""licenseNo"",
                         -- case 
                         	-- when p.""itemList"" is not null then pv.t_provincename
                         	-- else null
                         -- end	as province,
                         -- case 
                         	-- when p.""itemList"" is not null then m.""chassisNo""
                         	-- else null
                         -- end	as ""chassisNo"",
                         m.""voluntaryCode"", m.""compulsoryCode"" , m.""unregisterflag"",
                         m.""licenseNo"" as ""licenseNo"", pv.t_provincename as province, m.""chassisNo"" as ""chassisNo"",
                         m.""addition_access"" as ""addition_access"", m.""engineNo"", m.""modelYear"",
                         m.""brand"", m.""model"", m.""specname"", m.cc, m.seat, m.gvw, 
                         t.""dueDate"",
                         bjd.grossprem as ""grossPrem"", p.specdiscrate as ""specDiscRate"", jupgr.specdiscamt as ""specDiscAmt"", jupgr.netgrossprem as ""netGrossPrem"", 
                         jupgr.duty, jupgr.tax, jupgr.totalprem as ""totalPrem"", p.commout1_rate as ""commOutRate1"", jupgr.commout1_amt as ""commOutAmt1"",  
                         p.ovout1_rate as ""ovOutRate1"", jupgr.ovout1_amt as ""ovOutAmt1"", p.commout2_rate as ""commOutRate2"", jupgr.commout2_amt as ""commOutAmt2"", 
                         p.ovout2_rate as ""ovOutRate2"", jupgr.ovout2_amt as ""ovOutAmt2"", bjd.""comm-out%"" as ""commOutRate"", bjd.""comm-out-amt"" as ""commOutAmt"", 
                         bjd.""ov-out%"" as ""ovOutRate"", bjd.""ov-out-amt"" as ""ovOutAmt"", bjd.netflag as ""netFlag"", bjd.billpremium as ""billPremium"",
                         bj.billadvisorno as ""billAdvisorNo"",
                         bj.billdate as ""billDate"", p.""beneficiary""
                         from static_data.b_jabilladvisors bj
                         join static_data.b_jabilladvisordetails bjd on  bj.id = bjd.keyidm 
                        join static_data.""b_jupgrs"" jupgr on bjd.""policyNo""  = jupgr.""policyNo"" and  bjd.dftxno = jupgr.dftxno   and bjd.""seqno"" = jupgr.""seqNo"" 
                        join static_data.""Transactions"" t on t.""policyNo"" = jupgr.""policyNo""  and t.dftxno = jupgr.dftxno and t.""seqNo"" = jupgr.""seqNo"" 
                        join  static_data.""Policies"" p on p.id = jupgr.polid and p.""lastVersion"" = 'Y'
                        left join static_data.""Motors"" m on m.id = p.""itemList""
                        left join static_data.provinces pv on m.""motorprovinceID"" = pv.provinceid
                        left join static_data.""Insurees"" i on i.""insureeCode"" = p.""insureeCode"" and i.lastversion = 'Y'
                        left join static_data.""Entities"" e_i on i.""entityID"" = e_i.id
                        left join static_data.""Titles"" t_i on e_i.""titleID"" = t_i.""TITLEID""
                        left join static_data.""InsureTypes"" it on it.id = p.""insureID"" 
                         where
                        bj.billadvisorno = t.billadvisorno 
                        and t.""transType"" = 'PREM-IN'
                        and t.""status"" = 'N'
                        and bj.""active"" = 'Y'
                        and jupgr.installmenttype ='A'
                        -- and bjd.polid = p.id 
                         and bjd.customerid = i.id 
                         and t.billadvisorno is not null 
                        order by  bj.billadvisorno, p.""policyNo"",  t.""endorseNo"" ) as query where true ";

            if (!string.IsNullOrEmpty(data.InsurerCode))
            {
                sql += $@"and ""insurerCode"" = '{data.InsurerCode}' ";
            }
            if (!string.IsNullOrEmpty(data.AgentCode1))
            {
                sql += $@"and ""agentCode1"" = '{data.AgentCode1}' ";
            }
            if (!string.IsNullOrEmpty(data.AgentCode2))
            {
                sql += $@"and ""agentCode2"" = '{data.AgentCode2}' ";
            }
            if (!string.IsNullOrEmpty(data.StartBillAdvisorNo) && !string.IsNullOrEmpty(data.EndBillAdvisorNo))
            {
                sql += $@"and ""billAdvisorNo"" between '{data.StartBillAdvisorNo}' and '{data.EndBillAdvisorNo}' ";
            }
            string currentDate = DateTime.Now.ToString("yyyy-MM-dd", new System.Globalization.CultureInfo("en-US"));
            if (!string.IsNullOrEmpty(data.StartBillDate?.ToString()))
            {
                if (!string.IsNullOrEmpty(data.EndBillDate?.ToString()))
                {
                    sql += $@"and ""billDate"" between '{data.StartBillDate}' and '{data.EndBillDate}' ";
                }
                else
                {
                    sql += $@"and ""billDate"" between '{data.StartBillDate}' and '{currentDate}' ";
                }
            }
            //sql += $@";";
            var sql1 = "select * from " + sql;

            if ((data.Limit).HasValue && (data.Pagecount).HasValue)
            {
                sql1 += $@" LIMIT {data.Limit}  OFFSET ({data.Pagecount} - 1) * {data.Limit} ;";
            }
            else { sql1 += " ;"; }

            var sql2 = "select count(*) as count from " + sql + " ;";

            var json = await _dataContext.BillDatas.FromSqlRaw(sql1).ToListAsync();
            var count = await _dataContext.DatasCount.FromSqlRaw(sql2).ToListAsync();
            return new BillReportResult(json, count[0].Count);
        }

        public async Task<List<PolicyGroupBillReportResult>?> GetPolicyGroupBillReportJson(PolicyGroupBillReportInput data)
        {
            var sql = $@"select * from (
                         select 
                         case 
                         	when p.""itemList"" is not null then m.""licenseNo""
                         	else null
                         end	
                         as ""licenseNo"",
                         case 
                         	when p.""itemList"" is not null then m.brand
                         	else null
                         end	
                         as ""brand"",
                         case 
                         	when p.""itemList"" is not null then m.model
                         	else null
                         end	
                         as ""model"",
                         case 
                         	when e_i.""personType"" = 'O' then concat(t_i.""TITLETHAIBEGIN"", ' ', e_i.""t_ogName"", ' ', t_i.""TITLETHAIEND"") 
                         	when e_i.""personType"" = 'P' then concat(t_i.""TITLETHAIBEGIN"", ' ', e_i.""t_firstName"", ' ', e_i.""t_lastName"", ' ', t_i.""TITLETHAIEND"") 
                         	else null
                         end	
                         as ""insureeName"",
                         case 
                         	when p.""itemList"" is not null then m.""modelYear"" 
                         	else null
                         end	
                         as ""modelYear"",
                         case 
                         	when p.""itemList"" is not null then m.""chassisNo""
                         	else null
                         end	
                         as ""chassisNo"",
                         p.cover_amt as ""coverAmt"",
                         p.""policyNo"",
                         p.""actDate"",
                         p.""expDate"",
                         sum(p.netgrossprem) as ""netGrossPrem"",
                         sum(bad.duty) as ""duty"",
                         sum(p.netgrossprem) + sum(bad.duty) as ""netGrossPremBeforeTax"",
                         sum(bad.tax) as ""tax"",
                         sum(bad.totalprem) as ""totalPrem"",
                         sum(bad.withheld) as ""withHeld"",
                         sum(bad.billpremium) as ""billPremium"",
                         bad.netflag as ""netFlag"",
                         ba.billadvisorno as ""billAdvisorNo"",
                         ba.createusercode as ""createUserCode"",
                         ba.billdate as ""billDate""
                         from static_data.b_jabilladvisors ba
                         left join static_data.b_jabilladvisordetails bad on ba.id = bad.keyidm  
                         left join static_data.""Policies"" p on bad.polid = p.id
                         left join static_data.""Insurees"" i on p.""insureeCode"" = i.""insureeCode""
                         left join static_data.""Entities"" e_i on i.""entityID"" = e_i.id
                         left join static_data.""Titles"" t_i on e_i.""titleID"" = t_i.""TITLEID""
                         left join static_data.""Motors"" m on m.id = p.""itemList""
                         where ba.active = 'Y'
                         group by p.""policyNo"", 
                         p.""itemList"", 
                         m.""licenseNo"", 
                         m.brand, 
                         m.model, 
                         m.""modelYear"", 
                         e_i.""personType"", 
                         t_i.""TITLETHAIBEGIN"", 
                         e_i.""t_ogName"", 
                         t_i.""TITLETHAIEND"", 
                         e_i.""t_firstName"", 
                         e_i.""t_lastName"", 
                         m.""chassisNo"",
                         p.cover_amt,
                         p.""policyNo"",
                         p.""actDate"",
                         p.""expDate"",
                         bad.netflag,
                         ba.billadvisorno,
                         ba.createusercode,
                         ba.billdate
                         order by p.""policyNo"" asc) as query
                         where true ";
            if (!string.IsNullOrEmpty(data.ListBillAdvisorNo))
            {
                sql += $@"and ""billAdvisorNo"" = '{data.ListBillAdvisorNo}' ";
            }
            if (!string.IsNullOrEmpty(data.CreateUserCode))
            {
                sql += $@"and ""createUserCode"" = '{data.CreateUserCode}' ";
            }
            string currentDate = DateTime.Now.ToString("yyyy-MM-dd", new System.Globalization.CultureInfo("en-US"));
            if (!string.IsNullOrEmpty(data.StartBillDate?.ToString()))
            {
                if (!string.IsNullOrEmpty(data.EndBillDate?.ToString()))
                {
                    sql += $@"and ""billDate"" between '{data.StartBillDate}' and '{data.EndBillDate}' ";
                }
                else
                {
                    sql += $@"and ""billDate"" between '{data.StartBillDate}' and '{currentDate}' ";
                }
            }
            sql += $@"order by ""insureeName"";";
            var json = await _dataContext.PolicyGroupBillReportResults.FromSqlRaw(sql).ToListAsync();
            return json;
        }

        public async Task<InvoiceReportResult>? GetInvoiceJson(InvoiceReportInput data )
        {
            var sql = $@"select p.""agentCode"", p.""insurerCode"", p.""applicationNo"",
       t.""policyNo"", t.""endorseNo"", j.""invoiceNo"", t.""seqNo"" ,
        t.dfrpreferno ,t.rprefdate ,
      (select ""id"" from static_data.""Insurees"" where ""insureeCode"" = p.""insureeCode"" and lastversion = 'Y') as customerid, 
      p.""insureeCode"",
     getname(a.""entityID"" ) as agentname , getname(i.""entityID"" ) as insureename , 
     -- (case when e.""personType"" ='P' then  t2.""TITLETHAIBEGIN"" || ' ' || e.""t_firstName""||' '||e.""t_lastName"" else 
     --   t2.""TITLETHAIBEGIN""|| ' '|| e.""t_ogName""|| COALESCE(' สาขา '|| e.""t_branchName"",'' )  || ' '|| t2.""TITLETHAIEND"" end) as agentname ,
     -- (case when e_i.""personType"" ='P' then  t2_i.""TITLETHAIBEGIN"" || ' ' || e_i.""t_firstName""||' '||e_i.""t_lastName"" else 
     --   t2_i.""TITLETHAIBEGIN""|| ' '|| e_i.""t_ogName""|| COALESCE(' สาขา '|| e_i.""t_branchName"",'' )  || ' '|| t2_i.""TITLETHAIEND"" end) as insureename ,
      j.polid, 
      j.grossprem, j.specdiscrate, j.specdiscamt, j.netgrossprem, j.duty, j.tax, j.totalprem, j.withheld, 
      (j.totalprem - j.specdiscamt - j.withheld) as ""totalamt"",
      j.commout_rate,j.commout_amt, j.ovout_rate, j.ovout_amt,
      j.commout1_rate,j.commout1_amt, j.ovout1_rate, j.ovout1_amt, 
      t.""dueDate""
      from  static_data.b_jupgrs j 
      join static_data.""Policies"" p on p.id = j.polid
      join static_data.""Transactions"" t on j.""policyNo"" = t.""policyNo"" and  t.dftxno = j.dftxno and t.""seqNo"" = j.""seqNo"" and t.""transType"" = 'PREM-IN'
      left join static_data.""Agents"" a on a.""agentCode"" = p.""agentCode"" and a.lastversion ='Y'
      left join static_data.""Insurees"" i on i.""insureeCode"" =p.""insureeCode""  and i.lastversion  ='Y'
      -- left join static_data.""Entities"" e on e.id = a.""entityID"" 
      -- left join static_data.""Entities"" e_i on e_i.id = i.""entityID"" 
      -- left join static_data.""Titles"" t2 on t2.""TITLEID"" = e.""titleID""  
      -- left join static_data.""Titles"" t2_i on t2_i.""TITLEID"" = e_i.""titleID"" 
      left join static_data.""Motors"" m on m.id = p.""itemList""  -- only งานรายย่อย
      where j.installmenttype = 'A'
      and t.""status"" = 'N'
      and p.""lastVersion"" = 'Y'";

            var sqlcount = $@"select count(*) as count
      from  static_data.b_jupgrs j 
      join static_data.""Policies"" p on p.id = j.polid
      join static_data.""Transactions"" t on j.""policyNo"" = t.""policyNo"" and  t.dftxno = j.dftxno and t.""seqNo"" = j.""seqNo"" and t.""transType"" = 'PREM-IN'
      left join static_data.""Agents"" a on a.""agentCode"" = p.""agentCode"" and a.lastversion ='Y'
      left join static_data.""Insurees"" i on i.""insureeCode"" =p.""insureeCode""  and i.lastversion  ='Y'
      left join static_data.""Entities"" e on e.id = a.""entityID"" 
      left join static_data.""Entities"" e_i on e_i.id = i.""entityID""  
      left join static_data.""Titles"" t2 on t2.""TITLEID"" = e.""titleID""
      left join static_data.""Titles"" t2_i on t2_i.""TITLEID"" = e_i.""titleID"" 
        left join static_data.""Motors"" m on m.id = p.""itemList""   -- only งานรายย่อย
      where j.installmenttype = 'A'
      and t.""status"" = 'N'
      and p.""lastVersion"" = 'Y'";

            if (!string.IsNullOrEmpty(data.InsurerCode))
            {
                sql += $@"and p.""insurerCode"" = '{data.InsurerCode}' ";
                sqlcount += $@"and p.""insurerCode"" = '{data.InsurerCode}' ";
            }
            if (!string.IsNullOrEmpty(data.AgentCode))
            {
                sql += $@"and p.""agentCode"" = '{data.AgentCode}' ";
                sqlcount += $@"and p.""agentCode"" = '{data.AgentCode}' ";
            }
            if (!string.IsNullOrEmpty(data.StartInvoiceNo))
            {
                sql += $@"and j.""invoiceNo"" >=  '{data.StartInvoiceNo}' ";
                sqlcount += $@"and j.""invoiceNo"" >=  '{data.StartInvoiceNo}' ";
            }
            if (!string.IsNullOrEmpty(data.EndInvoiceNo))
            {
                sql += $@"and j.""invoiceNo"" <=  '{data.EndInvoiceNo}' ";
                sqlcount += $@"and j.""invoiceNo"" <=  '{data.EndInvoiceNo}' ";
            }

            string currentDate = DateTime.Now.ToString("yyyy-MM-dd", new System.Globalization.CultureInfo("en-US"));
            if (!string.IsNullOrEmpty(data.CreatedDateStart))
            {
                if (!string.IsNullOrEmpty(data.CreatedDateEnd))
                {
                    sql += $@"and date(p.""createdAt"") between '{data.CreatedDateStart}' and '{data.CreatedDateEnd}' ";
                    sqlcount += $@"and date(p.""createdAt"") between '{data.CreatedDateStart}' and '{data.CreatedDateEnd}' ";
                }
                else
                {
                    sql += $@"and date(p.""createdAt"") between '{data.CreatedDateStart}' and '{currentDate}' ";
                    sqlcount += $@"and date(p.""createdAt"") between '{data.CreatedDateStart}' and '{currentDate}' ";
                }
            }
            if (!string.IsNullOrEmpty(data.Status))
            {
                if (data.Status == "I")
                {
                    sql += $@"and t.dfrpreferno is  null ";
                    sqlcount += $@"and t.dfrpreferno is  null ";
                }
                else if (data.Status == "A")
                {
                    sql += $@"and t.dfrpreferno is not null ";
                    sqlcount += $@"and t.dfrpreferno is not null ";
                }
                
            }
            if (!string.IsNullOrEmpty(data.DueDateStatus))
            {
                if (data.DueDateStatus == "IN")
                {
                    sql += $@" and t.""dueDate"" >= CURRENT_DATE ";
                    sqlcount += $@" and t.""dueDate"" >= CURRENT_DATE ";
                }
                else if (data.DueDateStatus == "OV")
                {
                    sql += $@" and t.""dueDate"" < CURRENT_DATE ";
                    sqlcount += $@" and t.""dueDate"" < CURRENT_DATE ";
                }
            }

            // add 09-04-2025 find by endorseNo, insureeName, licenseNo, policyNo
            if (!string.IsNullOrEmpty(data.EndorseNo))
            {
                sql += $@"and t.""endorseNo"" =  '{data.EndorseNo}' ";
                sqlcount += $@"and t.""endorseNo"" =  '{data.EndorseNo}' ";
            }
            if (!string.IsNullOrEmpty(data.PolicyNo))
            {
                sql += $@"and t.""policyNo"" =  '{data.PolicyNo}' ";
                sqlcount += $@"and t.""policyNo"" =  '{data.PolicyNo}' ";
            }
            if (!string.IsNullOrEmpty(data.LicenseNo)) // only MO งานรายย่อย
            {
                sql += $@"and m.""licenseNo"" =  '{data.LicenseNo}' ";
                sqlcount += $@"and m.""licenseNo"" =  '{data.LicenseNo}' ";
            }
            if (!string.IsNullOrEmpty(data.PersonType)) 
            {
                if (data.PersonType == "P")
                {
                    if (!string.IsNullOrEmpty(data.Insuree_fn))
                    {
                        sql += $@"and e_i.""t_firstName"" =  '{data.Insuree_fn}' ";
                        sqlcount += $@"and e_i.""t_firstName"" =  '{data.Insuree_fn}' ";
                    }
                    if (!string.IsNullOrEmpty(data.Insuree_ln))
                    {
                        sql += $@"and e_i.""t_lastName"" =  '{data.Insuree_ln}' ";
                        sqlcount += $@"and e_i.""t_lastName"" =  '{data.Insuree_ln}' ";
                    }
                }else if (data.PersonType == "O")
                {
                    if (!string.IsNullOrEmpty(data.Insuree_ogn))
                    {
                        sql += $@"and e_i.""t_ogName"" =  '{data.Insuree_ogn}' ";
                        sqlcount += $@"and e_i.""t_ogName"" =  '{data.Insuree_ogn}' ";
                    }
                }

            }


            sql += $@" order by j.""invoiceNo"", j.""seqNo"" ";
            if ((data.Limit).HasValue && (data.Pagecount).HasValue)
            {
                sql += $@"LIMIT {data.Limit}  OFFSET ({data.Pagecount} - 1) * {data.Limit} ;";
            }
            else { sql += $@" ;"; }

            var json = await _dataContext.InvoiceDatas.FromSqlRaw(sql).ToListAsync();
            var count = await _dataContext.DatasCount.FromSqlRaw(sqlcount).ToListAsync();
            return new InvoiceReportResult(json, count[0].Count);
        }

        public async Task<InvoiceRpt>? GetInvoicerpt(string invoiceNo, string usercode)
        {
            var sql = $@"select ju.id as ""Id"" , ju.""invoiceNo"" ,t.""dueDate"" ,ju.""seqNo"",
            (case when e.""personType"" ='P' then  t2.""TITLETHAIBEGIN"" || ' ' || e.""t_firstName""||' '||e.""t_lastName"" else 
                    t2.""TITLETHAIBEGIN""|| ' '|| e.""t_ogName""|| COALESCE(' สาขา '|| e.""t_branchName"",'' ) || ' '|| t2.""TITLETHAIEND"" end) as ""insureeName"" ,
             (l.t_location_1||' '||l.t_location_2||' หมู่ '||l.t_location_3||' ซอย '||l.t_location_4||' ถนน '||l.t_location_5||' ต.'||t3.t_tambonname||' อ.'||a.t_amphurname||' จ.'||p2.t_provincename||' '||l.zipcode) as ""insureeLocation"",
             p.""insureeCode"" ,
             (case when e_ins.""personType"" ='P' then  tt_ins.""TITLETHAIBEGIN"" || ' ' || e_ins.""t_firstName""||' '||e_ins.""t_lastName"" else 
                    tt_ins.""TITLETHAIBEGIN""|| ' '|| e_ins.""t_ogName""|| COALESCE(' สาขา '|| e_ins.""t_branchName"",'' ) || ' '|| tt_ins.""TITLETHAIEND"" end) as ""insurerName"" ,
             (select it.""insureName""  from static_data.""InsureTypes"" it where it.id = p.""insureID"") as ""insureName"",
             p.""policyNo"" ,p3.""endorseNo"" ,p3.cover_amt ,p.""actDate"" ,p.""expDate"" ,
            ( case when epm.id  is null or edt2.edtypecode like 'MT%' then p3.netgrossprem else epm.diffnetgrossprem  end ) as netgrossprem,
            ( case when epm.id  is null or edt2.edtypecode like 'MT%' then p3.duty else epm.diffduty  end ) as duty,
            ( case when epm.id  is null or edt2.edtypecode like 'MT%' then p3.tax else epm.difftax  end ) as tax,
            ( case when epm.id  is null or edt2.edtypecode like 'MT%' then p3.totalprem else epm.difftotalprem  end ) as totalprem,
            ( case when epm.id  is null or edt2.edtypecode like 'MT%' then p3.specdiscamt else epm.discinamt  end ) as specdiscamt,
            ( case when epm.id  is null or edt2.edtypecode like 'MT%' then (p3.totalprem - p3.specdiscamt) else (epm.difftotalprem - epm.discinamt)  end ) as totalamt,
              (ju.totalprem - ju.specdiscamt) as seqamt
             from  static_data.b_jupgrs ju  
            join static_data.""Policies"" p on ju.polid =p.id and p.""lastVersion"" ='Y'
            --and ju.""endorseNo"" =p.""endorseNo""
            left join static_data.""Transactions"" t on ju.""policyNo"" = t.""policyNo"" and t.dftxno = ju.dftxno and ju.""seqNo"" =t.""seqNo""
            --and ju.""endorseNo"" =t.""endorseNo""  
            left join static_data.""Insurees"" i on p.""insureeCode""  = i.""insureeCode"" and i.lastversion = 'Y'
            left join static_data.""Entities"" e on e.id = i.""entityID"" 
            left join static_data.""Titles"" t2 on t2.""TITLEID"" =e.""titleID"" 
            left join static_data.""Insurers""  ins on p.""insurerCode""  = ins.""insurerCode""  and ins.lastversion  ='Y'
            left join static_data.""Entities"" e_ins on e_ins.id = ins.""entityID"" 
            left join static_data.""Titles"" tt_ins on tt_ins.""TITLEID"" =e_ins.""titleID"" 
            left join static_data.""Locations"" l on l.""entityID"" =e.id and l.lastversion = 'Y'
            join static_data.provinces p2 on p2.provinceid =l.""provinceID"" 
            join  static_data.""Amphurs"" a on a.amphurid =l.""districtID"" 
            join static_data.""Tambons"" t3 on t3.tambonid =l.""subDistrictID"" 
            left join static_data.b_juepms epm on epm.polid = t.polid 
            left join static_data.b_juedts edt2 on edt2.polid= t.polid 
            left join static_data.""Policies"" p3 on p3.id = t.polid 
            where ju.""invoiceNo"" = '{invoiceNo}'
            and t.status ='N'
            and t.""transType"" ='PREM-IN' ;";

         

            string currentDate = DateTime.Now.ToString("yyyy-MM-dd", new System.Globalization.CultureInfo("en-US"));
           

            

            var json = await _dataContext.InvoiceRptDatas.FromSqlRaw(sql).ToListAsync();
            //string sql_update = $@"update static_data.b_jupgrs
            //                      set
            //                      lastprintuser =  @usercode,
            //                      lastprintdate = now()
            //                      where ""invoiceNo"" = @invoiceNo;";
            //int rowsAffected = await _dataContext.Database.ExecuteSqlRawAsync(sql_update,
            //    new Npgsql.NpgsqlParameter("usercode", usercode), // Example for PostgreSQL (Npgsql)
            //    new Npgsql.NpgsqlParameter("currentDate", currentDate),
            //    new Npgsql.NpgsqlParameter("invoiceNo", invoiceNo));



            return json[0];
        }
    }
}
