using BestPolicyReport.Data;
using BestPolicyReport.Models.BillReport;
using BestPolicyReport.Models.DailyPolicyReport;
using Microsoft.EntityFrameworkCore;

namespace BestPolicyReport.Services.DailyPolicyService
{
    public class DailyPolicyService : IDailyPolicyServiceService
    {
        private readonly DataContext _dataContext;

        public DailyPolicyService(DataContext dataContext)
        {
            _dataContext = dataContext;
        }

        public async Task<DailyPolicyReportResult?> GetDailyPolicyReportJson(DailyPolicyReportInput data)
        {

            var sql = $@"from static_data.""Policies"" p
                         left join static_data.""Users"" u on u.""empCode"" = p.createusercode
                         left join static_data.""Agents"" a1 on (a1.""agentCode"" = p.""agentCode"" and a1.lastversion = 'Y')
                         left join static_data.""Entities"" e_a1 on e_a1.id = a1.""entityID""
                         left join static_data.""Titles"" t_a1 on e_a1.""titleID"" = t_a1.""TITLEID""
                         left join static_data.""Entities"" e_cp1 on e_cp1.id = a1.""contactPersonID""
                         left join static_data.""Titles"" t_cp1 on e_cp1.""titleID"" = t_cp1.""TITLEID""
                         left join static_data.""Agents"" a2 on (a2.""agentCode"" = p.""agentCode2"" and a2.lastversion = 'Y')
                         left join static_data.""Entities"" e_a2 on e_a2.id = a2.""entityID""
                         left join static_data.""Titles"" t_a2 on e_a2.""titleID"" = t_a2.""TITLEID""
                         left join static_data.""Entities"" e_cp2 on e_cp2.id = a2.""contactPersonID""
                         left join static_data.""Titles"" t_cp2 on e_cp2.""titleID"" = t_cp2.""TITLEID""
                         left join static_data.""Insurees"" i on p.""insureeCode"" = i.""insureeCode"" and i.lastversion  ='Y'
                         left join static_data.""Entities"" e_i on i.""entityID"" = e_i.id
                         left join static_data.""Titles"" t_i on e_i.""titleID"" = t_i.""TITLEID""
                         left join static_data.""InsureTypes"" it on it.id = p.""insureID""
                         left join static_data.""Motors"" m on m.id = p.""itemList""
                         left join static_data.provinces pv on m.""motorprovinceID"" = pv.provinceid
                         where  p.""endorseNo"" is null ";
            string currentDate = DateTime.Now.ToString("yyyy-MM-dd", new System.Globalization.CultureInfo("en-US"));
            if (!string.IsNullOrEmpty(data.StartPolicyDate?.ToString()))
            {
                if (!string.IsNullOrEmpty(data.EndPolicyDate?.ToString()))
                {
                    sql += $@"and p.""policyDate"" between '{data.StartPolicyDate}' and '{data.EndPolicyDate}' ";
                }
                else
                {
                    sql += $@"and p.""policyDate"" between '{data.StartPolicyDate}' and '{currentDate}' ";
                }
            }
            if (!string.IsNullOrEmpty(data.CreateUserCode?.ToString()))
            {
                sql += $@"and p.createusercode = '{data.CreateUserCode}' ";
            }
            if (!string.IsNullOrEmpty(data.ContactPersonId1?.ToString()))
            {
                sql += $@"and a1.""contactPersonID"" = {data.ContactPersonId1} ";
            }
            if (!string.IsNullOrEmpty(data.ContactPersonId2?.ToString()))
            {
                sql += $@"and a2.""contactPersonID"" = {data.ContactPersonId2} ";
            }
            if (!string.IsNullOrEmpty(data.AgentCode1?.ToString()))
            {
                sql += $@"and p.""agentCode"" = '{data.AgentCode1}' ";
            }
            if (!string.IsNullOrEmpty(data.AgentCode2?.ToString()))
            {
                sql += $@"and p.""agentCode2"" = '{data.AgentCode2}' ";
            }
            if (!string.IsNullOrEmpty(data.InsurerCode?.ToString()))
            {
                sql += $@"and p.""insurerCode"" = '{data.InsurerCode}' ";
            }
            if (!string.IsNullOrEmpty(data.Insurancestatus?.ToString()))
            {
                sql += $@"and p.insurancestatus = '{data.Insurancestatus}' ";

               
            }else
            {
                sql += $@"and p.insurancestatus != 'CC' ";
            }
            if (!string.IsNullOrEmpty(data.Class?.ToString()))
            {
                sql += $@"and it.""class"" = '{data.Class}' ";
            }
            if (!string.IsNullOrEmpty(data.SubClass?.ToString()))
            {
                sql += $@"and it.""subClass"" = '{data.SubClass}' ";
            }
            if (!string.IsNullOrEmpty(data.OrderBy?.ToString()))
            {
                if (data.OrderBy.ToString() == "createusercode")
                {
                    sql += $@"order by p.createusercode asc";
                }
                else if (data.OrderBy.ToString() == "employeecode")
                {
                    sql += $@"order by a1.""contactPersonID"" asc, a2.""contactPersonID"" asc";
                }
                else if (data.OrderBy.ToString() == "advisorcode")
                {
                    sql += $@"order by p.""agentCode"" asc, p.""agentCode2"" asc";
                }
                else if (data.OrderBy.ToString() == "insurercode")
                {
                    sql += $@"order by p.""insurerCode"" asc ";
                }
            }

            string sql1 = @"select p.""applicationNo"", 
                         p.""policyNo"",
                         p.""policyDate"", 
                         p.""actDate"", 
                         p.""expDate"", 
                         p.""issueDate"", 
                         p.createusercode as ""CreateUserCode"", 
                         u.""userName"" as ""Username"",
                         a1.""contactPersonID"" as ""contactPersonId1"",
                         case
                         	when e_cp1.""personType"" = 'O' then concat(t_cp1.""TITLETHAIBEGIN"", ' ', e_cp1.""t_ogName"", ' ', t_cp1.""TITLETHAIEND"") 
                         	when e_cp1.""personType"" = 'P' then concat(t_cp1.""TITLETHAIBEGIN"", ' ', e_cp1.""t_firstName"", ' ', e_cp1.""t_lastName"", ' ', t_cp1.""TITLETHAIEND"") 
                         	else null
                         end
                         as ""contactPersonName1"",
                         a2.""contactPersonID"" as ""contactPersonId2"",
                         case 
                         	when e_cp2.""personType"" = 'O' then concat(t_cp2.""TITLETHAIBEGIN"", ' ', e_cp2.""t_ogName"", ' ', t_cp2.""TITLETHAIEND"") 
                         	when e_cp2.""personType"" = 'P' then concat(t_cp2.""TITLETHAIBEGIN"", ' ', e_cp2.""t_firstName"", ' ', e_cp2.""t_lastName"", ' ', t_cp2.""TITLETHAIEND"") 
                         	else null
                         end	
                         as ""contactPersonName2"",
                         p.""agentCode"" as ""agentCode1"", 
                         case 
                         	when e_a1.""personType"" = 'O' then concat(t_a1.""TITLETHAIBEGIN"", ' ', e_a1.""t_ogName"", ' ', t_a1.""TITLETHAIEND"") 
                         	when e_a1.""personType"" = 'P' then concat(t_a1.""TITLETHAIBEGIN"", ' ', e_a1.""t_firstName"", ' ', e_a1.""t_lastName"", ' ', t_a1.""TITLETHAIEND"") 
                         	else null
                         end	
                         as ""agentName1"",
                         p.""agentCode2"", 
                         case 
                         	when e_a2.""personType"" = 'O' then concat(t_a2.""TITLETHAIBEGIN"", ' ', e_a2.""t_ogName"", ' ', t_a2.""TITLETHAIEND"") 
                         	when e_a2.""personType"" = 'P' then concat(t_a2.""TITLETHAIBEGIN"", ' ', e_a2.""t_firstName"", ' ', e_a2.""t_lastName"", ' ', t_a2.""TITLETHAIEND"") 
                         	else null
                         end	
                         as ""agentName2"",
                         p.""insureeCode"",
                         case 
                         	when e_i.""personType"" = 'O' then concat(t_i.""TITLETHAIBEGIN"", ' ', e_i.""t_ogName"", ' ', t_i.""TITLETHAIEND"") 
                         	when e_i.""personType"" = 'P' then concat(t_i.""TITLETHAIBEGIN"", ' ', e_i.""t_firstName"", ' ', e_i.""t_lastName"", ' ', t_i.""TITLETHAIEND"") 
                         	else null
                         end	
                         as ""insureeName"",
                         it.""class"",
                         it.""subClass"",
                         case 
                         	when p.""itemList"" is not null then m.""licenseNo""
                         	else null
                         end	
                         as ""licenseNo"",
                         case 
                         	when p.""itemList"" is not null then pv.t_provincename
                         	else null
                         end	
                         as province,
                         case 
                         	when p.""itemList"" is not null then m.""chassisNo""
                         	else null
                         end	
                         as ""chassisNo"",
                         p.grossprem as ""grossPrem"", 
                         p.specdiscrate as ""specDiscRate"", 
                         p.specdiscamt as ""specDiscAmt"", 
                         p.netgrossprem as ""netGrossPrem"", 
                         p.duty, 
                         p.tax, 
                         p.totalprem as ""totalPrem"", 
                         p.withheld as ""withheld"",
                         (p.totalprem - p.withheld - p.specdiscamt) as ""totalAmt"",
                         p.commin_rate as ""commInRate"", 
                         p.commin_amt as ""commInAmt"", 
                         p.commin_taxamt as ""commInTaxAmt"", 
                         p.ovin_rate as ""ovInRate"", 
                         p.ovin_amt as ""ovInAmt"", 
                         p.ovin_taxamt as ""ovInTaxAmt"", 
                         p.commout_rate as ""commOutRate"", 
                         p.commout_amt as ""commOutAmt"", 
                         p.ovout_rate as ""ovOutRate"", 
                         p.ovout_amt as ""ovOutAmt"",
                         p.""insurerCode"" " + sql;
            if ((data.Limit).HasValue && (data.Pagecount).HasValue)
            {
                sql1 += $@" LIMIT {data.Limit}  OFFSET ({data.Pagecount} - 1) * {data.Limit} ;";
            }
            else { sql1 += " ;"; }

            string sql2 = "select count(*) " + sql + " ;";
            var json = await _dataContext.DailyPolicyDatas.FromSqlRaw(sql1).ToListAsync();
            var count = await _dataContext.DatasCount.FromSqlRaw(sql2).ToListAsync();
            return new DailyPolicyReportResult(json, count[0].Count);
        }
       
        public async Task<EndorseReportResult?> GetEndorseReportJson(EndorseReportInput data)
        {


            var sql = $@"from static_data.""b_juepcs"" juepc
                         left join static_data.""b_juedts"" juedt on juepc.polid = juedt.polid
                         left join static_data.""b_tuedts"" tuedt on tuedt.edtypecode = juedt.edtypecode
                         left join static_data.""b_juepms"" juepm on juepc.polid = juepm.polid
                         left join static_data.""Policies"" p on juepc.polid = p.id
                         left join static_data.""InsureTypes"" it on it.id = p.""insureID""
                         where  1=1 ";

            string currentDate = DateTime.Now.ToString("yyyy-MM-dd", new System.Globalization.CultureInfo("en-US"));
            if (!string.IsNullOrEmpty(data.StartedDate?.ToString()))
            {
                if (!string.IsNullOrEmpty(data.EndedDate?.ToString()))
                {
                    sql += $@"and date(juepc.""createdAt"") between '{data.StartedDate}' and '{data.EndedDate}' ";
                }
                else
                {
                    sql += $@"and date(juepc.""createdAt"") between '{data.StartedDate}' and '{currentDate}' ";
                }
            }
            if (!string.IsNullOrEmpty(data.Class?.ToString()))
            {
                sql += $@"and it.""class"" = '{data.Class}' ";
            }
            if (!string.IsNullOrEmpty(data.SubClass?.ToString()))
            {
                sql += $@"and it.""subClass"" = '{data.SubClass}' ";
            }

            if (!string.IsNullOrEmpty(data.Edtypecode?.ToString()))
            {
                sql += $@"and tuedt.""edtypecode"" = '{data.Edtypecode}' ";
            }

            if (!string.IsNullOrEmpty(data.StartpolicyNo?.ToString()))
            {
                sql += $@"and p.""policyNo"" >= '{data.StartpolicyNo}' ";
            }
            if (!string.IsNullOrEmpty(data.EndpolicyNo?.ToString()))
            {
                sql += $@"and p.""policyNo"" <= '{data.EndpolicyNo}' ";
            }
            if (!string.IsNullOrEmpty(data.StartendorseNo?.ToString()))
            {
                sql += $@"and juepc.""endorseNo"" >= '{data.StartendorseNo}' ";
            }
            if (!string.IsNullOrEmpty(data.EndendorseNo?.ToString()))
            {
                sql += $@"and juepc.""endorseNo"" <= '{data.EndendorseNo}' ";
            }

            string sql1 = @"select 
                         p.""applicationNo"", p.""policyNo"", juepc.""endorseNo"", p.""endorseseries"",
juepc.""createdAt"" as ""EndorseDate"" , 
(case when juedt.edtypecode like 'MT%' then 'แก้ไขภายใน' else 'สลักหลังภายนอก' end ) as ""EndorseType"" ,
juedt.""edtypecode"",
                         tuedt.""t_title"" as ""edtitle"",   
                         juedt.""detail"" as ""eddetail"",
                         juepm.""diffnetgrossprem"" as ""ednetgrossprem"",
                         juepm.""diffduty"" as ""edduty"",
                         juepm.""difftax"" as ""edtax"",
                         juepm.""difftotalprem"" as ""edtotalprem"" " + sql + $@"order by p.""policyNo"",juepc.""endorseNo"",p.""endorseseries""" ;

            if ((data.Limit).HasValue && (data.Pagecount).HasValue)
            {
                sql1 += $@" LIMIT {data.Limit}  OFFSET ({data.Pagecount} - 1) * {data.Limit} ;";
            }
            else { sql1 += " ;"; }


            string sql2 = "select count(*) as count " + sql + " ;";
            var json = await _dataContext.EndorseDatas.FromSqlRaw(sql1).ToListAsync();

            var count = await _dataContext.DatasCount.FromSqlRaw(sql2).ToListAsync();
            return new EndorseReportResult(json, count[0].Count);
        }
    }
}

    
        

