using BestPolicyReport.Models.DailyPolicyReport;
using BestPolicyReport.Services.DailyPolicyService;
using ClosedXML.Excel;
using Microsoft.AspNetCore.Authorization;
using Microsoft.AspNetCore.Mvc;
using Microsoft.IdentityModel.Tokens;
using System.IdentityModel.Tokens.Jwt;
using System.Text;
using System.Globalization;

using System.Security.Claims;

namespace BestPolicyReport.Controllers
{
    [Authorize]
    [Route("api/[controller]")]
    [ApiController]
    public class DailyPolicyController : ControllerBase
    {
        private TimeZoneInfo bangkokTimeZone = TimeZoneInfo.FindSystemTimeZoneById("Asia/Bangkok");
        private readonly IOICService _dailyPolicyService;

        public DailyPolicyController(IOICService dailyPolicyService)
        {
            _dailyPolicyService = dailyPolicyService;
        }

        [HttpPost("json")]
        public async Task<ActionResult<DailyPolicyReportResult?>> GetDailyPolicyReportJson(DailyPolicyReportInput data)
        {
            var result = await _dailyPolicyService.GetDailyPolicyReportJson(data);
            if (result.Count == 0)
            {
                return Ok(new DailyPolicyReportResult());
            }
            //var sckey = new SigningCredentials(new SymmetricSecurityKey(Encoding.UTF8.GetBytes("CA118BA49B699264C25AC66D5382F"), SecurityAlgorithms.HmacSha256Signature);
            //var jwtToken = new JwtSecurityToken(
            //notBefore: DateTime.UtcNow,
            //expires: DateTime.UtcNow.AddDays(30),
            //signingCredentials: sckey);
        
            return Ok(result);
        }

        [HttpPost("excel")]
        public async Task<IActionResult?> GetDailyPolicyReportExcel(DailyPolicyReportInput data)
        {
            var result = await _dailyPolicyService.GetDailyPolicyReportJson(data);
            if (result.Count == 0)
            {
                return BadRequest("sql result = null");
            }
            using var workbook = new XLWorkbook();
            var sheetName = "บันทึกกธประจำวัน";
            if (!string.IsNullOrEmpty(data.OrderBy?.ToString()))
            {
                sheetName += $"_ตาม{data.OrderBy}";
            }
            var worksheet = workbook.Worksheets.Add(sheetName);

            // Headers
            var headers = new string[]
             {
                 "ApplicationNo", 
                 "หมายเลขกรมธรรม์", 
                 "วันที่นำข้อมูลเข้า", 
                 "วันที่เริ่มคุ้มครอง", 
                 "วันที่สิ้นสุดคุ้มครอง",
                 "วันที่ทำสัญญา", 
                 "รหัสผู้บันทึก", 
                 "Username ผู้บันทึก", 
                 "รหัสผู้ดูแล 1", 
                 "ชื่อผู้ดูแล 1",
                 "รหัสผู้ดูแล 2", 
                 "ชื่อผู้ดูแล 2", 
                 "รหัสผู้แนะนำ 1", 
                 "ชื่อผู้แนะนำ 1", 
                 "รหัสผู้แนะนำ 2", 
                 "ชื่อผู้แนะนำ 2",
                 "รหัสผู้เอาประกัน", 
                 "ชื่อผู้เอาประกัน", 
                 "ประเภทประกัน", 
                 "ประเภทย่อยประกัน", 
                 "ทะเบียนรถ", 
                 "จังหวัด", 
                 "เลขตัวถัง",
                 "เบี้ยรวม", 
                 "อัตราส่วนลด", 
                 "มูลค่าส่วนลด", 
                 "เบี้ยสุทธิ", 
                 "อากร", 
                 "ภาษี",
                 "เบี้ยประกันภัยรับรวม", 
                 "อัตราคอมมิชชั่นรับ", 
                 "ยอดคอมมิชชั่นรับ", 
                 "ยอดภาษีคอมมิชชั่นรับ", 
                 "อัตรา OV รับ", 
                 "ยอด OV รับ", 
                 "ยอดภาษี OV รับ",
                 "อัตราคอมมิชชั่นจ่าย", 
                 "ยอดคอมมิชชั่นจ่าย", 
                 "อัตรา OV จ่าย", 
                 "ยอด OV จ่าย", 
                 "บริษัทประกัน"
            };

            for (int col = 1; col <= headers.Length; col++)
            {
                worksheet.Cell(3, col).Value = headers[col - 1];
            }

            // Data
            int row = 4;
            foreach (var dailyPolicy in result.Datas)
            {
                int col = 1;
                worksheet.Cell(row, col++).Value = dailyPolicy.ApplicationNo;
                worksheet.Cell(row, col++).Value = dailyPolicy.PolicyNo;
                worksheet.Cell(row, col++).Value = dailyPolicy.PolicyDate;
                worksheet.Cell(row, col++).Value = dailyPolicy.ActDate;
                worksheet.Cell(row, col++).Value = dailyPolicy.ExpDate;
                worksheet.Cell(row, col++).Value = dailyPolicy.IssueDate;
                worksheet.Cell(row, col++).Value = dailyPolicy.CreateUserCode;
                worksheet.Cell(row, col++).Value = dailyPolicy.Username;
                worksheet.Cell(row, col++).Value = dailyPolicy.ContactPersonId1;
                worksheet.Cell(row, col++).Value = dailyPolicy.ContactPersonName1;
                worksheet.Cell(row, col++).Value = dailyPolicy.ContactPersonId2;
                worksheet.Cell(row, col++).Value = dailyPolicy.ContactPersonName2;
                worksheet.Cell(row, col++).Value = dailyPolicy.AgentCode1;
                worksheet.Cell(row, col++).Value = dailyPolicy.AgentName1;
                worksheet.Cell(row, col++).Value = dailyPolicy.AgentCode2;
                worksheet.Cell(row, col++).Value = dailyPolicy.AgentName2;
                worksheet.Cell(row, col++).Value = dailyPolicy.InsureeCode;
                worksheet.Cell(row, col++).Value = dailyPolicy.InsureeName;
                worksheet.Cell(row, col++).Value = dailyPolicy.Class;
                worksheet.Cell(row, col++).Value = dailyPolicy.SubClass;
                worksheet.Cell(row, col++).Value = dailyPolicy.LicenseNo;
                worksheet.Cell(row, col++).Value = dailyPolicy.Province;
                worksheet.Cell(row, col++).Value = dailyPolicy.ChassisNo;
                worksheet.Cell(row, col++).Value = dailyPolicy.GrossPrem;
                worksheet.Cell(row, col++).Value = dailyPolicy.SpecDiscRate;
                worksheet.Cell(row, col++).Value = dailyPolicy.SpecDiscAmt;
                worksheet.Cell(row, col++).Value = dailyPolicy.NetGrossPrem;
                worksheet.Cell(row, col++).Value = dailyPolicy.Duty;
                worksheet.Cell(row, col++).Value = dailyPolicy.Tax;
                worksheet.Cell(row, col++).Value = dailyPolicy.TotalPrem;
                worksheet.Cell(row, col++).Value = dailyPolicy.CommInRate;
                worksheet.Cell(row, col++).Value = dailyPolicy.CommInAmt;
                worksheet.Cell(row, col++).Value = dailyPolicy.CommInTaxAmt;
                worksheet.Cell(row, col++).Value = dailyPolicy.OvInRate;
                worksheet.Cell(row, col++).Value = dailyPolicy.OvInAmt;
                worksheet.Cell(row, col++).Value = dailyPolicy.OvInTaxAmt;
                worksheet.Cell(row, col++).Value = dailyPolicy.CommOutRate;
                worksheet.Cell(row, col++).Value = dailyPolicy.CommOutAmt;
                worksheet.Cell(row, col++).Value = dailyPolicy.OvOutRate;
                worksheet.Cell(row, col++).Value = dailyPolicy.OvOutAmt;
                worksheet.Cell(row, col).Value = dailyPolicy.InsurerCode;

                row++;
            }

            var tableRange = worksheet.Range(3, 1, row - 1, headers.Length);
            var table = tableRange.AsTable();

            // You can set the table name and style here if needed
            table.Name = "Table";
            table.ShowAutoFilter = true;
            worksheet.Columns().AdjustToContents();

            DateTime bangkokTime = TimeZoneInfo.ConvertTimeFromUtc(DateTime.UtcNow, bangkokTimeZone);
            string currentTime = bangkokTime.ToString("dd/MM/yyyy HH:mm", CultureInfo.InvariantCulture);
            string username = User.Claims.FirstOrDefault(c => c.Type == "USERNAME")?.Value;
            String head = string.Format("รายงาน{0} และวันที่ {1} ถึงวันที่ {2} ,วันที่พิมพ์ {3} จำนวน {4} รายการ export by {5}",
                sheetName, data.StartPolicyDate, data.EndPolicyDate, currentTime, result.Count, username);
            worksheet.Cell(1, 1).Value = head;

            using var stream = new MemoryStream();
            workbook.SaveAs(stream);
            var content = stream.ToArray();

            return File(
                content,
                "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                $"รายงาน{sheetName}.xlsx");
        }


        [HttpPost("endorse/json")]
        public async Task<ActionResult<EndorseReportResult?>> GetEndorseReportJson(EndorseReportInput data)
        {
            var result = await _dailyPolicyService.GetEndorseReportJson(data);
            if (result.Count == 0)
            {
                return Ok(new EndorseReportResult());
            }
            return Ok(result);
        }

        [HttpPost("endorse/excel")]
        public async Task<IActionResult?> GetEndorseReportExcel(EndorseReportInput data)
        {
            var result = await _dailyPolicyService.GetEndorseReportJson(data);
            if (result.Count == 0)
            {
                return BadRequest("sql result = null");
            }
            using var workbook = new XLWorkbook();
            var sheetName = "บันทึกสลักหลัง";
            
            var worksheet = workbook.Worksheets.Add(sheetName);

            // Headers
            var headers = new string[]
             {
                 "เลขที่ใบคำขอ",
                 "เลขที่กรมธรรม์",
                 "เลขที่สลักหลัง",
                 "วันที่บันทึกสลักหลัง",
                 "ประเภทสลักหลัง",
                 "สลักหลังลำดับที่",
                 "ชื่อสลักหลัง",
                 "รายละเอียดที่เปลี่ยนแปลง",
                 "เบี้ย (สลักหลัง)",
                 "อากร (สลักหลัง)",
                 "ภาษี (สลักหลัง)",
                "เบี้ยรวม (สลักหลัง)",
                
            };

            for (int col = 1; col <= headers.Length; col++)
            {
                worksheet.Cell(3, col).Value = headers[col - 1];
            }

            // Data
            int row = 4;
            foreach (var dailyPolicy in result.Datas)
            {
                int col = 1;
                worksheet.Cell(row, col++).Value = dailyPolicy.ApplicationNo;
                worksheet.Cell(row, col++).Value = dailyPolicy.PolicyNo;
                worksheet.Cell(row, col++).Value = dailyPolicy.EndorseNo;
                worksheet.Cell(row, col++).Value = dailyPolicy.EndorseDate;
                worksheet.Cell(row, col++).Value = dailyPolicy.EndorseType;
                worksheet.Cell(row, col++).Value = dailyPolicy.Endorseseries;
                //worksheet.Cell(row, col++).Value = dailyPolicy.Edtypecode;
                worksheet.Cell(row, col++).Value = dailyPolicy.Edtitle;
                worksheet.Cell(row, col++).Value = dailyPolicy.Eddetail;
                worksheet.Cell(row, col++).Value = dailyPolicy.Ednetgrossprem;
                worksheet.Cell(row, col++).Value = dailyPolicy.Edduty;
                worksheet.Cell(row, col++).Value = dailyPolicy.Edtax;
                worksheet.Cell(row, col++).Value = dailyPolicy.Edtotalprem;

                row++;
            }

            var tableRange = worksheet.Range(3, 1, row - 1, headers.Length);
            var table = tableRange.AsTable();

            // You can set the table name and style here if needed
            table.Name = "Table";
            table.ShowAutoFilter = true;
            worksheet.Columns().AdjustToContents();

            DateTime bangkokTime = TimeZoneInfo.ConvertTimeFromUtc(DateTime.UtcNow, bangkokTimeZone);
            string currentTime = bangkokTime.ToString("dd/MM/yyyy HH:mm", CultureInfo.InvariantCulture);
            string username = User.Claims.FirstOrDefault(c => c.Type == "USERNAME")?.Value;
            String head = string.Format("รายงาน{0} และวันที่ {1} ถึงวันที่ {2} ,วันที่พิมพ์ {3} จำนวน {4} รายการ export by {5}",
                sheetName, data.StartedDate, data.EndedDate, currentTime, result.Count,username);
            worksheet.Cell(1, 1).Value = head;

            using var stream = new MemoryStream();
            workbook.SaveAs(stream);
            var content = stream.ToArray();

            return File(
                content,
                "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                $"รายงาน{sheetName}.xlsx");
        }

       
    }
}
