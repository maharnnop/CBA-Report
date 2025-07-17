using BestPolicyReport.Models.PremInDirectReport;
using BestPolicyReport.Services.PremInDirectService;
using ClosedXML.Excel;
using Microsoft.AspNetCore.Authorization;
using Microsoft.AspNetCore.Mvc;
using System.Globalization;

namespace BestPolicyReport.Controllers
{
    [Authorize]
    [Route("api/[controller]")]
    [ApiController]
    public class PremInDirectController : ControllerBase
    {
        private TimeZoneInfo bangkokTimeZone = TimeZoneInfo.FindSystemTimeZoneById("Asia/Bangkok");
        private readonly IPremInDirectService _premInDirectService;

        public PremInDirectController(IPremInDirectService premInDirectService)
        {
            _premInDirectService = premInDirectService;
        }

        [HttpPost("json")]
        public async Task<ActionResult<PremInDirectReportResult?>> GetPremInDirectReportJson(PremInDirectReportInput data)
        {
            var result = await _premInDirectService.GetPremInDirectReportJson(data);
            if (result.Count == 0)
            {
                return Ok(new PremInDirectReportResult());
            }
            return Ok(result);
        }

        [HttpPost("excel")]
        public async Task<IActionResult?> GetPremInDirectReportExcel(PremInDirectReportInput data)
        {
            var result = await _premInDirectService.GetPremInDirectReportJson(data);
            if (result.Count == 0)
            {
                return BadRequest("sql result = null");
            }
            using var workbook = new XLWorkbook();
            var sheetName = "ลูกค้าจ่ายเบี้ยที่ประกันโดยตรง";
            var worksheet = workbook.Worksheets.Add(sheetName);

            // Headers
            var headers = new string[]
             {
                 "รหัสบริษัทประกัน",
                 "รหัสผู้แนะนำ 1",
                 "วันที่กำหนดชำระ",
                 "หมายเลขกรมธรรม์",
                 "เลขสลักหลัง",
                 "เลขที่ใบแจ้งหนี้",
                 "เลขที่งวด",
                 "รหัสผู้เอาประกัน",
                 "ชื่อผู้เอาประกัน",
                 "ทะเบียนรถ",
                 "จังหวัด",
                 "เลขตัวถัง",
                 "เบี้ยรวม",
                 "อากร",
                 "ภาษี",
                 "เบี้ยประกันภัยรับรวม",
                 "อัตราคอมมิชชั่นจ่าย",
                 "ยอดคอมมิชชั่นจ่าย",
                 "อัตรา OV จ่าย",
                 "ยอด OV จ่าย",
                 "netFlag",
                 "billPremium"
            };

            for (int col = 1; col <= headers.Length; col++)
            {
                worksheet.Cell(3, col).Value = headers[col - 1];
            }

            // Data
            int row = 4;
            foreach (var i in result.Datas)
            {
                int col = 1;
                worksheet.Cell(row, col++).Value = i.InsurerCode;
                worksheet.Cell(row, col++).Value = i.AgentCode1;
                worksheet.Cell(row, col++).Value = i.DueDate;
                worksheet.Cell(row, col++).Value = i.PolicyNo;
                worksheet.Cell(row, col++).Value = i.EndorseNo;
                worksheet.Cell(row, col++).Value = i.InvoiceNo;
                worksheet.Cell(row, col++).Value = i.SeqNo;
                worksheet.Cell(row, col++).Value = i.InsureeCode;
                worksheet.Cell(row, col++).Value = i.InsureeName;
                worksheet.Cell(row, col++).Value = i.LicenseNo;
                worksheet.Cell(row, col++).Value = i.Province;
                worksheet.Cell(row, col++).Value = i.ChassisNo;
                worksheet.Cell(row, col++).Value = i.GrossPrem;
                worksheet.Cell(row, col++).Value = i.Duty;
                worksheet.Cell(row, col++).Value = i.Tax;
                worksheet.Cell(row, col++).Value = i.TotalPrem;
                worksheet.Cell(row, col++).Value = i.CommOutRate;
                worksheet.Cell(row, col++).Value = i.CommOutAmt;
                worksheet.Cell(row, col++).Value = i.OvOutRate;
                worksheet.Cell(row, col++).Value = i.OvOutAmt;
                worksheet.Cell(row, col++).Value = i.NetFlag;
                worksheet.Cell(row, col++).Value = i.BillPremium;
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
                sheetName, data.StartRpRefDate, data.EndRpRefDate, currentTime, result.Count, username);
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
