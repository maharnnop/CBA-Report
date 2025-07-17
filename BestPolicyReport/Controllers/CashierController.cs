using BestPolicyReport.Models.CashierReport;
using BestPolicyReport.Services.CashierService;
using ClosedXML.Excel;
using Microsoft.AspNetCore.Authorization;
using Microsoft.AspNetCore.Mvc;
using System.Globalization;

namespace BestPolicyReport.Controllers
{
    [Authorize]
    [Route("api/[controller]")]
    [ApiController]
    public class CashierController : ControllerBase
    {
        private TimeZoneInfo bangkokTimeZone = TimeZoneInfo.FindSystemTimeZoneById("Asia/Bangkok");
        private readonly ICashierService _cashierService;

        public CashierController(ICashierService cashierService)
        {
            _cashierService = cashierService;
        }

        [HttpPost("json")]
        public async Task<ActionResult<CashierReportResult?>> GetCashierReportJson(CashierReportInput data)
        {
            var result = await _cashierService.GetCashierReportJson(data);
            if (result.Count == 0)
            {
                return Ok(new CashierReportResult());
            }
            return Ok(result);
        }
        
        [HttpPost("excel")]
        public async Task<IActionResult?> GetCashierReportExcel(CashierReportInput data)
        {
            var result = await _cashierService.GetCashierReportJson(data);
            if (result.Count == 0)
            {
                return BadRequest("sql result = null");
            }
            using var workbook = new XLWorkbook();
            var sheetName = "ใบรับเงิน";
            var worksheet = workbook.Worksheets.Add(sheetName);

            // Headers
            var headers = new string[]
            {
            "เลขใบรับเงิน", "วันที่รับเงิน", 
            "เลขที่กรมธรรม์", "ผู้เอาประกัน",                
                "BillAdvisorNo", "วันที่ออกใบวางบิล", "ReceiveFrom", "ReceiveName", "ReceiveType",
            "RefNo", "RefDate", "TransactionType", "CashierAmt", "เลขที่ตัดชำระ", "วันที่ตัดชำระ",
            "จำนวนเงินตัดชำระ", "จำนวนเงินคงเหลือ", "สถานะรายการ"
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
                worksheet.Cell(row, col++).Value = i.CashierReceiveNo;
                worksheet.Cell(row, col++).Value = i.CashierDate;
                worksheet.Cell(row, col++).Value = i.PolicyNo;
                worksheet.Cell(row, col++).Value = i.InsureeName;
                worksheet.Cell(row, col++).Value = i.BillAdvisorNo;
                worksheet.Cell(row, col++).Value = i.BillDate;
                worksheet.Cell(row, col++).Value = i.ReceiveFrom;
                worksheet.Cell(row, col++).Value = i.ReceiveName;
                worksheet.Cell(row, col++).Value = i.ReceiveType;
                worksheet.Cell(row, col++).Value = i.RefNo;
                worksheet.Cell(row, col++).Value = i.RefDate;
                worksheet.Cell(row, col++).Value = i.TransactionType;
                worksheet.Cell(row, col++).Value = i.CashierAmt;
                worksheet.Cell(row, col++).Value = i.DfRpReferNo;
                worksheet.Cell(row, col++).Value = i.RpRefDate;
                worksheet.Cell(row, col++).Value = i.ActualValue;
                worksheet.Cell(row, col++).Value = i.DiffAmt;
                worksheet.Cell(row, col++).Value = i.Status;
                row++;
            }

            var tableRange = worksheet.Range(3, 1, row - 1, headers.Length);
            var table = tableRange.AsTable();
            table.Name = "Table";
            table.ShowAutoFilter = true;
            worksheet.Columns().AdjustToContents();

            DateTime bangkokTime = TimeZoneInfo.ConvertTimeFromUtc(DateTime.UtcNow, bangkokTimeZone);
            string currentTime = bangkokTime.ToString("dd/MM/yyyy HH:mm", CultureInfo.InvariantCulture);
            string username = User.Claims.FirstOrDefault(c => c.Type == "USERNAME")?.Value;
            String head = string.Format("รายงาน{0} และวันที่ {1} ถึงวันที่ {2} ,วันที่พิมพ์ {3} จำนวน {4} รายการ export by {5}",
                sheetName, data.StartCashierDate, data.EndCashierDate, currentTime, result.Count, username);
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
