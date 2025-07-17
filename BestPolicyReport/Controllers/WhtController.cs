﻿using BestPolicyReport.Models.WhtReport;
using BestPolicyReport.Services.WhtService;
using ClosedXML.Excel;
using Microsoft.AspNetCore.Authorization;
using Microsoft.AspNetCore.Mvc;
using System.Globalization;

namespace BestPolicyReport.Controllers
{
    [Authorize]
    [Route("api/[controller]")]
    [ApiController]
    public class WhtController : ControllerBase
    {
        private TimeZoneInfo bangkokTimeZone = TimeZoneInfo.FindSystemTimeZoneById("Asia/Bangkok");
        private readonly IWhtService _whtService;

        public WhtController(IWhtService whtService)
        {
            _whtService = whtService;
        }

        [HttpPost("whtCommOutOvOut/json")]
        public async Task<ActionResult<List<WhtCommOutOvOutReportResult>?>> GetWhtCommOutOvOutReportJson(WhtReportInput data)
        {
            var result = await _whtService.GetWhtCommOutOvOutReportJson(data);
            if (result == null)
            {
                return Ok(new List<WhtCommOutOvOutReportResult>());
            }
            return Ok(result);
        }

        [HttpPost("whtCommOut/excel")]
        public async Task<IActionResult?> GetWhtCommOutReportExcel(WhtReportInput data)
        {
            var result = await _whtService.GetWhtCommOutOvOutReportJson(data);
            if (result == null)
            {
                return BadRequest("sql result = null");
            }
            using var workbook = new XLWorkbook();
            var sheetName = "WHT_CommOut";
            var worksheet = workbook.Worksheets.Add(sheetName);

            // Headers
            var headers = new string[]
            {
                "เลขที่ตัดรับ",
                "วันที่ตัดรับ",
                "รหัสผู้แนะนำ",
                "ชื่อผู้แนะนำ",
                "ยอดคอมมิชชั่นจ่าย",
                "ยอดภาษีหัก ณ ที่จ่ายคอมมิชชั่นจ่าย",
            };

            for (int col = 1; col <= headers.Length; col++)
            {
                worksheet.Cell(3, col).Value = headers[col - 1];
            }

            // Data
            int row = 4;
            foreach (var i in result)
            {
                int col = 1;
                worksheet.Cell(row, col++).Value = i.DfRpReferNo;
                worksheet.Cell(row, col++).Value = i.RpRefDate;
                worksheet.Cell(row, col++).Value = i.AgentCode;
                worksheet.Cell(row, col++).Value = i.AgentName;
                worksheet.Cell(row, col++).Value = i.CommOutAmt;
                worksheet.Cell(row, col++).Value = i.WhtCommOutAmt;
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

        [HttpPost("whtOvOut/excel")]
        public async Task<IActionResult?> GetWhtOvOutReportExcel(WhtReportInput data)
        {
            var result = await _whtService.GetWhtCommOutOvOutReportJson(data);
            if (result == null)
            {
                return BadRequest("sql result = null");
            }
            using var workbook = new XLWorkbook();
            var sheetName = "WHT_OvOut";
            var worksheet = workbook.Worksheets.Add(sheetName);

            // Headers
            var headers = new string[]
            {
                "เลขที่ตัดรับ",
                "วันที่ตัดรับ",
                "รหัสผู้แนะนำ",
                "ชื่อผู้แนะนำ",
                "ยอด OV จ่าย",
                "ยอดภาษีหัก ณ ที่จ่าย OV จ่าย",
            };

            for (int col = 1; col <= headers.Length; col++)
            {
                worksheet.Cell(3, col).Value = headers[col - 1];
            }

            // Data
            int row = 4;
            foreach (var i in result)
            {
                int col = 1;
                worksheet.Cell(row, col++).Value = i.DfRpReferNo;
                worksheet.Cell(row, col++).Value = i.RpRefDate;
                worksheet.Cell(row, col++).Value = i.AgentCode;
                worksheet.Cell(row, col++).Value = i.AgentName;
                worksheet.Cell(row, col++).Value = i.OvOutAmt;
                worksheet.Cell(row, col++).Value = i.WhtOvOutAmt;
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
