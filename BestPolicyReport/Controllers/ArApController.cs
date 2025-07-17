﻿using BestPolicyReport.Models.ArApReport;
using BestPolicyReport.Services.ArApService;
using ClosedXML.Excel;
using DocumentFormat.OpenXml.Wordprocessing;
using Microsoft.AspNetCore.Authorization;
using Microsoft.AspNetCore.Mvc;
using OfficeOpenXml;
using System;
using System.Diagnostics.Contracts;
using System.Globalization;

namespace BestPolicyReport.Controllers
{
    [Authorize]
    [Route("api/[controller]")]
    [ApiController]
    public class ArApController : ControllerBase
    {
        private readonly IArApService _arApService;
        private TimeZoneInfo bangkokTimeZone = TimeZoneInfo.FindSystemTimeZoneById("Asia/Bangkok");
        public ArApController(IArApService arApService)
        {
            _arApService = arApService;
        }

        [HttpPost("premInOpenItem/json")]
        public async Task<ActionResult<PremInReportResult?>> GetPremInOpenItemReportJson(ArApReportInput data)
        {
            var result = await _arApService.GetPremInOpenItemReportJson(data);
            if (result.Count == 0)
            {
                return Ok(new PremInReportResult());
            }
            return Ok(result);
        }

        [HttpPost("premInOpenItem/excel")]
        public async Task<IActionResult?> GetPremInOpenItemReportExcel(ArApReportInput data)
        {
            var result = await _arApService.GetPremInOpenItemReportJson(data);
            if (result.Count == 0)
            {
                return BadRequest("sql result = null");
            }
            using var workbook = new XLWorkbook();
            var sheetName = "ตัดหนี้_PremIn_ตัวตั้ง";
            var worksheet = workbook.Worksheets.Add(sheetName);

            // Headers
            var headers = new string[]
             {
                 "หมายเลขกรมธรรม์",
                 "หมายเลขสลักหลัง",
                 "หมายเลขใบแจ้งหนี้",
                 "เลขที่งวด",
                 "เลขที่แคชเชียร์",
                 "วันที่แคชเชียร์",
                 "ยอดแคชเชียร์",
                 "CashierReceiveType",
                 "CashierRefNo",
                 "CashierRefDate",
                 "เลขที่ตัดหนี้ PremIn",
                 "วันที่ตัดหนี้",
                 "NetFlag",
                 "จำนวนเงินตัดหนี้",
                 "จำนวนเงินตัดหนี้คงเหลือ",
                 "สถานะกรมธรรม์",
                 "วันที่เริ่มคุ้มครอง",
                 "วันที่สิ้นสุดคุ้มครอง",
                 "รหัส Main Account",
                 "ชื่อ Main Account",
                 "รหัสผู้เอาประกัน",
                 "ชื่อผู้เอาประกัน",
                 "ประเภทประกัน",
                 "ประเภทย่อยประกัน",
                 "อากร",
                 "ภาษี",
                 "เบี้ยประกันภัยรับรวม",
                 "อัตราคอมมิชชั่นจ่าย",
                 "ยอดคอมมิชชั่นจ่าย",
                 "อัตรา OV จ่าย",
                 "ยอด OV จ่าย"
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
                worksheet.Cell(row, col++).Value = i.PolicyNo;
                worksheet.Cell(row, col++).Value = i.EndorseNo;
                worksheet.Cell(row, col++).Value = i.InvoiceNo;
                worksheet.Cell(row, col++).Value = i.SeqNo;
                worksheet.Cell(row, col++).Value = i.CashierReceiveNo;
                worksheet.Cell(row, col++).Value = i.CashierDate;
                worksheet.Cell(row, col++).Value = i.CashierAmt;
                worksheet.Cell(row, col++).Value = i.CashierReceiveType;
                worksheet.Cell(row, col++).Value = i.CashierRefNo;
                worksheet.Cell(row, col++).Value = i.CashierRefDate;
                worksheet.Cell(row, col++).Value = i.PremInDfRpReferNo;
                worksheet.Cell(row, col++).Value = i.RpRefDate;
                worksheet.Cell(row, col++).Value = i.NetFlag;
                worksheet.Cell(row, col++).Value = i.PremInPaidAmt;
                worksheet.Cell(row, col++).Value = i.PremInDiffAmt;
                worksheet.Cell(row, col++).Value = i.Insurancestatus;
                worksheet.Cell(row, col++).Value = i.ActDate;
                worksheet.Cell(row, col++).Value = i.ExpDate;
                worksheet.Cell(row, col++).Value = i.MainAccountCode;
                worksheet.Cell(row, col++).Value = i.MainAccountName;
                worksheet.Cell(row, col++).Value = i.InsureeCode;
                worksheet.Cell(row, col++).Value = i.InsureeName;
                worksheet.Cell(row, col++).Value = i.Class;
                worksheet.Cell(row, col++).Value = i.SubClass;
                worksheet.Cell(row, col++).Value = i.Duty;
                worksheet.Cell(row, col++).Value = i.Tax;
                worksheet.Cell(row, col++).Value = i.TotalPrem;
                worksheet.Cell(row, col++).Value = i.CommOutRate;
                worksheet.Cell(row, col++).Value = i.CommOutAmt;
                worksheet.Cell(row, col++).Value = i.OvOutRate;
                worksheet.Cell(row, col++).Value = i.OvOutAmt;

                row++;
            }

            // Dynamically construct the table range using Cells[row, col]
            
            var tableRange = worksheet.Range(3, 1, row - 1, headers.Length);
            //var tableRange = worksheet.RangeUsed();
            var table = tableRange.AsTable();

            // You can set the table name and style here if needed
            table.Name = "Table";
            table.ShowAutoFilter = true;
            worksheet.Columns().AdjustToContents();

            DateTime bangkokTime = TimeZoneInfo.ConvertTimeFromUtc(DateTime.UtcNow, bangkokTimeZone);
            string currentTime = bangkokTime.ToString("dd/MM/yyyy HH:mm", CultureInfo.InvariantCulture);
            string username = User.Claims.FirstOrDefault(c => c.Type == "USERNAME")?.Value;
            String head = string.Format("รายงาน{0} และวันที่ {1} ถึงวันที่ {2} ,วันที่พิมพ์ {3} จำนวน {4} รายการ export by {5}",
                sheetName, data.StartPolicyIssueDate, data.EndPolicyIssueDate, currentTime, result.Count, username);
            worksheet.Cell(1, 1).Value = head;

            using var stream = new MemoryStream();
            workbook.SaveAs(stream);
            var content = stream.ToArray();

            return File(
                content,
                "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                $"รายงาน{sheetName}.xlsx");
        }

        [HttpPost("premInClearing/json")]
        public async Task<ActionResult<PremInReportResult?>> GetPremInClearingReportJson(ArApReportInput data)
        {
            var result = await _arApService.GetPremInClearingReportJson(data);
            if (result.Count == 0)
            {
                return Ok(new PremInReportResult());
            }
            return Ok(result);
        }

        [HttpPost("premInClearing/excel")]
        public async Task<IActionResult?> GetPremInClearingReportExcel(ArApReportInput data)
        {
            var result = await _arApService.GetPremInClearingReportJson(data);
            if (result.Count == 0)
            {
                return BadRequest("sql result = null");
            }
            using var workbook = new XLWorkbook();
            var sheetName = "ตัดหนี้_PremIn_ตัวตัด";
            var worksheet = workbook.Worksheets.Add(sheetName);

            // Headers
            var headers = new string[]
             {
                 "หมายเลขกรมธรรม์",
                 "หมายเลขสลักหลัง",
                 "หมายเลขใบแจ้งหนี้",
                 "เลขที่งวด",
                 "เลขที่แคชเชียร์",
                 "วันที่แคชเชียร์",
                 "ยอดแคชเชียร์",
                 "CashierReceiveType",
                 "CashierRefNo",
                 "CashierRefDate",
                 "เลขที่ตัดหนี้ PremIn",
                 "วันที่ตัดหนี้",
                 "NetFlag",
                 "จำนวนเงินตัดหนี้",
                 "จำนวนเงินตัดหนี้คงเหลือ",
                 "สถานะกรมธรรม์",
                 "วันที่เริ่มคุ้มครอง",
                 "วันที่สิ้นสุดคุ้มครอง",
                 "รหัส Main Account",
                 "ชื่อ Main Account",
                 "รหัสผู้เอาประกัน",
                 "ชื่อผู้เอาประกัน",
                 "ประเภทประกัน",
                 "ประเภทย่อยประกัน",
                 "อากร",
                 "ภาษี",
                 "เบี้ยประกันภัยรับรวม",
                 "อัตราคอมมิชชั่นจ่ายของผู้แนะนำ 1", 
                 "ยอดคอมมิชชั่นจ่ายของผู้แนะนำ 1",
                 "อัตรา OV จ่ายของผู้แนะนำ 1",
                 "ยอด OV จ่ายของผู้แนะนำ 1",
                 "อัตราคอมมิชชั่นจ่ายของผู้แนะนำ 2",
                 "ยอดคอมมิชชั่นจ่ายของผู้แนะนำ 2"
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
                worksheet.Cell(row, col++).Value = i.PolicyNo;
                worksheet.Cell(row, col++).Value = i.EndorseNo;
                worksheet.Cell(row, col++).Value = i.InvoiceNo;
                worksheet.Cell(row, col++).Value = i.SeqNo;
                worksheet.Cell(row, col++).Value = i.CashierReceiveNo;
                worksheet.Cell(row, col++).Value = i.CashierDate;
                worksheet.Cell(row, col++).Value = i.CashierAmt;
                worksheet.Cell(row, col++).Value = i.CashierReceiveType;
                worksheet.Cell(row, col++).Value = i.CashierRefNo;
                worksheet.Cell(row, col++).Value = i.CashierRefDate;
                worksheet.Cell(row, col++).Value = i.PremInDfRpReferNo;
                worksheet.Cell(row, col++).Value = i.RpRefDate;
                worksheet.Cell(row, col++).Value = i.NetFlag;
                worksheet.Cell(row, col++).Value = i.PremInPaidAmt;
                worksheet.Cell(row, col++).Value = i.PremInDiffAmt;
                worksheet.Cell(row, col++).Value = i.Insurancestatus;
                worksheet.Cell(row, col++).Value = i.ActDate;
                worksheet.Cell(row, col++).Value = i.ExpDate;
                worksheet.Cell(row, col++).Value = i.MainAccountCode;
                worksheet.Cell(row, col++).Value = i.MainAccountName;
                worksheet.Cell(row, col++).Value = i.InsureeCode;
                worksheet.Cell(row, col++).Value = i.InsureeName;
                worksheet.Cell(row, col++).Value = i.Class;
                worksheet.Cell(row, col++).Value = i.SubClass;
                worksheet.Cell(row, col++).Value = i.Duty;
                worksheet.Cell(row, col++).Value = i.Tax;
                worksheet.Cell(row, col++).Value = i.TotalPrem;
                worksheet.Cell(row, col++).Value = i.CommOutRate1;
                worksheet.Cell(row, col++).Value = i.CommOutAmt1;
                worksheet.Cell(row, col++).Value = i.OvOutRate1;
                worksheet.Cell(row, col++).Value = i.OvOutAmt1;
                worksheet.Cell(row, col++).Value = i.CommOutRate2;
                worksheet.Cell(row, col++).Value = i.CommOutAmt2;

                row++;
            }

            var tableRange = worksheet.Range(3, 1, row - 1, headers.Length);
            //var tableRange = worksheet.RangeUsed();
            var table = tableRange.AsTable();

            // You can set the table name and style here if needed
            table.Name = "Table";
            table.ShowAutoFilter = true;
            worksheet.Columns().AdjustToContents();

            DateTime bangkokTime = TimeZoneInfo.ConvertTimeFromUtc(DateTime.UtcNow, bangkokTimeZone);
            string currentTime = bangkokTime.ToString("dd/MM/yyyy HH:mm", CultureInfo.InvariantCulture);
            string username = User.Claims.FirstOrDefault(c => c.Type == "USERNAME")?.Value;
            String head = string.Format("รายงาน{0} และวันที่ {1} ถึงวันที่ {2} ,วันที่พิมพ์ {3} จำนวน {4} รายการ export by {5}",
                sheetName, data.StartPolicyIssueDate, data.EndPolicyIssueDate, currentTime, result.Count, username);
            worksheet.Cell(1, 1).Value = head;

            using var stream = new MemoryStream();
            workbook.SaveAs(stream);
            var content = stream.ToArray();

            return File(
                content,
                "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                $"รายงาน{sheetName}.xlsx");
        }

        [HttpPost("premInOutstanding/json")]
        public async Task<ActionResult<PremInReportResult?>> GetPremInOutstandingReportJson(ArApReportInput data)
        {
            var result = await _arApService.GetPremInOutstandingReportJson(data);
            if (result.Count == 0)
            {
                return Ok(new PremInReportResult());
            }
            return Ok(result);
        }

        [HttpPost("premInOutstanding/excel")]
        public async Task<IActionResult?> GetPremInOutstandingReportExcel(ArApReportInput data)
        {
            var result = await _arApService.GetPremInOutstandingReportJson(data);
            if (result.Count == 0)
            {
                return BadRequest("sql result = null");
            }
            using var workbook = new XLWorkbook();
            var sheetName = "ตัดหนี้_PremIn_ตัวคงเหลือ";
            var worksheet = workbook.Worksheets.Add(sheetName);

            // Headers
            var headers = new string[]
             {
                 "หมายเลขกรมธรรม์",
                 "หมายเลขสลักหลัง",
                 "หมายเลขใบแจ้งหนี้",
                 "เลขที่งวด",
                 "เลขที่แคชเชียร์",
                 "วันที่แคชเชียร์",
                 "ยอดแคชเชียร์",
                 "CashierReceiveType",
                 "CashierRefNo",
                 "CashierRefDate",
                 "เลขที่ตัดหนี้ PremIn",
                 "วันที่ตัดหนี้",
                 "NetFlag",
                 "จำนวนเงินตัดหนี้",
                 "จำนวนเงินตัดหนี้คงเหลือ",
                 "สถานะกรมธรรม์",
                 "วันที่เริ่มคุ้มครอง",
                 "วันที่สิ้นสุดคุ้มครอง",
                 "รหัส Main Account",
                 "ชื่อ Main Account",
                 "รหัสผู้เอาประกัน",
                 "ชื่อผู้เอาประกัน",
                 "ประเภทประกัน",
                 "ประเภทย่อยประกัน",
                 "อากร",
                 "ภาษี",
                 "เบี้ยประกันภัยรับรวม",
                 "อัตราคอมมิชชั่นจ่าย",
                 "ยอดคอมมิชชั่นจ่าย",
                 "อัตรา OV จ่าย",
                 "ยอด OV จ่าย"
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
                worksheet.Cell(row, col++).Value = i.PolicyNo;
                worksheet.Cell(row, col++).Value = i.EndorseNo;
                worksheet.Cell(row, col++).Value = i.InvoiceNo;
                worksheet.Cell(row, col++).Value = i.SeqNo;
                worksheet.Cell(row, col++).Value = i.CashierReceiveNo;
                worksheet.Cell(row, col++).Value = i.CashierDate;
                worksheet.Cell(row, col++).Value = i.CashierAmt;
                worksheet.Cell(row, col++).Value = i.CashierReceiveType;
                worksheet.Cell(row, col++).Value = i.CashierRefNo;
                worksheet.Cell(row, col++).Value = i.CashierRefDate;
                worksheet.Cell(row, col++).Value = i.PremInDfRpReferNo;
                worksheet.Cell(row, col++).Value = i.RpRefDate;
                worksheet.Cell(row, col++).Value = i.NetFlag;
                worksheet.Cell(row, col++).Value = i.PremInPaidAmt;
                worksheet.Cell(row, col++).Value = i.PremInDiffAmt;
                worksheet.Cell(row, col++).Value = i.Insurancestatus;
                worksheet.Cell(row, col++).Value = i.ActDate;
                worksheet.Cell(row, col++).Value = i.ExpDate;
                worksheet.Cell(row, col++).Value = i.MainAccountCode;
                worksheet.Cell(row, col++).Value = i.MainAccountName;
                worksheet.Cell(row, col++).Value = i.InsureeCode;
                worksheet.Cell(row, col++).Value = i.InsureeName;
                worksheet.Cell(row, col++).Value = i.Class;
                worksheet.Cell(row, col++).Value = i.SubClass;
                worksheet.Cell(row, col++).Value = i.Duty;
                worksheet.Cell(row, col++).Value = i.Tax;
                worksheet.Cell(row, col++).Value = i.TotalPrem;
                worksheet.Cell(row, col++).Value = i.CommOutRate;
                worksheet.Cell(row, col++).Value = i.CommOutAmt;
                worksheet.Cell(row, col++).Value = i.OvOutRate;
                worksheet.Cell(row, col++).Value = i.OvOutAmt;

                row++;
            }

            var tableRange = worksheet.Range(3, 1, row - 1, headers.Length);
            //var tableRange = worksheet.RangeUsed();
            var table = tableRange.AsTable();

            // You can set the table name and style here if needed
            table.Name = "Table";
            table.ShowAutoFilter = true;
            worksheet.Columns().AdjustToContents();

            DateTime bangkokTime = TimeZoneInfo.ConvertTimeFromUtc(DateTime.UtcNow, bangkokTimeZone);
            string currentTime = bangkokTime.ToString("dd/MM/yyyy HH:mm", CultureInfo.InvariantCulture);
            string username = User.Claims.FirstOrDefault(c => c.Type == "USERNAME")?.Value;
            String head = string.Format("รายงาน{0} และวันที่ {1} ถึงวันที่ {2} ,วันที่พิมพ์ {3} จำนวน {4} รายการ export by {5}",
                sheetName, data.StartPolicyIssueDate, data.EndPolicyIssueDate, currentTime, result.Count, username);
            worksheet.Cell(1, 1).Value = head;

            using var stream = new MemoryStream();
            workbook.SaveAs(stream);
            var content = stream.ToArray();

            return File(
                content,
                "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                $"รายงาน{sheetName}.xlsx");
        }

        [HttpPost("commOutOvOutOpenItem/json")]
        public async Task<ActionResult<CommOutOvOutReportResult?>> GetCommOutOvOutOpenItemReportJson(ArApReportInput data)
        {
            var result = await _arApService.GetCommOutOvOutOpenItemReportJson(data);
            if (result.Count == 0)
            {
                return Ok(new CommOutOvOutReportResult());
            }
            return Ok(result);
        }

        [HttpPost("commOutOvOutOpenItem/excel")]
        public async Task<IActionResult?> GetCommOutOvOutOpenItemReportExcel(ArApReportInput data)
        {
            var result = await _arApService.GetCommOutOvOutOpenItemReportJson(data);
            if (result.Count == 0)
            {
                return BadRequest("sql result = null");
            }
            using var workbook = new XLWorkbook();
            var sheetName = "ตัดจ่าย_CommOutOvOut_ตัวตั้ง";
            var worksheet = workbook.Worksheets.Add(sheetName);

            // Headers
            var headers = new string[]
             {
                 "หมายเลขกรมธรรม์", 
                 "หมายเลขสลักหลัง", 
                 "หมายเลขใบแจ้งหนี้", 
                 "เลขที่งวด", 
                 "เลขที่แคชเชียร์", 
                 "วันที่แคชเชียร์", 
                 "ยอดแคชเชียร์", 
                 "CashierReceiveType", 
                 "CashierRefNo", 
                 "CashierRefDate", 
                 "เลขที่ตัดหนี้ PremIn", 
                 "วันที่ตัดหนี้", 
                 "เบี้ยรวม", 
                 "อัตราส่วนลด", 
                 "มูลค่าส่วนลด", 
                 "เบี้ยสุทธิ", 
                 "อากร", 
                 "ภาษี", 
                 "เบี้ยประกันภัยรับรวม", 
                 "NetFlag", 
                 "วันที่เริ่มคุ้มครอง", 
                 "วันที่สิ้นสุดคุ้มครอง", 
                 "รหัส Main Account", 
                 "ชื่อ Main Account", 
                 "รหัสผู้เอาประกัน", 
                 "ชื่อผู้เอาประกัน", 
                 "ประเภทประกัน", 
                 "ประเภทย่อยประกัน", 
                 "ทะเบียนรถ", 
                 "จังหวัด", 
                 "เลขตัวถัง", 
                 "อัตราคอมมิชชั่นจ่าย", 
                 "อัตรา OV จ่าย", 
                 "ยอด OV จ่าย", 
                 "CommOutDfRpReferNo", 
                 "CommOutRpRefDate", 
                 "CommOutPaidAmt", 
                 "CommOutDiffAmt", 
                 "OvOutPaidAmt", 
                 "OvOutDiffAmt"
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
                worksheet.Cell(row, col++).Value = i.PolicyNo;
                worksheet.Cell(row, col++).Value = i.EndorseNo;
                worksheet.Cell(row, col++).Value = i.InvoiceNo;
                worksheet.Cell(row, col++).Value = i.SeqNo;
                worksheet.Cell(row, col++).Value = i.CashierReceiveNo;
                worksheet.Cell(row, col++).Value = i.CashierDate;
                worksheet.Cell(row, col++).Value = i.CashierAmt;
                worksheet.Cell(row, col++).Value = i.CashierReceiveType;
                worksheet.Cell(row, col++).Value = i.CashierRefNo;
                worksheet.Cell(row, col++).Value = i.CashierRefDate;
                worksheet.Cell(row, col++).Value = i.PremInDfRpReferNo;
                worksheet.Cell(row, col++).Value = i.RpRefDate;
                worksheet.Cell(row, col++).Value = i.GrossPrem;
                worksheet.Cell(row, col++).Value = i.SpecDiscRate;
                worksheet.Cell(row, col++).Value = i.SpecDiscAmt;
                worksheet.Cell(row, col++).Value = i.NetGrossPrem;
                worksheet.Cell(row, col++).Value = i.Duty;
                worksheet.Cell(row, col++).Value = i.Tax;
                worksheet.Cell(row, col++).Value = i.TotalPrem;
                worksheet.Cell(row, col++).Value = i.NetFlag;
                worksheet.Cell(row, col++).Value = i.ActDate;
                worksheet.Cell(row, col++).Value = i.ExpDate;
                worksheet.Cell(row, col++).Value = i.MainAccountCode;
                worksheet.Cell(row, col++).Value = i.MainAccountName;
                worksheet.Cell(row, col++).Value = i.InsureeCode;
                worksheet.Cell(row, col++).Value = i.InsureeName;
                worksheet.Cell(row, col++).Value = i.Class;
                worksheet.Cell(row, col++).Value = i.SubClass;
                worksheet.Cell(row, col++).Value = i.LicenseNo;
                worksheet.Cell(row, col++).Value = i.Province;
                worksheet.Cell(row, col++).Value = i.ChassisNo;
                worksheet.Cell(row, col++).Value = i.CommOutRate;
                worksheet.Cell(row, col++).Value = i.OvOutRate;
                worksheet.Cell(row, col++).Value = i.OvOutAmt;
                worksheet.Cell(row, col++).Value = i.CommOutDfRpReferNo;
                worksheet.Cell(row, col++).Value = i.CommOutRpRefDate;
                worksheet.Cell(row, col++).Value = i.CommOutPaidAmt;
                worksheet.Cell(row, col++).Value = i.CommOutDiffAmt;
                worksheet.Cell(row, col++).Value = i.OvOutPaidAmt;
                worksheet.Cell(row, col).Value = i.OvOutDiffAmt;

                row++;
            }

            var tableRange = worksheet.Range(3, 1, row - 1, headers.Length);
            //var tableRange = worksheet.RangeUsed();
            var table = tableRange.AsTable();

            // You can set the table name and style here if needed
            table.Name = "Table";
            table.ShowAutoFilter = true;
            worksheet.Columns().AdjustToContents();

            DateTime bangkokTime = TimeZoneInfo.ConvertTimeFromUtc(DateTime.UtcNow, bangkokTimeZone);
            string currentTime = bangkokTime.ToString("dd/MM/yyyy HH:mm", CultureInfo.InvariantCulture);
            string username = User.Claims.FirstOrDefault(c => c.Type == "USERNAME")?.Value;
            String head = string.Format("รายงาน{0} และวันที่ {1} ถึงวันที่ {2} ,วันที่พิมพ์ {3} จำนวน {4} รายการ export by {5}",
                sheetName, data.StartPolicyIssueDate, data.EndPolicyIssueDate, currentTime, result.Count,username);
            worksheet.Cell(1, 1).Value = head;

            using var stream = new MemoryStream();
            workbook.SaveAs(stream);
            var content = stream.ToArray();

            return File(
                content,
                "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                $"รายงาน{sheetName}.xlsx");
        }

        [HttpPost("commOutOvOutClearing/json")]
        public async Task<ActionResult<CommOutOvOutReportResult?>> GetCommOutOvOutClearingReportJson(ArApReportInput data)
        {
            var result = await _arApService.GetCommOutOvOutClearingReportJson(data);
            if (result.Count == 0)
            {
                return Ok(new CommOutOvOutReportResult());
            }
            return Ok(result);
        }

        [HttpPost("commOutOvOutClearing/excel")]
        public async Task<IActionResult?> GetCommOutOvOutClearingReportExcel(ArApReportInput data)
        {
            var result = await _arApService.GetCommOutOvOutClearingReportJson(data);
            if (result.Count == 0)
            {
                return BadRequest("sql result = null");
            }
            using var workbook = new XLWorkbook();
            var sheetName = "ตัดจ่าย_CommOutOvOut_ตัวตัด";
            var worksheet = workbook.Worksheets.Add(sheetName);

            // Headers
            var headers = new string[]
             {
                 "หมายเลขกรมธรรม์",
                 "หมายเลขสลักหลัง",
                 "หมายเลขใบแจ้งหนี้",
                 "เลขที่งวด",
                 "เลขที่แคชเชียร์",
                 "วันที่แคชเชียร์",
                 "ยอดแคชเชียร์",
                 "CashierReceiveType",
                 "CashierRefNo",
                 "CashierRefDate",
                 "เลขที่ตัดหนี้ PremIn",
                 "วันที่ตัดหนี้",
                 "เบี้ยรวม",
                 "อัตราส่วนลด",
                 "มูลค่าส่วนลด",
                 "เบี้ยสุทธิ",
                 "อากร",
                 "ภาษี",
                 "เบี้ยประกันภัยรับรวม",
                 "NetFlag",
                 "วันที่เริ่มคุ้มครอง",
                 "วันที่สิ้นสุดคุ้มครอง",
                 "รหัส Main Account",
                 "ชื่อ Main Account",
                 "รหัสผู้เอาประกัน",
                 "ชื่อผู้เอาประกัน",
                 "ประเภทประกัน",
                 "ประเภทย่อยประกัน",
                 "ทะเบียนรถ",
                 "จังหวัด",
                 "เลขตัวถัง",
                 "อัตราคอมมิชชั่นจ่าย",
                 "อัตรา OV จ่าย",
                 "ยอด OV จ่าย",
                 "CommOutDfRpReferNo",
                 "CommOutRpRefDate",
                 "CommOutPaidAmt",
                 "CommOutDiffAmt",
                 "OvOutPaidAmt",
                 "OvOutDiffAmt"
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
                worksheet.Cell(row, col++).Value = i.PolicyNo;
                worksheet.Cell(row, col++).Value = i.EndorseNo;
                worksheet.Cell(row, col++).Value = i.InvoiceNo;
                worksheet.Cell(row, col++).Value = i.SeqNo;
                worksheet.Cell(row, col++).Value = i.CashierReceiveNo;
                worksheet.Cell(row, col++).Value = i.CashierDate;
                worksheet.Cell(row, col++).Value = i.CashierAmt;
                worksheet.Cell(row, col++).Value = i.CashierReceiveType;
                worksheet.Cell(row, col++).Value = i.CashierRefNo;
                worksheet.Cell(row, col++).Value = i.CashierRefDate;
                worksheet.Cell(row, col++).Value = i.PremInDfRpReferNo;
                worksheet.Cell(row, col++).Value = i.RpRefDate;
                worksheet.Cell(row, col++).Value = i.GrossPrem;
                worksheet.Cell(row, col++).Value = i.SpecDiscRate;
                worksheet.Cell(row, col++).Value = i.SpecDiscAmt;
                worksheet.Cell(row, col++).Value = i.NetGrossPrem;
                worksheet.Cell(row, col++).Value = i.Duty;
                worksheet.Cell(row, col++).Value = i.Tax;
                worksheet.Cell(row, col++).Value = i.TotalPrem;
                worksheet.Cell(row, col++).Value = i.NetFlag;
                worksheet.Cell(row, col++).Value = i.ActDate;
                worksheet.Cell(row, col++).Value = i.ExpDate;
                worksheet.Cell(row, col++).Value = i.MainAccountCode;
                worksheet.Cell(row, col++).Value = i.MainAccountName;
                worksheet.Cell(row, col++).Value = i.InsureeCode;
                worksheet.Cell(row, col++).Value = i.InsureeName;
                worksheet.Cell(row, col++).Value = i.Class;
                worksheet.Cell(row, col++).Value = i.SubClass;
                worksheet.Cell(row, col++).Value = i.LicenseNo;
                worksheet.Cell(row, col++).Value = i.Province;
                worksheet.Cell(row, col++).Value = i.ChassisNo;
                worksheet.Cell(row, col++).Value = i.CommOutRate;
                worksheet.Cell(row, col++).Value = i.OvOutRate;
                worksheet.Cell(row, col++).Value = i.OvOutAmt;
                worksheet.Cell(row, col++).Value = i.CommOutDfRpReferNo;
                worksheet.Cell(row, col++).Value = i.CommOutRpRefDate;
                worksheet.Cell(row, col++).Value = i.CommOutPaidAmt;
                worksheet.Cell(row, col++).Value = i.CommOutDiffAmt;
                worksheet.Cell(row, col++).Value = i.OvOutPaidAmt;
                worksheet.Cell(row, col).Value = i.OvOutDiffAmt;

                row++;
            }

            var tableRange = worksheet.Range(3, 1, row - 1, headers.Length);
            //var tableRange = worksheet.RangeUsed();
            var table = tableRange.AsTable();

            // You can set the table name and style here if needed
            table.Name = "Table";
            table.ShowAutoFilter = true;
            worksheet.Columns().AdjustToContents();

            DateTime bangkokTime = TimeZoneInfo.ConvertTimeFromUtc(DateTime.UtcNow, bangkokTimeZone);
            string currentTime = bangkokTime.ToString("dd/MM/yyyy HH:mm", CultureInfo.InvariantCulture);
            string username = User.Claims.FirstOrDefault(c => c.Type == "USERNAME")?.Value;
            String head = string.Format("รายงาน{0} และวันที่ {1} ถึงวันที่ {2} ,วันที่พิมพ์ {3} จำนวน {4} รายการ export by {5}",
                sheetName, data.StartPolicyIssueDate, data.EndPolicyIssueDate, currentTime, result.Count, username);
            worksheet.Cell(1, 1).Value = head;

            using var stream = new MemoryStream();
            workbook.SaveAs(stream);
            var content = stream.ToArray();

            return File(
                content,
                "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                $"รายงาน{sheetName}.xlsx");
        }

        [HttpPost("commOutOvOutOutstanding/json")]
        public async Task<ActionResult<CommOutOvOutReportResult?>> GetCommOutOvOutOutstandingReportJson(ArApReportInput data)
        {
            var result = await _arApService.GetCommOutOvOutOutstandingReportJson(data);
            if (result.Count == 0)
            {
                return Ok(new CommOutOvOutReportResult());
            }
            return Ok(result);
        }

        [HttpPost("commOutOvOutOutstanding/excel")]
        public async Task<IActionResult?> GetCommOutOvOutOutstandingReportExcel(ArApReportInput data)
        {
            var result = await _arApService.GetCommOutOvOutOutstandingReportJson(data);
            if (result.Count == 0)
            {
                return BadRequest("sql result = null");
            }
            using var workbook = new XLWorkbook();
            var sheetName = "ตัดจ่าย_CommOutOvOut_ตัวคงเหลือ";
            var worksheet = workbook.Worksheets.Add(sheetName);

            // Headers
            var headers = new string[]
             {
                 "หมายเลขกรมธรรม์",
                 "หมายเลขสลักหลัง",
                 "หมายเลขใบแจ้งหนี้",
                 "เลขที่งวด",
                 "เลขที่แคชเชียร์",
                 "วันที่แคชเชียร์",
                 "ยอดแคชเชียร์",
                 "CashierReceiveType",
                 "CashierRefNo",
                 "CashierRefDate",
                 "เลขที่ตัดหนี้ PremIn",
                 "วันที่ตัดหนี้",
                 "เบี้ยรวม",
                 "อัตราส่วนลด",
                 "มูลค่าส่วนลด",
                 "เบี้ยสุทธิ",
                 "อากร",
                 "ภาษี",
                 "เบี้ยประกันภัยรับรวม",
                 "NetFlag",
                 "วันที่เริ่มคุ้มครอง",
                 "วันที่สิ้นสุดคุ้มครอง",
                 "รหัส Main Account",
                 "ชื่อ Main Account",
                 "รหัสผู้เอาประกัน",
                 "ชื่อผู้เอาประกัน",
                 "ประเภทประกัน",
                 "ประเภทย่อยประกัน",
                 "ทะเบียนรถ",
                 "จังหวัด",
                 "เลขตัวถัง",
                 "ประเภทตัดจ่าย",
                 "อัตราคอมมิชชั่นจ่าย",
                 "อัตรา OV จ่าย",
                 "ยอด OV จ่าย",
                 "CommOutDfRpReferNo",
                 "CommOutRpRefDate",
                 "CommOutPaidAmt",
                 "CommOutDiffAmt",
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
                worksheet.Cell(row, col++).Value = i.PolicyNo;
                worksheet.Cell(row, col++).Value = i.EndorseNo;
                worksheet.Cell(row, col++).Value = i.InvoiceNo;
                worksheet.Cell(row, col++).Value = i.SeqNo;
                worksheet.Cell(row, col++).Value = i.CashierReceiveNo;
                worksheet.Cell(row, col++).Value = i.CashierDate;
                worksheet.Cell(row, col++).Value = i.CashierAmt;
                worksheet.Cell(row, col++).Value = i.CashierReceiveType;
                worksheet.Cell(row, col++).Value = i.CashierRefNo;
                worksheet.Cell(row, col++).Value = i.CashierRefDate;
                worksheet.Cell(row, col++).Value = i.PremInDfRpReferNo;
                worksheet.Cell(row, col++).Value = i.RpRefDate;
                worksheet.Cell(row, col++).Value = i.GrossPrem;
                worksheet.Cell(row, col++).Value = i.SpecDiscRate;
                worksheet.Cell(row, col++).Value = i.SpecDiscAmt;
                worksheet.Cell(row, col++).Value = i.NetGrossPrem;
                worksheet.Cell(row, col++).Value = i.Duty;
                worksheet.Cell(row, col++).Value = i.Tax;
                worksheet.Cell(row, col++).Value = i.TotalPrem;
                worksheet.Cell(row, col++).Value = i.NetFlag;
                worksheet.Cell(row, col++).Value = i.ActDate;
                worksheet.Cell(row, col++).Value = i.ExpDate;
                worksheet.Cell(row, col++).Value = i.MainAccountCode;
                worksheet.Cell(row, col++).Value = i.MainAccountName;
                worksheet.Cell(row, col++).Value = i.InsureeCode;
                worksheet.Cell(row, col++).Value = i.InsureeName;
                worksheet.Cell(row, col++).Value = i.Class;
                worksheet.Cell(row, col++).Value = i.SubClass;
                worksheet.Cell(row, col++).Value = i.LicenseNo;
                worksheet.Cell(row, col++).Value = i.Province;
                worksheet.Cell(row, col++).Value = i.ChassisNo;
                worksheet.Cell(row, col++).Value = i.TransactionType;
                worksheet.Cell(row, col++).Value = i.CommOutRate;
                worksheet.Cell(row, col++).Value = i.OvOutRate;
                worksheet.Cell(row, col++).Value = i.OvOutAmt;
                worksheet.Cell(row, col++).Value = i.CommOutDfRpReferNo;
                worksheet.Cell(row, col++).Value = i.CommOutRpRefDate;
                worksheet.Cell(row, col++).Value = i.CommOutPaidAmt;
                worksheet.Cell(row, col++).Value = i.CommOutDiffAmt;

                row++;
            }

            var tableRange = worksheet.Range(3, 1, row - 1, headers.Length);
            //var tableRange = worksheet.RangeUsed();
            var table = tableRange.AsTable();

            // You can set the table name and style here if needed
            table.Name = "Table";
            table.ShowAutoFilter = true;
            worksheet.Columns().AdjustToContents();

            DateTime bangkokTime = TimeZoneInfo.ConvertTimeFromUtc(DateTime.UtcNow, bangkokTimeZone);
            string currentTime = bangkokTime.ToString("dd/MM/yyyy HH:mm", CultureInfo.InvariantCulture);
            string username = User.Claims.FirstOrDefault(c => c.Type == "USERNAME")?.Value;
            String head = string.Format("รายงาน{0} และวันที่ {1} ถึงวันที่ {2} ,วันที่พิมพ์ {3} จำนวน {4} รายการ export by {5}",
                sheetName, data.StartPolicyIssueDate, data.EndPolicyIssueDate, currentTime, result.Count, username);
            worksheet.Cell(1, 1).Value = head;

            using var stream = new MemoryStream();
            workbook.SaveAs(stream);
            var content = stream.ToArray();

            return File(
                content,
                "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                $"รายงาน{sheetName}.xlsx");
        }

        [HttpPost("premOutOpenItem/json")]
        public async Task<ActionResult<PremOutCommInOvInReportResult?>> GetPremOutOpenItemReportJson(ArApReportInput data)
        {
            var result = await _arApService.GetPremOutOpenItemReportJson(data);
            if (result.Count == 0)
            {
                return Ok(new PremOutCommInOvInReportResult());
            }
            return Ok(result);
        }

        [HttpPost("premOutOpenItem/excel")]
        public async Task<IActionResult?> GetPremOutOpenItemReportExcel(ArApReportInput data)
        {
            var result = await _arApService.GetPremOutOpenItemReportJson(data);
            if (result.Count == 0)
            {
                return BadRequest("sql result = null");
            }
            using var workbook = new XLWorkbook();
            var sheetName = "ตัดจ่าย_PremOut_ตัวตั้ง";
            var worksheet = workbook.Worksheets.Add(sheetName);

            // Headers
            var headers = new string[]
             {
                 "หมายเลขกรมธรรม์",
                 "หมายเลขสลักหลัง",
                 "หมายเลขใบแจ้งหนี้",
                 "เลขที่งวด",
                 "เลขที่ตัดหนี้ PremIn",
                 "วันที่ตัดหนี้ PremIn",
                 "เลขที่ตัดหนี้ PremOut",
                 "วันที่ตัดหนี้ PremOut",
                 "วันที่เริ่มคุ้มครอง",
                 "วันที่สิ้นสุดคุ้มครอง",
                 "รหัสบริษัทประกัน",
                 "ชื่อบริษัทประกัน",
                 "รหัสผู้เอาประกัน",
                 "ชื่อผู้เอาประกัน",
                 "ประเภทประกัน",
                 "ประเภทย่อยประกัน",
                 "เบี้ยรวม",
                 "อัตราส่วนลด",
                 "มูลค่าส่วนลด",
                 "เบี้ยสุทธิ",
                 "อากร",
                 "ภาษี",
                 "เบี้ยประกันภัยรับรวม",
                 "NetFlag",
                 "อัตราคอมมิชชั่นรับ",
                 "ยอดคอมมิชชั่นรับ",
                 "เลขที่ตัดหนี้ CommIn",
                 "วันที่ตัดหนี้ CommIn",
                 "จำนวนเงิน CommIn ตัดรับ",
                 "จำนวนเงิน CommIn คงเหลือ",
                 "อัตรา OV รับ",
                 "ยอด OV รับ",
                 "เลขที่ตัดหนี้ OvIn",
                 "วันที่ตัดหนี้ OvIn",
                 "จำนวนเงิน OvIn ตัดรับ",
                 "จำนวนเงิน OvIn คงเหลือ",
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
                worksheet.Cell(row, col++).Value = i.PolicyNo;
                worksheet.Cell(row, col++).Value = i.EndorseNo;
                worksheet.Cell(row, col++).Value = i.InvoiceNo;
                worksheet.Cell(row, col++).Value = i.SeqNo;
                worksheet.Cell(row, col++).Value = i.PremInDfRpReferNo;
                worksheet.Cell(row, col++).Value = i.PremInRpRefDate;
                worksheet.Cell(row, col++).Value = i.PremOutDfRpReferNo;
                worksheet.Cell(row, col++).Value = i.PremOutRpRefDate;
                worksheet.Cell(row, col++).Value = i.ActDate;
                worksheet.Cell(row, col++).Value = i.ExpDate;
                worksheet.Cell(row, col++).Value = i.InsurerCode;
                worksheet.Cell(row, col++).Value = i.InsurerName;
                worksheet.Cell(row, col++).Value = i.InsureeCode;
                worksheet.Cell(row, col++).Value = i.InsureeName;
                worksheet.Cell(row, col++).Value = i.Class;
                worksheet.Cell(row, col++).Value = i.SubClass;
                worksheet.Cell(row, col++).Value = i.GrossPrem;
                worksheet.Cell(row, col++).Value = i.SpecDiscRate;
                worksheet.Cell(row, col++).Value = i.SpecDiscAmt;
                worksheet.Cell(row, col++).Value = i.NetGrossPrem;
                worksheet.Cell(row, col++).Value = i.Duty;
                worksheet.Cell(row, col++).Value = i.Tax;
                worksheet.Cell(row, col++).Value = i.TotalPrem;
                worksheet.Cell(row, col++).Value = i.NetFlag;
                worksheet.Cell(row, col++).Value = i.CommInRate;
                worksheet.Cell(row, col++).Value = i.CommInAmt;
                worksheet.Cell(row, col++).Value = i.CommInDfRpReferNo;
                worksheet.Cell(row, col++).Value = i.CommInRpRefDate;
                worksheet.Cell(row, col++).Value = i.CommInPaidAmt;
                worksheet.Cell(row, col++).Value = i.CommInDiffAmt;
                worksheet.Cell(row, col++).Value = i.OvInRate;
                worksheet.Cell(row, col++).Value = i.OvInAmt;
                worksheet.Cell(row, col++).Value = i.OvInDfRpReferNo;
                worksheet.Cell(row, col++).Value = i.OvInRpRefDate;
                worksheet.Cell(row, col++).Value = i.OvInPaidAmt;
                worksheet.Cell(row, col).Value = i.OvInDiffAmt;

                row++;
            }

            var tableRange = worksheet.Range(3, 1, row - 1, headers.Length);
            //var tableRange = worksheet.RangeUsed();
            var table = tableRange.AsTable();

            // You can set the table name and style here if needed
            table.Name = "Table";
            table.ShowAutoFilter = true;
            worksheet.Columns().AdjustToContents();

            DateTime bangkokTime = TimeZoneInfo.ConvertTimeFromUtc(DateTime.UtcNow, bangkokTimeZone);
            string currentTime = bangkokTime.ToString("dd/MM/yyyy HH:mm", CultureInfo.InvariantCulture);
            string username = User.Claims.FirstOrDefault(c => c.Type == "USERNAME")?.Value;
            String head = string.Format("รายงาน{0} และวันที่ {1} ถึงวันที่ {2} ,วันที่พิมพ์ {3} จำนวน {4} รายการ export by {5}",
                sheetName, data.StartPolicyIssueDate, data.EndPolicyIssueDate, currentTime, result.Count, username);
            worksheet.Cell(1, 1).Value = head;

            using var stream = new MemoryStream();
            workbook.SaveAs(stream);
            var content = stream.ToArray();

            return File(
                content,
                "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                $"รายงาน{sheetName}.xlsx");
        }

        [HttpPost("premOutClearing/json")]
        public async Task<ActionResult<PremOutCommInOvInReportResult?>> GetPremOutClearingReportJson(ArApReportInput data)
        {
            var result = await _arApService.GetPremOutClearingReportJson(data);
            if (result.Count == 0)
            {
                return Ok(new PremOutCommInOvInReportResult());
            }
            return Ok(result);
        }

        [HttpPost("premOutClearing/excel")]
        public async Task<IActionResult?> GetPremOutClearingReportExcel(ArApReportInput data)
        {
            var result = await _arApService.GetPremOutClearingReportJson(data);
            if (result.Count == 0)
            {
                return BadRequest("sql result = null");
            }
            using var workbook = new XLWorkbook();
            var sheetName = "ตัดจ่าย_PremOut_ตัวตัด";
            var worksheet = workbook.Worksheets.Add(sheetName);

            // Headers
            var headers = new string[]
             {
                 "หมายเลขกรมธรรม์",
                 "หมายเลขสลักหลัง",
                 "หมายเลขใบแจ้งหนี้",
                 "เลขที่งวด",
                 "เลขที่ตัดหนี้ PremIn",
                 "วันที่ตัดหนี้ PremIn",
                 "เลขที่ตัดหนี้ PremOut",
                 "วันที่ตัดหนี้ PremOut",
                 "วันที่เริ่มคุ้มครอง",
                 "วันที่สิ้นสุดคุ้มครอง",
                 "รหัสบริษัทประกัน",
                 "ชื่อบริษัทประกัน",
                 "รหัสผู้เอาประกัน",
                 "ชื่อผู้เอาประกัน",
                 "ประเภทประกัน",
                 "ประเภทย่อยประกัน",
                 "เบี้ยรวม",
                 "อัตราส่วนลด",
                 "มูลค่าส่วนลด",
                 "เบี้ยสุทธิ",
                 "อากร",
                 "ภาษี",
                 "เบี้ยประกันภัยรับรวม",
                 "NetFlag",
                 "อัตราคอมมิชชั่นรับ",
                 "ยอดคอมมิชชั่นรับ",
                 "เลขที่ตัดหนี้ CommIn",
                 "วันที่ตัดหนี้ CommIn",
                 "จำนวนเงิน CommIn ตัดรับ",
                 "จำนวนเงิน CommIn คงเหลือ",
                 "อัตรา OV รับ",
                 "ยอด OV รับ",
                 "เลขที่ตัดหนี้ OvIn",
                 "วันที่ตัดหนี้ OvIn",
                 "จำนวนเงิน OvIn ตัดรับ",
                 "จำนวนเงิน OvIn คงเหลือ",
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
                worksheet.Cell(row, col++).Value = i.PolicyNo;
                worksheet.Cell(row, col++).Value = i.EndorseNo;
                worksheet.Cell(row, col++).Value = i.InvoiceNo;
                worksheet.Cell(row, col++).Value = i.SeqNo;
                worksheet.Cell(row, col++).Value = i.PremInDfRpReferNo;
                worksheet.Cell(row, col++).Value = i.PremInRpRefDate;
                worksheet.Cell(row, col++).Value = i.PremOutDfRpReferNo;
                worksheet.Cell(row, col++).Value = i.PremOutRpRefDate;
                worksheet.Cell(row, col++).Value = i.ActDate;
                worksheet.Cell(row, col++).Value = i.ExpDate;
                worksheet.Cell(row, col++).Value = i.InsurerCode;
                worksheet.Cell(row, col++).Value = i.InsurerName;
                worksheet.Cell(row, col++).Value = i.InsureeCode;
                worksheet.Cell(row, col++).Value = i.InsureeName;
                worksheet.Cell(row, col++).Value = i.Class;
                worksheet.Cell(row, col++).Value = i.SubClass;
                worksheet.Cell(row, col++).Value = i.GrossPrem;
                worksheet.Cell(row, col++).Value = i.SpecDiscRate;
                worksheet.Cell(row, col++).Value = i.SpecDiscAmt;
                worksheet.Cell(row, col++).Value = i.NetGrossPrem;
                worksheet.Cell(row, col++).Value = i.Duty;
                worksheet.Cell(row, col++).Value = i.Tax;
                worksheet.Cell(row, col++).Value = i.TotalPrem;
                worksheet.Cell(row, col++).Value = i.NetFlag;
                worksheet.Cell(row, col++).Value = i.CommInRate;
                worksheet.Cell(row, col++).Value = i.CommInAmt;
                worksheet.Cell(row, col++).Value = i.CommInDfRpReferNo;
                worksheet.Cell(row, col++).Value = i.CommInRpRefDate;
                worksheet.Cell(row, col++).Value = i.CommInPaidAmt;
                worksheet.Cell(row, col++).Value = i.CommInDiffAmt;
                worksheet.Cell(row, col++).Value = i.OvInRate;
                worksheet.Cell(row, col++).Value = i.OvInAmt;
                worksheet.Cell(row, col++).Value = i.OvInDfRpReferNo;
                worksheet.Cell(row, col++).Value = i.OvInRpRefDate;
                worksheet.Cell(row, col++).Value = i.OvInPaidAmt;
                worksheet.Cell(row, col).Value = i.OvInDiffAmt;

                row++;
            }

            var tableRange = worksheet.Range(3, 1, row - 1, headers.Length);
            //var tableRange = worksheet.RangeUsed();
            var table = tableRange.AsTable();

            // You can set the table name and style here if needed
            table.Name = "Table";
            table.ShowAutoFilter = true;
            worksheet.Columns().AdjustToContents();

            DateTime bangkokTime = TimeZoneInfo.ConvertTimeFromUtc(DateTime.UtcNow, bangkokTimeZone);
            string currentTime = bangkokTime.ToString("dd/MM/yyyy HH:mm", CultureInfo.InvariantCulture);
            string username = User.Claims.FirstOrDefault(c => c.Type == "USERNAME")?.Value;
            String head = string.Format("รายงาน{0} และวันที่ {1} ถึงวันที่ {2} ,วันที่พิมพ์ {3} จำนวน {4} รายการ export by {5}",
                sheetName, data.StartPolicyIssueDate, data.EndPolicyIssueDate, currentTime, result.Count, username);
            worksheet.Cell(1, 1).Value = head;

            using var stream = new MemoryStream();
            workbook.SaveAs(stream);
            var content = stream.ToArray();

            return File(
                content,
                "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                $"รายงาน{sheetName}.xlsx");
        }

        [HttpPost("premOutOutstanding/json")]
        public async Task<ActionResult<PremOutCommInOvInReportResult?>> GetPremOutOutstandingReportJson(ArApReportInput data)
        {
            var result = await _arApService.GetPremOutOutstandingReportJson(data);
            if (result.Count == 0)
            {
                return Ok(new PremOutCommInOvInReportResult());
            }
            return Ok(result);
        }

        [HttpPost("premOutOutstanding/excel")]
        public async Task<IActionResult?> GetPremOutOutstandingReportExcel(ArApReportInput data)
        {
            var result = await _arApService.GetPremOutOutstandingReportJson(data);
            if (result.Count == 0)
            {
                return BadRequest("sql result = null");
            }
            using var workbook = new XLWorkbook();
            var sheetName = "ตัดจ่าย_PremOut_ตัวคงเหลือ";
            var worksheet = workbook.Worksheets.Add(sheetName);

            // Headers
            var headers = new string[]
             {
                 "หมายเลขกรมธรรม์",
                 "หมายเลขสลักหลัง",
                 "หมายเลขใบแจ้งหนี้",
                 "เลขที่งวด",
                 "เลขที่ตัดหนี้ PremIn",
                 "วันที่ตัดหนี้ PremIn",
                 "เลขที่ตัดหนี้ PremOut",
                 "วันที่ตัดหนี้ PremOut",
                 "วันที่เริ่มคุ้มครอง",
                 "วันที่สิ้นสุดคุ้มครอง",
                 "รหัสบริษัทประกัน",
                 "ชื่อบริษัทประกัน",
                 "รหัสผู้เอาประกัน",
                 "ชื่อผู้เอาประกัน",
                 "ประเภทประกัน",
                 "ประเภทย่อยประกัน",
                 "เบี้ยรวม",
                 "อัตราส่วนลด",
                 "มูลค่าส่วนลด",
                 "เบี้ยสุทธิ",
                 "อากร",
                 "ภาษี",
                 "เบี้ยประกันภัยรับรวม",
                 "NetFlag",
                 "อัตราคอมมิชชั่นรับ",
                 "ยอดคอมมิชชั่นรับ",
                 "เลขที่ตัดหนี้ CommIn",
                 "วันที่ตัดหนี้ CommIn",
                 "จำนวนเงิน CommIn ตัดรับ",
                 "จำนวนเงิน CommIn คงเหลือ",
                 "อัตรา OV รับ",
                 "ยอด OV รับ",
                 "เลขที่ตัดหนี้ OvIn",
                 "วันที่ตัดหนี้ OvIn",
                 "จำนวนเงิน OvIn ตัดรับ",
                 "จำนวนเงิน OvIn คงเหลือ",
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
                worksheet.Cell(row, col++).Value = i.PolicyNo;
                worksheet.Cell(row, col++).Value = i.EndorseNo;
                worksheet.Cell(row, col++).Value = i.InvoiceNo;
                worksheet.Cell(row, col++).Value = i.SeqNo;
                worksheet.Cell(row, col++).Value = i.PremInDfRpReferNo;
                worksheet.Cell(row, col++).Value = i.PremInRpRefDate;
                worksheet.Cell(row, col++).Value = i.PremOutDfRpReferNo;
                worksheet.Cell(row, col++).Value = i.PremOutRpRefDate;
                worksheet.Cell(row, col++).Value = i.ActDate;
                worksheet.Cell(row, col++).Value = i.ExpDate;
                worksheet.Cell(row, col++).Value = i.InsurerCode;
                worksheet.Cell(row, col++).Value = i.InsurerName;
                worksheet.Cell(row, col++).Value = i.InsureeCode;
                worksheet.Cell(row, col++).Value = i.InsureeName;
                worksheet.Cell(row, col++).Value = i.Class;
                worksheet.Cell(row, col++).Value = i.SubClass;
                worksheet.Cell(row, col++).Value = i.GrossPrem;
                worksheet.Cell(row, col++).Value = i.SpecDiscRate;
                worksheet.Cell(row, col++).Value = i.SpecDiscAmt;
                worksheet.Cell(row, col++).Value = i.NetGrossPrem;
                worksheet.Cell(row, col++).Value = i.Duty;
                worksheet.Cell(row, col++).Value = i.Tax;
                worksheet.Cell(row, col++).Value = i.TotalPrem;
                worksheet.Cell(row, col++).Value = i.NetFlag;
                worksheet.Cell(row, col++).Value = i.CommInRate;
                worksheet.Cell(row, col++).Value = i.CommInAmt;
                worksheet.Cell(row, col++).Value = i.CommInDfRpReferNo;
                worksheet.Cell(row, col++).Value = i.CommInRpRefDate;
                worksheet.Cell(row, col++).Value = i.CommInPaidAmt;
                worksheet.Cell(row, col++).Value = i.CommInDiffAmt;
                worksheet.Cell(row, col++).Value = i.OvInRate;
                worksheet.Cell(row, col++).Value = i.OvInAmt;
                worksheet.Cell(row, col++).Value = i.OvInDfRpReferNo;
                worksheet.Cell(row, col++).Value = i.OvInRpRefDate;
                worksheet.Cell(row, col++).Value = i.OvInPaidAmt;
                worksheet.Cell(row, col).Value = i.OvInDiffAmt;

                row++;
            }

            var tableRange = worksheet.Range(3, 1, row - 1, headers.Length);
            //var tableRange = worksheet.RangeUsed();
            var table = tableRange.AsTable();

            // You can set the table name and style here if needed
            table.Name = "Table";
            table.ShowAutoFilter = true;
            worksheet.Columns().AdjustToContents();

            DateTime bangkokTime = TimeZoneInfo.ConvertTimeFromUtc(DateTime.UtcNow, bangkokTimeZone);
            string currentTime = bangkokTime.ToString("dd/MM/yyyy HH:mm", CultureInfo.InvariantCulture);
            string username = User.Claims.FirstOrDefault(c => c.Type == "USERNAME")?.Value;
            String head = string.Format("รายงาน{0} และวันที่ {1} ถึงวันที่ {2} ,วันที่พิมพ์ {3} จำนวน {4} รายการ export by {5}",
                sheetName, data.StartPolicyIssueDate, data.EndPolicyIssueDate, currentTime, result.Count, username);
            worksheet.Cell(1, 1).Value = head;

            using var stream = new MemoryStream();
            workbook.SaveAs(stream);
            var content = stream.ToArray();

            return File(
                content,
                "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                $"รายงาน{sheetName}.xlsx");
        }

        [HttpPost("commInOvInOpenItem/json")]
        public async Task<ActionResult<PremOutCommInOvInReportResult?>> GetCommInOvInOpenItemReportJson(ArApReportInput data)
        {
            var result = await _arApService.GetCommInOvInOpenItemReportJson(data);
            if (result.Count == 0)
            {
                return Ok(new PremOutCommInOvInReportResult());
            }
            return Ok(result);
        }

        [HttpPost("commInOvInOpenItem/excel")]
        public async Task<IActionResult?> GetCommInOvInOpenItemReportExcel(ArApReportInput data)
        {
            var result = await _arApService.GetCommInOvInOpenItemReportJson(data);
            if (result.Count == 0)
            {
                return BadRequest("sql result = null");
            }
            using var workbook = new XLWorkbook();
            var sheetName = "ตัดหนี้_CommInOvIn_ตัวตั้ง";
            var worksheet = workbook.Worksheets.Add(sheetName);

            // Headers
            var headers = new string[]
             {
                 "หมายเลขกรมธรรม์",
                 "หมายเลขสลักหลัง",
                 "หมายเลขใบแจ้งหนี้",
                 "เลขที่งวด",
                 "เลขที่ตัดหนี้ PremIn",
                 "วันที่ตัดหนี้ PremIn",
                 "เลขที่ตัดหนี้ PremOut",
                 "วันที่ตัดหนี้ PremOut",
                 "วันที่เริ่มคุ้มครอง",
                 "วันที่สิ้นสุดคุ้มครอง",
                 "รหัสบริษัทประกัน",
                 "ชื่อบริษัทประกัน",
                 "รหัสผู้เอาประกัน",
                 "ชื่อผู้เอาประกัน",
                 "ประเภทประกัน",
                 "ประเภทย่อยประกัน",
                 "เบี้ยรวม",
                 "อัตราส่วนลด",
                 "มูลค่าส่วนลด",
                 "เบี้ยสุทธิ",
                 "อากร",
                 "ภาษี",
                 "เบี้ยประกันภัยรับรวม",
                 "NetFlag",
                 "อัตราคอมมิชชั่นรับ",
                 "ยอดคอมมิชชั่นรับ",
                 "เลขที่ตัดหนี้ CommIn",
                 "วันที่ตัดหนี้ CommIn",
                 "จำนวนเงิน CommIn ตัดรับ",
                 "จำนวนเงิน CommIn คงเหลือ",
                 "อัตรา OV รับ",
                 "ยอด OV รับ",
                 "เลขที่ตัดหนี้ OvIn",
                 "วันที่ตัดหนี้ OvIn",
                 "จำนวนเงิน OvIn ตัดรับ",
                 "จำนวนเงิน OvIn คงเหลือ",
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
                worksheet.Cell(row, col++).Value = i.PolicyNo;
                worksheet.Cell(row, col++).Value = i.EndorseNo;
                worksheet.Cell(row, col++).Value = i.InvoiceNo;
                worksheet.Cell(row, col++).Value = i.SeqNo;
                worksheet.Cell(row, col++).Value = i.PremInDfRpReferNo;
                worksheet.Cell(row, col++).Value = i.PremInRpRefDate;
                worksheet.Cell(row, col++).Value = i.PremOutDfRpReferNo;
                worksheet.Cell(row, col++).Value = i.PremOutRpRefDate;
                worksheet.Cell(row, col++).Value = i.ActDate;
                worksheet.Cell(row, col++).Value = i.ExpDate;
                worksheet.Cell(row, col++).Value = i.InsurerCode;
                worksheet.Cell(row, col++).Value = i.InsurerName;
                worksheet.Cell(row, col++).Value = i.InsureeCode;
                worksheet.Cell(row, col++).Value = i.InsureeName;
                worksheet.Cell(row, col++).Value = i.Class;
                worksheet.Cell(row, col++).Value = i.SubClass;
                worksheet.Cell(row, col++).Value = i.GrossPrem;
                worksheet.Cell(row, col++).Value = i.SpecDiscRate;
                worksheet.Cell(row, col++).Value = i.SpecDiscAmt;
                worksheet.Cell(row, col++).Value = i.NetGrossPrem;
                worksheet.Cell(row, col++).Value = i.Duty;
                worksheet.Cell(row, col++).Value = i.Tax;
                worksheet.Cell(row, col++).Value = i.TotalPrem;
                worksheet.Cell(row, col++).Value = i.NetFlag;
                worksheet.Cell(row, col++).Value = i.CommInRate;
                worksheet.Cell(row, col++).Value = i.CommInAmt;
                worksheet.Cell(row, col++).Value = i.CommInDfRpReferNo;
                worksheet.Cell(row, col++).Value = i.CommInRpRefDate;
                worksheet.Cell(row, col++).Value = i.CommInPaidAmt;
                worksheet.Cell(row, col++).Value = i.CommInDiffAmt;
                worksheet.Cell(row, col++).Value = i.OvInRate;
                worksheet.Cell(row, col++).Value = i.OvInAmt;
                worksheet.Cell(row, col++).Value = i.OvInDfRpReferNo;
                worksheet.Cell(row, col++).Value = i.OvInRpRefDate;
                worksheet.Cell(row, col++).Value = i.OvInPaidAmt;
                worksheet.Cell(row, col).Value = i.OvInDiffAmt;

                row++;
            }

            var tableRange = worksheet.Range(3, 1, row - 1, headers.Length);
            //var tableRange = worksheet.RangeUsed();
            var table = tableRange.AsTable();

            // You can set the table name and style here if needed
            table.Name = "Table";
            table.ShowAutoFilter = true;
            worksheet.Columns().AdjustToContents();

            DateTime bangkokTime = TimeZoneInfo.ConvertTimeFromUtc(DateTime.UtcNow, bangkokTimeZone);
            string currentTime = bangkokTime.ToString("dd/MM/yyyy HH:mm", CultureInfo.InvariantCulture);
            string username = User.Claims.FirstOrDefault(c => c.Type == "USERNAME")?.Value;
            String head = string.Format("รายงาน{0} และวันที่ {1} ถึงวันที่ {2} ,วันที่พิมพ์ {3} จำนวน {4} รายการ export by {5}",
                sheetName, data.StartPolicyIssueDate, data.EndPolicyIssueDate, currentTime, result.Count, username);
            worksheet.Cell(1, 1).Value = head;

            using var stream = new MemoryStream();
            workbook.SaveAs(stream);
            var content = stream.ToArray();

            return File(
                content,
                "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                $"รายงาน{sheetName}.xlsx");
        }

        [HttpPost("commInOvInClearing/json")]
        public async Task<ActionResult<PremOutCommInOvInReportResult?>> GetCommInOvInClearingReportJson(ArApReportInput data)
        {
            var result = await _arApService.GetCommInOvInClearingReportJson(data);
            if (result.Count == 0)
            {
                return Ok(new PremOutCommInOvInReportResult());
            }
            return Ok(result);
        }

        [HttpPost("commInOvInClearing/excel")]
        public async Task<IActionResult?> GetCommInOvInClearingReportExcel(ArApReportInput data)
        {
            var result = await _arApService.GetCommInOvInClearingReportJson(data);
            if (result.Count == 0)
            {
                return BadRequest("sql result = null");
            }
            using var workbook = new XLWorkbook();
            var sheetName = "ตัดหนี้_CommInOvIn_ตัวตัด";
            var worksheet = workbook.Worksheets.Add(sheetName);

            // Headers
            var headers = new string[]
             {
                 "หมายเลขกรมธรรม์",
                 "หมายเลขสลักหลัง",
                 "หมายเลขใบแจ้งหนี้",
                 "เลขที่งวด",
                 "เลขที่ตัดหนี้ PremIn",
                 "วันที่ตัดหนี้ PremIn",
                 "เลขที่ตัดหนี้ PremOut",
                 "วันที่ตัดหนี้ PremOut",
                 "วันที่เริ่มคุ้มครอง",
                 "วันที่สิ้นสุดคุ้มครอง",
                 "รหัสบริษัทประกัน",
                 "ชื่อบริษัทประกัน",
                 "รหัสผู้เอาประกัน",
                 "ชื่อผู้เอาประกัน",
                 "ประเภทประกัน",
                 "ประเภทย่อยประกัน",
                 "เบี้ยรวม",
                 "อัตราส่วนลด",
                 "มูลค่าส่วนลด",
                 "เบี้ยสุทธิ",
                 "อากร",
                 "ภาษี",
                 "เบี้ยประกันภัยรับรวม",
                 "NetFlag",
                 "อัตราคอมมิชชั่นรับ",
                 "ยอดคอมมิชชั่นรับ",
                 "เลขที่ตัดหนี้ CommIn",
                 "วันที่ตัดหนี้ CommIn",
                 "จำนวนเงิน CommIn ตัดรับ",
                 "จำนวนเงิน CommIn คงเหลือ",
                 "อัตรา OV รับ",
                 "ยอด OV รับ",
                 "เลขที่ตัดหนี้ OvIn",
                 "วันที่ตัดหนี้ OvIn",
                 "จำนวนเงิน OvIn ตัดรับ",
                 "จำนวนเงิน OvIn คงเหลือ",
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
                worksheet.Cell(row, col++).Value = i.PolicyNo;
                worksheet.Cell(row, col++).Value = i.EndorseNo;
                worksheet.Cell(row, col++).Value = i.InvoiceNo;
                worksheet.Cell(row, col++).Value = i.SeqNo;
                worksheet.Cell(row, col++).Value = i.PremInDfRpReferNo;
                worksheet.Cell(row, col++).Value = i.PremInRpRefDate;
                worksheet.Cell(row, col++).Value = i.PremOutDfRpReferNo;
                worksheet.Cell(row, col++).Value = i.PremOutRpRefDate;
                worksheet.Cell(row, col++).Value = i.ActDate;
                worksheet.Cell(row, col++).Value = i.ExpDate;
                worksheet.Cell(row, col++).Value = i.InsurerCode;
                worksheet.Cell(row, col++).Value = i.InsurerName;
                worksheet.Cell(row, col++).Value = i.InsureeCode;
                worksheet.Cell(row, col++).Value = i.InsureeName;
                worksheet.Cell(row, col++).Value = i.Class;
                worksheet.Cell(row, col++).Value = i.SubClass;
                worksheet.Cell(row, col++).Value = i.GrossPrem;
                worksheet.Cell(row, col++).Value = i.SpecDiscRate;
                worksheet.Cell(row, col++).Value = i.SpecDiscAmt;
                worksheet.Cell(row, col++).Value = i.NetGrossPrem;
                worksheet.Cell(row, col++).Value = i.Duty;
                worksheet.Cell(row, col++).Value = i.Tax;
                worksheet.Cell(row, col++).Value = i.TotalPrem;
                worksheet.Cell(row, col++).Value = i.NetFlag;
                worksheet.Cell(row, col++).Value = i.CommInRate;
                worksheet.Cell(row, col++).Value = i.CommInAmt;
                worksheet.Cell(row, col++).Value = i.CommInDfRpReferNo;
                worksheet.Cell(row, col++).Value = i.CommInRpRefDate;
                worksheet.Cell(row, col++).Value = i.CommInPaidAmt;
                worksheet.Cell(row, col++).Value = i.CommInDiffAmt;
                worksheet.Cell(row, col++).Value = i.OvInRate;
                worksheet.Cell(row, col++).Value = i.OvInAmt;
                worksheet.Cell(row, col++).Value = i.OvInDfRpReferNo;
                worksheet.Cell(row, col++).Value = i.OvInRpRefDate;
                worksheet.Cell(row, col++).Value = i.OvInPaidAmt;
                worksheet.Cell(row, col).Value = i.OvInDiffAmt;

                row++;
            }

            var tableRange = worksheet.Range(3, 1, row - 1, headers.Length);
            //var tableRange = worksheet.RangeUsed();
            var table = tableRange.AsTable();

            // You can set the table name and style here if needed
            table.Name = "Table";
            table.ShowAutoFilter = true;
            worksheet.Columns().AdjustToContents();

            DateTime bangkokTime = TimeZoneInfo.ConvertTimeFromUtc(DateTime.UtcNow, bangkokTimeZone);
            string currentTime = bangkokTime.ToString("dd/MM/yyyy HH:mm", CultureInfo.InvariantCulture);
            string username = User.Claims.FirstOrDefault(c => c.Type == "USERNAME")?.Value;
            String head = string.Format("รายงาน{0} และวันที่ {1} ถึงวันที่ {2} ,วันที่พิมพ์ {3} จำนวน {4} รายการ export by {5}",
                sheetName, data.StartPolicyIssueDate, data.EndPolicyIssueDate, currentTime, result.Count, username);
            worksheet.Cell(1, 1).Value = head;

            using var stream = new MemoryStream();
            workbook.SaveAs(stream);
            var content = stream.ToArray();

            return File(
                content,
                "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                $"รายงาน{sheetName}.xlsx");
        }

        [HttpPost("commInOvInOutstanding/json")]
        public async Task<ActionResult<PremOutCommInOvInReportResult?>> GetCommInOvInOutstandingReportJson(ArApReportInput data)
        {
            var result = await _arApService.GetCommInOvInOutstandingReportJson(data);
            if (result.Count == 0)
            {
                return Ok(new PremOutCommInOvInReportResult());
            }
            return Ok(result);
        }

        [HttpPost("commInOvInOutstanding/excel")]
        public async Task<IActionResult?> GetCommInOvInOutstandingReportExcel(ArApReportInput data)
        {
            var result = await _arApService.GetCommInOvInOutstandingReportJson(data);
            if (result.Count == 0)
            {
                return BadRequest("sql result = null");
            }
            using var workbook = new XLWorkbook();
            var sheetName = "ตัดหนี้_CommInOvIn_ตัวคงเหลือ";
            var worksheet = workbook.Worksheets.Add(sheetName);

            // Headers
            var headers = new string[]
             {
                 "หมายเลขกรมธรรม์",
                 "หมายเลขสลักหลัง",
                 "หมายเลขใบแจ้งหนี้",
                 "เลขที่งวด",
                 "เลขที่ตัดหนี้ PremIn",
                 "วันที่ตัดหนี้ PremIn",
                 "เลขที่ตัดหนี้ PremOut",
                 "วันที่ตัดหนี้ PremOut",
                 "วันที่เริ่มคุ้มครอง",
                 "วันที่สิ้นสุดคุ้มครอง",
                 "รหัสบริษัทประกัน",
                 "ชื่อบริษัทประกัน",
                 "รหัสผู้เอาประกัน",
                 "ชื่อผู้เอาประกัน",
                 "ประเภทประกัน",
                 "ประเภทย่อยประกัน",
                 "เบี้ยรวม",
                 "อัตราส่วนลด",
                 "มูลค่าส่วนลด",
                 "เบี้ยสุทธิ",
                 "อากร",
                 "ภาษี",
                 "เบี้ยประกันภัยรับรวม",
                 "NetFlag",
                 "อัตราคอมมิชชั่นรับ",
                 "ยอดคอมมิชชั่นรับ",
                 "เลขที่ตัดหนี้ CommIn",
                 "วันที่ตัดหนี้ CommIn",
                 "จำนวนเงิน CommIn ตัดรับ",
                 "จำนวนเงิน CommIn คงเหลือ",
                 "อัตรา OV รับ",
                 "ยอด OV รับ",
                 "เลขที่ตัดหนี้ OvIn",
                 "วันที่ตัดหนี้ OvIn",
                 "จำนวนเงิน OvIn ตัดรับ",
                 "จำนวนเงิน OvIn คงเหลือ",
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
                worksheet.Cell(row, col++).Value = i.PolicyNo;
                worksheet.Cell(row, col++).Value = i.EndorseNo;
                worksheet.Cell(row, col++).Value = i.InvoiceNo;
                worksheet.Cell(row, col++).Value = i.SeqNo;
                worksheet.Cell(row, col++).Value = i.PremInDfRpReferNo;
                worksheet.Cell(row, col++).Value = i.PremInRpRefDate;
                worksheet.Cell(row, col++).Value = i.PremOutDfRpReferNo;
                worksheet.Cell(row, col++).Value = i.PremOutRpRefDate;
                worksheet.Cell(row, col++).Value = i.ActDate;
                worksheet.Cell(row, col++).Value = i.ExpDate;
                worksheet.Cell(row, col++).Value = i.InsurerCode;
                worksheet.Cell(row, col++).Value = i.InsurerName;
                worksheet.Cell(row, col++).Value = i.InsureeCode;
                worksheet.Cell(row, col++).Value = i.InsureeName;
                worksheet.Cell(row, col++).Value = i.Class;
                worksheet.Cell(row, col++).Value = i.SubClass;
                worksheet.Cell(row, col++).Value = i.GrossPrem;
                worksheet.Cell(row, col++).Value = i.SpecDiscRate;
                worksheet.Cell(row, col++).Value = i.SpecDiscAmt;
                worksheet.Cell(row, col++).Value = i.NetGrossPrem;
                worksheet.Cell(row, col++).Value = i.Duty;
                worksheet.Cell(row, col++).Value = i.Tax;
                worksheet.Cell(row, col++).Value = i.TotalPrem;
                worksheet.Cell(row, col++).Value = i.NetFlag;
                worksheet.Cell(row, col++).Value = i.CommInRate;
                worksheet.Cell(row, col++).Value = i.CommInAmt;
                worksheet.Cell(row, col++).Value = i.CommInDfRpReferNo;
                worksheet.Cell(row, col++).Value = i.CommInRpRefDate;
                worksheet.Cell(row, col++).Value = i.CommInPaidAmt;
                worksheet.Cell(row, col++).Value = i.CommInDiffAmt;
                worksheet.Cell(row, col++).Value = i.OvInRate;
                worksheet.Cell(row, col++).Value = i.OvInAmt;
                worksheet.Cell(row, col++).Value = i.OvInDfRpReferNo;
                worksheet.Cell(row, col++).Value = i.OvInRpRefDate;
                worksheet.Cell(row, col++).Value = i.OvInPaidAmt;
                worksheet.Cell(row, col).Value = i.OvInDiffAmt;

                row++;
            }

            var tableRange = worksheet.Range(3, 1, row - 1, headers.Length);
            //var tableRange = worksheet.RangeUsed();
            var table = tableRange.AsTable();

            // You can set the table name and style here if needed
            table.Name = "Table";
            table.ShowAutoFilter = true;
            worksheet.Columns().AdjustToContents();

            DateTime bangkokTime = TimeZoneInfo.ConvertTimeFromUtc(DateTime.UtcNow, bangkokTimeZone);
            string currentTime = bangkokTime.ToString("dd/MM/yyyy HH:mm", CultureInfo.InvariantCulture);
            string username = User.Claims.FirstOrDefault(c => c.Type == "USERNAME")?.Value;
            String head = string.Format("รายงาน{0} และวันที่ {1} ถึงวันที่ {2} ,วันที่พิมพ์ {3} จำนวน {4} รายการ export by {5}",
                sheetName, data.StartPolicyIssueDate, data.EndPolicyIssueDate, currentTime, result.Count, username);
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
