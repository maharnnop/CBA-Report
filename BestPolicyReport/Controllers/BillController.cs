using AspNetCore.Reporting;
using BestPolicyReport.Models.BillReport;
using BestPolicyReport.Services.BillService;
using BoldReports.RDL.DOM;
using ClosedXML;
using ClosedXML.Excel;
using DocumentFormat.OpenXml.Spreadsheet;
using FastReport;
using FastReport.Export.PdfSimple;
using FastReport.Web;
using Microsoft.AspNetCore.Authorization;
using Microsoft.AspNetCore.Hosting;
using Microsoft.AspNetCore.Mvc;
using Microsoft.VisualBasic;
using System.Globalization;
using System.Linq;
using System.Security.Claims;
using System.IdentityModel.Tokens.Jwt;
using Newtonsoft.Json.Linq;

namespace BestPolicyReport.Controllers
{

    [Authorize]
    [Route("api/[controller]")]
    [ApiController]
    public class BillController : ControllerBase
    {
      
        private TimeZoneInfo bangkokTimeZone = TimeZoneInfo.FindSystemTimeZoneById("Asia/Bangkok");
        private readonly IBillService _billService;
        private readonly IWebHostEnvironment _webHostEnvironment;

        public BillController(IBillService billService, IWebHostEnvironment webHostEnvironment)
        {
            _billService = billService;
            this._webHostEnvironment = webHostEnvironment;
            System.Text.Encoding.RegisterProvider(System.Text.CodePagesEncodingProvider.Instance);
    
        }

        [HttpPost("json")]
        public async Task<ActionResult<BillReportResult?>> GetBillReportJson(BillReportInput data)
        {
            var result = await _billService.GetBillReportJson(data);
            if (result.Count == 0)
            {
                return Ok(new BillReportResult());
            }
            return Ok(result);
        }

        [HttpPost("excel")]
        public async Task<IActionResult?> GetBillReportExcel(BillReportInput data)
        {
            var result = await _billService.GetBillReportJson(data);
            if (result.Count == 0)
            {
                return BadRequest("sql result = null");
            }
            using var workbook = new XLWorkbook();
            var sheetName = "ใบวางบิล";
            var worksheet = workbook.Worksheets.Add(sheetName);

            // Headers
            var headers = new string[]
            {
                 "เลขที่กรมธรรม์", "เลขที่สลักหลัง",
                "วันที่ทำสัญญา", "วันที่เริ่มต้นความคุ้มครอง", "วันที่สิ้นสุดความคุ้มครอง",
                "เลขที่ใบแจ้งหนี้", "เลขที่ใบกำกับภาษี", "ชื่อใบกำกับภาษี",
                "บริษัทประกัน", "Class", "SubClass", 
                "รหัสผู้แนะนำ 1", "รหัสผู้แนะนำ 2", "วันที่กำหนดชำระ", 
                "เลขที่งวด",
                "เบี้ยสุทธิ", "อากร", "ภาษี", "เบี้ยรวม",  "มูลค่าส่วนลด",
                "รหัสผู้เอาประกัน", "ชื่อผู้เอาประกัน", "ผู้รับผลประโยชน์",
                "รหัสรถ (V)", "รหัสรถ (C)", "ป้ายแดง", "ทะเบียนรถ", "จังหวัดที่จดทะเบียน", 
                "เลขตัวถัง", "เลขเครื่องยนต์", "ปีที่จดทะเบียน", "ยี่ห้อ", "รุ่น", "รุ่นย่อย", "ซีซี", "ที่นั่ง", "น้ำหนัก", "อุปกรณ์ตกแต่งเพิ่มเติม",


                "อัตราคอมมิชชั่นจ่ายของผู้แนะนำ 1", 
                "ยอดคอมมิชชั่นจ่ายของผู้แนะนำ 1", 
                "อัตรา OV จ่ายของผู้แนะนำ 1", 
                "ยอด OV จ่ายของผู้แนะนำ 1",
                "อัตราคอมมิชชั่นจ่ายของผู้แนะนำ 2", 
                "ยอดคอมมิชชั่นจ่ายของผู้แนะนำ 2", 
                "อัตรา OV จ่ายของผู้แนะนำ 2", 
                "ยอด OV จ่ายของผู้แนะนำ 2",
                "อัตราคอมมิชชั่นจ่าย", 
                "ยอดคอมมิชชั่นจ่าย", 
                "อัตรา OV จ่าย", 
                "ยอด OV จ่าย", 
                "netFlag", 
                "billPremium"
            };

            for (int col = 1; col <= headers.Length; col++)
            {
                worksheet.Cell(1, col).Value = headers[col - 1];
            }

            // Data
            int row = 2;
            foreach (var i in result.Datas)
            {
                int col = 1;
                worksheet.Cell(row, col++).Value = i.PolicyNo;
                worksheet.Cell(row, col++).Value = i.EndorseNo;
                worksheet.Cell(row, col++).Value = i.IssueDate;
                worksheet.Cell(row, col++).Value = i.ActDate;
                worksheet.Cell(row, col++).Value = i.ExpDate;
                worksheet.Cell(row, col++).Value = i.InvoiceNo;
                worksheet.Cell(row, col++).Value = i.TaxInvoiceNo;
                worksheet.Cell(row, col++).Value = i.InvoiceName;
                worksheet.Cell(row, col++).Value = i.InsurerCode;
                worksheet.Cell(row, col++).Value = i.Class;
                worksheet.Cell(row, col++).Value = i.SubClass;
                worksheet.Cell(row, col++).Value = i.AgentCode1;
                worksheet.Cell(row, col++).Value = i.AgentCode2;
                worksheet.Cell(row, col++).Value = i.DueDate;
                worksheet.Cell(row, col++).Value = i.SeqNo;
                worksheet.Cell(row, col++).Value = i.NetGrossPrem;
                worksheet.Cell(row, col++).Value = i.Duty;
                worksheet.Cell(row, col++).Value = i.Tax;
                worksheet.Cell(row, col++).Value = i.TotalPrem;
                worksheet.Cell(row, col++).Value = i.SpecDiscAmt;
                worksheet.Cell(row, col++).Value = i.InsureeCode;
                worksheet.Cell(row, col++).Value = i.InsureeName;
                worksheet.Cell(row, col++).Value = i.Beneficiary;

                worksheet.Cell(row, col++).Value = i.VoluntaryCode;
                worksheet.Cell(row, col++).Value = i.CompulsoryCode;
                worksheet.Cell(row, col++).Value = i.Unregisterflag;
                worksheet.Cell(row, col++).Value = i.LicenseNo;
                worksheet.Cell(row, col++).Value = i.Province;
                worksheet.Cell(row, col++).Value = i.ChassisNo;
                worksheet.Cell(row, col++).Value = i.EngineNo;
                worksheet.Cell(row, col++).Value = i.ModelYear;
                worksheet.Cell(row, col++).Value = i.Brand;
                worksheet.Cell(row, col++).Value = i.Model;
                worksheet.Cell(row, col++).Value = i.Specname;
                worksheet.Cell(row, col++).Value = i.Cc;
                worksheet.Cell(row, col++).Value = i.Seat;
                worksheet.Cell(row, col++).Value = i.Gvw;
                worksheet.Cell(row, col++).Value = i.Addition_access;

                worksheet.Cell(row, col++).Value = i.CommOutRate1;
                worksheet.Cell(row, col++).Value = i.CommOutAmt1;
                worksheet.Cell(row, col++).Value = i.OvOutRate1;
                worksheet.Cell(row, col++).Value = i.OvOutAmt1;
                worksheet.Cell(row, col++).Value = i.CommOutRate2;
                worksheet.Cell(row, col++).Value = i.CommOutAmt2;
                worksheet.Cell(row, col++).Value = i.OvOutRate2;
                worksheet.Cell(row, col++).Value = i.OvOutAmt2;
                worksheet.Cell(row, col++).Value = i.CommOutRate;
                worksheet.Cell(row, col++).Value = i.CommOutAmt;
                worksheet.Cell(row, col++).Value = i.OvOutRate;
                worksheet.Cell(row, col++).Value = i.OvOutAmt;
                worksheet.Cell(row, col++).Value = i.NetFlag;
                worksheet.Cell(row, col++).Value = i.BillPremium;
                row++;
            }

            var tableRange = worksheet.RangeUsed();
            var table = tableRange.AsTable();
            table.Name = "Table";
            table.ShowAutoFilter = true;
            worksheet.Columns().AdjustToContents();

            using var stream = new MemoryStream();
            workbook.SaveAs(stream);
            var content = stream.ToArray();

            return File(
                content,
                "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                $"รายงาน{sheetName}.xlsx");
        }

        [HttpPost("policyGroup/json")]
        public async Task<ActionResult<PolicyGroupBillReportResult?>> GetPolicyGroupBillReportJson(PolicyGroupBillReportInput data)
        {
            var result = await _billService.GetPolicyGroupBillReportJson(data);
            if (result.Count == 0)
            {
                return Ok(new PolicyGroupBillReportResult());
            }
            return Ok(result);
        }

        [HttpPost("policyGroup/excel")]
        public async Task<IActionResult?> GetPolicyGroupBillReportExcel(PolicyGroupBillReportInput data)
        {
            var result = await _billService.GetPolicyGroupBillReportJson(data);
            if (result.Count == 0)
            {
                return BadRequest("sql result = null");
            }
            using var workbook = new XLWorkbook();
            var sheetName = "ตารางแนบใบวางบิล";
            var worksheet = workbook.Worksheets.Add(sheetName);

            // Headers
            var headers = new string[]
            {
                "no.",
                "ทะเบียนรถ",
                "ยี่ห้อ",
                "รุ่น",
                "ผู้เอาประกัน",
                "ปี",
                "เลขถังรถ",
                "ทุนประกันภัย",
                "เลขที่กรมธรรม์",
                "วันที่เริ่มต้น",
                "วันที่สิ้นสุด",
                "เบี้ยประกัน",
                "อากร",
                "ก่อน Vat",
                "Vat 7%",
                "รวมเป็นเงิน",
                "หัก 1%",
                "ยอดจ่ายสุทธิ"
            };

            for (int col = 1; col <= headers.Length; col++)
            {
                worksheet.Cell(3, col).Value = headers[col - 1];
            }

            // Data
            int row = 4;
            int count = 1;
            foreach (var i in result)
            {
                int col = 1;
                worksheet.Cell(row, col++).Value = count++;
                worksheet.Cell(row, col++).Value = i.LicenseNo;
                worksheet.Cell(row, col++).Value = i.Brand;
                worksheet.Cell(row, col++).Value = i.Model;
                worksheet.Cell(row, col++).Value = i.InsureeName;
                worksheet.Cell(row, col++).Value = i.ModelYear;
                worksheet.Cell(row, col++).Value = i.ChassisNo;
                worksheet.Cell(row, col++).Value = i.CoverAmt;
                worksheet.Cell(row, col++).Value = i.PolicyNo;
                worksheet.Cell(row, col++).Value = i.ActDate;
                worksheet.Cell(row, col++).Value = i.ExpDate;
                worksheet.Cell(row, col++).Value = i.NetGrossPrem;
                worksheet.Cell(row, col++).Value = i.Duty;
                worksheet.Cell(row, col++).Value = i.NetGrossPremBeforeTax;
                worksheet.Cell(row, col++).Value = i.Tax;
                worksheet.Cell(row, col++).Value = i.TotalPrem;
                worksheet.Cell(row, col++).Value = i.WithHeld;
                worksheet.Cell(row, col++).Value = i.BillPremium;
                row++;
            }

            var tableRange = worksheet.Range(3, 1, row - 1, headers.Length);
            var table = tableRange.AsTable();
            table.Name = "Table";
            table.ShowAutoFilter = true;

            worksheet.Cell(row, 12).FormulaA1 = $"SUM(L2:L{row - 1})"; // NetGrossPrem
            worksheet.Cell(row, 13).FormulaA1 = $"SUM(M2:M{row - 1})"; // Duty
            worksheet.Cell(row, 14).FormulaA1 = $"SUM(N2:N{row - 1})"; // NetGrossPremBeforeTax
            worksheet.Cell(row, 15).FormulaA1 = $"SUM(O2:O{row - 1})"; // Tax
            worksheet.Cell(row, 16).FormulaA1 = $"SUM(P2:P{row - 1})"; // TotalPrem
            worksheet.Cell(row, 17).FormulaA1 = $"SUM(Q2:Q{row - 1})"; // WithHeld
            worksheet.Cell(row, 18).FormulaA1 = $"SUM(R2:R{row - 1})"; // BillPremium
            worksheet.Row(row).Style.Font.Bold = true;
            worksheet.Columns().AdjustToContents();

            DateTime bangkokTime = TimeZoneInfo.ConvertTimeFromUtc(DateTime.UtcNow, bangkokTimeZone);
            string currentTime = bangkokTime.ToString("dd/MM/yyyy HH:mm", CultureInfo.InvariantCulture);
            string username = User.Claims.FirstOrDefault(c => c.Type == "USERNAME")?.Value;
            String head = string.Format("รายงาน{0} และวันที่ {1} ถึงวันที่ {2} ,วันที่พิมพ์ {3} จำนวน {4} รายการ export by {5}",
                sheetName, data.StartBillDate, data.EndBillDate, currentTime, result.Count, username);
            worksheet.Cell(1, 1).Value = head;




            using var stream = new MemoryStream();
            workbook.SaveAs(stream);
            var content = stream.ToArray();

            return File(
                content,
                "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                $"รายงาน{sheetName}.xlsx");
        }
        
        [HttpPost("invoicelist/json")]
        public async Task<ActionResult<InvoiceReportResult?>> GetInvoiceReportJson(InvoiceReportInput data)
        {
            //var token = Request.Headers.Authorization Microsoft.AspNetCore.Http.IHeaderDictionary.Authorization   "Bearer eyJhbGciOiJIUzI1NiIsInR5cCI6IkpXVCJ9.eyJVU0VSTkFNRSI6ImFkbWluayIsIlJPTEUiOiJhZG1pbiIsImlhdCI6MTc0NTA4MDM2MCwiZXhwIjoxNzQ1MTA5MTYwfQ.QZ38iGkBuenihYohYvxu0XJbjbiPH75xpllpivhLbjc"    Microsoft.Extensions.Primitives.StringValues

            

            var result = await _billService.GetInvoiceJson(data);
            if (result.Count == 0)
            {
                return Ok(new InvoiceReportResult());
            }
            return Ok(result);
        }

        [HttpPost("invoicelist/excel")]
        public async Task<IActionResult?> GetInvoiceReportJsonExcel(InvoiceReportInput data)
        {
            var result = await _billService.GetInvoiceJson(data);
            if (result.Count == 0)
            {
                return BadRequest("sql result = null");
            }
            using var workbook = new XLWorkbook();
            var sheetName = "บันทึกใบแจ้งหนี้";
            var worksheet = workbook.Worksheets.Add(sheetName);

            // Headers
            var headers = new string[]
            {
                "no.",
                "เลขที่ใบคำขอ",
                "เลขที่กรมธรรม์",
                "เลขที่สลักหลัง",
                "ผู้เอาประกัน",
                "เลขที่ใบแจ้งหนี้อะมิตี้",
                "งวดที่",
                "เบี้ยรวม",
                "ภาษีหัก ณ ที่จ่าย 1%",
                "ส่วนลดลูกค้า",
                "เบี้ยรับจากลูกค้า",
                "วันที่ครบกำหนดชำระ",
                "รหัสผู้แนะนำ",
                "ชื่อผู้แนะนำ",
            };

            for (int col = 1; col <= headers.Length; col++)
            {
                worksheet.Cell(3, col).Value = headers[col - 1];
            }

            // Data
            int row = 4;
            int count = 1;
            foreach (var i in result.Datas)
            {
                int col = 1;
                worksheet.Cell(row, col++).Value = count++;
                worksheet.Cell(row, col++).Value = i.ApplicationNo;
                worksheet.Cell(row, col++).Value = i.PolicyNo;
                worksheet.Cell(row, col++).Value = i.EndorseNo;
                worksheet.Cell(row, col++).Value = i.InsureeName;
                worksheet.Cell(row, col++).Value = i.InvoiceNo;
                worksheet.Cell(row, col++).Value = i.SeqNo;
                worksheet.Cell(row, col++).Value = i.Totalprem;
                worksheet.Cell(row, col++).Value = i.Withheld;
                worksheet.Cell(row, col++).Value = i.Specdiscamt;
                worksheet.Cell(row, col++).Value = i.Totalamt;
                worksheet.Cell(row, col++).Value = i.DueDate;
                worksheet.Cell(row, col++).Value = i.AgentCode;
                worksheet.Cell(row, col++).Value = i.Agentname;
                row++;
            }

            var tableRange = worksheet.Range(3, 1, row - 1, headers.Length);
            var table = tableRange.AsTable();
            table.Name = "Table";
            table.ShowAutoFilter = true;
            worksheet.Columns().AdjustToContents();
            //worksheet.Cell(row, 12).FormulaA1 = $"SUM(L2:L{row - 1})"; // NetGrossPrem
            //worksheet.Cell(row, 13).FormulaA1 = $"SUM(M2:M{row - 1})"; // Duty
            //worksheet.Cell(row, 14).FormulaA1 = $"SUM(N2:N{row - 1})"; // NetGrossPremBeforeTax
            //worksheet.Cell(row, 15).FormulaA1 = $"SUM(O2:O{row - 1})"; // Tax
            //worksheet.Cell(row, 16).FormulaA1 = $"SUM(P2:P{row - 1})"; // TotalPrem
            //worksheet.Cell(row, 17).FormulaA1 = $"SUM(Q2:Q{row - 1})"; // WithHeld
            //worksheet.Cell(row, 18).FormulaA1 = $"SUM(R2:R{row - 1})"; // BillPremium
            //worksheet.Row(row).Style.Font.Bold = true;
            DateTime bangkokTime = TimeZoneInfo.ConvertTimeFromUtc(DateTime.UtcNow, bangkokTimeZone);
            string currentTime = bangkokTime.ToString("dd/MM/yyyy HH:mm", CultureInfo.InvariantCulture);
            string username = User.Claims.FirstOrDefault(c => c.Type == "USERNAME")?.Value;
            String head = string.Format("รายงาน{0} และวันที่ {1} ถึงวันที่ {2} ,วันที่พิมพ์ {3} จำนวน {4} รายการ export by {5}",
                sheetName, data.StartInvoiceNo, data.EndInvoiceNo, currentTime, result.Count, username);
            worksheet.Cell(1, 1).Value = head;


            using var stream = new MemoryStream();
            workbook.SaveAs(stream);
            var content = stream.ToArray();

            return File(
                content,
                "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                $"รายงาน{sheetName}.xlsx");
        }
        [HttpPost("invoiceReport/pdf")]
        public async Task<IActionResult?> GetInvoiceReportPDF(string invoiceNo)
        {
            var handler = new JwtSecurityTokenHandler();
            string? jwtToken = Request.Headers["Authorization"].ToString().Substring("Bearer ".Length).Trim();
            var jsonToken = handler.ReadToken(jwtToken) as JwtSecurityToken;
            var usernameClaim = jsonToken.Claims.FirstOrDefault(c => c.Type == "USERNAME")?.Value;

            var result = await _billService.GetInvoicerpt(invoiceNo, usernameClaim);
            WebReport web = new WebReport();
            var path = $"{this._webHostEnvironment.WebRootPath}/Reports/InvoiceReport.frx";
            web.Report.Load(path);
            web.Report.SetParameterValue("invoiceNo", result.invoiceNo);
            web.Report.SetParameterValue("dueDate", result.dueDate.ToString("dd/MM/yyyy"));
            web.Report.SetParameterValue("insureeName", result.insureeName);
            web.Report.SetParameterValue("insureeCode", result.insureeCode);
            web.Report.SetParameterValue("insureeLocation", result.insureeLocation);
            web.Report.SetParameterValue("insurerName", result.insurerName);
            web.Report.SetParameterValue("insureName", result.insureName);
            web.Report.SetParameterValue("policyNo", result.policyNo);
            web.Report.SetParameterValue("actDate", result.actDate.ToString("dd/MM/yyyy"));
            web.Report.SetParameterValue("expDate", result.expDate.ToString("dd/MM/yyyy"));
            web.Report.SetParameterValue("endorseNo", result.endorseNo);
            web.Report.SetParameterValue("cover_amt", result.cover_amt);
            web.Report.SetParameterValue("netgrossprem", result.netgrossprem);
            web.Report.SetParameterValue("duty", result.duty);
            web.Report.SetParameterValue("tax", result.tax);
            web.Report.SetParameterValue("totalprem", result.totalprem);
            web.Report.SetParameterValue("specdiscamt", result.specdiscamt);
            web.Report.SetParameterValue("totalamt", result.totalamt);
            web.Report.SetParameterValue("seqamt", result.seqamt);
            web.Report.Prepare();
            Stream stream = new MemoryStream();
            web.Report.Export(new PDFSimpleExport(), stream);
            stream.Position = 0;
            return File(stream, "application/pdf", invoiceNo + ".pdf");
            
            //return View(web);
        }
    }


}
