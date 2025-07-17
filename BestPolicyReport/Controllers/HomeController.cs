using FastReport.Web;
using FastReport.Data;
using FastReport.Export.PdfSimple;
using Microsoft.AspNetCore.Http;
using Microsoft.AspNetCore.Mvc;

namespace BestPolicyReport.Controllers
{
    [Route("api/[controller]")]
    [ApiController]
    public class HomeController : ControllerBase
    {
        private readonly IWebHostEnvironment _webHostEnvironment;
        public HomeController(IWebHostEnvironment webHostEnvironment) 
        {
            this._webHostEnvironment = webHostEnvironment;
        }
        [HttpPost("json")]
        public async Task<IActionResult?> Index()
        {
            WebReport web = new WebReport();
            var path = $"{this._webHostEnvironment.WebRootPath}\\Reports\\FastReport.frx";
            web.Report.Load(path);

            web.Report.Prepare();
            Stream stream = new MemoryStream();
            web.Report.Export(new PDFSimpleExport(), stream);
            stream.Position = 0;
            return File(stream, "application/zip", "report.pdf");

            //return View(web);
        }

    }
}
