using Microsoft.AspNetCore.Mvc;
using PCReports.Services;
using System.Globalization;
using log4net;
using Newtonsoft.Json;


namespace PCReports.Controllers
{
    [ApiController]
    public class GenerateFullReport : ControllerBase
    {
        private readonly IConfiguration _configuration;
        static readonly ILog log = LogManager.GetLogger(nameof(GenerateFullReport));

        public GenerateFullReport(IConfiguration configuration)
        {
            _configuration = configuration;
        }

        [HttpPost]
        [Route("api/GenerateFullReport/Facade")]
        public async Task<IActionResult> GenerateFacadeFullReport([FromBody] Models.Facade.PDFResult pdfResult)
        {
            try
            {
                CultureInfo cultureInfo = new CultureInfo(pdfResult.Language);
                Thread.CurrentThread.CurrentCulture = cultureInfo;

                Services.Facade.PdfService pdfServvice = new Services.Facade.PdfService();
                string pdfFilePath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, $"Content\\structural-result\\{Guid.NewGuid()}.pdf");
                string localReportPath = pdfServvice.GenerateReport(pdfFilePath, pdfResult);

                string projectGuid = pdfResult.ProjectGuid;
                string problemGuid = pdfResult.ProblemGuid;
                string reportUrl = projectGuid + @"/" + problemGuid + @"/structural_report.pdf";

                if (!string.IsNullOrEmpty(localReportPath) && !string.IsNullOrEmpty(projectGuid) && !string.IsNullOrEmpty(problemGuid))
                {
                    AmazonS3Service AWSS3Service = new AmazonS3Service(_configuration);
                    await AWSS3Service.uploadFileToS3Async(localReportPath, reportUrl);
                    System.IO.File.Delete(localReportPath);
                    return Ok();
                }
                else
                {
                    return BadRequest();
                }
            }
            catch (Exception ex)
            {
                //string jsonResult = JsonConvert.SerializeObject(pdfResult);
                //log.Error($" Facade Full Report Generation\n  Project:       {pdfResult.ProjectGuid}      Problem:  {pdfResult.ProblemGuid}    \n  Message:  {ex.Message} \n{ex.StackTrace} \n pdfResult: \n {jsonResult} \n\n");
                log.Error($" Facade Full Report Generation\n  Project:       {pdfResult.ProjectGuid}      Problem:  {pdfResult.ProblemGuid}    \n  Message:  {ex.Message} \n{ex.StackTrace} \n");
                throw;
            }
        }

        [HttpPost]
        [Route("api/GenerateFullReport/UDC")]
        public async Task<IActionResult> GenerateUDCFullReport([FromBody] Models.UDC.PDFResult pdfResult)
        {
            try
            {
                var cultureInfo = new CultureInfo(pdfResult.Language);
                Thread.CurrentThread.CurrentCulture = cultureInfo;

                Services.UDC.PdfService pdfServvice = new Services.UDC.PdfService();
                string pdfFilePath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, $"Content\\structural-result\\{Guid.NewGuid()}.pdf");
                string localReportPath = pdfServvice.GenerateReport(pdfFilePath, pdfResult);

                string projectGuid = pdfResult.ProjectGuid;
                string problemGuid = pdfResult.ProblemGuid;
                string reportUrl = projectGuid + @"/" + problemGuid + @"/structural_report.pdf";

                if (!string.IsNullOrEmpty(localReportPath) && !string.IsNullOrEmpty(projectGuid) && !string.IsNullOrEmpty(problemGuid))
                {
                    AmazonS3Service AWSS3Service = new AmazonS3Service(_configuration);
                    await AWSS3Service.uploadFileToS3Async(localReportPath, reportUrl);
                    System.IO.File.Delete(localReportPath);
                    return Ok();
                }
                else
                {
                    return BadRequest();
                }
            }
            catch (Exception ex)
            {
                //string jsonResult = JsonConvert.SerializeObject(pdfResult);
                //log.Error($" Facade Full Report Generation\n  Project:       {pdfResult.ProjectGuid}      Problem:  {pdfResult.ProblemGuid}    \n  Message:  {ex.Message} \n{ex.StackTrace} \n pdfResult: \n {jsonResult} \n\n");
                log.Error($" Facade Full Report Generation\n  Project:       {pdfResult.ProjectGuid}      Problem:  {pdfResult.ProblemGuid}    \n  Message:  {ex.Message} \n{ex.StackTrace} \n");
                throw;
            }

        }

        //[HttpPost]
        //[Route("api/GenerateFullReport/Window")]
        //public async Task<IActionResult> GenerateWindowFullReport([FromBody] Models.Window.PDFResult pdfResult)
        //{
        //    var cultureInfo = new CultureInfo(pdfResult.Language);
        //    Thread.CurrentThread.CurrentCulture = cultureInfo;

        //    Services.Window.PdfService pdfServvice = new Services.Window.PdfService();
        //    string pdfFilePath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, $"Content\\structural-result\\{Guid.NewGuid()}.pdf");
        //    string localReportPath = pdfServvice.GenerateReport(pdfFilePath, pdfResult);

        //    string projectGuid = pdfResult.ProjectGuid;
        //    string problemGuid = pdfResult.ProblemGuid;
        //    string reportUrl = projectGuid + @"/" + problemGuid + @"/structural_report.pdf";

        //    if (!string.IsNullOrEmpty(localReportPath) && !string.IsNullOrEmpty(projectGuid) && !string.IsNullOrEmpty(problemGuid))
        //    {
        //        AmazonS3Service AWSS3Service = new AmazonS3Service(_configuration);
        //        await AWSS3Service.uploadFileToS3Async(localReportPath, reportUrl);
        //        System.IO.File.Delete(localReportPath);
        //        return Ok();
        //    }
        //    else
        //    {
        //        return BadRequest();
        //    }
        //}

        //[HttpPost]
        //[Route("api/GenerateFullReport/ADS")]
        //public async Task<IActionResult> GenerateADSFullReport([FromBody] Models.ADS.PDFResult pdfResult)
        //{
        //    var cultureInfo = new CultureInfo(pdfResult.Language);
        //    Thread.CurrentThread.CurrentCulture = cultureInfo;

        //    Services.ADS.PdfService pdfServvice = new Services.ADS.PdfService();
        //    string pdfFilePath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, $"Content\\structural-result\\{Guid.NewGuid()}.pdf");
        //    string localReportPath = pdfServvice.GenerateReport(pdfFilePath, pdfResult);

        //    string projectGuid = pdfResult.ProjectGuid;
        //    string problemGuid = pdfResult.ProblemGuid;
        //    string reportUrl = projectGuid + @"/" + problemGuid + @"/structural_report.pdf";

        //    if (!string.IsNullOrEmpty(localReportPath) && !string.IsNullOrEmpty(projectGuid) && !string.IsNullOrEmpty(problemGuid))
        //    {
        //        AmazonS3Service AWSS3Service = new AmazonS3Service(_configuration);
        //        await AWSS3Service.uploadFileToS3Async(localReportPath, reportUrl);
        //        System.IO.File.Delete(localReportPath);
        //        return Ok();
        //    }
        //    else
        //    {
        //        return BadRequest();
        //    }
        //}
    }
}