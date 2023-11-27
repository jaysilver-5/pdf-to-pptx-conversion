
using System;
using System.IO;
using Aspose.Pdf;
using System.Web.Http;

namespace PDF_To_PPT.Controllers
{
    [RoutePrefix("api/pdf_to_ppt")]
    public class PDFtoPPTController : ApiController
    {
        [Route("editable_pptx")]
        [HttpGet]
        public IHttpActionResult ConvertToEditablePPTX()
        {
            string directory = System.Web.HttpContext.Current.Server.MapPath("~/TempPdf");

            Aspose.Pdf.Document doc = new Aspose.Pdf.Document(directory + "\\demo.pdf");

            Aspose.Pdf.PptxSaveOptions pptx_save = new Aspose.Pdf.PptxSaveOptions();

            pptx_save.SeparateImages = true;

            pptx_save.CustomProgressHandler = ShowProgressOnConsole;

            doc.Save(directory + "//result.pptx", pptx_save);

            return Ok(new { Message = "Congratulations!!! Your PDF file is converted to PPT" });
        }

        [Route("slides_as_images")]
        [HttpGet]
        public static void ConvertToSlidesAsImages()
        {
            string directory = System.Web.HttpContext.Current.Server.MapPath("~/TempPdf");

            Aspose.Pdf.Document doc = new Aspose.Pdf.Document(directory + "\\demo.pdf");

            Aspose.Pdf.PptxSaveOptions pptx_save = new Aspose.Pdf.PptxSaveOptions();

            pptx_save.SlidesAsImages = true;

            pptx_save.CustomProgressHandler = ShowProgressOnConsole;

            doc.Save(directory + "PDFToPPT_out_.pptx", pptx_save);
        }

        private static void ShowProgressOnConsole(PptxSaveOptions.ProgressEventHandlerInfo eventInfo)
        {
            switch (eventInfo.EventType)
            {
                case ProgressEventType.TotalProgress:
                    Console.WriteLine(String.Format("{0}  - Conversion progress : {1}% .", DateTime.Now.TimeOfDay, eventInfo.Value.ToString()));
                    break;
                case ProgressEventType.ResultPageCreated:
                    Console.WriteLine(String.Format("{0}  - Result page's {1} of {2} layout created.", DateTime.Now.TimeOfDay, eventInfo.Value.ToString(), eventInfo.MaxValue.ToString()));
                    break;
                case ProgressEventType.ResultPageSaved:
                    Console.WriteLine(String.Format("{0}  - Result page {1} of {2} exported.", DateTime.Now.TimeOfDay, eventInfo.Value.ToString(), eventInfo.MaxValue.ToString()));
                    break;
                case ProgressEventType.SourcePageAnalysed:
                    Console.WriteLine(String.Format("{0}  - Source page {1} of {2} analyzed.", DateTime.Now.TimeOfDay, eventInfo.Value.ToString(), eventInfo.MaxValue.ToString()));
                    break;

                default:
                    break;
            }
        }
    }
}
