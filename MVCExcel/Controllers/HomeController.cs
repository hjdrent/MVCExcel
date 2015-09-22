using System;
using System.IO;
using System.Net;
using System.Web.Mvc;
using System.Web.WebPages;
using ClosedXML.Excel;
using MVCExcel.Models;

namespace MVCExcel.Controllers
{
    public class HomeController : Controller
    {
        public ActionResult Index()
        {
            AutoModel autoModel = new AutoModel()
            {
                AantalWielen = 4,
                Naam = "Bentley"
            };
            return View(autoModel);
        }

        [HttpPost]
        public ActionResult Index(AutoModel model)
        {
            return View(model);
        }

        public ActionResult About(AutoModel model)
        {
            return View(model);
        }

        public ActionResult Contact(AutoModel model)
        {
            return View(model);
        }

        #region Excel Rapportage
        [HttpPost]
        public ActionResult CreateStandardReport(AutoModel model)
        {
            string filename = "testje.xlsx";

            XLWorkbook workbook = new XLWorkbook();
            IXLWorksheet worksheet = workbook.Worksheets.Add("Scheet");

            worksheet.ActiveCell = worksheet.Cell(string.Format("A1"));
            worksheet.ActiveCell.Value = model.Naam;
            worksheet.ActiveCell = worksheet.Cell(string.Format("B1"));
            worksheet.ActiveCell.Value = model.AantalWielen;

            //worksheet.Columns().AdjustToContents();
            worksheet.ActiveCell = worksheet.Cell("A1");

            try
            {
                DownloadExcel(workbook, filename);
                model.Message = "Succes!";
                return View(model);
            }
            catch (Exception ex)
            {
                model.Message = ex.Message;
                return View("About", model);
            }
        }

        private void DownloadExcel(XLWorkbook workbook, string ReportName)
        {
            string reportName = Server.UrlEncode(ReportName);
            var response = Response;
            response.AddHeader("content-disposition", "attachment; filename=" + reportName);
            response.ContentType = "vnd.openxmlformats-officedocument.spreadsheetml.sheet";
            using (var stream = new MemoryStream())
            {
                workbook.SaveAs(stream);
                stream.Position = 0;
                stream.WriteTo(response.OutputStream);
                stream.Close();
            }
            response.End();
        }
        #endregion

    }
}