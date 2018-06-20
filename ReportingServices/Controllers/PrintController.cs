using System;
using System.Collections;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Threading.Tasks;
using System.Web;
using System.Web.Mvc;

using iTextSharp.text;
using iTextSharp.text.pdf;
using Newtonsoft.Json;

using Implem.Pleasanter.Models;


using ReportingServices.Libraries;
using Newtonsoft.Json.Linq;

namespace ReportingServices.Controllers
{
    public class PrintController : Controller
    {

        // GET: Print
        public ActionResult Index()
        {
            return View();
        }

        public async Task<ActionResult> Excel(string template, string id)
        {
            string guid = await getGuid(template);
            string result = await ApiUtilities.GetByIDforAPI(id);
            var jsondata = JsonConvert.DeserializeObject<GetApiModel>(result);
            if (jsondata.StatusCode == 200)
            {
                var data = jsondata.Response.Data[0];
                data.Add("PrintDate", DateTime.Now.ToLongDateString());
                var xlsx = Reports.CreateXls(guid, data);

                //var log = new SysLogModel();
                //log.Finish();

                return File(xlsx, "application / vnd.openxmlformats - officedocument.spreadsheetml.sheet", "result.xlsx");
            }
            else
            {
                return View();
            }
        }

        public async Task<ActionResult> Pdf(string template, string id)
        {
            string result = await ApiUtilities.GetByIDforAPI(id);
            var jsondata = JsonConvert.DeserializeObject<GetApiModel>(result);
            var data = jsondata.Response.Data[0];
            data.Add("PrintDate", DateTime.Now.ToLongDateString());
            var pdf = Reports.CreatePdfResume(data);
            return new FileStreamResult(pdf, "application/pdf");
        }

        private async Task<string> getGuid(string template)
        {
            string result = await ApiUtilities.GetByIDforAPI(template);
            var jsondata = JsonConvert.DeserializeObject<GetApiModel>(result);
            var data = jsondata.Response.Data[0];
            var attachJson = JsonConvert.DeserializeObject<List<Attachments>>(data["AttachmentsA"]);
            return attachJson[0].Guid;

        }

    }
}