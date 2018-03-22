using Newtonsoft.Json;
using System;
using System.Collections;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Net;
using System.Net.Http;
using System.Text;
using System.Threading.Tasks;
using System.Web;
using System.Web.Mvc;
using System.Web.Script.Serialization;

namespace ReportingServices.Controllers
{
    public class SampleReportController : Controller
    {

        // GET: SampleReport
        public async Task<ActionResult> Index()
        {
            ViewBag.Result = await GetByIDforAPI();
            //ViewBag.Result = GetByIDforAPI2();
            return View();
        }

        [HttpPost]
        public async Task<ActionResult> Print()
        {
            string result = await GetByIDforAPI();
            Debug.WriteLine("No3");
            Debug.WriteLine(result);

            return View();

        }

        private async Task<string> GetByIDforAPI(string id = "98")
        {
            string uri = Parameters.Pleasanter.Uri.Replace("{id}", id);
            Dictionary<string, string> param = new Dictionary<string, string>()
            {
                [nameof(Parameters.Pleasanter.ApiKey)] = Parameters.Pleasanter.ApiKey
            };

            using (HttpClient httpClient = new HttpClient())
            {
                var json = new JavaScriptSerializer().Serialize(param);

                var content = new StringContent(json, Encoding.UTF8, "application/x-www-form-urlencoded");
                var response = await httpClient.PostAsync(uri, content);

                Debug.WriteLine("No1");
                Debug.WriteLine(response);
                var result = await response.Content.ReadAsStringAsync();
                Debug.WriteLine("No2");
                Debug.WriteLine(result);
                return result;
            }
        }
    }

}
