using Newtonsoft.Json;
using System;
using System.Collections;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
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
        private string baseUri = "https://39.110.246.5/pleasanter/api_items/33/get";
        //private string baseUri = "https://iis-pleasanter.azurewebsites.net/api_items/79/get";

        // GET: SampleReport
        public ActionResult Index()
        {
            var result = GetByIDforAPI();

            return View();
        }

        [HttpPost]
        public async Task<ActionResult> Print()
        {

            string result = await GetByIDforAPI();
            //string j = JsonConvert.
            return View();

        }

        private async Task<string> GetByIDforAPI()
        {
            string uri = baseUri;
            using (HttpClient httpClient = new HttpClient())
            {

                var param = new Hashtable();
                param["apikey"] = "2ff2c1afbe9667021a1fc828ac6307643526cea32ad1ecef28954b4ce61f3a6b194a40d55831a331bb11cc07eaf6d59c071774245bbcbdb464aa91ad8f5754a5";                           // 数値型のパラメータ
                //param["apikey"] = "495df7ed6df3cd77aeea7857cf4e77c39ccd6e6b5460b070c902a4b862f73380449143553ac53cad064886727d0b4f4a21830b0290f2847406937ec360f4a4be";                           // 数値型のパラメータ
                var serializer = new JavaScriptSerializer();

                //var json = "{ \"apikey\" : \"2ff2c1afbe9667021a1fc828ac6307643526cea32ad1ecef28954b4ce61f3a6b194a40d55831a331bb11cc07eaf6d59c071774245bbcbdb464aa91ad8f5754a5\" }";
                var json = serializer.Serialize(param);

                using (var client = new HttpClient())
                {
                    //var content = new StringContent(json, Encoding.UTF8, "application/json");
                    var content = new StringContent(json, Encoding.UTF8);

                    var response =  await client.PostAsync(uri, content);
                    Debug.WriteLine(response);
                    var result = await response.Content.ReadAsStringAsync();
                    Debug.WriteLine(result);
                    return result;
                }
            }
        }
    }
}
