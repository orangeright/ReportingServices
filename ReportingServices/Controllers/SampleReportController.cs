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
        //private string baseUri = "http://39.110.246.5/pleasanter/api_items/33/get";
        //private Hashtable param = new Hashtable
        //{
        //    ["ApiKey"] = "09ca5b781f694e4f7fae9299d9a62998cc1bcefa262e11be8299496491a96e4be857c0e838fa5365ee30c8266b96f2b1b1982a3aff88b9dbd2ea1c088f64766e"
        //};

        private string baseUri = "http://localhost/pleasanter/api_items/98/get";
        //private Hashtable param = new Hashtable
        //{
        //    ["ApiKey"] = "40a398c2c4ab22bda378b3650f6cb7ea72eeafa789dd9ef7c4dcbc3f056052db8e0d08c49d0efd3f5d028da5f5b7f36788d5836bc50783f5b1a1844124434496"
        //};
        private Dictionary<string, string> param = new Dictionary<string, string>();
        

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

        private async Task<string> GetByIDforAPI()
        {
            string uri = baseUri;
            using (HttpClient httpClient = new HttpClient())
            {
                param["ApiKey"] = "40a398c2c4ab22bda378b3650f6cb7ea72eeafa789dd9ef7c4dcbc3f056052db8e0d08c49d0efd3f5d028da5f5b7f36788d5836bc50783f5b1a1844124434496";
                var json = new JavaScriptSerializer().Serialize(param);

                //var json = "{\"ApiKey\":\"40a398c2c4ab22bda378b3650f6cb7ea72eeafa789dd9ef7c4dcbc3f056052db8e0d08c49d0efd3f5d028da5f5b7f36788d5836bc50783f5b1a1844124434496\"}";
                using (var client = new HttpClient())
                {
                    var content = new StringContent(json, Encoding.UTF8, "application/x-www-form-urlencoded");
                    //var content = new StringContent(json, Encoding.UTF8, "application/json");
                    //var content = new StringContent(json, Encoding.GetEncoding("utf-8"));

                    var response = await client.PostAsync(uri, content);

                    Debug.WriteLine("No1");
                    Debug.WriteLine(response);
                    var result = await response.Content.ReadAsStringAsync();
                    Debug.WriteLine("No2");
                    Debug.WriteLine(result);
                    return result;
                }
            }
        }

        private string GetByIDforAPI2()
        {
            HttpWebRequest req = (HttpWebRequest)WebRequest.Create(baseUri);
            req.ContentType = "application/json";
            req.Method = "POST";

            using (var streamwriter = new StreamWriter(req.GetRequestStream()))
            {
                string jsonPayload = new JavaScriptSerializer().Serialize(param);
                streamwriter.Write(jsonPayload);
            }

            HttpWebResponse res = (HttpWebResponse)req.GetResponse();

            using (res)
            {
                using (var resStream = res.GetResponseStream())
                {
                    StreamReader sr = new StreamReader(resStream, Encoding.UTF8);
                    string result = sr.ReadToEnd();
                    Debug.WriteLine("No11");
                    Debug.WriteLine(result);
                    return result;
                }
            }
        }

    }

}
