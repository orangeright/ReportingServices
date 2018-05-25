using Newtonsoft.Json;
using System.IO;
using System.Text;
using System.Web;

using ReportingServices.Controllers;
using System.Collections.Generic;

namespace ReportingServices
{
    public class Initializer
    {
        public static void Initialize()
        {
            string json;
            string path = Path.Combine(HttpContext.Current.Server.MapPath("./"), "App_Data", "ParamPleasanter.json");
            using (var reader = new StreamReader(path, Encoding.GetEncoding("utf-8")))
            {
                json = reader.ReadToEnd();
            }
            Parameters.Pleasanter = JsonConvert.DeserializeObject<Pleasanter>(json);

            Parameters.Pleasanter.TemplatePath = Path.Combine(HttpContext.Current.Server.MapPath("./"), "App_Data/Template", "template.xlsx");

            //path = Path.Combine(HttpContext.Current.Server.MapPath("./"), "App_Data", "testdata.json");
            //using (var reader = new StreamReader(path, Encoding.GetEncoding("utf-8")))
            //{
            //    json = reader.ReadToEnd();
            //}

            //JsonData.Jdata = JsonConvert.DeserializeObject<Dictionary<string, string>>(json);
        }
    }
}