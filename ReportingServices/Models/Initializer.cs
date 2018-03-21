using Newtonsoft.Json;
using System.IO;
using System.Text;
using System.Web;

namespace ReportingServices
{
    public static class Initializer
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

        }
    }
}