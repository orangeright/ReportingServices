using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Net.Http;
using System.Text;
using System.Threading.Tasks;
using System.Web;
using System.Web.Script.Serialization;


namespace ReportingServices.Libraries
{
    public static class ApiUtilities
    {
        public static async Task<string> GetByIDforAPI(string id)
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