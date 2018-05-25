using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;

namespace ReportingServices.Libraries
{
    public class GetApiModel
    {
        public int StatusCode { get; set; }
        public Response Response { get; set; }
    }

    public class Response
    {
        public List<Dictionary<string, string>> Data { get; set; }
        //public List<Datum> Data { get; set; }
    }


}