using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;

namespace ReportingServices
{
    public static class Parameters
    {
        public static Pleasanter Pleasanter;
    }

    public class Pleasanter
    {
        public string ApiKey;
        public string Uri;
        public string TemplateUri;
        public string ConnectionString;
        public string TemplatePath;
    }

}