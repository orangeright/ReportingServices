﻿using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;

namespace ReportingServices.Libraries
{
    public class ResponseApiModel
    {
        public int Id { get; set; }
        public int StatusCode { get; set; }
        public string Message { get; set; }
    }
}