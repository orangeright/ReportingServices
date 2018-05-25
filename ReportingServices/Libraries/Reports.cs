using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text.RegularExpressions;
using System.Web;
using ClosedXML.Excel;

namespace ReportingServices.Libraries
{
    public static class Reports
    {
        public static byte[] CreateXls(Dictionary<string, string> data)
        {
            using (var stream = new MemoryStream())
            {
                var workBook = new XLWorkbook(Parameters.Pleasanter.TemplatePath);

                var sheet = workBook.Worksheet(1);

                foreach(var cell in sheet.CellsUsed())
                {
                    foreach (KeyValuePair<string, string> kvp in data)
                    {
                        if (cell.Value.ToString().StartsWith("{{" + kvp.Key + "}}"))
                            cell.Value = cell.Value.ToString().Replace(kvp.Key, kvp.Value);
                    }
                }
                workBook.SaveAs(stream);

                byte[] byteInfo = stream.ToArray();
                //stream.Write(byteInfo, 0, byteInfo.Length);
                //stream.Position = 0;

                return byteInfo;
            }
        }
    }
}