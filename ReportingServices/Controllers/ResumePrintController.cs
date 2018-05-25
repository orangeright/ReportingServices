using System;
using System.Collections;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Threading.Tasks;
using System.Web;
using System.Web.Mvc;

using iTextSharp.text;
using iTextSharp.text.pdf;
using Newtonsoft.Json;

using ReportingServices.Libraries;

namespace ReportingServices.Controllers
{
    public class ResumePrintController : Controller
    {

        // GET: ResumePrint
        public ActionResult Index()
        {
            return View();
        }

        public ActionResult Pdfs()
        {
            var testdata = JsonConvert.DeserializeObject<Dictionary<string, string>>(JsonData.Jdata);

            var pdf = CreatePdf(testdata);
            return new FileStreamResult(pdf, "application/pdf");
            //return File(pdf, "application/pdf", "result.pdf");
        }

        public async Task<ActionResult> Pdf(string id)
        {
            string result = await ApiUtilities.GetByIDforAPI(id);
            var jsondata = JsonConvert.DeserializeObject<GetApiModel>(result);
            var data = jsondata.Response.Data[0];
            data.Add("PrintDate", DateTime.Now.ToLongDateString());
            var pdf = CreatePdf(data);
            return new FileStreamResult(pdf, "application/pdf");
        }

        public async Task<ActionResult> Excel(string id)
        {
            string result = await ApiUtilities.GetByIDforAPI(id);
            var jsondata = JsonConvert.DeserializeObject<GetApiModel>(result);
            var data = jsondata.Response.Data[0];
            data.Add("PrintDate", DateTime.Now.ToLongDateString());
            var xlsx = Reports.CreateXls(data);
            //return new FileStreamResult(pdf, "application/pdf");
            return File(xlsx, "application / vnd.openxmlformats - officedocument.spreadsheetml.sheet", "result.xlsx");
        }


        //[HttpPost]
        //public ActionResult Pdf(string json)
        //{
        //    var data = JsonConvert.DeserializeObject<Dictionary<string, string>>(json);
        //    var pdf = CreatePdf(data);
        //    return new FileStreamResult(pdf, "application/pdf");
        //}

        public MemoryStream CreatePdf(Dictionary<string, string> data)
        {
            var doc = new Document(PageSize.A4);
            var stream = new MemoryStream();
            var pw = PdfWriter.GetInstance(doc, stream);
            pw.CloseStream = false;

            doc.Open();

            Font normalFont = new Font(BaseFont.CreateFont("c:/windows/fonts/meiryo.ttc, 0", BaseFont.IDENTITY_H, true), 9);
            Font captionFont = new Font(BaseFont.CreateFont("c:/windows/fonts/meiryo.ttc, 0", BaseFont.IDENTITY_H, true), 11);
            Font titleFont = new Font(BaseFont.CreateFont("c:/windows/fonts/meiryo.ttc, 0", BaseFont.IDENTITY_H, true), 14);
            Font noticeFont = new Font(BaseFont.CreateFont("c:/windows/fonts/meiryo.ttc, 0", BaseFont.IDENTITY_H, true), 7);

            // 日付
            Paragraph printDate = new Paragraph("作成日:" + data["PrintDate"], normalFont);
            printDate.Alignment = Element.ALIGN_RIGHT;
            doc.Add(printDate);
            // タイトル
            Paragraph title = new Paragraph("業務経歴書", titleFont);
            title.Alignment = Element.ALIGN_CENTER;
            doc.Add(title);

            // 会社情報
            Paragraph companyName = new Paragraph("△△システム株式会社", normalFont);
            companyName.Alignment = Element.ALIGN_RIGHT;
            doc.Add(companyName);
            Paragraph zipCode = new Paragraph("〒101－1111", normalFont);
            zipCode.Alignment = Element.ALIGN_RIGHT;
            doc.Add(zipCode);
            Paragraph address = new Paragraph("東京都千代田区●●○○町9-9-9", normalFont);
            address.Alignment = Element.ALIGN_RIGHT;
            doc.Add(address);
            Paragraph tel = new Paragraph("TEL (03)9999－9999", normalFont);
            tel.Alignment = Element.ALIGN_RIGHT;
            doc.Add(tel);
            Paragraph fax = new Paragraph("FAX (03)9999－9999", normalFont);
            fax.Alignment = Element.ALIGN_RIGHT;
            doc.Add(fax);


            Paragraph nullRow = new Paragraph(".", normalFont);
            doc.Add(nullRow);
            doc.Add(nullRow);
            doc.Add(nullRow);

            // 名前テーブル
            Paragraph caption1 = new Paragraph("■ 基本情報", captionFont);
            doc.Add(caption1);
            PdfPTable nameTable = new PdfPTable(7);
            nameTable.WidthPercentage = 100;
            nameTable.SetWidths(new float[] { 0.1f, 0.3f, 0.05f, 0.05f, 0.1f, 0.1f, 0.3f });
            nameTable.DefaultCell.HorizontalAlignment = Element.ALIGN_CENTER;
            nameTable.DefaultCell.VerticalAlignment = Element.ALIGN_MIDDLE;
            nameTable.PaddingTop = 3;
            nameTable.SpacingBefore = 0;
            nameTable.SpacingAfter = 0;

            PdfPCell nameCell = new PdfPCell();
            nameCell = new PdfPCell(new Phrase("フリガナ", normalFont));
            CellStyling(nameCell);
            nameTable.AddCell(nameCell);
            nameCell = new PdfPCell(new Phrase(data["Class002"], normalFont));
            CellStyling(nameCell);
            nameTable.AddCell(nameCell);
            nameCell = new PdfPCell(new Phrase("年齢", normalFont));
            CellStyling(nameCell);
            nameTable.AddCell(nameCell);
            nameCell = new PdfPCell(new Phrase("性別", normalFont));
            CellStyling(nameCell);
            nameTable.AddCell(nameCell);
            nameCell = new PdfPCell(new Phrase("生年月", normalFont));
            CellStyling(nameCell);
            nameTable.AddCell(nameCell);
            nameCell = new PdfPCell(new Phrase("最終学歴", normalFont));
            CellStyling(nameCell);
            nameTable.AddCell(nameCell);
            nameCell = new PdfPCell(new Phrase("最寄駅", normalFont));
            CellStyling(nameCell);
            nameTable.AddCell(nameCell);

            nameCell = new PdfPCell(new Phrase("氏名", normalFont));
            CellStyling(nameCell);
            nameTable.AddCell(nameCell);
            nameCell = new PdfPCell(new Phrase(data["Class001"], normalFont));
            CellStyling(nameCell);
            nameTable.AddCell(nameCell);
            nameCell = new PdfPCell(new Phrase(data["NumA"], normalFont));
            CellStyling(nameCell);
            nameTable.AddCell(nameCell);
            nameCell = new PdfPCell(new Phrase(data["Class003"], normalFont));
            CellStyling(nameCell);
            nameTable.AddCell(nameCell);
            nameCell = new PdfPCell(new Phrase(data["NumB"] + "年" + data["NumC"] + "月", normalFont));
            CellStyling(nameCell);
            nameTable.AddCell(nameCell);
            nameCell = new PdfPCell(new Phrase(data["Class004"], normalFont));
            CellStyling(nameCell);
            nameTable.AddCell(nameCell);
            nameCell = new PdfPCell(new Phrase(data["Class005"] + " " + data["Class006"], normalFont));
            CellStyling(nameCell);
            nameTable.AddCell(nameCell);

            doc.Add(nameTable);


            doc.Add(nullRow);
            doc.Add(nullRow);

            // 資格テーブル
            Paragraph caption2 = new Paragraph("■ 資格情報", captionFont);
            doc.Add(caption2);
            PdfPTable qualificationTable = new PdfPTable(3);
            qualificationTable.WidthPercentage = 100;
            qualificationTable.SetWidths(new float[] { 0.1f, 0.6f, 0.3f });
            qualificationTable.DefaultCell.HorizontalAlignment = Element.ALIGN_CENTER;
            qualificationTable.DefaultCell.VerticalAlignment = Element.ALIGN_MIDDLE;
            qualificationTable.PaddingTop = 3;
            qualificationTable.SpacingBefore = 0;
            qualificationTable.SpacingAfter = 0;

            PdfPCell qualificationCell = new PdfPCell();
            qualificationCell = new PdfPCell(new Phrase("No", normalFont));
            CellStyling(qualificationCell);
            qualificationTable.AddCell(qualificationCell);
            qualificationCell = new PdfPCell(new Phrase("情報処理に関する資格", normalFont));
            CellStyling(qualificationCell);
            qualificationTable.AddCell(qualificationCell);
            qualificationCell = new PdfPCell(new Phrase("取得年月", normalFont));
            CellStyling(qualificationCell);
            qualificationTable.AddCell(qualificationCell);
            qualificationCell = new PdfPCell(new Phrase("1", normalFont));
            CellStyling(qualificationCell);
            qualificationTable.AddCell(qualificationCell);
            qualificationCell = new PdfPCell(new Phrase(data["Class007"], normalFont));
            CellStyling(qualificationCell);
            qualificationTable.AddCell(qualificationCell);
            qualificationCell = new PdfPCell(new Phrase(data["Class008"], normalFont));
            CellStyling(qualificationCell);
            qualificationTable.AddCell(qualificationCell);
            qualificationCell = new PdfPCell(new Phrase("2", normalFont));
            CellStyling(qualificationCell);
            qualificationTable.AddCell(qualificationCell);
            qualificationCell = new PdfPCell(new Phrase(data["Class009"], normalFont));
            CellStyling(qualificationCell);
            qualificationTable.AddCell(qualificationCell);
            qualificationCell = new PdfPCell(new Phrase(data["Class010"], normalFont));
            CellStyling(qualificationCell);
            qualificationTable.AddCell(qualificationCell);
            qualificationCell = new PdfPCell(new Phrase("3", normalFont));
            CellStyling(qualificationCell);
            qualificationTable.AddCell(qualificationCell);
            qualificationCell = new PdfPCell(new Phrase(data["Class011"], normalFont));
            CellStyling(qualificationCell);
            qualificationTable.AddCell(qualificationCell);
            qualificationCell = new PdfPCell(new Phrase(data["Class012"], normalFont));
            CellStyling(qualificationCell);
            qualificationTable.AddCell(qualificationCell);
            qualificationCell = new PdfPCell(new Phrase("4", normalFont));
            CellStyling(qualificationCell);
            qualificationTable.AddCell(qualificationCell);
            qualificationCell = new PdfPCell(new Phrase(data["Class013"], normalFont));
            CellStyling(qualificationCell);
            qualificationTable.AddCell(qualificationCell);
            qualificationCell = new PdfPCell(new Phrase(data["Class014"], normalFont));
            CellStyling(qualificationCell);
            qualificationTable.AddCell(qualificationCell);
            qualificationCell = new PdfPCell(new Phrase("5", normalFont));
            CellStyling(qualificationCell);
            qualificationTable.AddCell(qualificationCell);
            qualificationCell = new PdfPCell(new Phrase(data["Class015"], normalFont));
            CellStyling(qualificationCell);
            qualificationTable.AddCell(qualificationCell);
            qualificationCell = new PdfPCell(new Phrase(data["Class016"], normalFont));
            CellStyling(qualificationCell);
            qualificationTable.AddCell(qualificationCell);

            //ArrayList arrayList = new ArrayList();
            //arrayList.Add(new string[] { "1", "第二種情報処理技術者", "1997年12月" });
            //arrayList.Add(new string[] { "2", "ソフトウェア開発技術者", "2003年6月" });

            //foreach (string[] s in arrayList)
            //{
            //    qualificationCell = new PdfPCell(new Phrase(s[0], normalFont));
            //    qualificationTable.AddCell(qualificationCell);
            //    qualificationCell = new PdfPCell(new Phrase(s[1], normalFont));
            //    qualificationTable.AddCell(qualificationCell);
            //    qualificationCell = new PdfPCell(new Phrase(s[2], normalFont));
            //    qualificationTable.AddCell(qualificationCell);
            //}

            doc.Add(qualificationTable);

            doc.Add(nullRow);
            doc.Add(nullRow);
            doc.Add(nullRow);


            // スキルサマリテーブル
            Paragraph caption3 = new Paragraph("■ スキルサマリ", captionFont);
            doc.Add(caption3);
            PdfPTable skillsummaryTable = new PdfPTable(9);
            skillsummaryTable.WidthPercentage = 100;
            skillsummaryTable.SetWidths(new float[] { 0.12f, 0.11f, 0.11f, 0.11f, 0.11f, 0.11f, 0.11f, 0.11f, 0.11f });
            skillsummaryTable.DefaultCell.HorizontalAlignment = Element.ALIGN_CENTER;
            skillsummaryTable.DefaultCell.VerticalAlignment = Element.ALIGN_MIDDLE;
            skillsummaryTable.PaddingTop = 3;
            skillsummaryTable.SpacingBefore = 0;
            skillsummaryTable.SpacingAfter = 0;

            PdfPCell skillsummaryCell = new PdfPCell();
            skillsummaryCell = new PdfPCell(new Phrase("システム開発経験", normalFont));
            CellStyling(skillsummaryCell);
            skillsummaryCell.Colspan = 2;
            skillsummaryTable.AddCell(skillsummaryCell);
            skillsummaryCell = new PdfPCell(new Phrase(data["NumD"] + "年", normalFont));
            CellStyling(skillsummaryCell);
            skillsummaryTable.AddCell(skillsummaryCell);
            skillsummaryCell = new PdfPCell(new Phrase("5=精通している、4=自主的に出来る、3=指示通りできる、2=開発経験がある、1=教育・研修", noticeFont));
            CellStyling(skillsummaryCell);
            skillsummaryCell.Colspan = 6;
            skillsummaryTable.AddCell(skillsummaryCell);

            skillsummaryCell = new PdfPCell(new Phrase("OS", normalFont));
            CellStyling(skillsummaryCell);
            skillsummaryCell.Rowspan = 2;
            skillsummaryTable.AddCell(skillsummaryCell);
            skillsummaryCell = new PdfPCell(new Phrase("Windows", normalFont));
            CellStyling(skillsummaryCell);
            skillsummaryTable.AddCell(skillsummaryCell);
            skillsummaryCell = new PdfPCell(new Phrase("UNIX", normalFont));
            CellStyling(skillsummaryCell);
            skillsummaryTable.AddCell(skillsummaryCell);
            skillsummaryCell = new PdfPCell(new Phrase("Linux", normalFont));
            CellStyling(skillsummaryCell);
            skillsummaryTable.AddCell(skillsummaryCell);
            skillsummaryCell = new PdfPCell(new Phrase("Android", normalFont));
            CellStyling(skillsummaryCell);
            skillsummaryTable.AddCell(skillsummaryCell);
            skillsummaryCell = new PdfPCell(new Phrase("iOS", normalFont));
            CellStyling(skillsummaryCell);
            skillsummaryTable.AddCell(skillsummaryCell);
            skillsummaryCell = new PdfPCell(new Phrase("AS400", normalFont));
            CellStyling(skillsummaryCell);
            skillsummaryTable.AddCell(skillsummaryCell);
            skillsummaryCell = new PdfPCell(new Phrase("", normalFont));
            CellStyling(skillsummaryCell);
            skillsummaryTable.AddCell(skillsummaryCell);
            skillsummaryCell = new PdfPCell(new Phrase("", normalFont));
            CellStyling(skillsummaryCell);
            skillsummaryTable.AddCell(skillsummaryCell);
            skillsummaryCell = new PdfPCell(new Phrase(data["Num001"], normalFont));
            CellStyling(skillsummaryCell);
            skillsummaryTable.AddCell(skillsummaryCell);
            skillsummaryCell = new PdfPCell(new Phrase(data["Num002"], normalFont));
            CellStyling(skillsummaryCell);
            skillsummaryTable.AddCell(skillsummaryCell);
            skillsummaryCell = new PdfPCell(new Phrase(data["Num003"], normalFont));
            CellStyling(skillsummaryCell);
            skillsummaryTable.AddCell(skillsummaryCell);
            skillsummaryCell = new PdfPCell(new Phrase(data["Num004"], normalFont));
            CellStyling(skillsummaryCell);
            skillsummaryTable.AddCell(skillsummaryCell);
            skillsummaryCell = new PdfPCell(new Phrase(data["Num005"], normalFont));
            CellStyling(skillsummaryCell);
            skillsummaryTable.AddCell(skillsummaryCell);
            skillsummaryCell = new PdfPCell(new Phrase(data["Num006"], normalFont));
            CellStyling(skillsummaryCell);
            skillsummaryTable.AddCell(skillsummaryCell);
            skillsummaryCell = new PdfPCell(new Phrase("", normalFont));
            CellStyling(skillsummaryCell);
            skillsummaryTable.AddCell(skillsummaryCell);
            skillsummaryCell = new PdfPCell(new Phrase("", normalFont));
            CellStyling(skillsummaryCell);
            skillsummaryTable.AddCell(skillsummaryCell);

            skillsummaryCell = new PdfPCell(new Phrase("DB", normalFont));
            CellStyling(skillsummaryCell);
            skillsummaryCell.Rowspan = 2;
            skillsummaryTable.AddCell(skillsummaryCell);
            skillsummaryCell = new PdfPCell(new Phrase("Oracle", normalFont));
            CellStyling(skillsummaryCell);
            skillsummaryTable.AddCell(skillsummaryCell);
            skillsummaryCell = new PdfPCell(new Phrase("SQL Server", normalFont));
            CellStyling(skillsummaryCell);
            skillsummaryTable.AddCell(skillsummaryCell);
            skillsummaryCell = new PdfPCell(new Phrase("Access", normalFont));
            CellStyling(skillsummaryCell);
            skillsummaryTable.AddCell(skillsummaryCell);
            skillsummaryCell = new PdfPCell(new Phrase("MySQL", normalFont));
            CellStyling(skillsummaryCell);
            skillsummaryTable.AddCell(skillsummaryCell);
            skillsummaryCell = new PdfPCell(new Phrase("PostgreSQL", normalFont));
            CellStyling(skillsummaryCell);
            skillsummaryTable.AddCell(skillsummaryCell);
            skillsummaryCell = new PdfPCell(new Phrase("DB2", normalFont));
            CellStyling(skillsummaryCell);
            skillsummaryTable.AddCell(skillsummaryCell);
            skillsummaryCell = new PdfPCell(new Phrase("", normalFont));
            CellStyling(skillsummaryCell);
            skillsummaryTable.AddCell(skillsummaryCell);
            skillsummaryCell = new PdfPCell(new Phrase("", normalFont));
            CellStyling(skillsummaryCell);
            skillsummaryTable.AddCell(skillsummaryCell);
            skillsummaryCell = new PdfPCell(new Phrase(data["Num007"], normalFont));
            CellStyling(skillsummaryCell);
            skillsummaryTable.AddCell(skillsummaryCell);
            skillsummaryCell = new PdfPCell(new Phrase(data["Num008"], normalFont));
            CellStyling(skillsummaryCell);
            skillsummaryTable.AddCell(skillsummaryCell);
            skillsummaryCell = new PdfPCell(new Phrase(data["Num009"], normalFont));
            CellStyling(skillsummaryCell);
            skillsummaryTable.AddCell(skillsummaryCell);
            skillsummaryCell = new PdfPCell(new Phrase(data["Num010"], normalFont));
            CellStyling(skillsummaryCell);
            skillsummaryTable.AddCell(skillsummaryCell);
            skillsummaryCell = new PdfPCell(new Phrase(data["Num011"], normalFont));
            CellStyling(skillsummaryCell);
            skillsummaryTable.AddCell(skillsummaryCell);
            skillsummaryCell = new PdfPCell(new Phrase(data["Num012"], normalFont));
            CellStyling(skillsummaryCell);
            skillsummaryTable.AddCell(skillsummaryCell);
            skillsummaryCell = new PdfPCell(new Phrase("", normalFont));
            CellStyling(skillsummaryCell);
            skillsummaryTable.AddCell(skillsummaryCell);
            skillsummaryCell = new PdfPCell(new Phrase("", normalFont));
            CellStyling(skillsummaryCell);
            skillsummaryTable.AddCell(skillsummaryCell);

            skillsummaryCell = new PdfPCell(new Phrase("言語", normalFont));
            CellStyling(skillsummaryCell);
            skillsummaryCell.Rowspan = 4;
            skillsummaryTable.AddCell(skillsummaryCell);
            skillsummaryCell = new PdfPCell(new Phrase("VB.net", normalFont));
            CellStyling(skillsummaryCell);
            skillsummaryTable.AddCell(skillsummaryCell);
            skillsummaryCell = new PdfPCell(new Phrase("C#", normalFont));
            CellStyling(skillsummaryCell);
            skillsummaryTable.AddCell(skillsummaryCell);
            skillsummaryCell = new PdfPCell(new Phrase("Java", normalFont));
            CellStyling(skillsummaryCell);
            skillsummaryTable.AddCell(skillsummaryCell);
            skillsummaryCell = new PdfPCell(new Phrase("PHP", normalFont));
            CellStyling(skillsummaryCell);
            skillsummaryTable.AddCell(skillsummaryCell);
            skillsummaryCell = new PdfPCell(new Phrase("Ruby", normalFont));
            CellStyling(skillsummaryCell);
            skillsummaryTable.AddCell(skillsummaryCell);
            skillsummaryCell = new PdfPCell(new Phrase("Perl", normalFont));
            CellStyling(skillsummaryCell);
            skillsummaryTable.AddCell(skillsummaryCell);
            skillsummaryCell = new PdfPCell(new Phrase("Python", normalFont));
            CellStyling(skillsummaryCell);
            skillsummaryTable.AddCell(skillsummaryCell);
            skillsummaryCell = new PdfPCell(new Phrase("AWK", normalFont));
            CellStyling(skillsummaryCell);
            skillsummaryTable.AddCell(skillsummaryCell);
            skillsummaryCell = new PdfPCell(new Phrase(data["Num013"], normalFont));
            CellStyling(skillsummaryCell);
            skillsummaryTable.AddCell(skillsummaryCell);
            skillsummaryCell = new PdfPCell(new Phrase(data["Num014"], normalFont));
            CellStyling(skillsummaryCell);
            skillsummaryTable.AddCell(skillsummaryCell);
            skillsummaryCell = new PdfPCell(new Phrase(data["Num015"], normalFont));
            CellStyling(skillsummaryCell);
            skillsummaryTable.AddCell(skillsummaryCell);
            skillsummaryCell = new PdfPCell(new Phrase(data["Num016"], normalFont));
            CellStyling(skillsummaryCell);
            skillsummaryTable.AddCell(skillsummaryCell);
            skillsummaryCell = new PdfPCell(new Phrase(data["Num017"], normalFont));
            CellStyling(skillsummaryCell);
            skillsummaryTable.AddCell(skillsummaryCell);
            skillsummaryCell = new PdfPCell(new Phrase(data["Num018"], normalFont));
            CellStyling(skillsummaryCell);
            skillsummaryTable.AddCell(skillsummaryCell);
            skillsummaryCell = new PdfPCell(new Phrase(data["Num019"], normalFont));
            CellStyling(skillsummaryCell);
            skillsummaryTable.AddCell(skillsummaryCell);
            skillsummaryCell = new PdfPCell(new Phrase(data["Num020"], normalFont));
            CellStyling(skillsummaryCell);
            skillsummaryTable.AddCell(skillsummaryCell);
            skillsummaryCell = new PdfPCell(new Phrase("JavaScript", normalFont));
            CellStyling(skillsummaryCell);
            skillsummaryTable.AddCell(skillsummaryCell);
            skillsummaryCell = new PdfPCell(new Phrase("VBA", normalFont));
            CellStyling(skillsummaryCell);
            skillsummaryTable.AddCell(skillsummaryCell);
            skillsummaryCell = new PdfPCell(new Phrase("C", normalFont));
            CellStyling(skillsummaryCell);
            skillsummaryTable.AddCell(skillsummaryCell);
            skillsummaryCell = new PdfPCell(new Phrase("C++", normalFont));
            CellStyling(skillsummaryCell);
            skillsummaryTable.AddCell(skillsummaryCell);
            skillsummaryCell = new PdfPCell(new Phrase("COBOL", normalFont));
            CellStyling(skillsummaryCell);
            skillsummaryTable.AddCell(skillsummaryCell);
            skillsummaryCell = new PdfPCell(new Phrase("FORTRAN", normalFont));
            CellStyling(skillsummaryCell);
            skillsummaryTable.AddCell(skillsummaryCell);
            skillsummaryCell = new PdfPCell(new Phrase("", normalFont));
            CellStyling(skillsummaryCell);
            skillsummaryTable.AddCell(skillsummaryCell);
            skillsummaryCell = new PdfPCell(new Phrase("", normalFont));
            CellStyling(skillsummaryCell);
            skillsummaryTable.AddCell(skillsummaryCell);
            skillsummaryCell = new PdfPCell(new Phrase(data["Num021"], normalFont));
            CellStyling(skillsummaryCell);
            skillsummaryTable.AddCell(skillsummaryCell);
            skillsummaryCell = new PdfPCell(new Phrase(data["Num022"], normalFont));
            CellStyling(skillsummaryCell);
            skillsummaryTable.AddCell(skillsummaryCell);
            skillsummaryCell = new PdfPCell(new Phrase(data["Num023"], normalFont));
            CellStyling(skillsummaryCell);
            skillsummaryTable.AddCell(skillsummaryCell);
            skillsummaryCell = new PdfPCell(new Phrase(data["Num024"], normalFont));
            CellStyling(skillsummaryCell);
            skillsummaryTable.AddCell(skillsummaryCell);
            skillsummaryCell = new PdfPCell(new Phrase(data["Num025"], normalFont));
            CellStyling(skillsummaryCell);
            skillsummaryTable.AddCell(skillsummaryCell);
            skillsummaryCell = new PdfPCell(new Phrase(data["Num026"], normalFont));
            CellStyling(skillsummaryCell);
            skillsummaryTable.AddCell(skillsummaryCell);
            skillsummaryCell = new PdfPCell(new Phrase("", normalFont));
            CellStyling(skillsummaryCell);
            skillsummaryTable.AddCell(skillsummaryCell);
            skillsummaryCell = new PdfPCell(new Phrase("", normalFont));
            CellStyling(skillsummaryCell);
            skillsummaryTable.AddCell(skillsummaryCell);

            skillsummaryCell = new PdfPCell(new Phrase("フレームワーク", normalFont));
            CellStyling(skillsummaryCell);
            skillsummaryCell.Rowspan = 2;
            skillsummaryTable.AddCell(skillsummaryCell);
            skillsummaryCell = new PdfPCell(new Phrase("Spring", normalFont));
            CellStyling(skillsummaryCell);
            skillsummaryTable.AddCell(skillsummaryCell);
            skillsummaryCell = new PdfPCell(new Phrase("Struts", normalFont));
            CellStyling(skillsummaryCell);
            skillsummaryTable.AddCell(skillsummaryCell);
            skillsummaryCell = new PdfPCell(new Phrase("Hibernate", normalFont));
            CellStyling(skillsummaryCell);
            skillsummaryTable.AddCell(skillsummaryCell);
            skillsummaryCell = new PdfPCell(new Phrase("Seasar", normalFont));
            CellStyling(skillsummaryCell);
            skillsummaryTable.AddCell(skillsummaryCell);
            skillsummaryCell = new PdfPCell(new Phrase("XMDF(オリジナル)", normalFont));
            CellStyling(skillsummaryCell);
            skillsummaryTable.AddCell(skillsummaryCell);
            skillsummaryCell = new PdfPCell(new Phrase(".netFramework", normalFont));
            CellStyling(skillsummaryCell);
            skillsummaryTable.AddCell(skillsummaryCell);
            skillsummaryCell = new PdfPCell(new Phrase("", normalFont));
            CellStyling(skillsummaryCell);
            skillsummaryTable.AddCell(skillsummaryCell);
            skillsummaryCell = new PdfPCell(new Phrase("", normalFont));
            CellStyling(skillsummaryCell);
            skillsummaryTable.AddCell(skillsummaryCell);
            skillsummaryCell = new PdfPCell(new Phrase(data["Num027"], normalFont));
            CellStyling(skillsummaryCell);
            skillsummaryTable.AddCell(skillsummaryCell);
            skillsummaryCell = new PdfPCell(new Phrase(data["Num028"], normalFont));
            CellStyling(skillsummaryCell);
            skillsummaryTable.AddCell(skillsummaryCell);
            skillsummaryCell = new PdfPCell(new Phrase(data["Num029"], normalFont));
            CellStyling(skillsummaryCell);
            skillsummaryTable.AddCell(skillsummaryCell);
            skillsummaryCell = new PdfPCell(new Phrase(data["Num030"], normalFont));
            CellStyling(skillsummaryCell);
            skillsummaryTable.AddCell(skillsummaryCell);
            skillsummaryCell = new PdfPCell(new Phrase(data["Num031"], normalFont));
            CellStyling(skillsummaryCell);
            skillsummaryTable.AddCell(skillsummaryCell);
            skillsummaryCell = new PdfPCell(new Phrase(data["Num032"], normalFont));
            CellStyling(skillsummaryCell);
            skillsummaryTable.AddCell(skillsummaryCell);
            skillsummaryCell = new PdfPCell(new Phrase("", normalFont));
            CellStyling(skillsummaryCell);
            skillsummaryTable.AddCell(skillsummaryCell);
            skillsummaryCell = new PdfPCell(new Phrase("", normalFont));
            CellStyling(skillsummaryCell);
            skillsummaryTable.AddCell(skillsummaryCell);

            skillsummaryCell = new PdfPCell(new Phrase("開発ツール", normalFont));
            CellStyling(skillsummaryCell);
            skillsummaryCell.Rowspan = 2;
            skillsummaryTable.AddCell(skillsummaryCell);
            skillsummaryCell = new PdfPCell(new Phrase("Access", normalFont));
            CellStyling(skillsummaryCell);
            skillsummaryTable.AddCell(skillsummaryCell);
            skillsummaryCell = new PdfPCell(new Phrase("Visual Studio", normalFont));
            CellStyling(skillsummaryCell);
            skillsummaryTable.AddCell(skillsummaryCell);
            skillsummaryCell = new PdfPCell(new Phrase("PowerBuilder", normalFont));
            CellStyling(skillsummaryCell);
            skillsummaryTable.AddCell(skillsummaryCell);
            skillsummaryCell = new PdfPCell(new Phrase("FileMaker", normalFont));
            CellStyling(skillsummaryCell);
            skillsummaryTable.AddCell(skillsummaryCell);
            skillsummaryCell = new PdfPCell(new Phrase("Delphi", normalFont));
            CellStyling(skillsummaryCell);
            skillsummaryTable.AddCell(skillsummaryCell);
            skillsummaryCell = new PdfPCell(new Phrase("Web Performer", normalFont));
            CellStyling(skillsummaryCell);
            skillsummaryTable.AddCell(skillsummaryCell);
            skillsummaryCell = new PdfPCell(new Phrase("Eclipse", normalFont));
            CellStyling(skillsummaryCell);
            skillsummaryTable.AddCell(skillsummaryCell);
            skillsummaryCell = new PdfPCell(new Phrase("ActiveReports", normalFont));
            CellStyling(skillsummaryCell);
            skillsummaryTable.AddCell(skillsummaryCell);
            skillsummaryCell = new PdfPCell(new Phrase(data["Num033"], normalFont));
            CellStyling(skillsummaryCell);
            skillsummaryTable.AddCell(skillsummaryCell);
            skillsummaryCell = new PdfPCell(new Phrase(data["Num034"], normalFont));
            CellStyling(skillsummaryCell);
            skillsummaryTable.AddCell(skillsummaryCell);
            skillsummaryCell = new PdfPCell(new Phrase(data["Num035"], normalFont));
            CellStyling(skillsummaryCell);
            skillsummaryTable.AddCell(skillsummaryCell);
            skillsummaryCell = new PdfPCell(new Phrase(data["Num036"], normalFont));
            CellStyling(skillsummaryCell);
            skillsummaryTable.AddCell(skillsummaryCell);
            skillsummaryCell = new PdfPCell(new Phrase(data["Num037"], normalFont));
            CellStyling(skillsummaryCell);
            skillsummaryTable.AddCell(skillsummaryCell);
            skillsummaryCell = new PdfPCell(new Phrase(data["Num038"], normalFont));
            CellStyling(skillsummaryCell);
            skillsummaryTable.AddCell(skillsummaryCell);
            skillsummaryCell = new PdfPCell(new Phrase(data["Num039"], normalFont));
            CellStyling(skillsummaryCell);
            skillsummaryTable.AddCell(skillsummaryCell);
            skillsummaryCell = new PdfPCell(new Phrase(data["Num040"], normalFont));
            CellStyling(skillsummaryCell);
            skillsummaryTable.AddCell(skillsummaryCell);

            skillsummaryCell = new PdfPCell(new Phrase("サーブレット", normalFont));
            CellStyling(skillsummaryCell);
            skillsummaryCell.Rowspan = 2;
            skillsummaryTable.AddCell(skillsummaryCell);
            skillsummaryCell = new PdfPCell(new Phrase("Apache", normalFont));
            CellStyling(skillsummaryCell);
            skillsummaryTable.AddCell(skillsummaryCell);
            skillsummaryCell = new PdfPCell(new Phrase("Tomcat", normalFont));
            CellStyling(skillsummaryCell);
            skillsummaryTable.AddCell(skillsummaryCell);
            skillsummaryCell = new PdfPCell(new Phrase("IIS", normalFont));
            CellStyling(skillsummaryCell);
            skillsummaryTable.AddCell(skillsummaryCell);
            skillsummaryCell = new PdfPCell(new Phrase("WebLogic", normalFont));
            CellStyling(skillsummaryCell);
            skillsummaryTable.AddCell(skillsummaryCell);
            skillsummaryCell = new PdfPCell(new Phrase("WebSphere", normalFont));
            CellStyling(skillsummaryCell);
            skillsummaryTable.AddCell(skillsummaryCell);
            skillsummaryCell = new PdfPCell(new Phrase("iPlanet", normalFont));
            CellStyling(skillsummaryCell);
            skillsummaryTable.AddCell(skillsummaryCell);
            skillsummaryCell = new PdfPCell(new Phrase("", normalFont));
            CellStyling(skillsummaryCell);
            skillsummaryTable.AddCell(skillsummaryCell);
            skillsummaryCell = new PdfPCell(new Phrase("", normalFont));
            CellStyling(skillsummaryCell);
            skillsummaryTable.AddCell(skillsummaryCell);
            skillsummaryCell = new PdfPCell(new Phrase(data["Num041"], normalFont));
            CellStyling(skillsummaryCell);
            skillsummaryTable.AddCell(skillsummaryCell);
            skillsummaryCell = new PdfPCell(new Phrase(data["Num042"], normalFont));
            CellStyling(skillsummaryCell);
            skillsummaryTable.AddCell(skillsummaryCell);
            skillsummaryCell = new PdfPCell(new Phrase(data["Num043"], normalFont));
            CellStyling(skillsummaryCell);
            skillsummaryTable.AddCell(skillsummaryCell);
            skillsummaryCell = new PdfPCell(new Phrase(data["Num044"], normalFont));
            CellStyling(skillsummaryCell);
            skillsummaryTable.AddCell(skillsummaryCell);
            skillsummaryCell = new PdfPCell(new Phrase(data["Num045"], normalFont));
            CellStyling(skillsummaryCell);
            skillsummaryTable.AddCell(skillsummaryCell);
            skillsummaryCell = new PdfPCell(new Phrase(data["Num046"], normalFont));
            CellStyling(skillsummaryCell);
            skillsummaryTable.AddCell(skillsummaryCell);
            skillsummaryCell = new PdfPCell(new Phrase("", normalFont));
            CellStyling(skillsummaryCell);
            skillsummaryTable.AddCell(skillsummaryCell);
            skillsummaryCell = new PdfPCell(new Phrase("", normalFont));
            CellStyling(skillsummaryCell);
            skillsummaryTable.AddCell(skillsummaryCell);

            skillsummaryCell = new PdfPCell(new Phrase("工程※1", normalFont));
            CellStyling(skillsummaryCell);
            skillsummaryCell.Rowspan = 2;
            skillsummaryTable.AddCell(skillsummaryCell);
            skillsummaryCell = new PdfPCell(new Phrase("RS", normalFont));
            CellStyling(skillsummaryCell);
            skillsummaryTable.AddCell(skillsummaryCell);
            skillsummaryCell = new PdfPCell(new Phrase("SD", normalFont));
            CellStyling(skillsummaryCell);
            skillsummaryTable.AddCell(skillsummaryCell);
            skillsummaryCell = new PdfPCell(new Phrase("PD", normalFont));
            CellStyling(skillsummaryCell);
            skillsummaryTable.AddCell(skillsummaryCell);
            skillsummaryCell = new PdfPCell(new Phrase("CPT", normalFont));
            CellStyling(skillsummaryCell);
            skillsummaryTable.AddCell(skillsummaryCell);
            skillsummaryCell = new PdfPCell(new Phrase("ST", normalFont));
            CellStyling(skillsummaryCell);
            skillsummaryTable.AddCell(skillsummaryCell);
            skillsummaryCell = new PdfPCell(new Phrase("OM", normalFont));
            CellStyling(skillsummaryCell);
            skillsummaryTable.AddCell(skillsummaryCell);
            skillsummaryCell = new PdfPCell(new Phrase("◎=7年以上、○=3～6年、△=3年未満", noticeFont));
            CellStyling(skillsummaryCell);
            skillsummaryCell.Colspan = 2;
            skillsummaryCell.Rowspan = 4;
            skillsummaryTable.AddCell(skillsummaryCell);
            skillsummaryCell = new PdfPCell(new Phrase(data["Class017"], normalFont));
            CellStyling(skillsummaryCell);
            skillsummaryTable.AddCell(skillsummaryCell);
            skillsummaryCell = new PdfPCell(new Phrase(data["Class018"], normalFont));
            CellStyling(skillsummaryCell);
            skillsummaryTable.AddCell(skillsummaryCell);
            skillsummaryCell = new PdfPCell(new Phrase(data["Class019"], normalFont));
            CellStyling(skillsummaryCell);
            skillsummaryTable.AddCell(skillsummaryCell);
            skillsummaryCell = new PdfPCell(new Phrase(data["Class020"], normalFont));
            CellStyling(skillsummaryCell);
            skillsummaryTable.AddCell(skillsummaryCell);
            skillsummaryCell = new PdfPCell(new Phrase(data["Class021"], normalFont));
            CellStyling(skillsummaryCell);
            skillsummaryTable.AddCell(skillsummaryCell);
            skillsummaryCell = new PdfPCell(new Phrase(data["Class022"], normalFont));
            CellStyling(skillsummaryCell);
            skillsummaryTable.AddCell(skillsummaryCell);
            skillsummaryCell = new PdfPCell(new Phrase("役割※2", normalFont));
            CellStyling(skillsummaryCell);
            skillsummaryCell.Rowspan = 2;
            skillsummaryTable.AddCell(skillsummaryCell);
            skillsummaryCell = new PdfPCell(new Phrase("PM", normalFont));
            CellStyling(skillsummaryCell);
            skillsummaryTable.AddCell(skillsummaryCell);
            skillsummaryCell = new PdfPCell(new Phrase("PL", normalFont));
            CellStyling(skillsummaryCell);
            skillsummaryTable.AddCell(skillsummaryCell);
            skillsummaryCell = new PdfPCell(new Phrase("SE", normalFont));
            CellStyling(skillsummaryCell);
            skillsummaryTable.AddCell(skillsummaryCell);
            skillsummaryCell = new PdfPCell(new Phrase("PG", normalFont));
            CellStyling(skillsummaryCell);
            skillsummaryTable.AddCell(skillsummaryCell);
            skillsummaryCell = new PdfPCell(new Phrase("OP", normalFont));
            CellStyling(skillsummaryCell);
            skillsummaryTable.AddCell(skillsummaryCell);
            skillsummaryCell = new PdfPCell(new Phrase("", normalFont));
            CellStyling(skillsummaryCell);
            skillsummaryTable.AddCell(skillsummaryCell);
            skillsummaryCell = new PdfPCell(new Phrase(data["Class023"], normalFont));
            CellStyling(skillsummaryCell);
            skillsummaryTable.AddCell(skillsummaryCell);
            skillsummaryCell = new PdfPCell(new Phrase(data["Class024"], normalFont));
            CellStyling(skillsummaryCell);
            skillsummaryTable.AddCell(skillsummaryCell);
            skillsummaryCell = new PdfPCell(new Phrase(data["Class025"], normalFont));
            CellStyling(skillsummaryCell);
            skillsummaryTable.AddCell(skillsummaryCell);
            skillsummaryCell = new PdfPCell(new Phrase(data["Class026"], normalFont));
            CellStyling(skillsummaryCell);
            skillsummaryTable.AddCell(skillsummaryCell);
            skillsummaryCell = new PdfPCell(new Phrase(data["Class027"], normalFont));
            CellStyling(skillsummaryCell);
            skillsummaryTable.AddCell(skillsummaryCell);
            skillsummaryCell = new PdfPCell(new Phrase("", normalFont));
            CellStyling(skillsummaryCell);
            skillsummaryTable.AddCell(skillsummaryCell);

            doc.Add(skillsummaryTable);

            Paragraph notice1 = new Paragraph("※1　RS:要件定義、SD:基本設計、PD:詳細設計、CPT:開発・テスト、ST:システムテスト、OM:運用・保守", noticeFont);
            doc.Add(notice1);
            Paragraph notice2 = new Paragraph("※2　PM:プロジェクトマネージャー、PL:プロジェクトリーダー、SE:システムエンジニア、PG:プログラマー、OP:オペレーター（運用・保守等）", noticeFont);
            doc.Add(notice2);



            doc.Close();

            byte[] byteInfo = stream.ToArray();
            stream.Write(byteInfo, 0, byteInfo.Length);
            stream.Position = 0;

            return stream;
        }


        private void CellStyling(PdfPCell pdfPCell)
        {
            pdfPCell.HorizontalAlignment = Element.ALIGN_CENTER;
            pdfPCell.VerticalAlignment = Element.ALIGN_MIDDLE;
        }
    }
}