using ClosedXML.Excel;
using ExportExcelAspNetMVC.Models;
using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Linq;
using System.Web;
using System.Web.Mvc;

namespace ExportExcelAspNetMVC.Controllers
{
    public class ClosedXMLController : Controller
    {
        public List<SampleData> InitData()
        {
            List<SampleData> result = new List<SampleData>();
            result.Add(new SampleData() { Name = "Nguyen Van A", Email = "nva@gmail.com", Phone = "0938283", Address = "218 Hai Phong, Da Nang" });
            result.Add(new SampleData() { Name = "Nguyen Van B", Email = "nva@gmail.com", Phone = "8392832", Address = "189 Son Tra, Da Nang" });
            result.Add(new SampleData() { Name = "Nguyen Van C", Email = "nva@gmail.com", Phone = "9328928", Address = "98 Ngu Hanh Son, Da Nang" });
            result.Add(new SampleData() { Name = "Nguyen Van D", Email = "nva@gmail.com", Phone = "2517621", Address = "91 Bach Dang, Hai Chau, Da Nang" });
            result.Add(new SampleData() { Name = "Nguyen Van E", Email = "nva@gmail.com", Phone = "6903940", Address = "938 Cam Le, Da Nang" });
            return result;
        }
        // GET: ClosedXML
        public ActionResult Index()
        {
            return View(InitData());
        }

        public ActionResult ExportExcel()
        {
            DataTable dt = new DataTable();
            dt.Columns.AddRange(new DataColumn[4] { new DataColumn("Name", typeof(string)),
            new DataColumn("Email", typeof(string)),
            new DataColumn("Phone",typeof(string)),
            new DataColumn("Address",typeof(string))});

            var result = InitData();
            foreach (var item in result)
            {
                dt.Rows.Add(item.Name, item.Email, item.Phone, item.Address);
            }
            //Exporting to Excel
            //string folderPath = "C:\\Excel\\";
            //if (!Directory.Exists(folderPath))
            //{
            //    Directory.CreateDirectory(folderPath);
            //}
            //Codes for the Closed XML
            using (XLWorkbook wb = new XLWorkbook())
            {
                wb.Worksheets.Add(dt, "Customers");

                //wb.SaveAs(folderPath + "DataGridViewExport.xlsx");
                string myName = Server.UrlEncode("Test" + "_" + DateTime.Now.ToShortDateString() + ".xlsx");
                MemoryStream stream = GetStream(wb);// The method is defined below
                Response.Clear();
                Response.Buffer = true;
                Response.AddHeader("content-disposition", "attachment; filename=" + myName);
                Response.ContentType = "application/vnd.ms-excel";
                Response.BinaryWrite(stream.ToArray());
                Response.End();
            }

            return RedirectToAction("Index");
        }


        public MemoryStream GetStream(XLWorkbook excelWorkbook)
        {
            MemoryStream fs = new MemoryStream();
            excelWorkbook.SaveAs(fs);
            fs.Position = 0;
            return fs;
        }
    }
}