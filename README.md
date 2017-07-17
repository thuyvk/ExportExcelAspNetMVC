# Export Excel in AspNet MVC
Example export excel in asp.net mvc5

### Install ClosedXML to your project
To export the MVC view data into a Excel file, i'm using this library. You can install into you project by running following nuget command
> Install-Package ClosedXML -Version 0.87.1

### HomeController.cs
```
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
        public ActionResult Index()
        {
            
            return View(InitData());
        }

        [HttpPost, ValidateAntiForgeryToken]
        public ActionResult Index(List<SampleData> obj)
        {
            List<SampleData> result = InitData();
            var gv = new GridView();
            gv.DataSource = result;
            gv.DataBind();
            Response.ClearContent();
            Response.Buffer = true;
            Response.AddHeader("content-disposition", "attachment; filename=DemoExcel.xls");
            Response.ContentType = "application/ms-excel";
            Response.Charset = "";
            StringWriter objStringWriter = new StringWriter();
            HtmlTextWriter objHtmlTextWriter = new HtmlTextWriter(objStringWriter);
            gv.RenderControl(objHtmlTextWriter);
            Response.Output.Write(objStringWriter.ToString());
            Response.Flush();
            Response.End();
            return View(result);
        }

        public ActionResult About()
        {
            ViewBag.Message = "Your application description page.";

            return View();
        }

        public ActionResult Contact()
        {
            ViewBag.Message = "Your contact page.";

            return View();
        }
```
###  Index.cshtml
```
@model List<ExportExcelAspNetMVC.Models.SampleData>
@{
    ViewBag.Title = "Home Page";
}

<h3>Export Excel Example</h3>
@using (Html.BeginForm()) {
    @Html.AntiForgeryToken();
    <table class="table table-hover">
        <tr>
            <th>Name</th>
            <th>Email</th>
            <th>Phone</th>
            <th>Address</th>
        </tr>
        @foreach (var item in Model)
        {
            <tr>
                <td>@item.Name</td>
                <td>@item.Email</td>
                <td>@item.Phone</td>
                <td>@item.Address</td>
            </tr>
        }
    </table>
    <input type="submit" value="Export Excel" class="btn btn-primary" />
}
```
