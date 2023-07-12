using HtmltoWordConverter.Models;
using Microsoft.AspNetCore.Hosting;
using Microsoft.AspNetCore.Hosting.Server;
using Microsoft.AspNetCore.Mvc;
using Microsoft.AspNetCore.Mvc.Rendering;
using Microsoft.EntityFrameworkCore;
using Microsoft.Extensions.Logging;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using System;
using System.Collections.Generic;
using System.Drawing.Printing;
using System.IO;
using System.Linq;
using System.Text;
using System.Web;

namespace HtmltoWordConverter.Controllers
{
    public class HomeController : Controller
    {
        private readonly ILogger<HomeController> _logger;
        private IWebHostEnvironment _hostingEnvironment;
        public HomeController(ILogger<HomeController> logger, IWebHostEnvironment environment)
        {
            _hostingEnvironment = environment;
            _logger = logger;
        }

        public IActionResult Index()
        {
            return View();
        }
        public IActionResult ConvertToWord()
        {
            SaveDOCX("abc", null, false, 0, 0, 0, 0);

            string _baseURL = "http://localhost:1385/";
            string _filename = System.Guid.NewGuid().ToString() + ".doc";
            string htmlRaw = @"<table class='tbl'><thead><tr><th class='style0' colspan='2'> <img src='" + _baseURL + "img/logo.png' style='width: 180px;' /></th><th class='style1' colspan='4'><p style='font-size: 24px; padding-bottom: 2px; padding-top: 2px; font-weight: bold; margin-bottom: 1px;'>INVOICE</p> ID-2021-0024<br> Issue Date:21/09/2021<br> Delivery Date: 22/09/2021<br> Due Date:30/09/2021<br> <br><p style='font-size: 24px; padding-bottom: 2px; padding-top: 2px; font-weight: bold; margin-bottom: 1px;'>CLIENT DETAILS</p> Client 1<br> GST Number:XXXXXXXXXX</th></tr></thead><tbody><tr><td class='headstyle0' colspan='5' style='padding-top: 60px;'></td></tr><tr><td class='style3a'>ITEM</td><td class='style3a'>DESCRIPTION</td><td class='style3a'>QUANTITY</td><td class='style3a'>UNIT PRICE</td><td class='style3a'>TOTAL</td></tr><tr><td class='style3'>Item-1</td><td class='style3'>Description -1</td><td class='style3'>2 Pkt</td><td class='style3'>90.00</td><td class='style3b'>180.00</td></tr><tr><td class='style3'>Item-2</td><td class='style3'>Description-2</td><td class='style3'>5 Pkt</td><td class='style3'>35.00</td><td class='style3b'>175.00</td></tr><tr><td class='style3'>Item-3</td><td class='style3'>Description-3</td><td class='style3'>5 Kg</td><td class='style3'>50.00</td><td class='style3b'>250.00</td></tr><tr><td class='style3'>Item-4</td><td class='style3'>Description-4</td><td class='style3'>5 Kg</td><td class='style3'>150.00</td><td class='style3b'>750.00</td></tr><tr><td class='style3'>Item-5</td><td class='style3'>Description-5</td><td class='style3'>5 Kg</td><td class='style3'>100.00</td><td class='style3b'>500.00</td></tr><tr><td class='style0' colspan='2' rowspan='3'></td><td class='style3' colspan='2'>Total</td><td class='style3b'>1855.00</td></tr><tr><td class='style3' colspan='2'>GST@18%</td><td class='style3b'>333.90</td></tr><tr><td class='style3' colspan='2'>Net Payable Amount</td><td class='style3b'>2188.90</td></tr><tr><td class='style0' colspan='5' style='padding-top: 100px;'></td></tr><tr><td class='style0' colspan='5' style='background-color: aliceblue; border-radius: 2px;'><i>Note:Lorem ipsum dolor sit amet, consectetur adipiscing elit, sed do eiusmod tempor incididunt ut labore et dolore magna aliqua. Ut enim ad minim veniam, quis nostrud exercitation ullamco laboris nisi ut aliquip ex ea commodo consequat. Duis aute irure dolor in reprehenderit in voluptate velit esse cillum dolore eu fugiat nulla pariatur. Excepteur sint occaecat cupidatat non proident, sunt in culpa qui officia deserunt mollit anim id est laborum.</i></td></tr><tr><td class='style1' colspan='5' style='padding-top: 150px;'> Thank You<br> <b>CodeSample</b></td></tr></tbody></table>";

            StringBuilder strHTML = new StringBuilder("");
            strHTML.Append("<html " +
                " xmlns:o='urn:schemas-microsoft-com:office:office'" +
                " xmlns:w='urn:schemas-microsoft-com:office:word'" +
                " xmlns='http://www.w3.org/TR/REC-html40'>" +
                "<head><title>Invoice Sample</title>");

            strHTML.Append("<xml><w:WordDocument>" +
                " <w:View>Print</w:View>" +
                " <w:Zoom>100</w:Zoom>" +
                " <w:DoNotOptimizeForBrowser/>" +
                " </w:WordDocument>" +
                " </xml>");

            strHTML.Append("</head>");
            strHTML.Append("<body><div class='page-settings'>" + htmlRaw + "</div></body></html>");


            using (StreamReader Reader = new StreamReader("E:\\HTML to Word converter\\EmailTemplate.html"))
            {
                StringBuilder Sb = new StringBuilder();
                Sb.Append(Reader.ReadToEnd());


               
            }
            return View();
        }


        private static void SaveDOCX(string fileName, string BodyText, bool isLandScape, double rMargin, double lMargin, double bMargin, double tMargin)
        {
            WordprocessingDocument document = WordprocessingDocument.Open("E:\\HTML to Word converter\\sample.docx", true);
            MainDocumentPart mainDocumenPart = document.MainDocumentPart;

            //Place the HTML String into a MemoryStream Object
            using (StreamReader Reader = new StreamReader("E:\\HTML to Word converter\\Audit Report Tempalte (2).html"))
            {
                StringBuilder Sb = new StringBuilder();
                Sb.Append(Reader.ReadToEnd());
                MemoryStream ms = new MemoryStream(Encoding.UTF8.GetBytes(Sb.ToString()));

                //Assign an HTML Section for the String Text
                string htmlSectionID = "Sect1";

                // Create alternative format import part.
                AlternativeFormatImportPart formatImportPart = mainDocumenPart.AddAlternativeFormatImportPart(AlternativeFormatImportPartType.Html, htmlSectionID);

                // Feed HTML data into format import part (chunk).
                formatImportPart.FeedData(ms);
                AltChunk altChunk = new AltChunk();
                altChunk.Id = htmlSectionID;

                //Clear out the Document Body and Insert just the HTML string.  (This prevents an empty First Line)
                mainDocumenPart.Document.Body.RemoveAllChildren();
                mainDocumenPart.Document.Body.Append(altChunk);

                /*
                 Set the Page Orientation and Margins Based on Page Size
                 inch equiv = 1440 (1 inch margin)
                 */
                double width = 8.5 * 1440;
                double height = 11 * 1440;

                SectionProperties sectionProps = new SectionProperties();
                PageSize pageSize;
                if (isLandScape)
                    pageSize = new PageSize() { Width = (UInt32Value)height, Height = (UInt32Value)width, Orient = PageOrientationValues.Landscape };
                else
                    pageSize = new PageSize() { Width = (UInt32Value)width, Height = (UInt32Value)height, Orient = PageOrientationValues.Portrait };

                rMargin = rMargin * 1440;
                lMargin = lMargin * 1440;
                bMargin = bMargin * 1440;
                tMargin = tMargin * 1440;

                PageMargin pageMargin = new PageMargin() { Top = (Int32)tMargin, Right = (UInt32Value)rMargin, Bottom = (Int32)bMargin, Left = (UInt32Value)lMargin, Header = (UInt32Value)360U, Footer = (UInt32Value)360U, Gutter = (UInt32Value)0U };

                sectionProps.Append(pageSize);
                sectionProps.Append(pageMargin);
                mainDocumenPart.Document.Body.Append(sectionProps);

                //Saving/Disposing of the created word Document
                document.MainDocumentPart.Document.Save();
                document.Dispose();
            }
        }
        protected void Button1_Click(object sender, EventArgs e)
        {
            // SaveDOCX("abc", null, false, 0, 0, 0,0);
          
        }
        public IActionResult Privacy()
        {
            var FilePath = Path.Combine(_hostingEnvironment.WebRootPath, "EmailTemplate.html");
            StreamReader str = new StreamReader(FilePath);
            string MailText = str.ReadToEnd();
            str.Close();

            //Repalce [newusername] = signup user name   
            MailText = MailText.Replace("[RequestTitle]", "test req");
            MailText = MailText.Replace("[RequestDescription]", "test req");
            MailText = MailText.Replace("[link]", "test req");
            MailText = MailText.Replace("[imageLogo]", "test req");
            ViewBag.htmltext = MailText;
            return View();
        }

        [ResponseCache(Duration = 0, Location = ResponseCacheLocation.None, NoStore = true)]
        public IActionResult Error()
        {
            return View(new ErrorViewModel { RequestId = Activity.Current?.Id ?? HttpContext.TraceIdentifier });
        }



        public IActionResult Create()
        {
            var report = new Report();
            var deptlist = new List<Department>();
            deptlist.Add(new Department()
            {
                departmentId = 1,
                Name = "a"
            }
            );
            var a = new SelectListItem() { Value = "1", Text = "All" };
            var b = new List<SelectListItem>();
            b.Add(a);
          
            var abc = new Company();
            var abcList = new List<Company>();
            abc.companyId = 1;
            abc.Name = "a";
            abcList.Add(abc);
            abc.companyId = 2;
            abc.Name = "b";
            abcList.Add(abc);
            abc.companyId = 3;
            abc.Name = "c";
            abcList.Add(abc);
            abc.companyId = 4;
            abc.Name = "d";
            abcList.Add(abc);
            ViewBag.companyId = new SelectList(abcList.ToList(), "companyId", "Name");
            var abcd = new Client();
            var abcdList = new List<Client>();
            abcd.clientId = 1;
            abcd.Name = "a";
            abcdList.Add(abcd);
            abcd.clientId = 2;
            abcd.Name = "b";
            abcdList.Add(abcd);
            abcd.clientId = 3;
            abcd.Name = "c";
            abcdList.Add(abcd);
            abcd.clientId = 4;
            abcd.Name = "d";
            abcdList.Add(abcd);
            ViewBag.clientId = new SelectList(abcdList.ToList(), "clientId", "Name");
            var d = new Department();
            var dl = new List<Department>();
            d.departmentId = 1;
            d.Name = "d";
            dl.Add(d);

            ViewBag.dept = new SelectList(dl.ToList(), "departmentId", "Name");
            ViewBag.subDept = new SelectList(dl.ToList(), "departmentId", "Name");
            var oblist = new List<Observation1>
            {
 new Observation1 { observationId = 0, Name = "ob1" },
                new Observation1 { observationId = 0, Name = "ob2" },
            };
            ViewBag.Observation = new SelectList(oblist.ToList(), "observationId", "Name");
            return View(report);
        }

        // POST: Auditors/Create
        // To protect from overposting attacks, enable the specific properties you want to bind to, for 
        // more details, see http://go.microsoft.com/fwlink/?LinkId=317598.

        public IActionResult GetCities(int countryId)
        {
            // Retrieve cities based on the selected country from your data source
            var cities = new List<SubDepartment>
            {
                new SubDepartment { departmentId = 0, Name = "sub1", SubDepartmentId = 0 },
                new SubDepartment { departmentId = 0, Name = "sub2", SubDepartmentId = 0 },
                new SubDepartment { departmentId = 0, Name = "sub3", SubDepartmentId = 0 },
                new SubDepartment { departmentId = 0, Name = "sub4", SubDepartmentId = 0 }
            };
            // Create a list of SelectListItem objects for the cities
            var cityList = cities.Select(city => new SelectListItem
            {
                Value = city.SubDepartmentId.ToString(),
                Text = city.Name
            });

            return Json(cityList);
        }
     

        [HttpPost]
        [ValidateAntiForgeryToken]
        public async Task<IActionResult> Create( Report Report)
        {
            if (ModelState.IsValid)
            {
                
                return RedirectToAction(nameof(Privacy));
            }
            return View(Report);
        }

    }
}
