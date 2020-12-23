using Newtonsoft.Json;
using OfficeOpenXml;
using OpenQA.Selenium;
using OpenQA.Selenium.Chrome;
using OpenQA.Selenium.Remote;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Threading;

namespace Crunch
{
    class Program
    {
        static void Main(string[] args)
        {
            string json = File.ReadAllText("appsettings.json");
            AppSettings AppSettings = JsonConvert.DeserializeObject<AppSettings>(json);
            ChromeOptions options = new ChromeOptions
            {
                DebuggerAddress = AppSettings.DebuggerAddress
            };
            IWebDriver webDriver = new ChromeDriver(options)
            {
                Url = AppSettings.BaseUrl
            };

            List<string> pages = new List<string>();
            List<string> urls = new List<string>();
            List<Infomation> Infomations = new List<Infomation>();
            pages.Add(AppSettings.BaseUrl);
            for (int i = 1; i < AppSettings.PageNumber; i++)
            {
                var pageElement = webDriver.FindElement(By.ClassName("component--results-info"));
                var As = pageElement.FindElements(By.TagName("a"));
                var Next = As[1];
                string href = Next.GetAttribute("href");
                pages.Add(href);
                webDriver.Url = href;
            }

            foreach (string page in pages)
            {
                webDriver.Url = page;
                Thread.Sleep(2000);
                var elements = webDriver.FindElements(By.TagName("identifier-formatter"));
                foreach (var element in elements)
                {
                    var a = element.FindElement(By.TagName("a"));
                    string href = a.GetAttribute("href");
                    urls.Add(href);
                }
            }

            foreach (string url in urls)
            {
                webDriver.Url = url;
                Infomation Infomation = new Infomation();
                Infomations.Add(Infomation);
                Infomation.Url = url;
                Thread.Sleep(2000);
                var elements = webDriver.FindElements(By.TagName("profile-section"));
                var About = elements[0];
                var Highlights = elements[1];
                var RecentNews = elements[2];
                var Details = elements[3];
                // About
                {
                    var lis = About.FindElements(By.TagName("li"));
                    Infomation.Location = lis[0].Text;
                }

                //Highlights
                {
                    var As = Highlights.FindElements(By.TagName("a"));
                    foreach (var a in As)
                    {
                        var label_with_info = a.FindElement(By.TagName("label-with-info"));
                        var element = label_with_info.FindElement(By.XPath(".//span"));
                        if (element.Text.Contains("Total Funding"))
                        {
                            var field_formatter = a.FindElement(By.TagName("field-formatter"));
                            var span = field_formatter.FindElement(By.TagName("span"));
                            Infomation.TotalFunding = span.Text;
                            break;
                        }
                    }

                }

                //Details
                {
                    var lis = Details.FindElements(By.TagName("li"));
                    foreach (var li in lis)
                    {
                        try
                        {
                            var label_with_info = li.FindElement(By.TagName("label-with-info"));
                            var element = label_with_info.FindElement(By.XPath(".//span"));
                            if (element.Text.Contains("Industries"))
                            {
                                var mat_chip = li.FindElement(By.TagName("mat-chip"));
                                Infomation.Industry = mat_chip.Text;
                                continue;
                            }
                            if (element.Text.Contains("Operating Status"))
                            {
                                var field_formatter = li.FindElement(By.TagName("field-formatter"));
                                var span = field_formatter.FindElement(By.TagName("span"));
                                Infomation.OperatingStatus = span.Text;
                                continue;
                            }
                            if (element.Text.Contains("Contact Email"))
                            {
                                var field_formatter = li.FindElement(By.TagName("field-formatter"));
                                var span = field_formatter.FindElement(By.TagName("span"));
                                Infomation.ContactEmail = span.Text;
                                continue;
                            }
                            if (element.Text.Contains("Phone Number"))
                            {
                                var field_formatter = li.FindElement(By.TagName("field-formatter"));
                                var span = field_formatter.FindElement(By.TagName("span"));
                                Infomation.PhoneNumber = span.Text;
                                continue;
                            }
                        }
                        catch
                        {

                        }
                    }
                }
            }

            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
            using (ExcelPackage excel = new ExcelPackage())
            {
                excel.Workbook.Worksheets.Add("Worksheet1");
                var excelWorksheet = excel.Workbook.Worksheets["Worksheet1"];
               
                excelWorksheet.Cells[1, 1].Value = "URL";
                excelWorksheet.Cells[1, 2].Value = "TotalFunding";
                excelWorksheet.Cells[1, 3].Value = "Location";
                excelWorksheet.Cells[1, 4].Value = "Industry";
                excelWorksheet.Cells[1, 5].Value = "PhoneNumber";
                excelWorksheet.Cells[1, 6].Value = "ContactEmail";
                excelWorksheet.Cells[1, 7].Value = "OperatingStatus";
                for(int i=0; i< Infomations.Count; i++)
                {
                    Infomation Infomation = Infomations[i];
                    excelWorksheet.Cells[2 + i, 1].Value = Infomation.Url;
                    excelWorksheet.Cells[2 + i, 2].Value = Infomation.TotalFunding;
                    excelWorksheet.Cells[2 + i, 3].Value = Infomation.Location;
                    excelWorksheet.Cells[2 + i, 4].Value = Infomation.Industry;
                    excelWorksheet.Cells[2 + i, 5].Value = Infomation.PhoneNumber;
                    excelWorksheet.Cells[2 + i, 6].Value = Infomation.ContactEmail;
                    excelWorksheet.Cells[2 + i, 7].Value = Infomation.OperatingStatus;
                }    
                FileInfo excelFile = new FileInfo(@"result.xlsx");
                excel.SaveAs(excelFile);
            }
        }
    }

    public class Infomation
    {
        public string Url { get; set; }
        public string TotalFunding { get; set; }
        public string Location { get; set; }
        public string Industry { get; set; }
        public string PhoneNumber { get; set; }
        public string ContactEmail { get; set; }
        public string OperatingStatus { get; set; }
    }

    public class AppSettings
    {
        public long PageNumber { get; set; }
        public string BaseUrl { get; set; }
        public string DebuggerAddress { get; set; }
    }
}
