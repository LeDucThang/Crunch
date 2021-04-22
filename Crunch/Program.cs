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
                try
                {
                    var pageElements = webDriver.FindElements(By.ClassName("component--results-info"));
                    var pageElement = pageElements[1];
                    var As = pageElement.FindElements(By.TagName("a"));
                    var Next = As[1];
                    string href = Next.GetAttribute("href");
                    pages.Add(href);
                    webDriver.Url = href;
                }
                catch (Exception ex)
                {

                }
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
                    if (href.Contains("organization"))
                        urls.Add(href);
                }
            }

            foreach (string url in urls)
            {
                try
                {
                    Console.WriteLine(urls.IndexOf(url));
                    webDriver.Url = url;
                    Infomation Infomation = new Infomation();
                    Infomations.Add(Infomation);
                    Infomation.Url = url;
                    Thread.Sleep(2000);
                    var profileName = webDriver.FindElement(By.ClassName("profile-name"));
                    Infomation.CompanyName = profileName.Text;
                    var elements = webDriver.FindElements(By.TagName("profile-section"));
                    var About = elements[0];
                    var Highlights = elements[1];
                    var RecentNews = elements[2];
                    var Details = elements[3];
                    // About
                    {
                        var lis = About.FindElements(By.TagName("li"));
                        Infomation.Location = lis[0].Text;
                        try
                        {
                            Infomation.Website = lis[4].Text;
                        }
                        catch (Exception ex)
                        {

                        }
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
                catch (Exception)
                {

                }
            }


            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
            using (ExcelPackage excel = new ExcelPackage())
            {
                excel.Workbook.Worksheets.Add("Worksheet1");
                var excelWorksheet = excel.Workbook.Worksheets["Worksheet1"];

                int URLColumn = 1;
                int CompanyNameColumn = 2;
                int TotalFundingColumn = 3;
                int LocationColumn = 4;
                int WebsiteColumn = 5;
                int IndustryColumn = 6;
                int PhoneNumberColumn = 7;
                int ContactEmailColumn = 8;
                int OperatingStatusColumn = 9;

                excelWorksheet.Cells[1, URLColumn].Value = "URL";
                excelWorksheet.Cells[1, CompanyNameColumn].Value = "CompanyName";
                excelWorksheet.Cells[1, TotalFundingColumn].Value = "TotalFunding";
                excelWorksheet.Cells[1, LocationColumn].Value = "Location";
                excelWorksheet.Cells[1, WebsiteColumn].Value = "Website";
                excelWorksheet.Cells[1, IndustryColumn].Value = "Industry";
                excelWorksheet.Cells[1, PhoneNumberColumn].Value = "PhoneNumber";
                excelWorksheet.Cells[1, ContactEmailColumn].Value = "ContactEmail";
                excelWorksheet.Cells[1, OperatingStatusColumn].Value = "OperatingStatus";
                for (int i = 0; i < Infomations.Count; i++)
                {
                    Infomation Infomation = Infomations[i];
                    excelWorksheet.Cells[2 + i, URLColumn].Value = Infomation.Url;
                    excelWorksheet.Cells[2 + i, CompanyNameColumn].Value = Infomation.CompanyName;
                    excelWorksheet.Cells[2 + i, TotalFundingColumn].Value = Infomation.TotalFunding;
                    excelWorksheet.Cells[2 + i, LocationColumn].Value = Infomation.Location;
                    excelWorksheet.Cells[2 + i, WebsiteColumn].Value = Infomation.Website;
                    excelWorksheet.Cells[2 + i, IndustryColumn].Value = Infomation.Industry;
                    excelWorksheet.Cells[2 + i, PhoneNumberColumn].Value = Infomation.PhoneNumber;
                    excelWorksheet.Cells[2 + i, ContactEmailColumn].Value = Infomation.ContactEmail;
                    excelWorksheet.Cells[2 + i, OperatingStatusColumn].Value = Infomation.OperatingStatus;
                }
                FileInfo excelFile = new FileInfo(DateTime.Now.ToString("yyyyMMdd-hhmmss") + ".xlsx");
                excel.SaveAs(excelFile);
            }
        }
    }

    public class Infomation
    {
        public string CompanyName { get; set; }
        public string Url { get; set; }
        public string TotalFunding { get; set; }
        public string Location { get; set; }
        public string Website { get; set; }
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
