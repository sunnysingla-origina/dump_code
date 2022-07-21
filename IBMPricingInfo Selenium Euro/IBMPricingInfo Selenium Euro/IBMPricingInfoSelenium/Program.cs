using Microsoft.AspNetCore.Http;
using Newtonsoft.Json.Linq;
using OpenQA.Selenium;
using OpenQA.Selenium.Chrome;
using OpenQA.Selenium.Support.UI;
using System;
using System.Collections.Generic;
using System.Configuration;
using System.Data;
using System.Data.SqlClient;
using System.IO;
using System.Linq;
using System.Net;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using System.Web;

namespace IBMPricingInfoSelenium
{
    class Program
    {
        static string App_ID = ConfigurationManager.AppSettings["App_ID"].ToString();
        static string App_KEY = ConfigurationManager.AppSettings["App_KEY"].ToString();
        static string FetIBMProducts = "https://eu-central-1-renderer-read.knack.com/v1/scenes/scene_3889/views/view_7115/records?page={r}&rows_per_page=1000&sort_field=field_197&sort_order=asc&_=1614837355882";        
        static SqlConnection conn = new SqlConnection(ConfigurationManager.ConnectionStrings["connectionstring"].ToString());
        static WebDriverWait wait;
        static IWebElement Result;
        static TimeSpan interval = new TimeSpan(0, 0, 2);
        static TimeSpan Longinterval = new TimeSpan(0, 0, 4);
        static void Main(string[] args)
        {
            IWebDriver driver;
            var driverService = ChromeDriverService.CreateDefaultService();
            driverService.HideCommandPromptWindow = true;
            
            ChromeOptions options = new ChromeOptions();
            options.AddArgument("--incognito");
            options.AddArgument("--safebrowsing-disable-download-protection");
            options.AddUserProfilePreference("safebrowsing", "enabled");            

            try
            {
                using (driver = new ChromeDriver(driverService, options))
                {
                    driver.Manage().Window.Maximize();
                    startScope(driver);
                }
            }
            catch (Exception ex)
            {
                //SendEmail(ConfigurationManager.AppSettings["AddEmailIDs"].ToString(), "Error please check email", ex.ToString());
            }
        }
        public static void startScope(IWebDriver driver)
        {
            try
            {
                DataTable dt = new DataTable();
                dt.Columns.Add("PartNumber", typeof(string));
                dt.Columns.Add("DropDownValue", typeof(string));
                dt.Columns.Add("PartDescription", typeof(string));
                dt.Columns.Add("SoftwareBrand", typeof(string));
                dt.Columns.Add("ProductGroup", typeof(string));
                dt.Columns.Add("ProductSubgroup", typeof(string));
                dt.Columns.Add("ProductTradeName", typeof(string));
                dt.Columns.Add("ProductName", typeof(string));
                dt.Columns.Add("ProductPlatform", typeof(string));
                dt.Columns.Add("PID", typeof(string));
                dt.Columns.Add("EncryptionLevelCode", typeof(string));
                dt.Columns.Add("MediaType", typeof(string));
                dt.Columns.Add("PartLanguage", typeof(string));
                dt.Columns.Add("PartVersion", typeof(string));
                dt.Columns.Add("PartType", typeof(string));
                dt.Columns.Add("ChargeUnit", typeof(string));
                dt.Columns.Add("Country_Region_Currency", typeof(string));
                dt.Columns.Add("PriceEffectiveDate", typeof(string));
                dt.Columns.Add("SVP_BL", typeof(string));
                dt.Columns.Add("SVP_D", typeof(string));
                dt.Columns.Add("SVP_E", typeof(string));
                dt.Columns.Add("SVP_F", typeof(string));
                dt.Columns.Add("SVP_G", typeof(string));
                dt.Columns.Add("SVP_H", typeof(string));
                dt.Columns.Add("SVP_ED", typeof(string));
                dt.Columns.Add("SVP_Points", typeof(string));
                dt.Columns.Add("SRP", typeof(string));
                DataRow dr = null;
                string error = "";
                int rowNumber = 0;

                start:
                string PartNumber = ""; string DropDownValue = ""; string PartDescription = ""; string SoftwareBrand = ""; string ProductGroup = "";
                string ProductSubgroup = ""; string ProductTradeName = ""; string ProductName = ""; string ProductPlatform = ""; string PID = "";
                string EncryptionLevelCode = ""; string MediaType = ""; string PartLanguage = ""; string PartVersion = ""; string PartType = "";
                string ChargeUnit = ""; string Country_Region_Currency = ""; string PriceEffectiveDate = ""; string SVP_BL = ""; string SVP_D = "";
                string SVP_E = ""; string SVP_F = ""; string SVP_G = ""; string SVP_H = ""; string SVP_ED = ""; string SVP_Points = ""; string SRP = "";
                string foundPIDE = "";

                wait = new WebDriverWait(driver, TimeSpan.FromSeconds(500));
                driver.Navigate().GoToUrl(@"https://www-112.ibm.com/software/howtobuy/passportadvantage/dswpricebook/PbCfgPublic?P1=1041&aid=4");
                wait.Until(driver1 => ((IJavaScriptExecutor)driver).ExecuteScript("return document.readyState").Equals("complete"));
                
                if (IsElementPresent(driver, By.Id("username")))
                {
                    driver.FindElement(By.Id("username")).SendKeys(ConfigurationManager.AppSettings["EmailID"].ToString());
                    driver.FindElement(By.Id("continue-button")).Click();
                    driver.Manage().Timeouts().ImplicitWait = TimeSpan.FromSeconds(10);
                    driver.FindElement(By.Id("password")).SendKeys(ConfigurationManager.AppSettings["Pass"].ToString());
                    driver.FindElement(By.Id("signinbutton")).Click();
                }

                Result = wait.Until(ExpectedConditions.ElementExists(By.Id("prodGrpSrch")));
                wait.Until(driver1 => ((IJavaScriptExecutor)driver).ExecuteScript("return document.readyState").Equals("complete"));
                driver.Manage().Timeouts().ImplicitWait = TimeSpan.FromSeconds(10);

                DataTable dt_dropdownvalue = new DataTable();
                using (var cmd = new SqlCommand("USP_GETDropDownValue", conn))
                using (var da = new SqlDataAdapter(cmd))
                {
                    cmd.Parameters.AddWithValue("@Country_currency", ConfigurationManager.AppSettings["Currency"].ToString());
                    cmd.CommandType = CommandType.StoredProcedure;
                    da.Fill(dt_dropdownvalue);
                }

                SelectElement oSelect = new SelectElement(driver.FindElement(By.Id("prodGrpSrch")));
                IList<IWebElement> elementCount = oSelect.Options;
                
                string foundDropDown = "";
                string foundDropDownElement = "";
                for (int x = 1; x < elementCount.Count; x++)
                {                    
                    try
                    {
                        for (int d = 0; d < dt_dropdownvalue.Rows.Count; d++)
                        {
                            oSelect = new SelectElement(driver.FindElement(By.Id("prodGrpSrch")));
                            foundDropDown = "";
                            if (oSelect.Options.ElementAt(x).Text == dt_dropdownvalue.Rows[d][0].ToString())
                            {
                                foundDropDown = "Yes";
                                break;
                            }
                        }
                        if (foundDropDown != "Yes")
                        {                            
                            dr = dt.NewRow();
                            string found = "";
                            string enter = "";
                            PartNumber = ""; DropDownValue = ""; PartDescription = ""; SoftwareBrand = ""; ProductGroup = ""; ProductSubgroup = ""; ProductTradeName = "";
                            ProductName = ""; ProductPlatform = ""; PID = ""; EncryptionLevelCode = ""; MediaType = ""; PartLanguage = ""; PartVersion = ""; PartType = "";
                            ChargeUnit = ""; Country_Region_Currency = ""; PriceEffectiveDate = ""; SVP_BL = ""; SVP_D = ""; SVP_E = ""; SVP_F = ""; SVP_G = ""; SVP_H = "";
                            SVP_ED = ""; SVP_Points = ""; SRP = ""; foundPIDE = "";

                             oSelect = new SelectElement(driver.FindElement(By.Id("prodGrpSrch")));

                            oSelect.SelectByIndex(x);
                            DropDownValue = oSelect.Options.ElementAt(x).Text;
                            SelectElement oSelectPrice = new SelectElement(driver.FindElement(By.Id("cntryCurr")));
                            oSelectPrice.SelectByText(ConfigurationManager.AppSettings["Currency"].ToString());

                            Thread.Sleep(interval);

                            driver.FindElement(By.ClassName("ibm-btn-arrow-pri")).Click();

                            Result = wait.Until(ExpectedConditions.ElementExists(By.XPath("//a[normalize-space() = 'Edit selection']")));
                            wait.Until(driver1 => ((IJavaScriptExecutor)driver).ExecuteScript("return document.readyState").Equals("complete"));
                            driver.Manage().Timeouts().ImplicitWait = TimeSpan.FromSeconds(10);

                            Thread.Sleep(interval);

                            if (driver.FindElement(By.Id("ibm-content-main")).Text.IndexOf("There are zero active parts") >= 0)
                            {
                                found = "1";
                                dr["PartNumber"] = PartNumber;
                                dr["DropDownValue"] = DropDownValue;
                                dr["PartDescription"] = PartDescription;
                                dr["SoftwareBrand"] = SoftwareBrand;
                                dr["ProductGroup"] = ProductGroup;
                                dr["ProductSubgroup"] = ProductSubgroup;
                                dr["ProductTradeName"] = ProductTradeName;
                                dr["ProductName"] = ProductName;
                                dr["ProductPlatform"] = ProductPlatform;
                                dr["PID"] = PID;
                                dr["EncryptionLevelCode"] = EncryptionLevelCode;
                                dr["MediaType"] = MediaType;
                                dr["PartLanguage"] = PartLanguage;
                                dr["PartVersion"] = PartVersion;
                                dr["PartType"] = PartType;
                                dr["ChargeUnit"] = ChargeUnit;
                                dr["Country_Region_Currency"] = ConfigurationManager.AppSettings["Currency"].ToString();
                                dr["PriceEffectiveDate"] = PriceEffectiveDate;
                                dr["SVP_BL"] = SVP_BL;
                                dr["SVP_D"] = SVP_D;
                                dr["SVP_E"] = SVP_E;
                                dr["SVP_F"] = SVP_F;
                                dr["SVP_G"] = SVP_G;
                                dr["SVP_H"] = SVP_H;
                                dr["SVP_ED"] = SVP_ED;
                                dr["SVP_Points"] = SVP_Points;
                                dr["SRP"] = SRP;
                                dt.Rows.Add(dr);
                            }
                            if (found == "" && IsElementPresent(driver, By.XPath("//input[@value = 'View parts associated with all products']")))
                            {
                                Thread.Sleep(interval);
                                driver.FindElement(By.XPath("//input[@value = 'View parts associated with all products']")).Click();

                                List<IWebElement> ALLPID = driver.FindElements(By.ClassName("ibm-ind-link")).ToList();

                                for (int i = 0; i < ALLPID.Count; i++)
                                {
                                    if (rowNumber != 0 && error == "error")
                                    {                                       
                                        i = rowNumber + 1;
                                        rowNumber = 0;
                                    }
                                    code:
                                    string Estring = ConfigurationManager.AppSettings["fetch"].ToString();
                                    List<string> str = Estring.Split(',').ToList();
                                    for (int e = 0; e < str.Count; e++)
                                    {
                                        enter = "No";
                                        string text = str[e].ToString();
                                        if (ALLPID[i].Text.Substring(0, 1) == text)
                                        {
                                            enter = "Yes";
                                            break;
                                        }
                                    }

                                    if (ALLPID[i].Text.Length == 7 && enter == "Yes")
                                    {
                                        try
                                        {
                                            foundPIDE = "found";
                                            string foundelement = "";

                                            driver.FindElement(By.LinkText(ALLPID[i].Text)).Click();
                                            driver.SwitchTo().Window(driver.WindowHandles.Last());

                                            if (IsElementPresent(driver, By.Id("partNumber")))
                                            {
                                                PartNumber = driver.FindElement(By.XPath("//td[@headers = 'partNumber']")).Text;
                                                PartDescription = driver.FindElement(By.XPath("//td[@headers = 'partDescription']")).Text;
                                                SoftwareBrand = driver.FindElement(By.XPath("//td[@headers = 'productBrand']")).Text;
                                                ProductGroup = driver.FindElement(By.XPath("//td[@headers = 'wwideProdSetDscr']")).Text;
                                                ProductSubgroup = driver.FindElement(By.XPath("//td[@headers = 'wwideProdSubGrpDscr']")).Text;
                                                ProductTradeName = driver.FindElement(By.XPath("//td[@headers = 'productTradename']")).Text;
                                                ProductName = driver.FindElement(By.XPath("//td[@headers = 'productName']")).Text;
                                                ProductPlatform = driver.FindElement(By.XPath("//td[@headers = 'productPlatform']")).Text;
                                                PID = driver.FindElement(By.XPath("//td[@headers = 'pid']")).Text;
                                                EncryptionLevelCode = driver.FindElement(By.XPath("//td[@headers = 'encryptionLevelCode']")).Text;
                                                MediaType = driver.FindElement(By.XPath("//td[@headers = 'mediaType']")).Text;
                                                PartLanguage = driver.FindElement(By.XPath("//td[@headers = 'partLanguage']")).Text;
                                                PartVersion = driver.FindElement(By.XPath("//td[@headers = 'partVersion']")).Text;
                                                PartType = driver.FindElement(By.XPath("//td[@headers = 'partType']")).Text;
                                                ChargeUnit = driver.FindElement(By.XPath("//td[@headers = 'chargeUnit']")).Text;
                                                Country_Region_Currency = driver.FindElement(By.XPath("//td[@headers = 'priceCountryRegionCurrency']")).Text;
                                                PriceEffectiveDate = driver.FindElement(By.XPath("//td[@headers = 'priceEffectivityDate']")).Text;
                                            }

                                            driver.FindElement(By.XPath("//*[@id='ppdResultForm']/div[1]/div/ul/li[2]/a")).Click();

                                            if (IsElementPresent(driver, By.XPath("//*[@id='ppdResultForm']/div[2]/div/div[1]/div[2]/table/tbody/tr/td[2]")))
                                            {
                                                foundelement = "1";
                                                SVP_BL = driver.FindElement(By.XPath("//*[@id='ppdResultForm']/div[2]/div/div[1]/div[1]/table/tbody/tr[1]/td[2]")).Text;
                                                SVP_D = driver.FindElement(By.XPath("//*[@id='ppdResultForm']/div[2]/div/div[1]/div[1]/table/tbody/tr[2]/td[2]")).Text;
                                                SVP_E = driver.FindElement(By.XPath("//*[@id='ppdResultForm']/div[2]/div/div[1]/div[1]/table/tbody/tr[3]/td[2]")).Text;
                                                SVP_F = driver.FindElement(By.XPath("//*[@id='ppdResultForm']/div[2]/div/div[1]/div[1]/table/tbody/tr[4]/td[2]")).Text;
                                                SVP_G = driver.FindElement(By.XPath("//*[@id='ppdResultForm']/div[2]/div/div[1]/div[1]/table/tbody/tr[5]/td[2]")).Text;
                                                SVP_H = driver.FindElement(By.XPath("//*[@id='ppdResultForm']/div[2]/div/div[1]/div[1]/table/tbody/tr[6]/td[2]")).Text;
                                                SVP_ED = driver.FindElement(By.XPath("//*[@id='ppdResultForm']/div[2]/div/div[1]/div[1]/table/tbody/tr[7]/td[2]")).Text;

                                                SVP_Points = "";
                                                SRP = driver.FindElement(By.XPath("//*[@id='ppdResultForm']/div[2]/div/div[1]/div[2]/table/tbody/tr/td[2]")).Text;
                                            }
                                            if (foundelement == "" && IsElementPresent(driver, By.XPath("//*[@id='ppdResultForm']/div[2]/div/div[1]/div[1]/table/tbody/tr/td")))
                                                SRP = driver.FindElement(By.XPath("//*[@id='ppdResultForm']/div[2]/div/div[1]/div[1]/table/tbody/tr/td")).Text;


                                            driver.Close();
                                            driver.SwitchTo().Window(driver.WindowHandles.First());
                                            foundelement = "";

                                            dr["PartNumber"] = PartNumber;
                                            dr["DropDownValue"] = DropDownValue;
                                            dr["PartDescription"] = PartDescription;
                                            dr["SoftwareBrand"] = SoftwareBrand;
                                            dr["ProductGroup"] = ProductGroup;
                                            dr["ProductSubgroup"] = ProductSubgroup;
                                            dr["ProductTradeName"] = ProductTradeName;
                                            dr["ProductName"] = ProductName;
                                            dr["ProductPlatform"] = ProductPlatform;
                                            dr["PID"] = PID;
                                            dr["EncryptionLevelCode"] = EncryptionLevelCode;
                                            dr["MediaType"] = MediaType;
                                            dr["PartLanguage"] = PartLanguage;
                                            dr["PartVersion"] = PartVersion;
                                            dr["PartType"] = PartType;
                                            dr["ChargeUnit"] = ChargeUnit;
                                            dr["Country_Region_Currency"] = Country_Region_Currency;
                                            dr["PriceEffectiveDate"] = PriceEffectiveDate;
                                            dr["SVP_BL"] = SVP_BL;
                                            dr["SVP_D"] = SVP_D;
                                            dr["SVP_E"] = SVP_E;
                                            dr["SVP_F"] = SVP_F;
                                            dr["SVP_G"] = SVP_G;
                                            dr["SVP_H"] = SVP_H;
                                            dr["SVP_ED"] = SVP_ED;
                                            dr["SVP_Points"] = SVP_Points;
                                            dr["SRP"] = SRP;
                                            dt.Rows.Add(dr);
                                            dr = dt.NewRow();
                                            rowNumber = i;
                                            error = "";
                                        }
                                        catch (Exception ex)
                                        {
                                            string errorPIDS = ALLPID[i].Text.ToString();
                                            i++;
                                            dr = dt.NewRow();
                                            using (SqlCommand cmd = new SqlCommand("USP_ErrorPIDs", conn))
                                            {
                                                cmd.CommandType = CommandType.StoredProcedure;
                                                cmd.Parameters.AddWithValue("@PID", errorPIDS);
                                                cmd.Parameters.AddWithValue("@DropDown", DropDownValue);
                                                cmd.Parameters.AddWithValue("@Error", ex.Message);
                                                cmd.Parameters.AddWithValue("@Currency", ConfigurationManager.AppSettings["Currency"].ToString());
                                                conn.Open();
                                                SqlDataAdapter da = new SqlDataAdapter();
                                                da.SelectCommand = cmd;
                                                cmd.ExecuteNonQuery();
                                                conn.Close();
                                            }
                                            goto code;
                                        }
                                    }
                                }
                            }

                            Result = wait.Until(ExpectedConditions.ElementExists(By.XPath("//a[normalize-space() = 'Edit selection']")));
                            wait.Until(driver1 => ((IJavaScriptExecutor)driver).ExecuteScript("return document.readyState").Equals("complete"));
                            driver.Manage().Timeouts().ImplicitWait = TimeSpan.FromSeconds(10);

                            driver.FindElement(By.XPath("//a[normalize-space() = 'Edit selection']")).Click();
                        }
                    }
                    catch (Exception ex)
                    {
                        driver.SwitchTo().Window(driver.WindowHandles.First());
                        error = "error";
                        goto start;
                    }

                    if (dt.Rows.Count == 0 && foundDropDown == "" && foundPIDE == "")
                    {
                        dr["PartNumber"] = PartNumber;
                        dr["DropDownValue"] = DropDownValue;
                        dr["PartDescription"] = PartDescription;
                        dr["SoftwareBrand"] = SoftwareBrand;
                        dr["ProductGroup"] = ProductGroup;
                        dr["ProductSubgroup"] = ProductSubgroup;
                        dr["ProductTradeName"] = ProductTradeName;
                        dr["ProductName"] = ProductName;
                        dr["ProductPlatform"] = ProductPlatform;
                        dr["PID"] = PID;
                        dr["EncryptionLevelCode"] = EncryptionLevelCode;
                        dr["MediaType"] = MediaType;
                        dr["PartLanguage"] = PartLanguage;
                        dr["PartVersion"] = PartVersion;
                        dr["PartType"] = PartType;
                        dr["ChargeUnit"] = ChargeUnit;
                        dr["Country_Region_Currency"] = ConfigurationManager.AppSettings["Currency"].ToString();
                        dr["PriceEffectiveDate"] = PriceEffectiveDate;
                        dr["SVP_BL"] = SVP_BL;
                        dr["SVP_D"] = SVP_D;
                        dr["SVP_E"] = SVP_E;
                        dr["SVP_F"] = SVP_F;
                        dr["SVP_G"] = SVP_G;
                        dr["SVP_H"] = SVP_H;
                        dr["SVP_ED"] = SVP_ED;
                        dr["SVP_Points"] = SVP_Points;
                        dr["SRP"] = SRP;
                        dt.Rows.Add(dr);
                        dr = dt.NewRow();
                    }
                    if (dt.Rows.Count > 0 && error != "error")
                    {
                        using (SqlCommand cmd = new SqlCommand("USP_InsertPricing", conn))
                        {
                            cmd.CommandType = CommandType.StoredProcedure;
                            cmd.Parameters.AddWithValue("@Details", dt);
                            conn.Open();
                            SqlDataAdapter da = new SqlDataAdapter();
                            da.SelectCommand = cmd;
                            cmd.ExecuteNonQuery();
                            conn.Close();
                        }
                    }
                    if (error != "error")
                    {
                        dt = new DataTable();
                        dt.Columns.Add("PartNumber", typeof(string));
                        dt.Columns.Add("DropDownValue", typeof(string));
                        dt.Columns.Add("PartDescription", typeof(string));
                        dt.Columns.Add("SoftwareBrand", typeof(string));
                        dt.Columns.Add("ProductGroup", typeof(string));
                        dt.Columns.Add("ProductSubgroup", typeof(string));
                        dt.Columns.Add("ProductTradeName", typeof(string));
                        dt.Columns.Add("ProductName", typeof(string));
                        dt.Columns.Add("ProductPlatform", typeof(string));
                        dt.Columns.Add("PID", typeof(string));
                        dt.Columns.Add("EncryptionLevelCode", typeof(string));
                        dt.Columns.Add("MediaType", typeof(string));
                        dt.Columns.Add("PartLanguage", typeof(string));
                        dt.Columns.Add("PartVersion", typeof(string));
                        dt.Columns.Add("PartType", typeof(string));
                        dt.Columns.Add("ChargeUnit", typeof(string));
                        dt.Columns.Add("Country_Region_Currency", typeof(string));
                        dt.Columns.Add("PriceEffectiveDate", typeof(string));
                        dt.Columns.Add("SVP_BL", typeof(string));
                        dt.Columns.Add("SVP_D", typeof(string));
                        dt.Columns.Add("SVP_E", typeof(string));
                        dt.Columns.Add("SVP_F", typeof(string));
                        dt.Columns.Add("SVP_G", typeof(string));
                        dt.Columns.Add("SVP_H", typeof(string));
                        dt.Columns.Add("SVP_ED", typeof(string));
                        dt.Columns.Add("SVP_Points", typeof(string));
                        dt.Columns.Add("SRP", typeof(string));
                    }
                }
            }
            catch (Exception ex)
            {
                //SendEmail(ConfigurationManager.AppSettings["AddEmailIDs"].ToString(), "Error please check email", ex.ToString());
            }
        }

        private static bool IsElementPresent(IWebDriver driver, By by)
        {
            try
            {
                driver.FindElement(by);
                return true;
            }
            catch (NoSuchElementException)
            {
                return false;
            }
        }
    }
}
