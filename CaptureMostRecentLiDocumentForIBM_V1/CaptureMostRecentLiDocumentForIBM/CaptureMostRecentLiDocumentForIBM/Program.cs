using OpenQA.Selenium;
using OpenQA.Selenium.Chrome;
using OpenQA.Selenium.Support.UI;
using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Threading;
using HtmlAgilityPack;
using System.Collections;
using System.Configuration;
using System.Text.RegularExpressions;
using System.Net;
using System.IO;
using Newtonsoft.Json.Linq;

namespace CaptureMostRecentLiDocumentForIBM
{
    class Program
    {

        static string App_ID = "537f13ff299a6cfc5685ac58";
        static string App_KEY = "bfc24698-e490-499f-b9c8-bbd255362473";
        static string GetIBMPID = "https://eu-api.knack.com/v1/objects/object_139/records?page=1&rows_per_page=1000&sort_field=field_2049&sort_order=asc&filters=%7B%22match%22%3A%22and%22%2C%22rules%22%3A%5B%7B%22field%22%3A%22field_2049%22%2C%22operator%22%3A%22is+not+blank%22%2C%22field_name%22%3A%22IBMPID%22%7D%5D%7D&_=1539668445202";
        static DataTable Dt;

        public static DataTable dt_TableData = null;
        public static DataTable dt_mail = null;
        public static string path = AppDomain.CurrentDomain.BaseDirectory + "Screenshots";
        public static string errorMsg = string.Empty;
        public static int errorFlag = 0;
        static DateTime startDate = new DateTime();
        static DateTime endDate = new DateTime();
        static void Main(string[] args)
        {
            
            ServicePointManager.SecurityProtocol = SecurityProtocolType.Tls12;


            startDate = DateTime.Now;

            clsBLL objbll = new clsBLL();
            WebDriverWait wait = null;
            DataHelper.getDataSetExecuteSPNoParam("USP_ClearALL");

            try
            {
                DataTable dtlatestAnnouncedate = objbll.GetAnnounceDate();

                string date = dtlatestAnnouncedate.Rows[0][0].ToString();
                

                dt_TableData = new DataTable();
                dt_TableData.Columns.AddRange(new DataColumn[7] 
                    { 
                        new DataColumn("ReferenceID", typeof(Int32)),
                        new DataColumn("SubReferenceID", typeof(Int32)),
                        new DataColumn("Announce", typeof(string)),
                        new DataColumn("Program_name(s)", typeof(string)),
                        new DataColumn("Prog#", typeof(string)),
                        new DataColumn("Comments", typeof(string)),
                        new DataColumn("insertDateTime", typeof(DateTime))
                    });



                dt_mail = new DataTable();
                dt_mail.Columns.AddRange(new DataColumn[5] 
                    { 
                        new DataColumn("Sno", typeof(Int32)),
                        new DataColumn("Announce Date", typeof(string)),
                        new DataColumn("Announce Date Count", typeof(string)),
                        new DataColumn("Status", typeof(string)),
                        new DataColumn("Comments", typeof(string))
                    });

                string siteUrl = ConfigurationManager.AppSettings["URL"].ToString();
                //int NoOfDays = Convert.ToInt32(ConfigurationManager.AppSettings["NoOfDays"]);
                DateTime fetch_Date = DateTime.Parse(date); //DateTime.Now.Date.AddDays(-(NoOfDays-1));



                var chromeDriverService = ChromeDriverService.CreateDefaultService();
                chromeDriverService.HideCommandPromptWindow = true;
                
                ChromeOptions optionsChrome = new ChromeOptions();

                optionsChrome.AddArgument("--start-maximized");

                using (IWebDriver driver = new ChromeDriver(chromeDriverService, optionsChrome))//  Create instance of chrome driver
                {

                    try
                    {
                        driver.Navigate().GoToUrl(siteUrl);
                        wait = new WebDriverWait(driver, TimeSpan.FromSeconds(120));
                        wait.Until(driver1 => ((IJavaScriptExecutor)driver).ExecuteScript("return document.readyState").Equals("complete"));
                        Thread.Sleep(2000);


                        string Announce = string.Empty;
                        string Program_name = string.Empty;
                        string Prog = string.Empty;
                        string[] lines_Program_name = null;
                        string[] lines_Prog = null;
                        
                        int counter = 0;
                        int finallyBreakCounter = 0;
                        int currentRowCount = 0;


                        IList<IWebElement> allElement_tableOuter1 = driver.FindElements(By.ClassName("table"));


                        try
                        {
                            //IWebElement OutDiv = driver.FindElement(By.ClassName("truste_box_overlay_inner"));

                            //IWebElement detailFrame = OutDiv.FindElement(By.XPath("//iframe"));
                            IWebDriver driver1 = driver;
                            //IWebElement detailFrame = driver.FindElement(By.XPath("//iframe"));
                            driver1.SwitchTo().Frame(Convert.ToInt32(ConfigurationManager.AppSettings["FrameIndex"]));

                            //alternatively, find the frame first, then use GetAttribute() to get frame name
                            try
                            {

                                //you are now in iframe "Details", then find the elements you want in the frame now
                                IWebElement foo = driver1.FindElement(By.ClassName("call"));
                                foo.Click();
                                driver.SwitchTo().DefaultContent();
                            }
                            catch
                            {
                                driver.SwitchTo().DefaultContent();
                            }
                            //switch back to main frame

                        }
                        catch
                        {
                            driver.SwitchTo().DefaultContent();
                        }
                        #region Get main table data with all clickable link

                        Next_Page:

                        counter = 0;
                        IList<IWebElement> allElement_tableOuter = driver.FindElements(By.TagName("table"));
                        foreach (IWebElement elementTableOuter in allElement_tableOuter)// Get main program list
                        {
                            IWebElement elementTHOuter = elementTableOuter.FindElements(By.TagName("th")).ToList().FirstOrDefault();
                            if (elementTHOuter != null)
                            {
                                if (elementTHOuter.Text == "Announce")
                                {   
                                    IList<IWebElement> allElement = elementTableOuter.FindElements(By.TagName("td"));
                                    foreach (IWebElement element in allElement)
                                    {

                                        if (counter == 0)
                                        {
                                            if (fetch_Date < Convert.ToDateTime(element.Text).Date)
                                            {
                                                Announce = element.Text;
                                                counter++;
                                                finallyBreakCounter = 0;
                                            }
                                            else
                                            {
                                                finallyBreakCounter = 1;
                                                break;
                                            }
                                        }
                                        else if (counter == 1)
                                        {
                                            if (element.Text.Contains("\r\n") || element.Text.Contains("\n") || element.Text.Contains("\r"))
                                            {
                                                lines_Program_name = element.Text.Split(
                                                      new[] { "\r\n", "\r", "\n" },
                                                      StringSplitOptions.None
                                                  );
                                                //Thread.Sleep(1000);
                                            }
                                            else
                                            {
                                                lines_Program_name = null;
                                                Program_name = element.Text;
                                            }
                                            counter++;
                                        }
                                        else if (counter == 2)
                                        {
                                            if (element.Text.Contains("\r\n") || element.Text.Contains("\n") || element.Text.Contains("\r"))
                                            {
                                                lines_Prog = element.Text.Split(
                                                      new[] { "\r\n", "\r", "\n" },
                                                      StringSplitOptions.None
                                                  );
                                                //Thread.Sleep(1000);
                                            }
                                            else
                                            {
                                                lines_Prog = null;
                                                Prog = element.Text;
                                            }

                                            currentRowCount = currentRowCount + 1;

                                            if (lines_Prog == null)
                                            {
                                                dt_TableData.Rows.Add(new Object[] { currentRowCount, 1, Announce, Program_name, Prog, "", DateTime.Now });
                                            }
                                            else
                                            {
                                                for (int i = 0; i < lines_Prog.Length;i++ )
                                                {
                                                    dt_TableData.Rows.Add(new Object[] { currentRowCount, (i + 1), Announce, lines_Program_name[i], lines_Prog[i], "", DateTime.Now });
                                                }
                                                    
                                            }
                                            counter = 0;
                                            Announce = string.Empty;
                                            Program_name = string.Empty;
                                            Prog = string.Empty;
                                            
                                        }
                                    }
                                    if (finallyBreakCounter == 0)
                                    {
                                        IWebElement A_Next = driver.FindElements(By.TagName("a")).ToList().Where(x => x.GetAttribute("innerText") == "Next").FirstOrDefault();

                                        if (A_Next != null)
                                        {
                                            Thread.Sleep(100);
                                            A_Next.Click();
                                            wait = new WebDriverWait(driver, TimeSpan.FromSeconds(120));
                                            wait.Until(driver1 => ((IJavaScriptExecutor)driver).ExecuteScript("return document.readyState").Equals("complete"));
                                            Thread.Sleep(1000);
                                            counter = 0;
                                            Announce = string.Empty;
                                            Program_name = string.Empty;
                                            Prog = string.Empty;
                                            goto Next_Page;
                                        }
                                        else
                                        {
                                            break;
                                        }
                                    }
                                    else
                                    {
                                        break;
                                    }
                                }
                            }
                        }

                        #endregion

                       


                        Thread.Sleep(1000);
                        driver.Dispose();

                        //======================Start : Save Data into database===============

                        objbll.BulkInsertToDataBase(dt_TableData, "tbl_MostRecentLiDocumentDataDetail");
                        endDate = DateTime.Now;
                        Post("https://eu-central-1-builder-write.knack.com/v1/objects/object_276/records", "field_4048=Supporting Script to Capturing Versoning Support Program&field_4049=" + startDate.ToString("MM/dd/yyyy h:mm tt") + "&field_4050=" + endDate.ToString("MM/dd/yyyy h:mm tt") + "&field_4052=" + dt_TableData.Rows.Count , "POST");

                       
                        //======================End : Save Data into database=================

                        if (dt_TableData != null && dt_TableData.Rows.Count > 0)
                        {
                            DataTable dt_distinct_Date = dt_TableData.DefaultView.ToTable(true, "Announce");

                            if (dt_distinct_Date != null && dt_distinct_Date.Rows.Count > 0)
                            {
                                foreach (DataRow dtRow in dt_distinct_Date.Rows)
                                {
                                    DataTable tblFiltered = dt_TableData.AsEnumerable()
                                                          .Where(row => row.Field<String>("Announce") == Convert.ToString(dtRow["Announce"]))
                                                          .OrderByDescending(row => row.Field<String>("Announce"))
                                                          .CopyToDataTable();

                                    dt_mail.Rows.Add(new Object[] { Convert.ToInt32(dt_mail.Rows.Count + 1), Convert.ToString(dtRow["Announce"]), tblFiltered.Rows.Count, "Success", "" });

                                }
                            }
                            else
                            {
                                dt_mail.Rows.Add(new Object[] { Convert.ToInt32(dt_mail.Rows.Count + 1), "", "", "Fail", "No Announce Date fall into fetch No Of Days filter criteria" });
                            }
                        }
                        else
                        {
                            dt_mail.Rows.Add(new Object[] { Convert.ToInt32(dt_mail.Rows.Count + 1), "", "", "Fail", "No Announce Date fall into fetch No Of Days filter criteria or No Table found to fetch Announce Date" });
                        }


                        

                    }
                    catch (Exception ex)
                    {
                        Thread.Sleep(1000);
                        driver.Dispose();
                        dt_mail.Rows.Add(new Object[] { Convert.ToInt32(dt_mail.Rows.Count + 1), "","", "Fail", ex.ToString() });

                    }

                }//End of Using

                //=================Start : Final Status Mail sent===================

                if (dt_mail.Rows.Count > 0)
                {
                    try
                    {
                        DataTable dt_temp = dt_mail.AsEnumerable().Where(x => x.Field<string>("Status") == "Fail").CopyToDataTable();
                        if (dt_temp != null)
                        {
                            DataHelper.SendScriptMail(ConvertDataTableToHTML(dt_mail, 1, ""), Convert.ToString(ConfigurationManager.AppSettings["ErrorSubject"]), null);
                        }
                        else
                        {
                            DataHelper.SendScriptMail("No", Convert.ToString(ConfigurationManager.AppSettings["SucessSubject"]), ConvertDataTableToHTML(dt_mail, 1, ""));
                        }
                    }
                    catch
                    {
                        DataHelper.SendScriptMail("No", Convert.ToString(ConfigurationManager.AppSettings["SucessSubject"]), ConvertDataTableToHTML(dt_mail, 1, ""));
                    }
                }
                else
                {
                    DataHelper.SendScriptMail("No", Convert.ToString(ConfigurationManager.AppSettings["SucessSubject"]), ConvertDataTableToHTML(dt_mail, 0, ""));
                }


                //=================End : Final Status Mail sent===================


            }
            catch (Exception ex)
            {
                DataHelper.SendScriptMail(ConvertDataTableToHTML(dt_mail, 4, ex.ToString()), Convert.ToString(ConfigurationManager.AppSettings["ErrorSubject"]), null);
            }
            finally
            {
                Dt = new DataTable();
                Dt.Columns.Add("IBMPID");
                Dt.Columns.Add("RCDID");
                GET(1);

             

                objbll.BulkInsertToDataBase(Dt, "typIBMPID_T");
                Post("https://eu-central-1-builder-write.knack.com/v1/objects/object_276/records", "field_4048=Internal Script for Capturing Most Recent LI Document&field_4049=" + startDate.ToString("MM/dd/yyyy h:mm tt") + "&field_4050=" + (DateTime.Now).ToString("MM/dd/yyyy h:mm tt") + "&field_4052=" + Dt.Rows.Count.ToString() + "&field_4054=0", "POST");
                DataHelper.getDataSetExecuteSPNoParam("USP_EXPORT");
                //DataSet ds = new DataSet();
                //GenralFunction gf = new GenralFunction();
                //SqlParameter[] PM = new SqlParameter[1];
                //PM[0] = new SqlParameter("@typIBMPID", SqlDbType.Structured);
                //PM[0].Value = Dt;

                //ds = gf.Filldatasetvalue(null, "USP_EXPORT", ds, null);
            }
        }
        /// <summary>
        /// Convert datatable to HTML string for mial body
        /// </summary>
        /// <param name="dt"></param>
        /// <param name="flag"></param>
        /// <param name="errors"></param>
        /// <returns></returns>
        private static string ConvertDataTableToHTML(DataTable dt, int flag, string errors)
        {
            string html = string.Empty;
            html += "<html>";
            html += "<head>";
            html += "</head>";
            html += "<body style='font-family:verdana;font-size: 10pt;'>";

            if (flag == 1)
            {
                html += "<div>Hi Team</div><br><div> Following are the status of Capture Most Recent LI Document script</div><br>";
                html += "<div>This script is run with last <b>" + Convert.ToString(ConfigurationManager.AppSettings["NoOfDays"]) + "</b> day(s) filter conditions including today's date for Announce Date</div><br>";
                html += "<div>Process URL : <b>" + Convert.ToString(ConfigurationManager.AppSettings["URL"]) + "</b></div><br>";

                if (dt.Rows.Count == 1 && Convert.ToString(dt.Rows[0]["Announce Date"]) == "")
                {
                    html += "<div>Error : <b>No Announce Date fall into fetch No Of Days filter criteria or No Table found to fetch Announce Date</b></div>";
                }
                else
                {
                    html += "<div><table border='1px' cellpadding='5' cellspacing='0' style='border: 1px solid Black; font-family:verdana;font-size: 10pt;'>";
                    //add header row
                    html += "<tr align='left' valign='top' style='border: 1px solid Black; font-family:verdana;font-size: 10pt;'>";
                    for (int i = 0; i < dt.Columns.Count; i++)
                        html += "<td align='left' valign='top' style='border: 1px solid Black; font-family:verdana;font-size: 10pt;'><b>" + dt.Columns[i].ColumnName + "</b></td>";
                    html += "</tr>";
                    //add rows
                    for (int i = 0; i < dt.Rows.Count; i++)
                    {
                        html += "<tr align='left' valign='top' style='border: 1px solid Black; font-family:verdana;font-size: 10pt;'>";
                        for (int j = 0; j < dt.Columns.Count; j++)
                            html += "<td align='left' valign='top' style='border: 1px solid Black; font-family:verdana;font-size: 10pt;'>" + dt.Rows[i][j].ToString() + "</td>";
                        html += "</tr>";
                    }
                    html += "</table></div>";
                }
            }
            else if (flag == 0)
            {
                html += "<div>Hi Team</div><br><div> Following are the status of Capture Most Recent LI Document script</div><br>";
                html += "<div>This script is run with last <b>" + Convert.ToString(ConfigurationManager.AppSettings["NoOfDays"]) + "</b> day(s) filter conditions including today's date for Announce Date</div><br>";
                html += "<div>Process URL : <b>" + Convert.ToString(ConfigurationManager.AppSettings["URL"]) + "</b></div><br>";

                html += "<div>Error : <b>There are no status create of Capture Most Recent LI Document script</b></div><br>";
            }
            else if (flag == 2)
            {
                html += "<div>Hi Team</div><br><div> Following are the status of Capture Most Recent LI Document script</div><br>";
                html += "<div>This script is run with last <b>" + Convert.ToString(ConfigurationManager.AppSettings["NoOfDays"]) + "</b> day(s) filter conditions including today's date for Announce Date</div><br>";
                html += "<div>Process URL : <b>" + Convert.ToString(ConfigurationManager.AppSettings["URL"]) + "</b></div><br>";

                html += "<div>Error : <b>There are below errors during save data of Capture Most Recent LI Document script</b></div><br><div>" + errors + "</div>";
            }
            else if (flag == 3)
            {
                html += "<div>Hi Team</div><br><div> Following are the status of Capture Most Recent LI Document script</div><br>";
                html += "<div>This script is run with last <b>" + Convert.ToString(ConfigurationManager.AppSettings["NoOfDays"]) + "</b> day(s) filter conditions including today's date for Announce Date</div><br>";
                html += "<div>Process URL : <b>" + Convert.ToString(ConfigurationManager.AppSettings["URL"]) + "</b></div><br>";

                html += "<div>Error : <b>There are below errors during Capture Most Recent LI Document script</b>  </div><br><div>" + errors + "</div>";
            }
            else if (flag == 4)// for generic error
            {
                html += "<div>Hi Team</div><br><div> Following are the status of Capture Most Recent LI Document script</div><br>";
                html += "<div>This script is run with last <b>" + Convert.ToString(ConfigurationManager.AppSettings["NoOfDays"]) + "</b> day(s) filter conditions including today's date for Announce Date</div><br>";
                html += "<div>Process URL : <b>" + Convert.ToString(ConfigurationManager.AppSettings["URL"]) + "</b></div><br>";

                html += "<div>Error : <b>There are below errors Capture Most Recent LI Document script</b>  </div><br><div>" + errors + "</div>";
            }

           
            html += "<br><div>Regards</div><br><div>Support Team</div>";

            html += "</body>";
            html += "</html>";

            return html;
        }
        /// <summary>
        /// Save current screen shot in drive
        /// </summary>
        /// <param name="path_temp"></param>
        /// <param name="fileName"></param>
        /// <param name="driver"></param>
        private static void SaveScreenShot(string path_temp, String fileName, IWebDriver driver)
        {
            try
            {
                fileName = fileName + string.Format("{0:yyyy-MM-dd_hh-mm-ss-tt}", DateTime.Now) + ".jpeg";
                Screenshot ss = ((ITakesScreenshot)driver).GetScreenshot();
                ss.SaveAsFile(path_temp + "//" + fileName);
            }
            catch (Exception ex)
            {
                errorMsg += "<b>Error occur during save screen shot : </b>" + ex.ToString() + "<br>";
            }
        }




        #region "get Post Methods"
        public static void GET(int page)
        {
            string vystup = null;
            try
            {
                ServicePointManager.SecurityProtocol = SecurityProtocolType.Tls12;

                string url = GetIBMPID.Replace("page=1", "page=" + page.ToString());
                //Our postvars
                byte[] buffer = System.Text.Encoding.UTF8.GetBytes("");
                //Initialisation, we use localhost, change if appliable
                HttpWebRequest WebReq = (HttpWebRequest)WebRequest.Create(url);
                //Our method is post, otherwise the buffer (postvars) would be useless
                WebReq.Method = "GET";
                //We use form contentType, for the postvars.
                WebReq.ContentType = "application/x-www-form-urlencoded";
                WebReq.Headers.Add("X-Knack-Application-Id", App_ID);
                WebReq.Headers.Add("X-Knack-REST-API-Key", App_KEY);
                //The length of the buffer (postvars) is used as contentlength.
                //WebReq.ContentLength = buffer.Length;
                //We open a stream for writing the postvars
                WebResponse response = WebReq.GetResponse();
                Stream stream = response.GetResponseStream();

                StreamReader _Answer = new StreamReader(stream);
                vystup = _Answer.ReadToEnd();
                JObject jObject = JObject.Parse(vystup);

                vystup = jObject["records"][0]["id"].ToString();
                var Objects = jObject["records"];

                for (int x = 0; x < Objects.Count(); x++)
                {
                    DataRow dr = Dt.NewRow();
                    dr[0] = jObject["records"][x]["field_2049"];
                    dr[1] = jObject["records"][x]["id"];
                    if (dr[0].ToString().Length == 7)
                    {
                        Dt.Rows.Add(dr);
                    }
                    Console.WriteLine(jObject["records"][x]["field_2049"]);


                }
                if (jObject["total_pages"].ToString() != jObject["current_page"].ToString())
                {
                    GET(page + 1);
                }

                //Congratulations, you just requested your first POST page, you
                //can now start logging into most login forms, with your application
                //Or other examples.
            }
            catch (Exception ex)
            {
                vystup = "";
                //Console.WriteLine(ex.ToString());
            }



        }

        public static List<string> GETList(string url, string data)
        {
            string vystup = null;
            List<string> strIds = new List<string>();

            try
            {
                //Our postvars
                byte[] buffer = System.Text.Encoding.UTF8.GetBytes(data);
                //Initialisation, we use localhost, change if appliable
                HttpWebRequest WebReq = (HttpWebRequest)WebRequest.Create(url);
                //Our method is post, otherwise the buffer (postvars) would be useless
                WebReq.Method = "GET";
                //We use form contentType, for the postvars.
                WebReq.ContentType = "application/x-www-form-urlencoded";
                WebReq.Headers.Add("X-Knack-Application-Id", App_ID);
                WebReq.Headers.Add("X-Knack-REST-API-Key", App_KEY);
                //The length of the buffer (postvars) is used as contentlength.
                //WebReq.ContentLength = buffer.Length;
                //We open a stream for writing the postvars
                WebResponse response = WebReq.GetResponse();
                Stream stream = response.GetResponseStream();

                StreamReader _Answer = new StreamReader(stream);
                vystup = _Answer.ReadToEnd();
                JObject jObject = JObject.Parse(vystup);

                vystup = jObject["records"][0]["id"].ToString();
                for (int i = 0; i < jObject["records"].Count(); i++)
                {
                    strIds.Add(jObject["records"][i]["id"].ToString().Trim());
                }
                //Congratulations, you just requested your first POST page, you
                //can now start logging into most login forms, with your application
                //Or other examples.
            }
            catch (Exception ex)
            {
                vystup = "";
                strIds.Add("");
                //Console.WriteLine(ex.ToString());
            }
            return strIds;
        }

        public static string GETValue(string url, string data, string FieldId)
        {
            string vystup = null;
            try
            {
                //Our postvars
                byte[] buffer = System.Text.Encoding.UTF8.GetBytes(data);
                //Initialisation, we use localhost, change if appliable
                HttpWebRequest WebReq = (HttpWebRequest)WebRequest.Create(url);
                //Our method is post, otherwise the buffer (postvars) would be useless
                WebReq.Method = "GET";
                //We use form contentType, for the postvars.
                WebReq.ContentType = "application/x-www-form-urlencoded";
                WebReq.Headers.Add("X-Knack-Application-Id", App_ID);
                WebReq.Headers.Add("X-Knack-REST-API-Key", App_KEY);
                //The length of the buffer (postvars) is used as contentlength.
                //WebReq.ContentLength = buffer.Length;
                //We open a stream for writing the postvars
                WebResponse response = WebReq.GetResponse();
                Stream stream = response.GetResponseStream();

                StreamReader _Answer = new StreamReader(stream);
                vystup = _Answer.ReadToEnd();
                JObject jObject = JObject.Parse(vystup);

                vystup = jObject["records"][0][FieldId].ToString();
                //Congratulations, you just requested your first POST page, you
                //can now start logging into most login forms, with your application
                //Or other examples.
            }
            catch (Exception ex)
            {
                vystup = "";
                //Console.WriteLine(ex.ToString());
            }
            return vystup.Trim();
        }

        public static List<string> GETValue(string url, string data, string FieldId1, string FieldId2, string FieldId3)
        {
            string vystup = null;
            List<string> li = new List<string>();
            try
            {
                //Our postvars
                byte[] buffer = System.Text.Encoding.UTF8.GetBytes(data);
                //Initialisation, we use localhost, change if appliable
                HttpWebRequest WebReq = (HttpWebRequest)WebRequest.Create(url);
                //Our method is post, otherwise the buffer (postvars) would be useless
                WebReq.Method = "GET";
                //We use form contentType, for the postvars.
                WebReq.ContentType = "application/x-www-form-urlencoded";
                WebReq.Headers.Add("X-Knack-Application-Id", App_ID);
                WebReq.Headers.Add("X-Knack-REST-API-Key", App_KEY);
                //The length of the buffer (postvars) is used as contentlength.
                //WebReq.ContentLength = buffer.Length;
                //We open a stream for writing the postvars
                WebResponse response = WebReq.GetResponse();
                Stream stream = response.GetResponseStream();

                StreamReader _Answer = new StreamReader(stream);
                vystup = _Answer.ReadToEnd();
                JObject jObject = JObject.Parse(vystup);

                vystup = jObject["records"][0][FieldId1].ToString();

                li.Add(jObject["records"][0][FieldId1].ToString());
                li.Add(jObject["records"][0][FieldId2].ToString());
                li.Add(jObject["records"][0][FieldId3].ToString());

                //Congratulations, you just requested your first POST page, you
                //can now start logging into most login forms, with your application
                //Or other examples.
            }
            catch (Exception ex)
            {
                vystup = "";
                //Console.WriteLine(ex.ToString());
            }
            return li;
        }
        public static string Post(string url, string data, string Post_Put)
        {
            string vystup = null;
            try
            {
                //Our postvars
                byte[] buffer = System.Text.Encoding.UTF8.GetBytes(data);
                //Initialisation, we use localhost, change if appliable
                HttpWebRequest WebReq = (HttpWebRequest)WebRequest.Create(url);
                //Our method is post, otherwise the buffer (postvars) would be useless
                WebReq.Method = Post_Put;
                //We use form contentType, for the postvars.
                WebReq.ContentType = "application/x-www-form-urlencoded";
                WebReq.Headers.Add("X-Knack-Application-Id", App_ID);
                WebReq.Headers.Add("X-Knack-REST-API-Key", App_KEY);
                //The length of the buffer (postvars) is used as contentlength.
                WebReq.ContentLength = buffer.Length;
                //We open a stream for writing the postvars
                Stream PostData = WebReq.GetRequestStream();
                //Now we write, and afterwards, we close. Closing is always important!
                PostData.Write(buffer, 0, buffer.Length);
                PostData.Close();
                //Get the response handle, we have no true response yet!
                HttpWebResponse WebResp = (HttpWebResponse)WebReq.GetResponse();
                //Now, we read the response (the string), and output it.
                Stream Answer = WebResp.GetResponseStream();
                StreamReader _Answer = new StreamReader(Answer);
                vystup = _Answer.ReadToEnd();
                JObject jObject = JObject.Parse(vystup);

                vystup = jObject["records"][0]["id"].ToString();
                //Congratulations, you just requested your first POST page, you
                //can now start logging into most login forms, with your application
                //Or other examples.
            }
            catch (Exception ex)
            {
                //Console.WriteLine(ex.ToString());
            }
            return vystup.Trim() + "\n";
        }
        #endregion
    }
}
