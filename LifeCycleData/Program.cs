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
using OpenQA.Selenium.Interactions;
using System.Data.SqlClient;
using System.Net.Mail;

namespace ConsoleApp1
{
    class Program
    {
        static string App_ID = ConfigurationManager.AppSettings["App_ID"].ToString();
        static string App_KEY = ConfigurationManager.AppSettings["App_KEY"].ToString();
        static string FetLifeCycle = "https://eu-api.knack.com/v1/objects/object_64/records?page=1&rows_per_page=25&sort_field=field_2345&sort_order=desc&filters=%7B%22match%22%3A%22and%22%2C%22rules%22%3A%5B%7B%22field%22%3A%22field_2345%22%2C%22operator%22%3A%22is%22%2C%22value%22%3A%22[IBMPID]%22%2C%22field_name%22%3A%22IBMID_Connection%22%7D%2C%7B%22match%22%3A%22and%22%2C%22field%22%3A%22field_2057%22%2C%22operator%22%3A%22is%22%2C%22value%22%3A%22[KEY4]%22%2C%22field_name%22%3A%22synKey4%22%7D%5D%7D&_=1529604135745";
        static string IBMPID = "https://eu-api.knack.com/v1/objects/object_139/records?page=1&rows_per_page=25&filters=%7B%22match%22%3A%22and%22%2C%22rules%22%3A%5B%7B%22field%22%3A%22field_2049%22%2C%22operator%22%3A%22is%22%2C%22value%22%3A%22[ID]%22%2C%22field_name%22%3A%22IBMPID%22%7D%5D%7D&_=1529607191975";
        static string OroginaSWBD = "https://eu-api.knack.com/v1/objects/object_20/records?page=1&rows_per_page=1000&sort_field=field_2051-field_2051&sort_order=desc&filters=%7B%22match%22%3A%22and%22%2C%22rules%22%3A%5B%7B%22field%22%3A%22field_444%22%2C%22operator%22%3A%22is%22%2C%22value%22%3A%22[PID]%22%2C%22field_name%22%3A%22ibm_pid%22%7D%5D%7D&_=1529659943387";

        static string BackupOldData = "https://eu-api.knack.com/v1/scenes/scene_1384/views/view_2463/records/";

        static string FetLifeCycle_versoning = "https://eu-api.knack.com/v1/objects/object_216/records?page=1&rows_per_page=1000";

        static string getVersoningFromLifeCycle = "https://eu-api.knack.com/v1/scenes/scene_2344/views/view_4078/records?page=1&rows_per_page=1000&sort_field=field_826&sort_order=desc";


        static List<string> ListIds = new List<string>();
        static DataTable Dt2 = new DataTable();
        static DataTable Dt3 = new DataTable();
        public static DataTable dt_TableData = new DataTable();

        static DateTime starttime = new DateTime();
        static DateTime endtime = new DateTime();

        static void Main(string[] args)
        {
            System.Net.ServicePointManager.SecurityProtocol = SecurityProtocolType.Tls12 | SecurityProtocolType.Tls11 | SecurityProtocolType.Tls;
            clsBLL objbll = new clsBLL();
            WebDriverWait wait = null;

            try
            {                
                starttime = DateTime.Now;
                SqlConnection conn = new SqlConnection(ConfigurationManager.ConnectionStrings["connectionstring"].ToString());
                SqlCommand cmd = new SqlCommand("USP_Truncatetable", conn);
                cmd.CommandType = CommandType.StoredProcedure;
                SqlDataAdapter da = new SqlDataAdapter();
                da.SelectCommand = cmd;
                da.Fill(dt_TableData);

                DataTable Dt1 = new DataTable();
                Dt1.Columns.Add("ScrappingId");
                Dt1.Columns.Add("Recordid");


                List<string> Columns1 = new List<string>();
                Columns1.Add("field_2049");
                Columns1.Add("id");

                GET(1, ConfigurationManager.AppSettings["object_139"].ToString(), Dt1, Columns1);

                objbll.BulkInsertToDataBase(Dt1, "tbl_Scrapping");

                DataTable Dt = new DataTable();
                Dt.Columns.Add("RCDID");
                Dt.Columns.Add("Uniqueness");
                Dt.Columns.Add("SupportEnd");
                Dt.Columns.Add("Versioning");

                List<string> Columns = new List<string>();
                Columns.Add("id");
                Columns.Add("field_3628");
                Columns.Add("field_2058");
                Columns.Add("field_844");

                GET(1, ConfigurationManager.AppSettings["object_64"].ToString(), Dt, Columns);

                //-----------------------------------------
                DataTable dt = new DataTable();
                
                SqlCommand cmd1 = new SqlCommand("USP_Scrapping", conn);
                cmd1.CommandType = CommandType.StoredProcedure;
                SqlDataAdapter da1 = new SqlDataAdapter();
                da1.SelectCommand = cmd1;
                da1.Fill(dt);

                var chromeDriverService = ChromeDriverService.CreateDefaultService();
                chromeDriverService.HideCommandPromptWindow = true;               

                ChromeOptions optionsChrome = new ChromeOptions();
                List<string> lst = new List<string>();
                optionsChrome.AddArgument("--start-maximized");
                IWebDriver driver = new ChromeDriver(chromeDriverService, optionsChrome);
                int errorCountt = 0;
                if (dt.Rows.Count > 0)
                {                                        
                    int x = 1;

                    try
                    {
                        int ik = 0;
                        for (int i = 0; i < dt.Rows.Count; i++)
                        {                        
                            if (x == 0)
                                driver = new ChromeDriver(chromeDriverService, optionsChrome);

                            x = 1;
                            string siteUrl = ConfigurationManager.AppSettings["URL"].ToString();
                            siteUrl = siteUrl.Replace("q=", "q=" + dt.Rows[i][0].ToString());

                            wait = new WebDriverWait(driver, TimeSpan.FromSeconds(500));
                            driver.Navigate().GoToUrl(siteUrl);
                            wait.Until(driver1 => ((IJavaScriptExecutor)driver).ExecuteScript("return document.readyState").Equals("complete"));
                            try
                            {
                                //IWebElement detailFrame = driver.FindElement(By.XPath("//iframe"));
                                IWebDriver driver1 = driver;
                                driver1.SwitchTo().Frame(Convert.ToInt32(ConfigurationManager.AppSettings["FrameIndex"]));

                                //alternatively, find the frame first, then use GetAttribute() to get frame name
                                try
                                {
                                    //you are now in iframe "Details", then find the elements you want in the frame now
                                    Thread.Sleep(2000);
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

                            try
                            {
                                wait.Until(driver1 => ((IJavaScriptExecutor)driver).ExecuteScript("return document.readyState").Equals("complete"));
                                SelectElement drpPageCount = new SelectElement(driver.FindElement(By.XPath("//*[@id='ibm-plc-results-table_length']/label/select")));
                                drpPageCount.SelectByValue("100");

                                IWebElement table_element = driver.FindElement(By.Id("ibm-plc-results-table"));

                                table_element.FindElement(By.XPath("//*[@id='ibm-plc-results-table']/thead/tr/th[6]/a")).Click();
                                wait = new WebDriverWait(driver, TimeSpan.FromSeconds(1));
                                table_element.FindElement(By.XPath("//*[@id='ibm-plc-results-table']/thead/tr/th[6]/a")).Click();

                                IList<IWebElement> tableRow = table_element.FindElements(By.XPath("//*[@id='ibm-plc-results-table']/tbody/tr"));
                                IList<IWebElement> rowTD;
                                int t = 0;

                                foreach (IWebElement row in tableRow)
                                {
                                    rowTD = row.FindElements(By.TagName("td"));
                                    string Uniqueness = rowTD[4].Text.Replace("-", "") + "-" + rowTD[1].FindElements(By.TagName("a"))[0].GetAttribute("href").Split('=')[1];
                                    string SupportEnd = rowTD[6].Text;
                                    string Version = rowTD[2].Text;
                                    if (rowTD[6].Text != "")
                                        SupportEnd = Convert.ToDateTime(rowTD[6].Text).ToString("dd/MM/yyyy");


                                    var result = Dt.AsEnumerable().Where(myRow => myRow.Field<string>("Uniqueness").ToLower() == Uniqueness.ToLower() && myRow.Field<string>("SupportEnd") == SupportEnd && myRow.Field<string>("Versioning") == Version);
                                    DataTable view = result.AsDataView().ToTable().Copy();

                                    if (view.Rows.Count == 0)
                                    {
                                        ik++;
                                        int found = 0;
                                        string fo = rowTD[0].FindElement(By.TagName("label")).GetAttribute("for").ToString();
                                        for (int f = 0; f < lst.Count; f++)
                                        {
                                            if (fo == lst[f])
                                            {
                                                found = 1;
                                            }
                                        }
                                        if (found == 0)
                                        {
                                            lst.Add(fo);
                                            rowTD[0].FindElement(By.TagName("label")).Click();
                                        }


                                    }

                                    t++;

                                }
                            }
                            catch (Exception ex)
                            {
                                //SendEmail(ConfigurationManager.AppSettings["ErrorAddEmailIDs"].ToString(), ConfigurationManager.AppSettings["ErrorSubject"].ToString(), ex.ToString());
                                errorCountt++;
                                if (errorCountt <= int.Parse(ConfigurationManager.AppSettings["ErrorCount"].ToString()))
                                {
                                    //i--;
                                   // x = 0;
                                }
                                else {
                                    SendEmail(ConfigurationManager.AppSettings["ErrorAddEmailIDs"].ToString(), "Driver issue or no record found in query"+ ConfigurationManager.AppSettings["ErrorSubject"].ToString(), ex.ToString());
                                }
                                //driver.Close();
                            }

                            if (ik > 80 || i + 1 == dt.Rows.Count)
                            {
                                if (ik == 0)
                                {

                                    string HTML = "<table><tr><td colspan='2'>Hello team,</td></tr><tr><td colspan='2'>Script run successfully! Please check details.</td></tr><tr><td></td><td></td></tr>";
                                    HTML += "<tr><td>Start time</td><td>" + starttime + "</td></tr><tr><td></td><td></td></tr><tr><td>End time</td><td>" + DateTime.Now.ToString() + "</td></tr>";
                                    HTML += "<tr><td></td><td></td></tr><tr><td>Inserted Records</td><td>0</td></tr><tr><td></td><td></td></tr>";
                                    HTML += "<tr><td>Updated Records</td><td>0</td></tr><tr><td colspan='2'></td></tr><tr><td colspan='2'>";
                                    HTML += "Thanks,</td></tr><tr><td colspan='2'>Software Team</td></tr></table>";

                                    SendEmail(ConfigurationManager.AppSettings["AddEmailIDs"].ToString(),"No more record found "+ ConfigurationManager.AppSettings["SucessSubject"].ToString(), HTML.ToString());
                                }
                                else
                                {
                                    ik = 0;
                                    driver.Navigate().GoToUrl(ConfigurationManager.AppSettings["DownloadURL"].ToString());
                                    //wait = new WebDriverWait(driver, TimeSpan.FromSeconds(120));
                                    wait.Until(driver1 => ((IJavaScriptExecutor)driver).ExecuteScript("return document.readyState").Equals("complete"));

                                    driver.FindElement(By.Name("ibm-submit")).Click();

                                    string strXmlUrl = driver.Url;
                                    DemonstrateReadWriteXMLDocumentWithString_Daily(strXmlUrl);
                                    x = 0;
                                    if (i + 1 != dt.Rows.Count)
                                        driver.Close();
                                }
                            }
                        }
                        driver.Close();
                    }
                    catch (Exception ex)
                    {
                        SendEmail(ConfigurationManager.AppSettings["AddEmailIDs"].ToString(),ConfigurationManager.AppSettings["ErrorSubject"].ToString(), ex.ToString());
                        driver.Close();
                    }
                }
            }
            catch (Exception ex)
            {
                SendEmail(ConfigurationManager.AppSettings["AddEmailIDs"].ToString(), "Final catch " + ConfigurationManager.AppSettings["ErrorSubject"].ToString(), ex.ToString());                
            }
        }

        public void BulkInsertToDataBase(DataTable dt, String TableName)
        {
            try
            {
                SqlConnection con = new SqlConnection();
                DataHelper.OpenCon(con);
                //creating object of SqlBulkCopy  
                SqlBulkCopy objbulk = new SqlBulkCopy(con);
                //assigning Destination table name             
                objbulk.DestinationTableName = TableName;
                objbulk.BulkCopyTimeout = 3600;
                objbulk.BatchSize = 5000;
                foreach (DataColumn item in dt.Columns)
                {
                    item.ColumnName = item.ColumnName.TrimEnd();
                    string columnName = item.ColumnName;
                    if (!string.IsNullOrEmpty(columnName))
                    {
                        objbulk.ColumnMappings.Add(columnName, columnName);
                    }
                }

                objbulk.WriteToServer(dt);
                DataHelper.CloseCon(con);


            }
            catch (Exception ex)
            {
                throw ex;
            }

        }

        public static void GET(int page, string GetURl, DataTable Dt, List<string> Columns)
        {
            string vystup = null;
            try
            {
                string url = GetURl.Replace("page=1", "page=" + page.ToString());
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
                    for (int y = 0; y < Columns.Count; y++)
                    {
                        if (Columns[y].IndexOf("raw_id") > 0)
                        {
                            try
                            {
                                dr[y] = jObject["records"][x][Columns[y].Replace("raw_id", "raw")][0]["id"];
                            }
                            catch
                            {
                                dr[y] = 0;
                            }
                        }
                        else if (Columns[y].IndexOf("_iso_timestamp") > 0)
                        {
                            try
                            {
                                dr[y] = jObject["records"][x][Columns[y].Replace("iso_timestamp", "raw")]["timestamp"].ToString().Substring(0, 10).TrimEnd();
                            }
                            catch
                            {
                                dr[y] = DateTime.Now;
                            }
                        }
                        else if (Columns[y].IndexOf("raw_role") > 0)
                        {
                            try
                            {
                                dr[y] = jObject["records"][x][Columns[y].Replace("raw_role", "raw")];
                            }
                            catch
                            {
                                dr[y] = 0;
                            }
                        }
                        else if (Columns[y].IndexOf("raw_name") > 0)
                        {
                            try
                            {
                                dr[y] = jObject["records"][x][Columns[y].Replace("raw_name", "raw")][0]["identifier"];
                            }
                            catch
                            {
                                dr[y] = 0;
                            }
                        }
                        else if (Columns[y].IndexOf("raw_email") > 0)
                        {
                            try
                            {
                                dr[y] = jObject["records"][x][Columns[y].Replace("raw_email", "raw")]["email"];
                            }
                            catch
                            {
                                dr[y] = 0;
                            }
                        }
                        else if (Columns[y].IndexOf("raw_fname") > 0)
                        {
                            try
                            {
                                dr[y] = jObject["records"][x][Columns[y].Replace("raw_fname", "raw")]["first"];
                            }
                            catch
                            {
                                dr[y] = 0;
                            }
                        }
                        else if (Columns[y].IndexOf("raw_lname") > 0)
                        {
                            try
                            {
                                dr[y] = jObject["records"][x][Columns[y].Replace("raw_lname", "raw")]["last"];
                            }
                            catch
                            {
                                dr[y] = 0;
                            }
                        }
                        else
                        {
                            dr[y] = jObject["records"][x][Columns[y]];
                        }
                    }
                    Dt.Rows.Add(dr);

                    Console.WriteLine(jObject["records"][x]["field_2049"]);


                }
                if (jObject["total_pages"].ToString() != jObject["current_page"].ToString())
                {
                    GET(page + 1, GetURl, Dt, Columns);
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

        private static void DemonstrateReadWriteXMLDocumentWithString_Daily(string url)
        {
     
            int InsertCount = 0;
            int UpdateCount = 0;
            //DataTable table = CreateTestTable("XmlDemo");
            // PrintValues(table, "Original table");
            // Create a request for the URL.   
            WebRequest request = WebRequest.Create(
              url);
            // If required by the server, set the credentials.  
            request.Credentials = CredentialCache.DefaultCredentials;
            // Get the response.  
            WebResponse response = request.GetResponse();
            // Display the status.  
            Console.WriteLine(((HttpWebResponse)response).StatusDescription);
            // Get the stream containing content returned by the server.  
            Stream dataStream = response.GetResponseStream();
            // Open the stream using a StreamReader for easy access.  
            StreamReader reader = new StreamReader(dataStream);
            // Read the content.  
            string responseFromServer = reader.ReadToEnd();
            // Display the content.  
            File.WriteAllText(@"XMLFile/PLCDailyXMLDownload.xml", responseFromServer);

            // Open the file to read from.


            // Clean up the streams and the response.  
            reader.Close();
            response.Close();



            string fileName = "XMLFile/PLCDailyXMLDownload.xml";
            //table.WriteXml(fileName, XmlWriteMode.WriteSchema);

            DataSet ds = new DataSet();
            ds.ReadXml(fileName);
            if (ds.Tables.Count <= 2)
            {
                Console.WriteLine("No Data found");
                return;
            }
            var results = from table1 in ds.Tables[1].AsEnumerable()
                          join table2 in ds.Tables[2].AsEnumerable() on (int)table1[2] equals (int)table2[1]
                          join table3 in ds.Tables[4].AsEnumerable() on (int)table2[1] equals (int)table3[5]
                          join table4 in ds.Tables[6].AsEnumerable() on (int)table3[1] equals (int)table4[5]
                          join table5 in ds.Tables[7].AsEnumerable() on (int)table4[2] equals (int)table5[8]
                          select new
                          {
                              Key1 = (string)table1["synKey"],
                              key2 = (string)table2["synKey"],
                              key3 = (string)table3["synKey"],
                              key4 = (string)table4["synKey"],
                              Title = (string)table1[1],
                              Type = (string)table3["offeringType"],
                              Pid = ((string)table3["pid"]),
                              url = (string)table4["PLCMoreInfoURL"],
                              vNum = (string)table3["versionNumber"],
                              rNum = (string)table4["releaseNumber"],
                              Level = (string)table4["modLevelNumber"],
                              lPolicy = (string)table5["lifecyclePolicy"],
                              exc = (string)table5["exception"],
                              eosAnnLetNo = (string)table5["eosAnnLetNo"],
                              eosDate = (((string)table5["eosDate"] != "") ? Convert.ToDateTime((string)table5["eosDate"]).ToString("dd/MM/yyyy") : ""),
                              eomAnnLetNo = (string)table5["eomAnnLetNo"],
                              eomDate = (((string)table5["eomDate"] != "") ? Convert.ToDateTime((string)table5["eomDate"]).ToString("dd/MM/yyyy") : ""),
                              qaAnnLetNo = (string)table5["gaAnnLetNo"],
                              qaDate = (((string)table5["gaDate"] != "") ? Convert.ToDateTime((string)table5["gaDate"]).ToString("dd/MM/yyyy") : ""),
                          };






            Console.WriteLine(results.Count().ToString());
            foreach (var item in results)
            {

                DateTime Date1 = new DateTime(2011, 1, 1);
                DateTime.TryParse(item.eosDate.ToString(), out Date1);

                DateTime Date2 = new DateTime(2011, 1, 1);
                DateTime.TryParse(item.eosDate.ToString(), out Date2);

                DateTime Date3 = new DateTime(2011, 1, 1);
                DateTime.TryParse(item.eosDate.ToString(), out Date3);

                DateTime dtfinal = new DateTime();

                dtfinal = ((Date1 >= Date2) ? ((Date1 >= Date3) ? Date1 : Date3) : ((Date2 >= Date3) ? Date2 : Date3));


                string RecordID = GET(FetLifeCycle.Replace("[IBMPID]", item.Pid.Replace("-", "").Replace(" ", "")).Replace("[KEY4]", item.key4), "");

                string PID = GET(IBMPID.Replace("[ID]", item.Pid.Replace("-", "").Replace(" ", "")), "");

                string postObject1 = "field_3562=" + item.vNum + "." + item.rNum + "." + item.Level;
                try
                {
                    Post("https://eu-api.knack.com/v1/objects/object_216/records", postObject1, "POST", Dt2, item.vNum + "." + item.rNum + "." + item.Level);
                }
                catch
                { }

                if (RecordID == "")
                {

                    if (PID == "")
                    {
                        string postObject = "field_2049=" + item.Pid.Replace("-", "").Replace(" ", "");

                        PID = Post("https://eu-api.knack.com/v1/objects/object_139/records", postObject, "POST");

                    }

                }

                string postObject_Main = "";
                List<string> _listSWBDId = GETList(OroginaSWBD.Replace("[PID]", item.Pid.Replace("-", "").Replace(" ", "")), "");

                for (int x = 0; x < _listSWBDId.Count(); x++)
                {
                    postObject_Main = postObject_Main + ((postObject_Main == "") ? "" : "&") + "field_2238[" + x.ToString() + "]=" + _listSWBDId[x].ToString();
                }
                postObject_Main = postObject_Main + ((postObject_Main == "") ? "" : "&") + "field_2053[0]=" + PID;
                postObject_Main = postObject_Main + ((postObject_Main == "") ? "" : "&") + "field_824=" + item.Title;
                postObject_Main = postObject_Main + ((postObject_Main == "") ? "" : "&") + "field_825=" + item.Type;

                postObject_Main = postObject_Main + ((postObject_Main == "") ? "" : "&") + "field_826=" + item.vNum;
                postObject_Main = postObject_Main + ((postObject_Main == "") ? "" : "&") + "field_827=" + item.rNum;
                postObject_Main = postObject_Main + ((postObject_Main == "") ? "" : "&") + "field_830=" + item.exc;
                postObject_Main = postObject_Main + ((postObject_Main == "") ? "" : "&") + "field_828=" + item.Level;
                postObject_Main = postObject_Main + ((postObject_Main == "") ? "" : "&") + "field_829=" + item.lPolicy;

                // postObject_Main = postObject_Main + ((postObject_Main == "") ? "" : "&") + "field_831=" + item.eomAnnLetNo;
                postObject_Main = postObject_Main + ((postObject_Main == "") ? "" : "&") + "field_833=" + item.eomAnnLetNo;
                postObject_Main = postObject_Main + ((postObject_Main == "") ? "" : "&") + "field_834=" + item.eomDate;

                postObject_Main = postObject_Main + ((postObject_Main == "") ? "" : "&") + "field_831=" + item.eosAnnLetNo;
                postObject_Main = postObject_Main + ((postObject_Main == "") ? "" : "&") + "field_2058=" + item.eosDate;

                postObject_Main = postObject_Main + ((postObject_Main == "") ? "" : "&") + "field_835=" + item.qaAnnLetNo;
                postObject_Main = postObject_Main + ((postObject_Main == "") ? "" : "&") + "field_836=" + item.qaDate;

                postObject_Main = postObject_Main + ((postObject_Main == "") ? "" : "&") + "field_837=" + item.url;

                postObject_Main = postObject_Main + ((postObject_Main == "") ? "" : "&") + "field_2054=" + item.Key1;
                postObject_Main = postObject_Main + ((postObject_Main == "") ? "" : "&") + "field_2055=" + item.key2;
                postObject_Main = postObject_Main + ((postObject_Main == "") ? "" : "&") + "field_2056=" + item.key3;
                postObject_Main = postObject_Main + ((postObject_Main == "") ? "" : "&") + "field_2057=" + item.key4;
                postObject_Main = postObject_Main + ((postObject_Main == "") ? "" : "&") + "field_2361=" + DateTime.Now.ToString("dd/MM/yyyy");

                //unComment below Line later
                if (RecordID == "")
                {
                    InsertCount++;
                }
                else
                {
                    UpdateCount++;
                }
                if (RecordID != "")
                {
                    string DataforBackup = "field_2392=" + RecordID;
                    Post(BackupOldData, DataforBackup, "POST");
                }
                Post("https://eu-api.knack.com/v1/objects/object_64/records" + ((RecordID == "") ? "" : "/" + RecordID), postObject_Main, ((RecordID == "") ? "POST" : "PUT"));
            }

            endtime = DateTime.Now;

            string HTML = "<table><tr><td colspan='2'>Hello team,</td></tr><tr><td colspan='2'>Script run successfully! Please check details.</td></tr><tr><td></td><td></td></tr>";
            HTML += "<tr><td>Start time</td><td>" + starttime + "</td></tr><tr><td></td><td></td></tr><tr><td>End time</td><td>" + endtime + "</td></tr>";
            HTML += "<tr><td></td><td></td></tr><tr><td>Inserted Records</td><td>" + InsertCount.ToString() + "</td></tr><tr><td></td><td></td></tr>";
            HTML += "<tr><td>Updated Records</td><td>" + UpdateCount.ToString() + "</td></tr><tr><td colspan='2'></td></tr><tr><td colspan='2'>";
            HTML += "Thanks,</td></tr><tr><td colspan='2'>Software Team</td></tr></table>";
            
            Post("https://eu-central-1-builder-write.knack.com/v1/objects/object_276/records", "field_4048=[PI] IBMLifeCycle&field_4049=" + starttime.ToString("MM/dd/yyyy h:mm tt") + "&field_4050=" + endtime.ToString("MM/dd/yyyy h:mm tt") + "&field_4052=" + InsertCount.ToString() + "&field_4054=" + UpdateCount.ToString(), "POST");
            SendEmail(ConfigurationManager.AppSettings["AddEmailIDs"].ToString(), ConfigurationManager.AppSettings["SucessSubject"].ToString(), HTML.ToString());

            // Print out values in the table.
            // PrintValues(newTable, "New table");
        }

        #region "get Post Methods"
        public static string GET(string url, string data)
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

                vystup = jObject["records"][0]["id"].ToString();
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
        public static DataTable Post(string url, string data, string Post_Put, DataTable dt, string value)
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
                DataRow dr = dt.NewRow();
                dr[0] = vystup;
                dr[1] = value;
                dt.Rows.Add(dr);
                //Congratulations, you just requested your first POST page, you
                //can now start logging into most login forms, with your application
                //Or other examples.
            }
            catch (Exception ex)
            {
                //Console.WriteLine(ex.ToString());
            }
            return dt;
        }

        public static void GET(int page, string GetURl, DataTable Dt, List<string> Columns, int removeduplicates)
        {
            string vystup = null;
            try
            {
                string url = GetURl.Replace("page=1", "page=" + page.ToString());
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
                    for (int y = 0; y < Columns.Count; y++)
                    {
                        if (Columns[y].IndexOf("raw_id") > 0)
                        {
                            try
                            {
                                dr[y] = jObject["records"][x][Columns[y].Replace("raw_id", "raw")][0]["id"];
                            }
                            catch
                            {
                                dr[y] = 0;
                            }
                        }
                        else if (Columns[y].IndexOf("raw_role") > 0)
                        {
                            try
                            {
                                dr[y] = jObject["records"][x][Columns[y].Replace("raw_role", "raw")];
                            }
                            catch
                            {
                                dr[y] = 0;
                            }
                        }
                        else if (Columns[y].IndexOf("raw_name") > 0)
                        {
                            try
                            {
                                dr[y] = jObject["records"][x][Columns[y].Replace("raw_name", "raw")][0]["identifier"];
                            }
                            catch
                            {
                                dr[y] = 0;
                            }
                        }
                        else if (Columns[y].IndexOf("raw_email") > 0)
                        {
                            try
                            {
                                dr[y] = jObject["records"][x][Columns[y].Replace("raw_email", "raw")]["email"];
                            }
                            catch
                            {
                                dr[y] = 0;
                            }
                        }
                        else if (Columns[y].IndexOf("raw_fname") > 0)
                        {
                            try
                            {
                                dr[y] = jObject["records"][x][Columns[y].Replace("raw_fname", "raw")]["first"];
                            }
                            catch
                            {
                                dr[y] = 0;
                            }
                        }
                        else if (Columns[y].IndexOf("raw_lname") > 0)
                        {
                            try
                            {
                                dr[y] = jObject["records"][x][Columns[y].Replace("raw_lname", "raw")]["last"];
                            }
                            catch
                            {
                                dr[y] = 0;
                            }
                        }
                        else
                        {
                            dr[y] = jObject["records"][x][Columns[y]];
                        }
                    }
                    int found = 0;
                    if (removeduplicates == 1)
                    {
                        for (int t = 0; t < Dt.Rows.Count; t++)
                        {
                            if (Dt.Rows[t]["Versoning"].ToString() == dr["Versoning"].ToString())
                                found = 1;
                        }
                    }
                    if (found == 0)
                        Dt.Rows.Add(dr);

                    //Console.WriteLine(jObject["records"][x]["field_2049"]);


                }
                if (jObject["total_pages"].ToString() != jObject["current_page"].ToString())
                {
                    GET(page + 1, GetURl, Dt, Columns, removeduplicates);
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
        #endregion


        // This function contains the logic to send mail.    
        public static void SendEmail(String ToEmail, String Subj, string Message)
        {
            //Reading sender Email credential from web.config file  

            string HostAdd = ConfigurationManager.AppSettings["smtpClient"].ToString();
            string FromEmailid = ConfigurationManager.AppSettings["SmtpUserID"].ToString();
            string Pass = ConfigurationManager.AppSettings["SmtpPWD"].ToString();

            //creating the object of MailMessage  
            MailMessage mailMessage = new MailMessage();
            mailMessage.From = new MailAddress(FromEmailid); //From Email Id  
            mailMessage.Subject = Subj; //Subject of Email  
            mailMessage.Body = Message; //body or message of Email  
            mailMessage.IsBodyHtml = true;

            string[] ToMuliId = ToEmail.Split(',');
            foreach (string ToEMailId in ToMuliId)
            {
                mailMessage.To.Add(new MailAddress(ToEMailId)); //adding multiple TO Email Id  
            }



            SmtpClient smtp = new SmtpClient();  // creating object of smptpclient  
            smtp.Host = HostAdd;              //host of emailaddress for example smtp.gmail.com etc  

            //network and security related credentials  

            smtp.EnableSsl = true;
            NetworkCredential NetworkCred = new NetworkCredential();
            NetworkCred.UserName = mailMessage.From.Address;
            NetworkCred.Password = Pass;
            smtp.UseDefaultCredentials = true;
            smtp.Credentials = NetworkCred;
           
            smtp.Port = Convert.ToInt32(ConfigurationManager.AppSettings["port"].ToString()); ;
            smtp.Send(mailMessage); //sending Email  
        }

    }
}