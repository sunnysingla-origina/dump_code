using System;
using System.Collections.Generic;
using System.Configuration;
using System.Data;
using System.Data.SqlClient;
using System.Linq;
using System.Net;
using System.Net.Mail;
using System.Text;
using System.Threading.Tasks;

namespace ConsoleApp1
{
    public static class DataHelper
    {
        static DataHelper()
        {
            //
            // TODO: Add constructor logic here
            //
        }
        public static SqlConnection OpenCon(SqlConnection _cnn)
        {
            if (_cnn.State == ConnectionState.Closed)
            {
                try
                {
                    _cnn.ConnectionString = ConfigClass.con;
                    _cnn.Open();
                }
                catch (Exception)
                {
                    throw;

                }
            }
            return _cnn;
        }

        public static string conString()
        {
            return ConfigClass.con;
        }

        public static void CloseCon(SqlConnection _cnn)
        {
            if (_cnn.State == ConnectionState.Open)
            {
                try
                {
                    _cnn.Close();
                }
                catch (Exception)
                {
                    throw;
                }
            }
        }

        public static void DisposeCon(SqlConnection _cnn)
        {
            if (_cnn.State == ConnectionState.Open)
            {
                try
                {
                    _cnn.Dispose();
                }
                catch (Exception)
                {
                    throw;
                }
            }

        }

        public static DataTable getDataTable(string query)
        {
            DataTable dt = new DataTable();
            SqlConnection con = new SqlConnection();
            try
            {
                SqlDataAdapter da = new SqlDataAdapter(query, DataHelper.OpenCon(con));
                if (ConfigurationManager.AppSettings["commandTimeout"] != null)
                    da.SelectCommand.CommandTimeout = Convert.ToInt16(ConfigurationManager.AppSettings["commandTimeout"].ToString());
                da.Fill(dt);
            }
            catch (Exception)
            {
                throw;
            }
            finally
            {
                //if (con.State != ConnectionState.Closed)
                DataHelper.CloseCon(con);

            }
            return dt;
        }

        public static DataTable getDataTableExecuteSP(string strStoredProc, SqlParameter[] sqlParams)
        {
            SqlConnection con = new SqlConnection();
            OpenCon(con);
            SqlCommand SqlComm = new SqlCommand(strStoredProc, con);
            SqlComm.CommandType = CommandType.StoredProcedure;
            SqlComm.CommandTimeout = 2000;
            SqlDataAdapter da = new SqlDataAdapter(SqlComm);
            DataTable dt = new DataTable();

            if (sqlParams != null)
            {
                foreach (SqlParameter sqlParam in sqlParams)
                {

                    SqlComm.Parameters.Add(sqlParam);
                }
            }

            try
            {
                da.Fill(dt);
            }
            catch (Exception)
            {
                throw;
            }
            finally
            {
                DisposeCon(con);
                SqlComm.Dispose();
                sqlParams = null;
            }
            return dt;
        }

        public static int execQuery(string query)
        {
            int i = -1;
            SqlConnection con = new SqlConnection();
            SqlCommand com = null;
            try
            {
                DataHelper.OpenCon(con);
                com = new SqlCommand(query, con);
                i = com.ExecuteNonQuery();

            }
            catch (Exception)
            {
                throw;
            }
            finally
            {
                DataHelper.DisposeCon(con);
                com.Dispose();
            }
            return i;
        }

        public static DataSet getDataSet(string query)
        {
            DataSet ds = new DataSet();
            SqlConnection con = new SqlConnection();
            try
            {
                SqlDataAdapter da = new SqlDataAdapter(query, DataHelper.OpenCon(con));
                if (ConfigurationManager.AppSettings["commandTimeout"] != null)
                    da.SelectCommand.CommandTimeout = Convert.ToInt16(ConfigurationManager.AppSettings["commandTimeout"].ToString());
                da.Fill(ds);
            }
            catch (Exception)
            {
                throw;
            }
            finally
            {
                DisposeCon(con);
            }
            return ds;
        }

        public static DataSet getDataSetExecuteSP(string strStoredProc, SqlParameter[] sqlParams)
        {
            SqlConnection con = new SqlConnection();
            OpenCon(con);
            SqlCommand SqlComm = new SqlCommand(strStoredProc, con);
            SqlComm.CommandType = CommandType.StoredProcedure;
            SqlComm.CommandTimeout = 2000;
            SqlDataAdapter da = new SqlDataAdapter(SqlComm);
            DataSet ds = new DataSet();

            if (sqlParams != null)
            {

                foreach (SqlParameter sqlParam in sqlParams)
                {

                    SqlComm.Parameters.Add(sqlParam);
                }
            }

            try
            {
                da.Fill(ds);
            }
            catch (Exception)
            {
                throw;
            }
            finally
            {
                DisposeCon(con);
                SqlComm.Dispose();
                sqlParams = null;
            }
            return ds;

        }

        public static int ExecuteSP(string strStoredProc, SqlParameter[] sqlParams)
        {
            SqlConnection con = new SqlConnection();
            OpenCon(con);
            SqlCommand SqlComm = new SqlCommand(strStoredProc, con);
            SqlComm.CommandType = CommandType.StoredProcedure;
            int result = -1;
            if (sqlParams != null)
            {

                foreach (SqlParameter sqlParam in sqlParams)
                {

                    SqlComm.Parameters.Add(sqlParam);
                }
            }

            try
            {
                result = SqlComm.ExecuteNonQuery();
            }
            catch (Exception)
            {
                throw;
            }
            finally
            {
                DisposeCon(con);
                SqlComm.Dispose();
                sqlParams = null;
            }
            return result;

        }

        /// <summary>
        /// execute sp with custom connection
        /// </summary>
        /// <param name="strStoredProc"></param>
        /// <param name="sqlParams"></param>
        /// <returns></returns>
        public static int ExecuteSP(string strStoredProc, SqlParameter[] sqlParams, string ConnectionName)
        {
            String connectionString = ConfigurationManager.ConnectionStrings[ConnectionName].ConnectionString.ToString();

            SqlConnection con = new SqlConnection(connectionString);
            //  OpenCon(con);
            if (con.State == ConnectionState.Closed)
            {
                try
                {
                    // con.ConnectionString = ConfigClass.con;
                    con.Open();
                }
                catch (Exception)
                {
                    throw;

                }
            }

            SqlCommand SqlComm = new SqlCommand(strStoredProc, con);
            SqlComm.CommandType = CommandType.StoredProcedure;
            int result = -1;
            if (sqlParams != null)
            {
                foreach (SqlParameter sqlParam in sqlParams)
                {

                    SqlComm.Parameters.Add(sqlParam);
                }
            }

            try
            {
                result = SqlComm.ExecuteNonQuery();
            }
            catch (Exception)
            {
                throw;
            }
            finally
            {
                DisposeCon(con);
                SqlComm.Dispose();
                sqlParams = null;
            }
            return result;

        }

        public static int ExecuteSPWithExecuteScalar(string strStoredProc, SqlParameter[] sqlParams)
        {
            SqlConnection con = new SqlConnection();
            OpenCon(con);
            SqlCommand SqlComm = new SqlCommand(strStoredProc, con);
            SqlComm.CommandType = CommandType.StoredProcedure;
            int result = -1;
            if (sqlParams != null)
            {

                foreach (SqlParameter sqlParam in sqlParams)
                {

                    SqlComm.Parameters.Add(sqlParam);
                }
            }

            try
            {
                result = Convert.ToInt32(SqlComm.ExecuteScalar());
            }
            catch (Exception)
            {
                throw;
            }
            finally
            {
                DisposeCon(con);
                SqlComm.Dispose();
                sqlParams = null;
            }
            return result;

        }

        public static DataSet getDataSetExecuteSPNoParam(string strStoredProc)
        {
            SqlConnection con = new SqlConnection();
            OpenCon(con);
            SqlCommand SqlComm = new SqlCommand(strStoredProc, con);
            SqlComm.CommandType = CommandType.StoredProcedure;
            // SqlComm.CommandTimeout = 2000;
            SqlDataAdapter da = new SqlDataAdapter(SqlComm);
            DataSet ds = new DataSet();

            try
            {
                da.Fill(ds);
            }
            catch (Exception)
            {
                throw;
            }
            finally
            {
                DisposeCon(con);
                SqlComm.Dispose();
                con.Close();

            }
            return ds;

        }

        public static DataTable getDataTableExecuteSPNoParam(string strStoredProc)
        {
            SqlConnection con = new SqlConnection();
            OpenCon(con);
            SqlCommand SqlComm = new SqlCommand(strStoredProc, con);
            SqlComm.CommandType = CommandType.StoredProcedure;
            SqlComm.CommandTimeout = 2000;
            SqlDataAdapter da = new SqlDataAdapter(SqlComm);
            DataTable dt = new DataTable();


            try
            {
                da.Fill(dt);
            }
            catch
            {
                throw;
            }
            finally
            {
                DisposeCon(con);
                SqlComm.Dispose();



            }
            return dt;
            // if null then fetch from the database

        }

        public static int execQueryScalar(string query)
        {
            int i = -1;
            SqlConnection con = new SqlConnection();
            try
            {
                DataHelper.OpenCon(con);
                SqlCommand com = new SqlCommand(query, con);
                i = Convert.ToInt32(com.ExecuteScalar());

            }
            catch (Exception)
            {
                throw;
            }
            finally
            {
                if (con.State != ConnectionState.Closed)
                    DataHelper.CloseCon(con);
                con = null;
            }
            return i;
        }

        public static void SendScriptMail(string errormsg, string subject, string EmailBody)
        {
            string AddEmailIDs = string.Empty;
            string AddEmailIDsCC = string.Empty;
            String fromEmailId = string.Empty;

            try
            {
                string Body = string.Empty;
                MailMessage Mesage2 = new MailMessage();

                SmtpClient smtp2 = new SmtpClient(ClsEmailServiceData.SmtpClient);
                smtp2.Credentials = new NetworkCredential(ClsEmailServiceData.SmtpUserID, ClsEmailServiceData.SmtpPassword);

                fromEmailId = ClsEmailServiceData.FromEmailID;
                if (string.IsNullOrEmpty(fromEmailId))
                    fromEmailId = "CaptureMostRecentLiDocumentForIBM@support.com";
                Mesage2.From = new MailAddress(fromEmailId);

                if (errormsg == "No")
                {
                    Body = EmailBody;
                    AddEmailIDs = ClsEmailServiceData.AddEmailIDs;
                    AddEmailIDsCC = ClsEmailServiceData.AddCCEmailIDs;
                    if (string.IsNullOrEmpty(subject))
                    subject = ClsEmailServiceData.SucessSubject;
                }
                else
                {
                    Body = errormsg;
                    AddEmailIDs = ClsEmailServiceData.ErrorAddEmailIDs;
                    if (string.IsNullOrEmpty(subject))
                    {
                        subject = ClsEmailServiceData.ErrorSubject;
                    }
                }
                if (string.IsNullOrEmpty(AddEmailIDs))
                {
                    AddEmailIDs = ConfigurationManager.AppSettings["AddEmailIDs"];
                }

                if (AddEmailIDs.Trim() != "")
                {
                    AddEmailIDs = AddEmailIDs.Replace("\r", "").Replace("\n", "").Replace(" ", "");
                    string[] bccIds = AddEmailIDs.Replace(',', ';').Split(';');
                    for (int ctrBCC = 0; ctrBCC < bccIds.Length; ctrBCC++)
                    {
                        Mesage2.To.Add(new MailAddress(bccIds[ctrBCC]));
                    }
                }

                if (AddEmailIDsCC.Trim() != "")
                {
                    AddEmailIDsCC = AddEmailIDsCC.Replace("\r", "").Replace("\n", "").Replace(" ", "");
                    string[] bccIds = AddEmailIDsCC.Replace(',', ';').Split(';');
                    for (int ctrBCC = 0; ctrBCC < bccIds.Length; ctrBCC++)
                    {
                        Mesage2.CC.Add(new MailAddress(bccIds[ctrBCC]));
                    }
                }

                Mesage2.IsBodyHtml = true;
                Mesage2.Body = Body;
                Mesage2.Subject = subject;
                Mesage2.Priority = MailPriority.High;
                smtp2.DeliveryMethod = SmtpDeliveryMethod.Network;
                smtp2.Send(Mesage2);

            }
            catch
            {

            }
        }

        /// <summary>
        /// Class EmailServiceData
        /// </summary>
        /// <param name="status">Add All Email Infromation and Store globle </param>
        /// <param name="message">Message</param>
        public static class ClsEmailServiceData
        {
            static ClsEmailServiceData()
            {
                AddEmailIDs = ConfigurationManager.AppSettings["AddEmailIDs"];
                FromEmailID = ConfigurationManager.AppSettings["fromEmailID"];
                AddCCEmailIDs = ConfigurationManager.AppSettings["AddCCEmailIDs"];
                ErrorAddEmailIDs = ConfigurationManager.AppSettings["ErrorAddEmailIDs"];
                ErrorSubject = ConfigurationManager.AppSettings["ErrorSubject"];
                SucessSubject = ConfigurationManager.AppSettings["SucessSubject"];
                SmtpClient = ConfigurationManager.AppSettings["smtpClient"];
                SmtpUserID = ConfigurationManager.AppSettings["SmtpUserID"];
                SmtpPassword = ConfigurationManager.AppSettings["SmtpPWD"];

            }
            public static String AddEmailIDs
            {
                get;
                set;
            }
            public static String FromEmailID
            {
                get;
                set;
            }
            public static String ErrorAddEmailIDs
            {
                get;
                set;
            }
            public static String AddCCEmailIDs
            {
                get;
                set;
            }
            public static String ErrorSubject
            {
                get;
                set;
            }
            public static String SucessSubject
            {
                get;
                set;
            }
            public static String SmtpClient
            {
                get;
                set;
            }
            public static String SmtpUserID
            {
                get;
                set;
            }
            public static String SmtpPassword
            {
                get;
                set;
            }

        }

    }
}
