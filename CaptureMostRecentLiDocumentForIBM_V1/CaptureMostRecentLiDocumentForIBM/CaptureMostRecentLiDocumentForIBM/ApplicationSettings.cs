using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Data;
using System.Configuration;



namespace CaptureMostRecentLiDocumentForIBM
{

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
