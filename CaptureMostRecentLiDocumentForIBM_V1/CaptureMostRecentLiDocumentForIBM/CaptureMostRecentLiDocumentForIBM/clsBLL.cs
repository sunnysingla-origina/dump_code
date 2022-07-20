using System;
using System.Collections;
using System.Collections.Generic;
using System.Data;
using System.Data.SqlClient;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading.Tasks;


namespace CaptureMostRecentLiDocumentForIBM
{
    public class clsBLL
    {
        public DataTable GetIBMPIDDetails()
        {
            try
            {
                DataTable dt = DataHelper.getDataTableExecuteSPNoParam("USP_GetIBMPIDDetails");

                return dt;
            }
            catch (Exception ex)
            {
                throw ex;
            }
        }
        public DataTable GetAnnounceDate()
        {
            try
            {
                DataTable dt = DataHelper.getDataTableExecuteSPNoParam("USP_GetLatestAnnounceDate");

                return dt;
            }
            catch (Exception ex)
            {
                throw ex;
            }
        }
        public DataTable FetchCurrentStatus()
        {
            try
            {
                DataTable dt = DataHelper.getDataTableExecuteSPNoParam("USP_FetchCurrentStatus");

                return dt;
            }
            catch (Exception ex)
            {
                throw ex;
            }
        }
        public void UpdateCurrentRun(string PID)
        {
            try
            {
                SqlParameter[] param = new SqlParameter[1];
                param[0] = new SqlParameter("@PID", PID);

                int i = DataHelper.ExecuteSPWithExecuteScalar("USP_UpdateCurrentRun", param);
                
            }
            catch (Exception)
            {
                
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

    }
}
