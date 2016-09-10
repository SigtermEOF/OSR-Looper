using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Data.Odbc;
using System.Data.OleDb;
using System.Data.Sql;
using System.Data.SqlClient;
using System.Diagnostics;
using System.Threading;
using System.IO;

namespace OSR_Looper
{
    class TaskMethods
    {
        string sOSRConnString = OSR_Looper.Properties.Settings.Default.OSRConnString.ToString();
        string sCDSConnString = OSR_Looper.Properties.Settings.Default.CDSConnString.ToString();

        public void SQLNonQuery(string sConnString, string sCommText)
        {
            try
            {
                SqlConnection myConn = new SqlConnection(sConnString);

                SqlCommand myCommand = myConn.CreateCommand();

                myCommand.CommandText = sCommText;

                myConn.Open();

                myCommand.ExecuteNonQuery();

                myCommand.Dispose();

                myConn.Close();
                myConn.Dispose();
            }
            catch (Exception ex)
            {
                this.SaveExceptionToDB(ex);
            }
        }

        public void SQLQuery(string sConnString, string sCommText, DataTable dt) 
        {
            try
            {
                SqlConnection myConn = new SqlConnection(sConnString);

                SqlCommand myCommand = myConn.CreateCommand();

                myCommand.CommandText = sCommText;

                myConn.Open();

                SqlDataReader myDataReader = myCommand.ExecuteReader();

                if (myDataReader.HasRows)
                {
                    dt.Clear();
                    dt.Load(myDataReader);
                }

                myDataReader.Close();
                myDataReader.Dispose();

                myCommand.Dispose();

                myConn.Close();
                myConn.Dispose();
            }
            catch (Exception ex)
            {
                this.SaveExceptionToDB(ex);
            }
        }

        public void SQLQueryGenericWithReturnValue(string sConnString, string sCommTextValue, DataTable dt, ref string sValueString, string sColumn)   
        {
            try
            {
                SqlConnection myConn = new SqlConnection(sConnString);

                SqlCommand myCommand = myConn.CreateCommand();

                myCommand.CommandText = sCommTextValue;

                myConn.Open();

                SqlDataReader myDataReader = myCommand.ExecuteReader();

                if (myDataReader.HasRows)
                {
                    dt.Clear();
                    dt.Load(myDataReader);

                    sValueString = string.Empty;
                    sValueString = Convert.ToString(dt.Rows[0][sColumn]).Trim();
                }

                myDataReader.Close();
                myDataReader.Dispose();

                myCommand.Dispose();

                myConn.Close();
                myConn.Dispose();
            }
            catch (Exception ex)
            {
                this.SaveExceptionToDB(ex);
            }
        }

        public void CDSQuery(string sConnString, string sCommText, DataTable dt)
        {
            try
            {
                OleDbConnection CDSconn = new OleDbConnection(sConnString);

                OleDbCommand CDScommand = CDSconn.CreateCommand();

                CDScommand.CommandText = sCommText;

                CDSconn.Open();

                CDScommand.CommandTimeout = 0;

                OleDbDataReader CDSreader = CDScommand.ExecuteReader();

                if (CDSreader.HasRows)
                {
                    dt.Clear();
                    dt.Load(CDSreader);
                }

                CDScommand.Dispose();

                CDSreader.Close();
                CDSreader.Dispose();

                CDSconn.Close();
                CDSconn.Dispose();
            }
            catch (Exception ex)
            {
                this.SaveExceptionToDB(ex);
            }
        }

        public void CDSNonQuery(string sConnString, string sCommText)
        {
            try
            {
                OleDbConnection CDSconn = new OleDbConnection(sConnString);

                OleDbCommand CDScommand = CDSconn.CreateCommand();

                CDScommand.CommandText = sCommText;

                CDSconn.Open();

                CDScommand.CommandTimeout = 0;

                OleDbDataReader CDSreader = CDScommand.ExecuteReader();

                CDScommand.Dispose();

                CDSreader.Close();
                CDSreader.Dispose();

                CDSconn.Close();
                CDSconn.Dispose();
            }
            catch (Exception ex)
            {
                this.SaveExceptionToDB(ex);
            }
        } 

        public void SaveExceptionToDB(Exception ex) 
        {
            string sException = ex.ToString().Trim();
            sException = sException.Replace(@"'", "");

            string sConnString = sOSRConnString;
            string sCommText = "INSERT INTO [OSR_Errors] VALUES ('" + sException + "', '" + DateTime.Now.ToString().Trim() + "', '0')";

            this.SQLNonQuery(sConnString, sCommText);
        }

        public void SendErrorFlagging(string sProdNum)
        {
            string sConnString = sOSRConnString;
            string sCommText = "UPDATE [OSRetouch].[dbo].[OSR_Orders] SET [OrderStatus_Code] = '4' WHERE [Orders_ProdNum] ='" + sProdNum + "' ";

            this.SQLNonQuery(sConnString, sCommText);
        }
        
        public void EmailVariables(ref string sEmailServer, ref string sEmailMyBccAdd, ref string sSendToAdd2, ref string sAPReportSendAddy, ref string sErrorSendToAddy, ref string sFromAddy)
        {
            try
            {
                string sConnString = sOSRConnString;
                string sCommText = "SELECT [Variables_Variable] FROM [OSR_Variables] WHERE [Variables_VariableName] = 'APS_Email_Server'"; // Email server name.
                DataTable dt = new DataTable();

                this.SQLQuery(sConnString, sCommText, dt);

                if (dt.Rows.Count > 0)
                {
                    sEmailServer = Convert.ToString(dt.Rows[0]["Variables_Variable"]).Trim();
                }

                sCommText = "SELECT [Variables_Variable] FROM [OSR_Variables] WHERE [Variables_VariableName] = 'APS_Email_My_BCC'"; // Blind copy all OSR related emails to my email addy.

                this.SQLQuery(sConnString, sCommText, dt);

                if (dt.Rows.Count > 0)
                {
                    sEmailMyBccAdd = Convert.ToString(dt.Rows[0]["Variables_Variable"]).Trim();
                }

                sCommText = "SELECT [Variables_Variable] FROM [OSR_Variables] WHERE [Variables_VariableName] = 'VPN_CC_Email_Adds'"; // List of in lab email addys for notification emails.

                this.SQLQuery(sConnString, sCommText, dt);

                if (dt.Rows.Count > 0)
                {
                    sSendToAdd2 = Convert.ToString(dt.Rows[0]["Variables_Variable"]).Trim();
                }

                sCommText = "SELECT [Variables_Variable] FROM [OSR_Variables] WHERE [Variables_VariableName] = 'OSR_Excel_SendToAdd'"; // Accounts payable email addy.

                this.SQLQuery(sConnString, sCommText, dt);

                if (dt.Rows.Count > 0)
                {
                    sAPReportSendAddy = Convert.ToString(dt.Rows[0]["Variables_Variable"]).Trim();
                }

                sCommText = "SELECT [Variables_Variable] FROM [OSR_Variables] WHERE [Variables_VariableName] = 'APS_Email_Error_Sendto'"; // Email addy to send error notifications to.

                this.SQLQuery(sConnString, sCommText, dt);

                if (dt.Rows.Count > 0)
                {
                    sErrorSendToAddy = Convert.ToString(dt.Rows[0]["Variables_Variable"]).Trim();
                }

                // This does not work as a single string or seperated after the fact.

                //sCommText = "SELECT [Variables_Variable] FROM [OSR_Variables] WHERE [Variables_VariableName] = 'APS_Email_Server_FromAdd'"; // Email addy and descriptor for sent emails.

                //this.SQLQueryGeneric(sConnString, sCommText, dt);

                //if (dt.Rows.Count > 0)
                //{
                //    sFromAddy = Convert.ToString(dt.Rows[0][@"Variables_Variable"]).Trim();
                //    string[] s = sFromAddy.Split(',');

                //    string sFirst = s[0].ToString().Trim();
                //    string sSecond = s[1].ToString().Trim();
                //}
            }
            catch (Exception ex)
            {
                this.SaveExceptionToDB(ex);
            }
        }

        public void RecErrorFlagging(string sRecProdNum)
        {
            string sConnString = sOSRConnString;
            string sCommText = "UPDATE [OSR_Orders] SET [OrderStatus_Code] = '4' WHERE [Orders_ProdNum] = '" + sRecProdNum + "' ";

            this.SQLNonQuery(sConnString, sCommText);
        }

    }
}
