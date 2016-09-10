//****************************
//#define dev
//****************************

using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.IO;
using System.Data.Odbc;
using System.Data.OleDb;
using System.Data.Sql;
using System.Data.SqlClient;
using System.Diagnostics;
using System.Timers;
using OfficeOpenXml;
using OfficeOpenXml.Drawing;
using ExcelLibrary;

namespace OSR_Looper
{
    public partial class Main : Form
    {
        // Public variables.

        TaskMethods TM = null;
        Email EM = null;
        string sOSRConnString = OSR_Looper.Properties.Settings.Default.OSRConnString.ToString();
        string sCDSConnString = OSR_Looper.Properties.Settings.Default.CDSConnString.ToString();
        bool bReturned = false;
        bool bResults = false;
        bool bHalt = false;
        bool bAPReportGenerated = false;

        public Main()
        {
            InitializeComponent();
            TM = new TaskMethods();
            EM = new Email();
        }

        #region Form events.

        private void Main_Load(object sender, EventArgs e)
        {
            // On load set the title on the form including version which is read from a table associated with the application.
            string sCommTextValue = "SELECT * FROM [OSR_ChangeLog] WHERE [App] = 'Looper'";
            DataTable dt = new DataTable();
            string sValueString = string.Empty;
            string sColumn = "Version";

            TM.SQLQueryGenericWithReturnValue(sOSRConnString, sCommTextValue, dt, ref sValueString, sColumn);
            
            this.Text = "O.S.R. Looper " + sValueString;

            Application.DoEvents();
        }

        private void btnStart_Click(object sender, EventArgs e)
        {
            bHalt = false;

            //string sProdNum = string.Empty;
            //string sServerDir = string.Empty;
            //string sRefNum = string.Empty;
            //string sFilePath = string.Empty;

            //this.WorkToDo(sProdNum, sServerDir, sRefNum, ref sFilePath);

            if (DateTime.Now.ToString("ddd").Trim() == "Tue")
            {
                bAPReportGenerated = false;
            }

            string sDateOnly = DateTime.Now.Date.ToString("M/dd/yy").Trim();
            string sTimeOnly = DateTime.Now.ToString("hh:mm:ss tt").Trim();

            this.DoAllWork();

            this.btnStart.Enabled = false;
            this.btnHalt.Enabled = true;
            Application.DoEvents();
        }

        private void btnHalt_Click(object sender, EventArgs e)
        {
            this.timer1.Stop();
            this.timer1.Enabled = false;

            this.btnHalt.Enabled = false;
            this.btnStart.Enabled = true;

            this.lblLastRan.Text = "Halted.";
            Application.DoEvents();

            bHalt = true;
        }

        private void timer1_Tick(object sender, EventArgs e)
        {
            bHalt = false;

            if (DateTime.Now.ToString("ddd").Trim() == "Tue")
            {
                bAPReportGenerated = false;
            }

            this.btnStart.Enabled = false;
            this.btnHalt.Enabled = false;
            Application.DoEvents();

            this.DoAllWork();

            this.btnHalt.Enabled = true;
            Application.DoEvents();
        }

        private void ClearFormVariables()
        {
            bReturned = false;
            bResults = false;
            bHalt = false;
        }

        #endregion

        public void DoAllWork()
        {
            if (bHalt != true)
            {
                this.timer1.Stop();
                this.timer1.Enabled = false;
            }

            this.SendWork();
            //this.RecWork();
            //this.MiscWork();
            //this.Reporting();
          

            this.lblLastRan.Text = "";
            this.lblLastRan.Text = "Last loop: " + DateTime.Now.ToString().Trim() + "" +
                Environment.NewLine + "Next loop: " + DateTime.Now.AddMinutes(15).ToString().Trim() + "";

            Application.DoEvents();

            this.timer1.Enabled = true;
            this.timer1.Start();
        }

        #region SendWork methods.

        private void SendWork()
        {
            this.OrdersSendQuery();
            this.ClearFormVariables();
        }

        private void OrdersSendQuery()
        {
            try
            {
                string sCommText = "SELECT * FROM [OSR_Orders] WHERE [OrderStatus_Code] = '1'";
                DataTable dt = new DataTable();

                TM.SQLQuery(sOSRConnString, sCommText, dt);

                if (dt.Rows.Count > 0)
                {
                    foreach (DataRow dr in dt.Rows)
                    {
                        bReturned = false;

                        string sOriginalPath = Convert.ToString(dr[@"Orders_FileLoc"]).Trim();
                        string sProdNum = Convert.ToString(dr["Orders_ProdNum"]).Trim();
                        string sDirectory = new FileInfo(sOriginalPath).Directory.Name;
                        string sServerDir = sOriginalPath.Trim() + "Original";
                        string sCount = Convert.ToString(dr["Orders_Count"]).Trim();
                        string sRefNum = Convert.ToString(dr["Orders_RefNum"]).Trim();
                        string sCntrctrIDSend = Convert.ToString(dr["Cntrctr_ID"]).Trim();

                        string sContractor = string.Empty;
                        string sCarrierID = string.Empty;
                        string sCarrierAdd = string.Empty;
                        int iOriginalFileCount = 0;
                        int iCopiedToServerFileCount = 0;
                        string sFilePath = string.Empty;
                        string sConPhoneNum = string.Empty;

                        this.GetContractorInfo(sCntrctrIDSend, ref sContractor, ref sCarrierID, ref sConPhoneNum);
                        this.JobsQuery(sProdNum, sServerDir, sOriginalPath);

                        if (bReturned != true)
                        {
                            this.WorkToDo(sProdNum, sServerDir, sRefNum, ref sFilePath);

                            if(bReturned != true)
                            {
                                this.FileCopyDestination(sContractor, sProdNum, sServerDir, iOriginalFileCount, ref iCopiedToServerFileCount);
                            }
                            if (bReturned != true)
                            {
                                this.UpdateOrderStatusSent(sProdNum);
                                this.GetCarrierInfo(sCarrierID, ref sCarrierAdd);
                                this.EmailSend(sProdNum, iCopiedToServerFileCount, sCntrctrIDSend, sFilePath, sCarrierAdd, sConPhoneNum);
                                this.UpdateStampsOnSend(sProdNum);
                            }
                            else if (bReturned == true)
                            {
                                return;
                            }
                        }
                        else if (bReturned == true)
                        {
                            return;
                        }
                    }
                }
                else if (dt.Rows.Count == 0)
                {
                    bResults = false;
                }
            }
            catch(Exception ex)
            {
                TM.SaveExceptionToDB(ex);
            }
        }

        private void GetContractorInfo(string sCntrctrIDSend, ref string sContractor, ref string sCarrierID, ref string sConPhoneNum)
        {
            try
            {
                string sCommText = "SELECT * FROM [OSR_Contractor] WHERE [Cntrctr_ID] ='" + sCntrctrIDSend + "' ";
                DataTable dt = new DataTable();

                TM.SQLQuery(sOSRConnString, sCommText, dt);

                if (dt.Rows.Count > 0)
                {
                    string sConFName = Convert.ToString(dt.Rows[0]["Cntrctr_FName"]).Trim();
                    string sConLName = Convert.ToString(dt.Rows[0]["Cntrctr_LName"]).Trim();
                    sContractor = sConFName + sConLName;

                    sConPhoneNum = Convert.ToString(dt.Rows[0]["Cntrctr_Phone1"]).Trim();

                    sCarrierID = Convert.ToString(dt.Rows[0]["Carriers_ID"]).Trim();
                }
            }
            catch (Exception ex)
            {
                TM.SaveExceptionToDB(ex);
            }
        }

        private void JobsQuery(string sProdNum, string sServerDir, string sOriginalPath)
        {
            try
            {
                string sCommText = "SELECT [Jobs_Filename] FROM [OSR_Jobs] WHERE [Orders_ProdNum] ='" + sProdNum + "' ";
                DataTable dt = new DataTable();

                TM.SQLQuery(sOSRConnString, sCommText, dt);

                if (dt.Rows.Count > 0)
                {
                    this.FileCopyToServer(sServerDir, sProdNum, dt, sOriginalPath);
                }
                else if (dt.Rows.Count == 0)
                {
                    bResults = false;
                }
            }
            catch(Exception ex)
            {
                TM.SaveExceptionToDB(ex);
            }
        }

        private void FileCopyToServer(string sServerDir, string sProdNum, DataTable dtJobs, string sOriginalPath)
        {
            try
            {
                if (Directory.Exists(sServerDir))
                {
                    TM.SendErrorFlagging(sProdNum);

                    string sException = "Original subdirectory exists in jobs directory for production number: " + sProdNum + ".";

                    string sCommText = "INSERT INTO [OSR_Errors] VALUES ('" + sException + "', '" + DateTime.Now.ToString().Trim() + "', '0')";

                    TM.SQLNonQuery(sOSRConnString, sCommText);

                    bReturned = true;
                    return;
                }

                DirectoryInfo dirInfo = Directory.CreateDirectory(sServerDir);

                foreach (DataRow dr in dtJobs.Rows)
                {
                    foreach (var vItem in dr.ItemArray)
                    {
                        string sPics = Convert.ToString(vItem).Trim();

                        DataTable dTbl = new DataTable();
                        string sCommText = "SELECT [Jobs_Descript] FROM [OSR_Jobs] WHERE [Jobs_FileName] = '" + sPics + "'";

                        TM.SQLQuery(sOSRConnString, sCommText, dTbl);

                        if (dTbl.Rows.Count > 0)
                        {
                            string sDescript = Convert.ToString(dTbl.Rows[0]["Jobs_Descript"]).Trim();

                            if (sDescript == "16x20 Painter Portrait without Frame")
                            {
                                sOriginalPath += @"Corrected\";

                                if (Directory.Exists(sOriginalPath))
                                {
                                    string sSourceFile = Path.Combine(sOriginalPath, sPics).Trim();
                                    string sDestFile = Path.Combine(sServerDir, sPics).Trim();
                                    File.Copy(sSourceFile, sDestFile, true);
                                }
                                else if (!Directory.Exists(sOriginalPath))
                                {
                                    string sError = "Corrected directory does not exist for image: " + sPics;
                                    MessageBox.Show(sError);
                                    sCommText = "INSERT INTO [OSR_Errors] VALUES ('" + sError + "', '" + DateTime.Now.ToString().Trim() + "', '0')";

                                    TM.SQLNonQuery(sOSRConnString, sCommText);

                                    sCommText = "UPDATE [OSR_Orders] SET [OrderStatus_Code] = '4' WHERE [Orders_ProdNum] = '" + sProdNum + "'";

                                    TM.SQLNonQuery(sOSRConnString, sCommText);
                                }
                            }
                            else if (sDescript != "16x20 Painter Portrait without Frame")
                            {
                                string sSourceFile = Path.Combine(sOriginalPath, sPics).Trim();
                                string sDestFile = Path.Combine(sServerDir, sPics).Trim();
                                File.Copy(sSourceFile, sDestFile, true);
                            }
                        }
                        else if (dTbl.Rows.Count == 0)
                        {

                        }
                    }
                }
            }
            catch (Exception ex)
            {
                TM.SaveExceptionToDB(ex);
            }
        }

        private void WorkToDo(string sProdNum, string sServerDir, string sRefNum, ref string sFilePath)
        {
            try
            {
                //sProdNum = "46162";
                //sRefNum = "252020";
                //sServerDir = @"\\cdsdata\vol1\jlett_tmp\";


                SqlDataReader reader;
                string sCommText = "SELECT [Jobs_Filename], [Jobs_Descript] FROM [OSR_Jobs] WHERE [Orders_ProdNum] = '" + sProdNum + "'" +
                " ORDER BY [Jobs_Filename] ASC";
                string sDelimiter = "                    ";
                sFilePath = sServerDir + @"\" + sProdNum + ".txt";
                int iOriginalFileCount = Directory.GetFiles(sServerDir, "*.jpg", SearchOption.TopDirectoryOnly).Length;

                using (SqlConnection sqlConn = new SqlConnection(sOSRConnString))
                {
                    sqlConn.Open();

                    using (reader = new SqlCommand(sCommText, sqlConn).ExecuteReader())
                    {
                        if (reader.HasRows)
                        {
                            StringBuilder sb = new StringBuilder();
                            Object[] items = new Object[reader.FieldCount];

                            sb.AppendFormat("Production #: " + sProdNum);
                            sb.Append("          ");
                            sb.AppendFormat("Reference #: " + sRefNum);
                            sb.Append(Environment.NewLine);

                            sb.AppendFormat("Images: " + iOriginalFileCount);
                            sb.Append(Environment.NewLine);

                            sb.AppendFormat("File generated: " + DateTime.Now.ToString().Trim());
                            sb.Append(Environment.NewLine);

                            sb.Append("------------------------------------------------------------");
                            sb.Append(Environment.NewLine);

                            sb.Append("  Image Name" + "                              " + "Retouch Type");
                            sb.Append(Environment.NewLine);

                            sb.Append("------------------------------------------------------------");
                            sb.Append(Environment.NewLine);

                            while (reader.Read())
                            {
                                reader.GetValues(items);

                                foreach (var vItem in items)
                                {
                                    sb.Append(vItem.ToString().Trim());
                                    sb.Append(sDelimiter);
                                }

                                if (sb.ToString().EndsWith(", "))
                                    sb = sb.Remove(sb.Length - 2, 2);

                                sb.Append(Environment.NewLine);
                            }

                            reader.Close();
                            reader.Dispose();
                            sqlConn.Dispose();
                            sqlConn.Close();

                            File.WriteAllText(sFilePath, sb.ToString());
                    }
                 }
              }
            }
            catch (Exception ex)
            {
                TM.SaveExceptionToDB(ex);
            }
        }

        private void FileCopyDestination(string sContractor, string sProdNum, string sServerDir, int iOriginalFileCount, ref int iCopiedToServerFileCount)
        {
            string sDestDir = string.Empty;

            string sCommText = "SELECT [Variables_Variable] FROM [OSR_Variables] WHERE [Variables_VariableName] = 'VPN_Server_Path'";
            DataTable dt = new DataTable();

            TM.SQLQuery(sOSRConnString, sCommText, dt);

            if (dt.Rows.Count > 0)
            {
                sDestDir = Convert.ToString(dt.Rows[0]["Variables_Variable"]).Trim();
            }

            try
            {
                string sCreateFinishedWorkDir = sDestDir + @"\" + sContractor + @"\Finished_Work\";
                sDestDir += @"\" + sContractor + @"\New_Work\" + sProdNum;

                if (!Directory.Exists(sDestDir))
                {
                    Directory.CreateDirectory(sDestDir);
                }
                else if (Directory.Exists(sDestDir))
                {
                    string sException = "Directory already exists on FTP for production number: " + sProdNum + ". Contractor: " + sContractor;

                    sCommText = "INSERT INTO [OSR_Errors] VALUES ('" + sException + "', '" + DateTime.Now.ToString().Trim() + "', '0')";

                    TM.SQLNonQuery(sOSRConnString, sCommText);

                    bReturned = true;
                    return;
                }
                if (!Directory.Exists(sCreateFinishedWorkDir))
                {
                    Directory.CreateDirectory(sCreateFinishedWorkDir);
                }
                else if (Directory.Exists(sCreateFinishedWorkDir))
                {
                    // Do nothing if "Finished_Work" exists.
                }

                string[] sFiles = Directory.GetFiles(sServerDir);

                foreach (string s in sFiles)
                {
                    string sFileName = Path.GetFileName(s);
                    string sDestFiles = Path.Combine(sDestDir, sFileName);
                    File.Copy(s, sDestFiles, true);
                }

                iOriginalFileCount = Directory.GetFiles(sServerDir, "*.jpg", SearchOption.TopDirectoryOnly).Length;
                iCopiedToServerFileCount = Directory.GetFiles(sDestDir, "*.jpg", SearchOption.TopDirectoryOnly).Length;

                if (iOriginalFileCount != iCopiedToServerFileCount)
                {
                    TM.SendErrorFlagging(sProdNum);

                    string sException = "File count mismatch for production number: " + sProdNum + ". Contractor: " + sContractor + ".";

                    sCommText = "INSERT INTO [OSR_Errors] VALUES ('" + sException + "', '" + DateTime.Now.ToString().Trim() + "', '0')";

                    TM.SQLNonQuery(sOSRConnString, sCommText);

                    bReturned = true;
                    return;
                }
            }
            catch (Exception ex)
            {
                TM.SaveExceptionToDB(ex);
            }
        }

        private void UpdateOrderStatusSent(string sProdNum)
        {
            try
            {
                string sCommText = "UPDATE [OSR_Orders] SET [OrderStatus_Code] = '2', [Orders_SentDate] ='" + DateTime.Now.ToString().Trim() + "' " +
                " WHERE [Orders_ProdNum] = '" + sProdNum + "' ";

                TM.SQLNonQuery(sOSRConnString, sCommText);
            }
            catch (Exception ex)
            {
                TM.SaveExceptionToDB(ex);
            }
        }

        private void GetCarrierInfo(string sCarrierID, ref string sCarrierAdd)
        {
            try
            {
                string sCommText = "SELECT * FROM [OSR_Carriers] WHERE [Carriers_ID] ='" + sCarrierID + "' ";
                DataTable dt = new DataTable();

                TM.SQLQuery(sOSRConnString, sCommText, dt);

                if (dt.Rows.Count > 0)
                {
                    sCarrierAdd = Convert.ToString(dt.Rows[0]["Carriers_Address"]).Trim();
                }
            }
            catch (Exception ex)
            {
                TM.SaveExceptionToDB(ex);
            }
        }

        private void EmailSend(string sProdNum, int iCopiedToServerFileCount, string sCntrctrIDSend, string sFilePath, string sCarrierAdd, string sConPhoneNum)
        {
            string sSubject = string.Format("APS Production number " + sProdNum + " has been uploaded to the FTP server for retouching" +
                " on " + DateTime.Now.Date.ToString("M/dd/yy").Trim() + " at " + DateTime.Now.ToString("hh:mm:ss tt").Trim() + ".");
            string sBody = string.Format("APS Production number " + sProdNum + " has been uploaded to the FTP server for retouching." + Environment.NewLine +
                "This job was uploaded on " + DateTime.Now.Date.ToString("M/dd/yy").Trim() + " at " + DateTime.Now.ToString("hh:mm:ss tt").Trim() + "." + Environment.NewLine +
                "Inside the New Work directory there is a directory called " + sProdNum + " and it contains " + iCopiedToServerFileCount + " images." + Environment.NewLine +
                "Inside the directory you will find a text file titled " + sProdNum + ".txt that contains the retouching instructions." + Environment.NewLine +
                "This file is also attached to this email for your convenience." + Environment.NewLine + Environment.NewLine +
                "After retouching has been completed, please create a new directory under the Finished Work directory and name it " + sProdNum + "." + Environment.NewLine +
                "Please make sure that all " + iCopiedToServerFileCount + " images are in this directory." + Environment.NewLine + Environment.NewLine +
                "These images are transferred via an automated system." + Environment.NewLine +
                "There is no need to email APS when the job is complete unless you wish to do so." + Environment.NewLine +
                "The automated system will check the Finished Work directory periodically and will handle this job accordingly." + Environment.NewLine +
                "Images are checked and counted prior to uploading and the same system checks and counts the images prior to placing them back into production at APS." + Environment.NewLine +
                "Any discrepancies in transferring will result in an error." + Environment.NewLine + Environment.NewLine +
                "This email was sent via an automated system, please do not reply to this email as this address is not monitored.");

            string sSendTo = string.Empty;
            string sFile = sFilePath;

            try
            {
                string sCommText = "SELECT * FROM [OSR_Contractor] WHERE [Cntrctr_ID] ='" + sCntrctrIDSend + "' ";
                DataTable dt = new DataTable();

                TM.SQLQuery(sOSRConnString, sCommText, dt);

                if (dt.Rows.Count > 0)
                {
                    sSendTo = Convert.ToString(dt.Rows[0]["Cntrctr_Email1"]).Trim();
                }

                string sSendToAdd = sConPhoneNum + sCarrierAdd;

                string sEmailServer = string.Empty;
                string sEmailMyBccAdd = string.Empty;
                string sSendToAdd2 = string.Empty;
                string sAPReportSendAddy = string.Empty;
                string sErrorSendToAddy = string.Empty;
                string sFromAddy = string.Empty;

                TM.EmailVariables(ref sEmailServer, ref sEmailMyBccAdd, ref sSendToAdd2, ref sAPReportSendAddy, ref sErrorSendToAddy, ref sFromAddy);

                EM.EmailSend(sEmailServer, sEmailMyBccAdd, sSendTo, sSendToAdd, sSendToAdd2, sSubject, sBody, sFile, sFromAddy);
            }
            catch (Exception ex)
            {
                TM.SaveExceptionToDB(ex);
            }
        }

        private void UpdateStampsOnSend(string sProdNum)
        {
            try
            {
                string sCommText = "INSERT INTO Stamps (User_id, Stationid, Lookupnum, Date, Time, Action, Wbs_task, Sequence, Framenum, Count," +
                    " Seconds, Wbs_plan, Wbs_track, Wbs_status, App_level, Processed) VALUES " +
                    "('AUTO IMG', 'AUTO IMGER', '" + sProdNum + "'," + " DATE(" + DateTime.Now.Date.ToString("yyyy,MM,dd").Trim() + "), '" + DateTime.Now.ToString("H:mm:ss").Trim() + "', 'RTCH_STRT', 'RTCH_STRT', 0, ' ', ' ', ' ', ' '," +
                    " .F., ' ', ' ', ' ' )";

                TM.CDSNonQuery(sCDSConnString, sCommText);
            }
            catch (Exception ex)
            {
                TM.SaveExceptionToDB(ex);
            }
        }

        #endregion

        #region RecWork methods.  

        private void RecWork()
        {
            this.OrdersRecQuery();
            this.ClearFormVariables();
        }

        private void OrdersRecQuery()
        {
            try
            {
                string sCommText = "SELECT * FROM [OSR_Orders] WHERE [OrderStatus_Code] = '2'";
                DataTable dt = new DataTable();

                TM.SQLQuery(sOSRConnString, sCommText, dt);

                if (dt.Rows.Count > 0)
                {
                    bReturned = false;

                    foreach (DataRow dr in dt.Rows)
                    {
                        bReturned = false;

                        string sRecOriginalPath = Convert.ToString(dr[@"Orders_FileLoc"]).Trim();
                        string sRecProdNum = Convert.ToString(dr["Orders_ProdNum"]).Trim();
                        string sRecCustNum = Convert.ToString(dr["Orders_AccNum"]).Trim();
                        string sRecRefNum = Convert.ToString(dr["Orders_RefNum"]).Trim();
                        string sRecDirectory = new FileInfo(sRecOriginalPath).Directory.Name;
                        string sRecServerDir = sRecOriginalPath.Trim() + "Original";
                        string sCntrctrIDRec = Convert.ToString(dr["Cntrctr_ID"]).Trim();
                        string sRecSentDate = Convert.ToString(dr["Orders_SentDate"]);

                        string sRecContractor = string.Empty;
                        string sCarrierIDRec = string.Empty;
                        string sConPhoneNumRec = string.Empty;

                        this.GetContractorRecInfo(sCntrctrIDRec, ref sRecContractor, ref sCarrierIDRec, ref sConPhoneNumRec);

                        int iFinishedWorkFileCount = 0;

                        this.JobsRecQuery(sRecProdNum, sRecContractor, sRecServerDir, sRecOriginalPath, ref iFinishedWorkFileCount);

                        if (bReturned != true)
                        {
                            string sCarrierAddRec = string.Empty;
                            this.GetCarrierInfoRec(sCarrierIDRec, ref sCarrierAddRec);

                            this.EmailRec(sRecProdNum, iFinishedWorkFileCount, sRecDirectory, sRecSentDate, sCntrctrIDRec, sConPhoneNumRec, sCarrierAddRec);

                            this.GenerateThumbs(sRecRefNum, sRecCustNum);

                            this.UpdateStampsOnRec(sRecProdNum);
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                TM.SaveExceptionToDB(ex);
            }
        }

        private void GetContractorRecInfo(string sCntrctrIDRec, ref string sRecContractor, ref string sCarrierIDRec, ref string sConPhoneNumRec)
        {
            try
            {
                string sCommText = "SELECT * FROM [OSR_Contractor] WHERE [Cntrctr_ID] = '" + sCntrctrIDRec + "' ";
                DataTable dt = new DataTable();

                TM.SQLQuery(sOSRConnString, sCommText, dt);

                if (dt.Rows.Count > 0)
                {
                    string sRecConFName = Convert.ToString(dt.Rows[0]["Cntrctr_FName"]).Trim();
                    string sRecConLName = Convert.ToString(dt.Rows[0]["Cntrctr_LName"]).Trim();
                    sRecContractor = sRecConFName + sRecConLName;
                    sConPhoneNumRec = Convert.ToString(dt.Rows[0]["Cntrctr_Phone1"]).Trim();
                    sCarrierIDRec = Convert.ToString(dt.Rows[0]["Carriers_ID"]).Trim();
                }
            }
            catch (Exception ex)
            {
                TM.SaveExceptionToDB(ex);
            }
        }

        private void JobsRecQuery(string sRecProdNum, string sRecContractor, string sRecServerDir, string sRecOriginalPath, ref int iFinishedWorkFileCount)
        {
            try
            {
                string sCommText = "SELECT [Jobs_Filename] FROM [OSR_Jobs] WHERE [Orders_ProdNum] = '" + sRecProdNum + "' ";
                DataTable dt = new DataTable();

                TM.SQLQuery(sOSRConnString, sCommText, dt);

                if (dt.Rows.Count > 0)
                {
                    this.FileCopyToJobs(sRecContractor, sRecProdNum, sRecServerDir, sRecOriginalPath, ref iFinishedWorkFileCount);
                }
                else if (dt.Rows.Count == 0)
                {
                    bResults = false;
                    return;
                }
            }
            catch (Exception ex)
            {
                TM.SaveExceptionToDB(ex);
            }
        }

        private void FileCopyToJobs(string sRecContractor, string sRecProdNum, string sRecServerDir, string sRecOriginalPath, ref int iFinishedWorkFileCount)
        {
            if (Directory.Exists(sRecServerDir))
            {

                string sServerPath = string.Empty;

                string sCommText = "SELECT [Variables_Variable] FROM [OSR_Variables] WHERE [Variables_VariableName] = 'VPN_Server_Path'";
                DataTable dt = new DataTable();

                TM.SQLQuery(sOSRConnString, sCommText, dt);

                if (dt.Rows.Count > 0)
                {
                    sServerPath = Convert.ToString(dt.Rows[0]["Variables_Variable"]).Trim();
                }

                string sRecSourceDir = sServerPath;
                string sRecSourceDir2 = sServerPath;

                try
                {
                    sRecSourceDir += @"\" + sRecContractor + @"\Finished_Work\" + sRecProdNum;
                    sRecSourceDir2 += @"\" + sRecContractor + @"\New_Work\" + sRecProdNum;

                    string sErrorPath = sRecServerDir + @"\" + "OSR_Error.txt";

                    if (!Directory.Exists(sRecSourceDir))
                    {
                        // The order hasn't been uploaded by the contractor and would not be available at this point. Will loop until ready.
                        bReturned = true;
                    }
                    else if (Directory.Exists(sRecSourceDir))
                    {
                        iFinishedWorkFileCount = Directory.GetFiles(sRecSourceDir, "*.jpg", SearchOption.TopDirectoryOnly).Length;
                        int iFileCountInOriginal = Directory.GetFiles(sRecServerDir, "*.jpg", SearchOption.TopDirectoryOnly).Length;

                        if (iFinishedWorkFileCount == iFileCountInOriginal)
                        {
                            // Pause the thread for 1 minute on match for finalizing the transferring of images if needed.
                            Thread.Sleep(1 * 60 * 1000);

                            DirectoryInfo dir1 = new DirectoryInfo(sRecServerDir);

                            IEnumerable<FileInfo> fileInfo = dir1.GetFiles("*.jpg", SearchOption.TopDirectoryOnly);

                            foreach (FileInfo fI in fileInfo)
                            {
                                if (!File.Exists(sRecSourceDir + @"\" + fI.Name))
                                {
                                    File.WriteAllText(sErrorPath, fI.Name);
                                    bReturned = true;

                                    string sException = "File name mismatch for production number: " + sRecProdNum + ". Contractor: " + sRecContractor + ". Error file: " + sErrorPath + ".";

                                    sCommText = "INSERT INTO [OSR_Errors] VALUES ('" + sException + "', '" + DateTime.Now.ToString().Trim() + "', '0')";

                                    TM.SQLNonQuery(sOSRConnString, sCommText);

                                    TM.RecErrorFlagging(sRecProdNum);

                                    return;
                                }
                            }

                            string[] sRecFiles = Directory.GetFiles(sRecSourceDir, "*.jpg");

                            try
                            {
                                foreach (string s in sRecFiles)
                                {
                                    string sRecFileName = Path.GetFileName(s);
                                    string sRecDestFiles = Path.Combine(sRecOriginalPath, sRecFileName);
                                    File.Copy(s, sRecDestFiles, true);
                                }

                                sCommText = "UPDATE [OSR_Orders] SET [OrderStatus_Code] = '3', [Orders_ReceivedDate] = '" + DateTime.Now.ToString().Trim() + "' WHERE [Orders_ProdNum] = '" + sRecProdNum + "' ";

                                TM.SQLNonQuery(sOSRConnString, sCommText);

                            }
                            catch (IOException)
                            {
                                bReturned = true;

                                string sException = "Error copying files to server for production number: " + sRecProdNum + ". Contractor: " + sRecContractor + ". Error file: " + sErrorPath + ".";
                                sCommText = "INSERT INTO [OSR_Errors] VALUES ('" + sException + "', '" + DateTime.Now.ToString() + "', '0')";

                                TM.SQLNonQuery(sOSRConnString, sCommText);

                                return;
                            }
                        }
                        else if (iFinishedWorkFileCount != iFileCountInOriginal)
                        {
                            string sErrorCountPath = sRecServerDir + @"\" + "OSR_Error_MISMATCH.txt";

                            DirectoryInfo dirInfo = new DirectoryInfo(sRecServerDir);

                            IEnumerable<FileInfo> fileInfo2 = dirInfo.GetFiles("*.jpg", SearchOption.TopDirectoryOnly);

                            foreach (FileInfo fInfo in fileInfo2)
                            {
                                if (!File.Exists(sRecSourceDir + @"\" + fInfo.Name))
                                {
                                    File.WriteAllText(sErrorCountPath, fInfo.Name);
                                }
                            }

                            bReturned = true;

                            string sException = "File count mismatch for production number: " + sRecProdNum + ". Contractor: " + sRecContractor + ". Error file: " + sErrorCountPath + ".";
                            sCommText = "INSERT INTO [OSR_Errors] VALUES ('" + sException + "', '" + DateTime.Now.ToString().Trim() + "', '0')";

                            TM.SQLNonQuery(sOSRConnString, sCommText);                            

                            return;
                        }
                    }
                }
                catch (Exception ex)
                {
                    TM.SaveExceptionToDB(ex);
                }
            }
            else if (!Directory.Exists(sRecServerDir))
            {
                bReturned = true;

                string sException = "Originals directory missing for production number: " + sRecProdNum + ". Contractor: " + sRecContractor + ".";
                string sCommText = "INSERT INTO [OSR_Errors] VALUES ('" + sException + "', '" + DateTime.Now.ToString().Trim() + "', '0')";

                TM.SQLNonQuery(sOSRConnString, sCommText);

                TM.RecErrorFlagging(sRecProdNum);

                return;
            }
        }

        private void GetCarrierInfoRec(string sCarrierIDRec, ref string sCarrierAddRec)
        {
            try
            {
                string sCommText = "SELECT * FROM [OSR_Carriers] WHERE [Carriers_ID] ='" + sCarrierIDRec + "' ";
                DataTable dt = new DataTable();

                TM.SQLQuery(sOSRConnString, sCommText, dt);

                if (dt.Rows.Count > 0)
                {
                    sCarrierAddRec = Convert.ToString(dt.Rows[0]["Carriers_Address"]).Trim();
                }
            }
            catch(Exception ex)
            {
                TM.SaveExceptionToDB(ex);
            }
        }

        private void EmailRec(string sRecProdNum, int iFinishedWorkFileCount, string sRecDirectory, string sRecSentDate, string sCntrctrIDRec, string sConPhoneNumRec, string sCarrierAddRec)
        {
            string sSubject = string.Format("APS Production number " + sRecProdNum + " has been downloaded from the FTP server on " + DateTime.Now.Date.ToString("M/dd/yy").Trim() + " at " + DateTime.Now.ToString("hh:mm:ss tt").Trim() + ".");
            string sBody = string.Format("APS Production number " + sRecProdNum + " has been downloaded from the FTP server." + Environment.NewLine +
                "This job was downloaded on " + DateTime.Now.Date.ToString("M/dd/yy").Trim() + " at " + DateTime.Now.ToString("hh:mm:ss tt").Trim() + "." + Environment.NewLine +
                +iFinishedWorkFileCount + " images have been copied to " + sRecDirectory + " in the Jobs directory." + Environment.NewLine +
                "This job was originally uploaded on " + sRecSentDate + ".");

            try
            {
                string sCommText = "SELECT * FROM [OSR_Contractor] WHERE [Cntrctr_ID] ='" + sCntrctrIDRec + "' ";
                DataTable dt = new DataTable();

                TM.SQLQuery(sOSRConnString, sCommText, dt);

                if (dt.Rows.Count > 0)
                {
                    string sSendTo = Convert.ToString(dt.Rows[0]["Cntrctr_Email1"]).Trim();
                    string sSendToAddRec = sConPhoneNumRec + sCarrierAddRec;

                    string sEmailServer = string.Empty;
                    string sEmailMyBccAdd = string.Empty;
                    string sSendToAdd2 = string.Empty;
                    string sAPReportSendAddy = string.Empty;
                    string sErrorSendToAddy = string.Empty;
                    string sFromAddy = string.Empty;

                    TM.EmailVariables(ref sEmailServer, ref sEmailMyBccAdd, ref sSendToAdd2, ref sAPReportSendAddy, ref sErrorSendToAddy, ref sFromAddy);

                    EM.EmailRec(sEmailServer, sEmailMyBccAdd, sSendTo, sSendToAddRec, sSendToAdd2, sSubject, sBody, sFromAddy);
                }
            }
            catch (Exception ex)
            {
                TM.SaveExceptionToDB(ex);
            }
        }

        private void GenerateThumbs(string sRecRefNum, string sRecCustNum)
        {
            try
            {
                string sCommText = "SELECT * FROM thumbgen WHERE Order = " + "'" + sRecRefNum + "'";
                DataTable dt = new DataTable();

                if (dt.Rows.Count > 0)
                {
                    sCommText = "UPDATE thumbgen SET Stmpdate = " + "DATE(" + DateTime.Now.Date.ToString("yyyy,MM,dd").Trim() + ")," + " Stmptime = '" + DateTime.Now.ToString("HH:mm:ss").Trim() + "', " +
                    "In_dp2 = .F., Thumbed = .F., Error = .F., Error_Desc = '' WHERE Order = " + "'" + sRecRefNum + "'";

                    TM.CDSNonQuery(sCDSConnString, sCommText);
                }
                else if (dt.Rows.Count == 0)
                {
                    sCommText = "INSERT INTO thumbgen ([Order], Customer, Stmpdate, Stmptime, In_dp2, Thumbed, [Error], Error_desc, Priority, Delay) VALUES " +
                    "('" + sRecRefNum + "', '" + sRecCustNum + "'," + " DATE(" + DateTime.Now.Date.ToString("yyyy,MM,dd").Trim() + "), '" + DateTime.Now.ToString("HH:mm:ss").Trim() + "', .F., .F., .F., ' ', 0, 0)";

                    TM.CDSNonQuery(sCDSConnString, sCommText);
                }
            }
            catch(Exception ex)
            {
                TM.SaveExceptionToDB(ex);
            }
        }

        private void UpdateStampsOnRec(string sRecProdNum)
        {
            try
            {
                string sCommText = "INSERT INTO Stamps (User_id, Stationid, Lookupnum, Date, Time, Action, Wbs_task, Sequence, Framenum, Count," +
                    " Seconds, Wbs_plan, Wbs_track, Wbs_status, App_level, Processed) VALUES " +
                    "('AUTO IMG', 'AUTO IMGER', " + "'" + sRecProdNum + "'," + " DATE(" + DateTime.Now.Date.ToString("yyyy,MM,dd").Trim() + "), '" + DateTime.Now.ToString("H:mm:ss").Trim() + "', 'RTCH_FINSH', 'RTCH_FINSH', 0, ' ', ' ', ' ', ' '," +
                    " .F., ' ', ' ', ' ' )";

                TM.CDSNonQuery(sCDSConnString, sCommText);
            }
            catch (Exception ex)
            {
                TM.SaveExceptionToDB(ex);
            }
        }

        #endregion

        #region Misc tasks unrelated to file sending or receiving. 

        private void MiscWork()
        {
            this.SendPW();
            this.ExceptionEmail();
            this.DeleteFinishedWork();
        }

        private void SendPW()
        {
            string sPWConFName = string.Empty;
            string sPWConLName = string.Empty;
            string sLogin = string.Empty;
            string sPassword = string.Empty;
            string sConEmail = string.Empty;
            string sConID = string.Empty;

            try
            {
                string sCommText = "SELECT * FROM [OSR_Contractor] WHERE [Cntrctr_PWSent] = 0 AND [Cntrctr_Password] IS NOT NULL";
                DataTable dt = new DataTable();

                TM.SQLQuery(sOSRConnString, sCommText, dt);

                if (dt.Rows.Count > 0)
                {
                    foreach (DataRow dr in dt.Rows)
                    {
                        sPWConFName = Convert.ToString(dr["Cntrctr_FName"]).Trim();
                        sPWConLName = Convert.ToString(dr["Cntrctr_LName"]).Trim();
                        sPassword = Convert.ToString(dr["Cntrctr_Password"]).Trim();
                        sConEmail = Convert.ToString(dr["Cntrctr_Email1"]).Trim();
                        sConID = Convert.ToString(dr["Cntrctr_ID"]).Trim();

                        sLogin = sPWConFName + sPWConLName;

                        string sSendTo = sConEmail;

                        sCommText = "SELECT [Variables_Variable] FROM [OSR_Variables] WHERE [Variables_VariableName] = 'VPN_Instructions.txt_Path'";
                        DataTable dt2 = new DataTable();
                        string sFile = string.Empty;

                        TM.SQLQuery(sOSRConnString, sCommText, dt2);

                        if (dt2.Rows.Count > 0)
                        {
                            sFile = Convert.ToString(dt2.Rows[0]["Variables_Variable"]).Trim();
                        }

                        string sEmailServer = string.Empty;
                        string sEmailMyBccAdd = string.Empty;
                        string sSendToAdd2 = string.Empty;
                        string sAPReportSendAddy = string.Empty;
                        string sErrorSendToAddy = string.Empty;
                        string sFromAddy = string.Empty;

                        TM.EmailVariables(ref sEmailServer, ref sEmailMyBccAdd, ref sSendToAdd2, ref sAPReportSendAddy, ref sErrorSendToAddy, ref sFromAddy);

                        string sSubject = string.Format("APS FTP Instructions");
                        string sBody = string.Format("The instructions to connect to the APS FTP have been attached to this email." + Environment.NewLine + Environment.NewLine +
                            "Listed below are your login and password for the APS FTP:" + Environment.NewLine +
                            "Login: " + sLogin + Environment.NewLine + "Password: " + sPassword + Environment.NewLine + Environment.NewLine +
                            "This email was sent via an automated system, please do not reply to this email as this address is not monitored." + Environment.NewLine + Environment.NewLine);

                        EM.EmailPW(sEmailServer, sEmailMyBccAdd, sSendToAdd2, sSubject, sBody, sFile, sFromAddy);

                        sCommText = "UPDATE [OSR_Contractor] SET [Cntrctr_PWSent] = '1' WHERE [Cntrctr_ID] = '" + sConID + "'";

                        TM.SQLNonQuery(sOSRConnString, sCommText);
                    }
                }
            }
            catch (Exception ex)
            {
                TM.SaveExceptionToDB(ex);
            }
        }

        private void ExceptionEmail()
        {
            try
            {
                string sCommText = "SELECT * FROM [OSR_Errors] WHERE [Errors_Email_Sent] = '0' OR [Errors_Email_Sent] IS NULL";
                DataTable dt = new DataTable();

                TM.SQLQuery(sOSRConnString, sCommText, dt);

                if (dt.Rows.Count > 0)
                {
                    foreach (DataRow dr in dt.Rows)
                    {
                        string sExceptionString = Convert.ToString(dr["Errors_String"]).Trim();
                        DateTime dateTimeException = Convert.ToDateTime(dr["Errors_DateTime"]);
                        string sDateTimeException = Convert.ToString(dr["Errors_DateTime"]);

                        string sEmailServer = string.Empty;
                        string sEmailMyBccAdd = string.Empty;
                        string sSendToAdd2 = string.Empty;
                        string sAPReportSendAddy = string.Empty;
                        string sErrorSendToAddy = string.Empty;
                        string sFromAddy = string.Empty;

                        TM.EmailVariables(ref sEmailServer, ref sEmailMyBccAdd, ref sSendToAdd2, ref sAPReportSendAddy, ref sErrorSendToAddy, ref sFromAddy);

                        string sSubject = string.Format("OSR Error Reporting");
                        string sBody = string.Format("An exception was recorded in the Errors database at " + dateTimeException + ":" + Environment.NewLine + Environment.NewLine + sExceptionString);

                        EM.EmailError(sEmailServer, sEmailMyBccAdd, sErrorSendToAddy, sSubject, sBody, sFromAddy);

                        sCommText = "UPDATE [OSR_Errors] SET [Errors_Email_Sent] = '1' WHERE [Errors_String] = '" + sExceptionString + "' AND ([Errors_Email_Sent] = '0' OR [Errors_Email_Sent] IS NULL)";

                        TM.SQLNonQuery(sOSRConnString, sCommText);
                    }
                }
            }
            catch (Exception ex)
            {
                TM.SaveExceptionToDB(ex);
            }
        }

        private void DeleteFinishedWork()
        {
            try
            {
                string sCommText = "SELECT [Orders_ProdNum], [OrderStatus_Code], [Orders_ReceivedDate], [Cntrctr_ID], [Orders_DelCheck] " +
                    "FROM [OSR_Orders] WHERE [OrderStatus_Code] = 3 AND " +
                    "[Orders_ReceivedDate] <= '" + DateTime.Now.AddDays(-14).ToString().Trim() + "' AND [Orders_DelCheck] IS NULL";
                DataTable dt = new DataTable();

                TM.SQLQuery(sOSRConnString, sCommText, dt);

                if (dt.Rows.Count > 0)
                {
                    foreach (DataRow dr in dt.Rows)
                    {
                        string sRecSourceDirPNum = Convert.ToString(dr["Orders_ProdNum"]).Trim();
                        string sRecDelFinWorkConID = Convert.ToString(dr["Cntrctr_ID"]).Trim();

                        sCommText = "SELECT * FROM [OSR_Contractor] WHERE [Cntrctr_ID] ='" + sRecDelFinWorkConID + "' ";
                        DataTable dt2 = new DataTable();

                        TM.SQLQuery(sOSRConnString, sCommText, dt2);

                        string sRecDelFinWorkConFName = string.Empty;
                        string sRecDelFinWorkConLName = string.Empty;
                        string sRecDelFinWorkConFullName = string.Empty;

                        if (dt2.Rows.Count > 0)
                        {
                            sRecDelFinWorkConFName = Convert.ToString(dt2.Rows[0]["Cntrctr_FName"]).Trim();
                            sRecDelFinWorkConLName = Convert.ToString(dt2.Rows[0]["Cntrctr_LName"]).Trim();
                            sRecDelFinWorkConFullName = sRecDelFinWorkConFName + sRecDelFinWorkConLName;

                            sCommText = "SELECT [Variables_Variable] FROM [OSR_Variables] WHERE [Variables_VariableName] = 'VPN_Server_Path'";
                            DataTable dt3 = new DataTable();
                            string sServerPath = string.Empty;

                            TM.SQLQuery(sOSRConnString, sCommText, dt3);

                            if (dt3.Rows.Count > 0)
                            {
                                sServerPath = Convert.ToString(dt3.Rows[0]["Variables_Variable"]).Trim();

                                string sRecDelFinWorkSourceDir = sServerPath;
                                string sRecDelFinWorkSourceDir2 = sServerPath;
                                sRecDelFinWorkSourceDir += @"\" + sRecDelFinWorkConFullName + @"\Finished_Work\" + sRecSourceDirPNum;
                                sRecDelFinWorkSourceDir2 += @"\" + sRecDelFinWorkConFullName + @"\New_Work\" + sRecSourceDirPNum;
                                string sRecDelFinWorkSourceDirNoPNum = sServerPath + @"\" + sRecDelFinWorkConFullName;
                                string sRecDelFinWorkSourceDir2NoPNum = sServerPath + @"\" + sRecDelFinWorkConFullName;

                                int iAttempts = 5;
                                bool bDeleted = false;
                                string sDir = string.Empty;

                                if (Directory.Exists(sRecDelFinWorkSourceDir) && Directory.Exists(sRecDelFinWorkSourceDir2))
                                {
                                    for (int i = 0; i < iAttempts; i++)
                                    {
                                        if (bDeleted != true)
                                        {
                                            try
                                            {
                                                sDir = sRecDelFinWorkSourceDir;
                                                Directory.Delete(sRecDelFinWorkSourceDir, true);
                                                sDir = sRecDelFinWorkSourceDir2;
                                                Directory.Delete(sRecDelFinWorkSourceDir2, true);
                                                bDeleted = true;

                                                sCommText = "UPDATE [OSR_Orders] SET [Orders_DelCheck] = '1' WHERE [Orders_ProdNum] ='" + sRecSourceDirPNum + "' "; // Null = not checked : 1= checked.

                                                TM.SQLNonQuery(sOSRConnString, sCommText);
                                            }
                                            catch (IOException)
                                            {
                                                bDeleted = false;

                                                if (i >= 2)
                                                {
                                                    TimeSpan tSpan = new TimeSpan(0, 1, 0); // Sleep thread 1 minute on attempt 2 through 5 before next attempt or giving up.
                                                    Thread.Sleep(tSpan);
                                                }

                                                if (i == 5)
                                                {
                                                    sCommText = "UPDATE [OSR_Orders] SET [Orders_DelCheck] = '1' WHERE [Orders_ProdNum] ='" + sRecSourceDirPNum + "' "; // Null = not checked : 1= checked.

                                                    TM.SQLNonQuery(sOSRConnString, sCommText);

                                                    string sException = "Unable to delete the following directory: " + sDir + " after 5 attempts.";

                                                    sCommText = "INSERT INTO [OSR_Errors] VALUES ('" + sException + "', '" + DateTime.Now.ToString().Trim() + "', '0')";

                                                    TM.SQLNonQuery(sOSRConnString, sCommText);
                                                }
                                            }
                                            catch (UnauthorizedAccessException)
                                            {
                                                bDeleted = false;

                                                if (i >= 2)
                                                {
                                                    TimeSpan tSpan = new TimeSpan(0, 1, 0); // Sleep thread 1 minute on attempt 2 through 5 before next attempt or giving up.
                                                    Thread.Sleep(tSpan);
                                                }

                                                if (i == 5)
                                                {
                                                    sCommText = "UPDATE [OSR_Orders] SET [Orders_DelCheck] = '1' WHERE [Orders_ProdNum] ='" + sRecSourceDirPNum + "' "; // Null = not checked : 1= checked.

                                                    TM.SQLNonQuery(sOSRConnString, sCommText);

                                                    string sException = "Unable to delete the following directory: " + sDir + " after 5 attempts.";

                                                    sCommText = "INSERT INTO [OSR_Errors] VALUES ('" + sException + "', '" + DateTime.Now.ToString().Trim() + "', '0')";

                                                    TM.SQLNonQuery(sOSRConnString, sCommText);
                                                }
                                            }
                                        }
                                    }
                                }
                            }
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                TM.SaveExceptionToDB(ex);
            }
        }

        #endregion

        #region Reporting methods.        

        private void Reporting()
        {
            if (bAPReportGenerated != true)
            {
                this.WeeklyTotalForWorkDone();
            }
        }

        #region Accounts payable report section.

        private void WeeklyTotalForWorkDone()
        {
            string sWeeklyEndDate = DateTime.Now.AddDays(-1).Date.ToString("MM/dd/yy").Trim();

            string sTempConName = string.Empty;
            List<string> listWeeklyTotal = new List<string>();
            string sWeeklyDistinctCount = string.Empty;

            List<string> Total = new List<string>();
            List<string> CodesAndQuantities = new List<string>();

            string sReportSavePath = string.Empty;

            try
            {
                if (DateTime.Now.ToString("ddd").Trim() == "Mon")
                {
                    string sCommText = "SELECT DISTINCT [Cntrctr_ID] FROM [OSR_Orders] WHERE Orders_ReceivedDate >= '" + DateTime.Now.Date.AddDays(-7).ToString("MM/dd/yy").Trim() + "'";
                    DataTable dt2 = new DataTable();

                    TM.SQLQuery(sOSRConnString, sCommText, dt2); // Select all distinct contractors for work done in the past week.

                    if (dt2.Rows.Count > 0)
                    {
                        foreach (DataRow dr in dt2.Rows)
                        {
                            string sWeeklyTotalConID = Convert.ToString(dr["Cntrctr_ID"]).Trim(); // Get the ID for each contractor that had completed work for the previous week.

                            string sCurrentLogFileDateTime = DateTime.Now.Date.ToString("MM/dd/yy").Trim();

                            sCommText = "SELECT * FROM [OSR_Reports] WHERE [Date] = '" + sCurrentLogFileDateTime + "' AND [ConID] = '" + sWeeklyTotalConID + "'";
                            DataTable dt = new DataTable();

                            TM.SQLQuery(sOSRConnString, sCommText, dt);

                            if (dt.Rows.Count > 0)
                            {
                                // Do nothing. Contractor report has already been generated.
                            }
                            else if (dt.Rows.Count == 0)
                            {
                                sCommText = "INSERT INTO [OSR_Reports] VALUES ('" + sCurrentLogFileDateTime + "', '" + sWeeklyTotalConID + "', '', '0', '0')";

                                TM.SQLNonQuery(sOSRConnString, sCommText); // Insert current datetime into the Reports table for this week.

                                sCommText = "SELECT * FROM [OSR_Contractor] WHERE [Cntrctr_ID] = '" + sWeeklyTotalConID + "' ";
                                DataTable dt3 = new DataTable();

                                TM.SQLQuery(sOSRConnString, sCommText, dt3); // For each distinct contractor that completed work for the previous week, set some variables.

                                if (dt3.Rows.Count > 0)
                                {
                                    string sWReportConFName = Convert.ToString(dt3.Rows[0]["Cntrctr_FName"]).Trim();
                                    string sWReportConLName = Convert.ToString(dt3.Rows[0]["Cntrctr_LName"]).Trim();
                                    string sWeeklyReportConName = sWReportConFName + " " + sWReportConLName;

                                    sCommText = "SELECT [Orders_ProdNum], [Orders_SentDate], [Orders_ReceivedDate] FROM OSR_Orders WHERE [Cntrctr_ID] = '" + sWeeklyTotalConID + "' AND Orders_ReceivedDate >= '" + DateTime.Now.Date.AddDays(-7).ToString("MM/dd/yy").Trim() + "'";
                                    DataTable dt4 = new DataTable();

                                    TM.SQLQuery(sOSRConnString, sCommText, dt4); // Select all orders for the current distinct contractor that had completed work for the previous week.

                                    if (dt4.Rows.Count > 0)
                                    {
                                        string sWeeklyProdNum = string.Empty;
                                        string sWeeklyOrderSentDate = string.Empty;
                                        string sWeeklyOrderRecDate = string.Empty;

                                        foreach (DataRow dr4 in dt4.Rows)
                                        {
                                            Total.Clear();
                                            CodesAndQuantities.Clear();

                                            sWeeklyProdNum = Convert.ToString(dr4["Orders_ProdNum"]).Trim();
                                            sWeeklyOrderSentDate = Convert.ToString(dr4["Orders_SentDate"]).Trim();
                                            sWeeklyOrderRecDate = Convert.ToString(dr4["Orders_ReceivedDate"]).Trim();

                                            int iHeadCount = 0;

                                            sCommText = "SELECT DISTINCT [Jobs_Descript], [Jobs_Headcount] FROM [OSR_Jobs] WHERE [Orders_ProdNum] ='" + sWeeklyProdNum + "' ";
                                            DataTable dt5 = new DataTable();

                                            TM.SQLQuery(sOSRConnString, sCommText, dt5); // Select the Jobs (image) data for the current order.

                                            if (dt5.Rows.Count > 0)
                                            {
                                                foreach (DataRow dr5 in dt5.Rows)
                                                {
                                                    string sWeeklyReportRetcodeDescript = Convert.ToString(dr5["Jobs_Descript"]).Trim();



                                                    sCommText = "SELECT [RetouchCodes_Code], [RetouchCodes_Price] FROM [OSR_RetouchCodes] WHERE [RetouchCodes_Description] = " +
                                                        "'" + sWeeklyReportRetcodeDescript + "'";
                                                    DataTable dt6 = new DataTable();

                                                    TM.SQLQuery(sOSRConnString, sCommText, dt6); // Select each retouch code in current order selected set some variables.

                                                    if (dt6.Rows.Count > 0)
                                                    {
                                                        string sWeeklyReportRetCode = Convert.ToString(dt6.Rows[0]["RetouchCodes_Code"]).Trim();
                                                        string sWeeklyReportCodePrice = Convert.ToString(dt6.Rows[0]["RetouchCodes_Price"]).Trim();



                                                        //************************************************************************************************
                                                        //Note: Possibly could drop this portion if paul gives me 2 new codes
                                                        //************************************************************************************************




                                                        if (sWeeklyReportRetcodeDescript == "16x20 Painter Portrait without Frame Add")
                                                        {
                                                            iHeadCount = Convert.ToInt32(dr5["Jobs_Headcount"]);
                                                            int iWeeklyReportCodePrice = Convert.ToInt32(dt6.Rows[0]["RetouchCodes_Price"]);
                                                            int iTotal = iWeeklyReportCodePrice * iHeadCount;

                                                            sWeeklyReportCodePrice = Convert.ToString(iTotal) + ".0000";
                                                        }


                                                        //************************************************************************************************
                                                        //************************************************************************************************
                                                        //************************************************************************************************




                                                        sWeeklyReportCodePrice = sWeeklyReportCodePrice.Remove(sWeeklyReportCodePrice.Length - 2);

                                                        this.WeeklyRetCodeCountsForAPReport(sWeeklyProdNum, sWeeklyReportRetcodeDescript, sWeeklyReportConName, sTempConName, listWeeklyTotal, sWeeklyDistinctCount, sWeeklyReportCodePrice, sWeeklyReportRetCode);                                                        

                                                        this.WeeklyGetRetCodeCountsForTxtReport(sWeeklyProdNum, sWeeklyReportRetcodeDescript, sWeeklyReportCodePrice, sWeeklyReportRetCode, Total, CodesAndQuantities);                                                        

                                                        bAPReportGenerated = true;
                                                    }
                                                    else if (dt6.Rows.Count == 0)
                                                    {
                                                        // In the instance of a non defined code coming through, save code to Errors table for notification.

                                                        string sException = "Code not defined: " + "[" + sWeeklyReportRetcodeDescript + "]" + " for production number: " + "[" + sWeeklyProdNum + "]";

                                                        sCommText = "INSERT INTO [OSR_Errors] VALUES ('" + sException + "', '" + DateTime.Now.ToString().Trim() + "', '0')";

                                                        TM.SQLNonQuery(sOSRConnString, sCommText);

                                                        return;
                                                    }
                                                }                                                

                                                this.GenerateWeeklyContractorJobBreakdownReport(sTempConName, sWeeklyReportConName, sWeeklyProdNum, sWeeklyOrderSentDate, sWeeklyOrderRecDate, Total, CodesAndQuantities, sWeeklyTotalConID, ref sReportSavePath);

                                            }
                                        }

                                        this.GenerateWeeklyAPReport(sWeeklyReportConName, listWeeklyTotal, sWReportConFName, sWReportConLName, sWeeklyEndDate, sCurrentLogFileDateTime, sWeeklyTotalConID);

                                        string sEmailServer = string.Empty;
                                        string sEmailMyBccAdd = string.Empty;
                                        string sSendToAdd2 = string.Empty;
                                        string sAPReportSendAddy = string.Empty;
                                        string sErrorSendToAddy = string.Empty;
                                        string sFromAddy = string.Empty;

                                        TM.EmailVariables(ref sEmailServer, ref sEmailMyBccAdd, ref sSendToAdd2, ref sAPReportSendAddy, ref sErrorSendToAddy, ref sFromAddy);

                                        EM.EmailReport(sEmailServer, sEmailMyBccAdd, sReportSavePath, sSendToAdd2, sFromAddy);

                                    }
                                }
                            }
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                TM.SaveExceptionToDB(ex);
            }
        }

        private void WeeklyRetCodeCountsForAPReport(string sWeeklyProdNum, string sWeeklyReportRetcodeDescript, string sWeeklyReportConName, string sTempConName, List<string> listWeeklyTotal, string sWeeklyDistinctCount, string sWeeklyReportCodePrice, string sWeeklyReportRetCode)
        {            
            string sWeeklyTotal = string.Empty;

            try
            {
                if (sTempConName != string.Empty) // If the previously assigned temp contractor record is not blank. (Would not be after running this method initially.)
                {
                    if (sTempConName != sWeeklyReportConName) // If the previosuly assigned temp contractor does not match the current contractor being processed.
                    {
                        listWeeklyTotal.Clear(); // Clear the list for the current contractor to keep report generation to a single contractor.
                    }
                }

                string sCommText = "SELECT COUNT(*) AS Count FROM [OSR_Jobs] WHERE [Jobs_Descript] = '" + sWeeklyReportRetcodeDescript +
                    "' AND [Orders_ProdNum] = '" + sWeeklyProdNum + "'";
                DataTable dt = new DataTable();

                TM.SQLQuery(sOSRConnString, sCommText, dt);

                if (dt.Rows.Count > 0)
                {
                    sWeeklyDistinctCount = Convert.ToString(dt.Rows[0]["Count"]).Trim();
                    






                    //************************************************************************************************
                    //Note: would need to change sWeeklyDistinctCount to = jobs_headcount if code = additional painter
                    //************************************************************************************************








                    decimal dWeeklysum = decimal.Parse(sWeeklyReportCodePrice) * int.Parse(sWeeklyDistinctCount);
                    if (dWeeklysum < 1)
                    {
                        sWeeklyTotal = dWeeklysum.ToString(".00");
                    }
                    if (dWeeklysum < 10 && dWeeklysum >= 1)
                    {
                        sWeeklyTotal = dWeeklysum.ToString("0.00");
                    }
                    if (dWeeklysum >= 10 && dWeeklysum < 100)
                    {
                        sWeeklyTotal = dWeeklysum.ToString("00.00");
                    }
                    else if (dWeeklysum >= 100)
                    {
                        sWeeklyTotal = dWeeklysum.ToString("000.00");
                    }

                    string sWeeklyList = sWeeklyDistinctCount + " x " + sWeeklyReportRetCode + " @ " + sWeeklyReportCodePrice + " = " + sWeeklyTotal;

                    var vWeeklyTotal = sWeeklyList.Substring(sWeeklyList.LastIndexOf('=') + 1);

                    listWeeklyTotal.Add(vWeeklyTotal);

                    sTempConName = sWeeklyReportConName;
                }

            }
            catch(Exception ex)
            {
                TM.SaveExceptionToDB(ex);
            }
        }

        private void GenerateWeeklyAPReport(string sWeeklyReportConName, List<string> listWeeklyTotal, string sWReportConFName, string sWReportConLName, string sWeeklyEndDate, string sCurrentLogFileDateTime, string sWeeklyTotalConID)
        {
            string sAPReportTemplate = @"\\nb\vol1\jobs\OSR\Reports\Template\OSR_Report.xlsx"; // Note: This file needs to be set as a variable in the Variables table.            
            string sAPReportSaveDir = @"\\nb\vol1\jobs\OSR\Reports\" + "[" + sWeeklyTotalConID + "]" + sWeeklyReportConName;
            string sAPReportConFile = sAPReportSaveDir + @"\OSR_Report-" + sWeeklyReportConName + "-" + DateTime.Now.ToString("M-dd-yyyy").Trim() + ".xlsx";

            if (!Directory.Exists(sAPReportSaveDir))
            {
                Directory.CreateDirectory(sAPReportSaveDir);
            }

            string sCommText = "UPDATE [OSR_Reports] SET [Report] = '" + sAPReportConFile + "', [Generated] = 0, [Sent] = 0 WHERE [Date] = '" + sCurrentLogFileDateTime + "' AND [ConID] = '" + sWeeklyTotalConID + "'";

            TM.SQLNonQuery(sOSRConnString, sCommText);

            try
            {
                decimal dWeeklySum = 0;

                for (int i = 0; i < listWeeklyTotal.Count; i++)
                {
                    dWeeklySum += Convert.ToDecimal(listWeeklyTotal[i]);
                }

                FileInfo xlsFile = new FileInfo(sAPReportTemplate);
                FileInfo xlsFileDone = new FileInfo(sAPReportConFile);

                ExcelPackage xlPkg = new ExcelPackage(xlsFile);
                ExcelWorkbook xlWB = xlPkg.Workbook;
                ExcelWorksheet xlWS = xlWB.Worksheets.First();

                xlWS.Cells[7, 2].Value = sWReportConFName;
                xlWS.Cells[7, 2].Style.Font.Bold = true;
                xlWS.Cells[7, 3].Value = sWReportConLName;
                xlWS.Cells[7, 3].Style.Font.Bold = true;
                xlWS.Cells[7, 6].Value = dWeeklySum;
                xlWS.Cells[7, 6].Style.Font.Bold = true;
                xlWS.Cells[14, 3].Value = sWeeklyEndDate;
                xlWS.Cells[14, 3].Style.Font.Bold = true;
                xlWS.Cells[15, 3].Value = DateTime.Now.AddDays(3).ToString("MM/dd/yy").Trim();
                xlWS.Cells[15, 3].Style.Font.Bold = true;
                xlPkg.SaveAs(xlsFileDone);

                string sEmailServer = string.Empty;
                string sEmailMyBccAdd = string.Empty;
                string sSendToAdd2 = string.Empty;
                string sAPReportSendAddy = string.Empty;
                string sErrorSendToAddy = string.Empty;
                string sFromAddy = string.Empty;

                TM.EmailVariables(ref sEmailServer, ref sEmailMyBccAdd, ref sSendToAdd2, ref sAPReportSendAddy, ref sErrorSendToAddy, ref sFromAddy);

#if (dev)
                string sAPSendAddy = "thegrump1976@gmail.com"; // Note: For testing only. Set this to sAPReportSendAddy for production use.
                EM.EmailAPReport(sEmailServer, sEmailMyBccAdd, sAPReportConFile, sAPSendAddy, sFromAddy);
#endif          
#if (!dev)
                EM.EmailAPReport(sEmailServer, sEmailMyBccAdd, sAPReportConFile, sAPReportSendAddy, sFromAddy);
#endif
                sCommText = "UPDATE [OSR_Reports] SET [Generated] = 1, [Sent] = 1 WHERE [Date] = '" + sCurrentLogFileDateTime + "' AND [ConID] = '" + sWeeklyTotalConID + "'";

                TM.SQLNonQuery(sOSRConnString, sCommText);                
            }
            catch(Exception ex)
            {
                TM.SaveExceptionToDB(ex);
            }
        }

        #endregion

        #region Job details report section.

        private void WeeklyGetRetCodeCountsForTxtReport(string sWeeklyProdNum, string sWeeklyReportRetcodeDescript, string sWeeklyReportCodePrice, string sWeeklyReportRetCode, List<string> Total, List<string> CodesAndQuantities)
        {
            string sDistinctCount = string.Empty;
            string sTotal = string.Empty;

            try
            {
                string sCommText = "SELECT COUNT(*) AS Count FROM [OSR_Jobs] WHERE [Jobs_Descript] = " + "'" + sWeeklyReportRetcodeDescript + "'" +
                    " AND [Orders_ProdNum] = " + "'" + sWeeklyProdNum + "'";
                DataTable dt = new DataTable();

                TM.SQLQuery(sOSRConnString, sCommText, dt);

                if (dt.Rows.Count > 0)
                {
                    sDistinctCount = Convert.ToString(dt.Rows[0]["Count"]).Trim();

                    decimal dsum = decimal.Parse(sWeeklyReportCodePrice) * int.Parse(sDistinctCount);
                    if (dsum < 1)
                    {
                        sTotal = dsum.ToString(".00");
                    }
                    if (dsum < 10 && dsum >= 1)
                    {
                        sTotal = dsum.ToString("0.00");
                    }
                    if (dsum >= 10 && dsum < 100)
                    {
                        sTotal = dsum.ToString("00.00");
                    }
                    else if (dsum >= 100)
                    {
                        sTotal = dsum.ToString("000.00");
                    }

                    string sList = sDistinctCount + " x " + sWeeklyReportRetCode + " @ " + sWeeklyReportCodePrice + " = " + sTotal;

                    var vTotal = sList.Substring(sList.LastIndexOf('=') + 1);

                    Total.Add(vTotal);

                    if (sWeeklyReportRetCode != "RFO") // Note: Create an exceptions table to filter codes here.
                    {
                        CodesAndQuantities.Add(sList);
                    }
                }
            }
            catch (Exception ex)
            {
                TM.SaveExceptionToDB(ex);
            }
        }

        private void GenerateWeeklyContractorJobBreakdownReport(string sTempConName, string sWeeklyReportConName, string sWeeklyProdNum, string sWeeklyOrderSentDate, string sWeeklyOrderRecDate, List<string> Total, List<string> CodesAndQuantities, string sWeeklyTotalConID, ref string sReportSavePath)
        {
            string sReportSaveFile = DateTime.Now.ToString("M-dd-yyyy").Trim() + ".txt";            
            string sReportDir = @"\\nb\vol1\jobs\OSR\Reports\" + "[" + sWeeklyTotalConID + "]" + sWeeklyReportConName;
            sReportSavePath = @"\\nb\vol1\jobs\OSR\Reports\" + "[" + sWeeklyTotalConID + "]" + sWeeklyReportConName + @"\" + sReportSaveFile;

            if (!Directory.Exists(sReportDir))
            {
                Directory.CreateDirectory(sReportDir);
            }

            try
            {
                StringBuilder sb = new StringBuilder();

                if (sTempConName != string.Empty)
                {
                    if (sTempConName != sWeeklyReportConName)
                    {
                        sb.AppendFormat("-----------------------------------------------");
                        sb.Append(Environment.NewLine);
                        sb.Append(Environment.NewLine);
                    }
                }

                sb.AppendFormat("Contractor: " + sWeeklyReportConName);
                sb.Append(Environment.NewLine);
                sb.AppendFormat("Production #: " + sWeeklyProdNum);
                sb.Append(Environment.NewLine);
                sb.AppendFormat("Sent Date: " + sWeeklyOrderSentDate);
                sb.Append(Environment.NewLine);
                sb.AppendFormat("Received Date: " + sWeeklyOrderRecDate);
                sb.Append(Environment.NewLine);

                foreach(string Codes in CodesAndQuantities)
                {
                    sb.AppendFormat("Codes: " + Codes);
                    sb.Append(Environment.NewLine);
                }

                decimal dSum = 0;

                for (int i = 0; i < Total.Count; i++)
                {
                    dSum += Convert.ToDecimal(Total[i]);
                }

                sb.AppendFormat("Total: $" + dSum);
                sb.Append(Environment.NewLine);
                sb.Append(Environment.NewLine);

                File.AppendAllText(sReportSavePath, sb.ToString());

                sTempConName = sWeeklyReportConName;
            }
            catch (Exception ex)
            {
                TM.SaveExceptionToDB(ex);
            }
        }

        #endregion

        #endregion
    }
}

