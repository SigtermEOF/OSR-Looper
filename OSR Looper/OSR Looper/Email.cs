using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Net;
using System.Net.Mail;
using System.Net.Mime;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Windows.Forms;
using System.Data.Odbc;
using System.Data.OleDb;
using System.Data.Sql;
using System.Data.SqlClient;
using System.Diagnostics;

namespace OSR_Looper
{
    class Email
    {
        TaskMethods TM = new TaskMethods();

        public void Emailer()
        {            
            string sSendto = "";
            string sMysubject = "";
            string sMybody = "";
            string sServer = "email";
            string sFile = "";
            string sBcc = "jlett@advancedphoto.com";

            MailAddress from = new MailAddress("APSAUTO@ADVANCEDPHOTO.COM", "APS");
            MailAddress to = new MailAddress(sSendto);
            MailMessage message = new MailMessage(from, to);
            message.Subject = sMysubject;
            message.Body = sMybody;
            MailAddress bcc = new MailAddress(sBcc);
            message.Bcc.Add(bcc);
            Attachment data = new Attachment(sFile, MediaTypeNames.Application.Octet);
            message.Attachments.Add(data);
            SmtpClient myclient = new SmtpClient(sServer);
            myclient.Credentials = CredentialCache.DefaultNetworkCredentials;

            try
            {
                myclient.Send(message);
            }
            catch (Exception ex)
            {
                TM.SaveExceptionToDB(ex);
            }
        }

        public void EmailSend(string sEmailServer, string sEmailMyBccAdd, string sSendTo, string sSendtoAdd, string sSendtoAdd2, string sMysubject, string sMybody, string sFile, string sFromAddy)
        {
            MailAddress from = new MailAddress("APSAUTO@ADVANCEDPHOTO.COM", "APS");
            MailAddress to = new MailAddress(sSendTo);
            MailMessage message = new MailMessage(from, to);
            message.To.Add(sSendtoAdd2);
            message.To.Add(sSendtoAdd);
            message.Subject = sMysubject;
            message.Body = sMybody;
            MailAddress bcc = new MailAddress(sEmailMyBccAdd);
            message.Bcc.Add(bcc);
            Attachment data = new Attachment(sFile, MediaTypeNames.Application.Octet);
            message.Attachments.Add(data);
            SmtpClient myclient = new SmtpClient(sEmailServer);
            myclient.Credentials = CredentialCache.DefaultNetworkCredentials;

            try
            {
                myclient.Send(message);
            }
            catch (Exception ex)
            {
                TM.SaveExceptionToDB(ex);
            }
        }

        public void EmailRec(string sEmailServer, string sEmailMyBccAdd, string sSendTo, string sSendToAddRec, string sSendtoAdd2, string sSubject, string sBody, string sFromAddy)
        {
            MailAddress from = new MailAddress("APSAUTO@ADVANCEDPHOTO.COM", "APS");
            MailAddress to = new MailAddress(sSendTo);
            MailMessage message = new MailMessage(from, to);
            message.To.Add(sSendtoAdd2);
            message.To.Add(sSendToAddRec);
            message.Subject = sSubject;
            message.Body = sBody;
            MailAddress bcc = new MailAddress(sEmailMyBccAdd);
            message.Bcc.Add(bcc);
            SmtpClient myclient = new SmtpClient(sEmailServer);
            myclient.Credentials = CredentialCache.DefaultNetworkCredentials;

            try
            {
                myclient.Send(message);
            }
            catch (Exception ex)
            {
                TM.SaveExceptionToDB(ex);
            }
        }

        public void EmailError(string sEmailServer, string sEmailMyBccAdd, string sEmailMyErrorSendAdd, string sSubject, string sBody, string sFromAddy)
        {
            MailAddress from = new MailAddress("APSAUTO@ADVANCEDPHOTO.COM", "APS");
            MailAddress to = new MailAddress(sEmailMyErrorSendAdd);
            MailMessage message = new MailMessage(from, to);
            message.Subject = sSubject;
            message.Body = sBody;
            MailAddress bcc = new MailAddress(sEmailMyBccAdd);
            message.Bcc.Add(bcc);
            SmtpClient myclient = new SmtpClient(sEmailServer);
            myclient.Credentials = CredentialCache.DefaultNetworkCredentials;

            try
            {
                myclient.Send(message);
            }
            catch (Exception ex)
            {
                TM.SaveExceptionToDB(ex);
            }
        }


        public void EmailPW(string sEmailServer, string sEmailMyBccAdd, string sSendTo, string sSubject, string sBody, string sFile, string sFromAddy)
        {
            MailAddress from = new MailAddress("APSAUTO@ADVANCEDPHOTO.COM", "APS");
            MailAddress to = new MailAddress(sSendTo);
            MailMessage message = new MailMessage(from, to);
            message.Subject = sSubject;
            message.Body = sBody;
            MailAddress bcc = new MailAddress(sEmailMyBccAdd);
            message.Bcc.Add(bcc);
            Attachment data = new Attachment(sFile, MediaTypeNames.Application.Octet);
            message.Attachments.Add(data);
            SmtpClient myclient = new SmtpClient(sEmailServer);
            myclient.Credentials = CredentialCache.DefaultNetworkCredentials;

            try
            {
                myclient.Send(message);
            }
            catch (Exception ex)
            {
                TM.SaveExceptionToDB(ex);
            }
        }

        public void EmailReport(string sEmailServer, string sEmailMyBccAdd, string sReportSavePath, string sSendTo, string sFromAddy)
        {            
            string sMysubject = "Weekly OSR Report";
            string sMybody = "Last weeks OSR report is attached.";

            MailAddress from = new MailAddress("APSAUTO@ADVANCEDPHOTO.COM", "APS");
            //MailAddress to = new MailAddress(sSendTo); // Send to internal lab email addys.
            MailAddress to = new MailAddress(sEmailMyBccAdd); // Send to my internal email addy.
            MailMessage message = new MailMessage(from, to);
            message.To.Add(sSendTo); // Send to internal lab email addys.
            message.Subject = sMysubject;
            message.Body = sMybody;
            //MailAddress bcc = new MailAddress(sEmailMyBccAdd); // BCC my internal email addy.
            //message.Bcc.Add(bcc);
            Attachment data = new Attachment(sReportSavePath, MediaTypeNames.Application.Octet);
            message.Attachments.Add(data);
            SmtpClient myclient = new SmtpClient(sEmailServer);
            myclient.Credentials = CredentialCache.DefaultNetworkCredentials;

            try
            {
                myclient.Send(message);
            }
            catch (Exception ex)
            {
                TM.SaveExceptionToDB(ex);
            }
        }

        public void EmailAPReport(string sEmailServer, string sEmailMyBccAdd, string sFile, string sAPSendAddy, string sFromAddy)
        {
            string sMysubject = "Weekly OSR report for AP.";
            string sMybody = "Last weeks OSR report(s).";

            MailAddress from = new MailAddress("APSAUTO@ADVANCEDPHOTO.COM", "APS");
            MailAddress to = new MailAddress(sAPSendAddy);
            MailMessage message = new MailMessage(from, to);
            message.Subject = sMysubject;
            message.Body = sMybody;
            MailAddress bcc = new MailAddress(sEmailMyBccAdd);
            message.Bcc.Add(bcc);
            Attachment data = new Attachment(sFile, MediaTypeNames.Application.Octet);
            message.Attachments.Add(data);
            SmtpClient myclient = new SmtpClient(sEmailServer);
            myclient.Credentials = CredentialCache.DefaultNetworkCredentials;

            try
            {
                myclient.Send(message);
            }
            catch (Exception ex)
            {
                TM.SaveExceptionToDB(ex);
            }
        }


    }
}
