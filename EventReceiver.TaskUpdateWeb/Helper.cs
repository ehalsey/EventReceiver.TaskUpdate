using Microsoft.SharePoint.Client;
using Microsoft.SharePoint.Client.Utilities;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Net;
using System.Net.Mail;
using System.Text;
using System.Web;

namespace EventReceiver.TaskUpdateWeb
{
    public static class Helper
    {
        public static System.Text.StringBuilder log = new System.Text.StringBuilder();

        public static void ClearLog()
        {
            log = new System.Text.StringBuilder();
        }

        public static void AddLog(string msg)
        {
            log.AppendLine("<BR/>" + DateTime.Now.ToString() + " -- " + msg);
            log.AppendLine("<HR/>");
        }
        public static void AddLog(string catagory, string msg)
        {
            log.AppendLine("<BR/>" + catagory);
            log.AppendLine("<BR/>" + DateTime.Now.ToString() + " -- " + msg);
            log.AppendLine("<HR/>");
        }

        public static string GetLog()
        {
            return log.ToString();
        }

        public static void SendEmail(List<string> emailTo, string subject, string body)
        {
            /*
            string smtpServer = "smtp.office365.com";
            string mailSender = "me@gauravgoyal.onmicrosoft.com";
            body = body.Replace("\n", "");
            StringBuilder html = new StringBuilder();
            html.AppendLine("<table width='100%' align='center' cellpadding='2' cellspacing='10'>");
            html.AppendLine("	<tr>");
            html.AppendLine("		<td colspan='2' style='font-family: Tahoma; font-size: 9pt;'>");
            html.AppendLine(body);
            html.AppendLine("		</td>");
            html.AppendLine("	</tr>");
            html.AppendLine("</table>");

            SmtpClient objClient = LoadSmtpInformation(smtpServer);
            objClient.Send(BuildMailMessage(emailTo, subject, html.ToString().Replace("\r\n", ""), mailSender));*/
        }

        private static SmtpClient LoadSmtpInformation(string smtpServer)
        {
            SmtpClient client = new SmtpClient(smtpServer);
            client.UseDefaultCredentials = false;
            ///////////client.Credentials = new NetworkCredential("user", "password");
            client.Port = 587;
            client.Host = "smtp.office365.com";
            client.DeliveryMethod = SmtpDeliveryMethod.Network;
            client.EnableSsl = true;
            return client;
        }

        private static MailMessage BuildMailMessage(List<string> emailTo, string subject, string msg, string smtpSender)
        {
            MailMessage message = new MailMessage();

            message.From = new MailAddress(smtpSender);
            foreach (var id in emailTo)
            {
                message.To.Add(id);
            }
            message.IsBodyHtml = true;
            message.Body = msg;
            message.Subject = subject;


            return message;
        }
    }

    public class TaskDetail
    {
        public string OldUserId { get; set; }
        public string OldUserName { get; set; }
        public string OldUserLoginId { get; set; }

        public string NewUserId { get; set; }
        public string NewUserName { get; set; }
        public string NewUserLoginId { get; set; }
        public string ListId { get; set; }
        public string WebId { get; set; }
        public string ItemId { get; set; }
        public string RelatedItem { get; set; }
    }

    public class WFRelatedItemDetails
    {
        public int ItemId { get; set; }
        public string ListId { get; set; }
        public string WebId { get; set; }
    }
}