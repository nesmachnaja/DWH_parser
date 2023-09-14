using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Net;
using System.Net.Mail;
using System.Net.Mime;

namespace robot
{
    class cl_Send_Email
    {
        public cl_Send_Email()
        {
            var fromAddress = new MailAddress("BIH_RISK@2pp.dev", "ETL_bot");
            var toAddress = new MailAddress("nesmachnaya@itexp.pro", "Liudmila");
            const string fromPassword = "w8VW8ntP3}";
            const string subject = "BOT_notification";
            const string body = "Hello! File was uploaded successfully.";

            var smtp = new SmtpClient
            {
                Host = "mx-1.2pp.dev",
                EnableSsl = true,
                Port = 587,
                DeliveryMethod = SmtpDeliveryMethod.Network,
                UseDefaultCredentials = false,
                Credentials = new NetworkCredential(fromAddress.Address, fromPassword)
            };
            using (var message = new MailMessage(fromAddress, toAddress)
            {
                Subject = subject,
                Body = body
            })
            {
                smtp.Send(message);
            }
        }
        //public cl_Send_Email()
        //{
        //    try
        //    {

        //        SmtpClient mySmtpClient = new SmtpClient("smtp.2pp.dev", 587);
        //        mySmtpClient.EnableSsl = true;

        //        // set smtp-client with basicAuthentication
        //        mySmtpClient.UseDefaultCredentials = false;
        //        System.Net.NetworkCredential basicAuthenticationInfo = new
        //           System.Net.NetworkCredential("MKD_RISK@2pp.dev", "M3&jh83AmE");
        //        mySmtpClient.Credentials = basicAuthenticationInfo;

        //        // add from,to mailaddresses
        //        MailAddress from = new MailAddress("MKD_RISK@2pp.dev", "ETL_bot");
        //        MailAddress to = new MailAddress("nesmachnaya@itexp.pro", "");
        //        MailMessage myMail = new System.Net.Mail.MailMessage(from, to);

        //        // add ReplyTo
        //        MailAddress replyTo = new MailAddress("reply@example.com");
        //        myMail.ReplyToList.Add(replyTo);

        //        // set subject and encoding
        //        myMail.Subject = "BOT_notification";
        //        myMail.SubjectEncoding = System.Text.Encoding.UTF8;

        //        // set body-message and encoding
        //        myMail.Body = "<b>Hello!</b><br>File <b>was uploaded successfully.</b>.";
        //        myMail.BodyEncoding = System.Text.Encoding.UTF8;
        //        // text or html
        //        myMail.IsBodyHtml = true;

        //        mySmtpClient.Send(myMail);
        //    }

        //    catch (SmtpException ex)
        //    {
        //        throw new ApplicationException
        //          ("SmtpException has occured: " + ex.Message);
        //    }
        //    catch (Exception ex)
        //    {
        //        throw ex;
        //    }
        //}
    }
}
