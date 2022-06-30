using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Net;
using System.Net.Mail;
using System.Net.Mime;
using Newtonsoft.Json.Linq;
using robot.DataSet1TableAdapters;
using System.Data;

namespace robot
{
    class cl_Send_Report
    {
        public cl_Send_Report(JToken account)
        {
            string country = account["country"].ToString();
            var from_address = new MailAddress(account["email"].ToString(), "ETL_bot");
            List <MailAddress> to_address_list = new List<MailAddress>();

            COUNTRY_contactsTableAdapter contacts = new COUNTRY_contactsTableAdapter();
            DataTable contact_data = contacts.GetCountryContacts(country);
            foreach (DataRow email in contact_data.Rows)
                to_address_list.Add(new MailAddress(email.ItemArray[2].ToString(), ""));

            string from_password = account["password"].ToString();
            const string subject = "BOT_notification";
            const string body = "Hello!<br> <b>File</b> was uploaded successfully.";

            var smtp = new SmtpClient
            {
                Host = account["server"].ToString(),
                EnableSsl = true,
                Port = 587,
                DeliveryMethod = SmtpDeliveryMethod.Network,
                UseDefaultCredentials = false,
                Credentials = new NetworkCredential(from_address.Address, from_password)
            };
            foreach (MailAddress to_address in to_address_list)
            {
                using (var message = new MailMessage(from_address, to_address)
                {
                    Subject = subject,
                    IsBodyHtml = true,
                    Body = body
                })
                {
                    smtp.Send(message);
                }
            }
        }
    }
}
