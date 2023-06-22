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
using System.IO;

namespace robot
{
    class cl_Send_Report
    {
        static JObject accounts;
        string _country;
        string _country_file;
        int _report_type;
        List<MailAddress> to_address_list = new List<MailAddress>();

        public cl_Send_Report(string country_file, int report_type)
        {
            _report_type = report_type;
            _country = country_file.Substring(0,country_file.IndexOf("_")).ToLower();
            //_country = "test";
            _country_file = country_file;

            GetContactList();
            //SendEmail(account);
        }

        private void SendEmail(JToken account)
        {
            //string country = account["country"].ToString();
            var from_address = new MailAddress(account["email"].ToString(), "ETL_bot");

            COUNTRY_contactsTableAdapter contacts = new COUNTRY_contactsTableAdapter();
            DataTable contact_data = contacts.GetCountryContacts(_country, _country_file);
            foreach (DataRow email in contact_data.Rows)
            {
                to_address_list.Add(new MailAddress(email.ItemArray[2].ToString(), ""));
            }

            string from_password = account["password"].ToString();

            COUNTRY_reportingTableAdapter report_data = new COUNTRY_reportingTableAdapter();
            DataRow message_row;
            message_row = report_data.GetMessageParameters(_report_type).Rows[0];

            string subject = message_row.ItemArray[1].ToString();
            string body = _country_file == "MD_SNAP" || _country_file.Contains("SMS") ? message_row.ItemArray[2].ToString().Insert(18, _country_file).Replace("File","Files").Replace("was","were") : message_row.ItemArray[2].ToString().Insert(18, _country_file);

            var smtp = new SmtpClient
            {
                Host = account["server"].ToString(),
                EnableSsl = true,
                Port = 587,
                DeliveryMethod = SmtpDeliveryMethod.Network,
                UseDefaultCredentials = false,
                Credentials = new NetworkCredential(from_address.Address, from_password)
            };
            ServicePointManager.SecurityProtocol = SecurityProtocolType.Tls12;
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

        private void GetContactList()
        {
            try
            {
                accounts = JObject.Parse(File.ReadAllText(@"js_Accounts.json"));
                JToken account_param;
                foreach (JObject account in accounts["accounts"])
                    if (account["name"].ToString().Equals(_country))
                    {
                        account_param = (JToken)account["transport"];
                        SendEmail(account_param);
                        //cl_Send_Report Report = new cl_Send_Report(account_param);

                        PrintReport();
                    }
            }
            catch (FileNotFoundException ex)
            {
                Console.WriteLine("Configuration file wasnt found.");
                Console.ReadLine();
                return;
            }
        }

        private void PrintReport()
        {
            Console.WriteLine("Report was sended to:");
            foreach (MailAddress email in to_address_list)
                Console.WriteLine(email.Address.ToString());
        }
    }
}
