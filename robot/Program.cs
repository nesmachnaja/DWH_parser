using Newtonsoft.Json.Linq;
using robot.Parsers;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace robot
{
    class Program
    {
        static JObject accounts; 
        static void Main(string[] args)
        {
            Console.WriteLine("Appoint a country: ");
            string country = Console.ReadLine();

            switch (country)
            {
                case "BIH":
                    {
                        cl_Parser_BIH Parser = new cl_Parser_BIH();
                        Parser.StartParsing();
                        break;
                    }
                case "LIGA":
                    {
                        cl_Parser_LIGA Parser = new cl_Parser_LIGA();
                        Parser.OpenFile();
                        break;
                    }
                case "MD":
                    {
                        cl_Parser_MD Parser = new cl_Parser_MD();
                        Parser.OpenFile();
                        break;
                    }
                case "MKD":
                    {
                        cl_Parser_MKD Parser = new cl_Parser_MKD();
                        Parser.OpenFile();
                        break;
                    }
                case "SMS":
                    {
                        cl_Parser_SMS Parser = new cl_Parser_SMS();
                        Parser.OpenFile();
                        break;
                    }
                case "MX":
                    {
                        cl_Parser_MX Parser = new cl_Parser_MX();
                        Parser.StartParsing();
                        break;
                    }

            }


            try
            {
                accounts = JObject.Parse(File.ReadAllText(@"js_Accounts.json"));
                JToken account_param;
                foreach (JObject account in accounts["accounts"])
                    if (account["name"].ToString().Equals(country))
                    {
                        account_param = (JToken)account["transport"];
                        cl_Send_Report Report = new cl_Send_Report(account_param);
                    }
            }
            catch (FileNotFoundException ex)
            {
                Console.WriteLine("Configuration file wasnt found.");
                Console.ReadLine();
                return;
            }


            Console.ReadKey();

        }

    }
}
