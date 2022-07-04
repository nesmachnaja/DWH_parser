using robot.Parsers;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace robot
{
    class Program
    {
        static void Main(string[] args)
        {
            //cl_Tasks tasks = new cl_Tasks();

            Console.WriteLine("Appoint a country: ");
            string country = Console.ReadLine();

            //cl_Send_Report report = new cl_Send_Report("test");


            switch (country.ToLower())
            {
                case "bih":
                    {
                        cl_Parser_BIH Parser = new cl_Parser_BIH();
                        Parser.StartParsing();
                        break;
                    }
                case "liga":
                    {
                        cl_Parser_LIGA Parser = new cl_Parser_LIGA();
                        Parser.OpenFile();
                        break;
                    }
                case "md":
                    {
                        cl_Parser_MD Parser = new cl_Parser_MD();
                        Parser.OpenFile();
                        break;
                    }
                case "mkd":
                    {
                        cl_Parser_MKD Parser = new cl_Parser_MKD();
                        Parser.StartParsing();
                        break;
                    }
                case "sms":
                    {
                        cl_Parser_SMS Parser = new cl_Parser_SMS();
                        Parser.StartParsing();
                        break;
                    }
                case "mx":
                    {
                        cl_Parser_MX Parser = new cl_Parser_MX();
                        Parser.StartParsing();
                        break;
                    }

            }

            Console.ReadKey();

        }

    }
}
