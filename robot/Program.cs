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
        private static string path;
        static void Main(string[] args)
        {
            Console.WriteLine("Appoint a country: ");
            string country = Console.ReadLine();

            switch (country)
            {
                case "BIH":
                    {
                        cl_Parser_BIH Parser = new cl_Parser_BIH();
                        Parser.OpenFile();
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
            }

            //cl_Parser_MD Parser = new cl_Parser_MD();
            //Parser.OpenFile();
            //cl_Parser_BIH Parser = new cl_Parser_BIH();
            //Parser.OpenFile();
            //cl_Parser_LIGA Parser = new cl_Parser_LIGA();
            //Parser.OpenFile();
            //cl_Parser_SMS Parser = new cl_Parser_SMS();
            //Parser.OpenFile();
            //cl_Parser_MKD Parser = new cl_Parser_MKD();
            //Parser.OpenFile();
        }

    }
}
