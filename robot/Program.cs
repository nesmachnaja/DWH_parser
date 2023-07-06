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
            //cl_Send_Report send_report = new cl_Send_Report("MX_DCA", 1);
            //cl_PQR_Forming pqr_test = new cl_PQR_Forming("md");

            //cl_Tasks task = new cl_Tasks("exec DWH_Risk.dbo.sp_MD_SNAP_raw @MD_SNAP_raw = ", new System.Data.DataTable());

            //cl_Tasks tasks = new cl_Tasks("exec Risk.dbo.sp_MD_TOTAL_SNAP_CFIELD");

            Console.WriteLine("Appoint a country: ");
            string country = Console.ReadLine();

            //cl_Send_Report report = new cl_Send_Report("kz_snap", 1);


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
                        Parser.StartParsing();
                        break;
                    }
                case "md":
                    {
                        /*cl_Parser_MD Parser = new cl_Parser_MD();
                        Parser.StartParsing();*/
                        cl_Loop_Files loop = new cl_Loop_Files(country);
                        break;
                    }
                case "mkd":
                    {
                        /*cl_Parser_MKD Parser = new cl_Parser_MKD();
                        Parser.StartParsing();*/
                        cl_Loop_Files loop = new cl_Loop_Files(country);
                        break;
                    }
                case "sms":
                    {
                        /*cl_Parser_SMS Parser = new cl_Parser_SMS();
                        Parser.StartParsing();*/
                        cl_Loop_Files loop = new cl_Loop_Files(country);
                        break;
                    }
                case "kz":
                    {
                        /*cl_Parser_SMS Parser = new cl_Parser_SMS();
                        Parser.StartParsing();*/
                        cl_Loop_Files loop = new cl_Loop_Files(country);
                        break;
                    }
                case "mx":
                    {
                        /*cl_Parser_MX Parser = new cl_Parser_MX();
                        Parser.StartParsing();*/
                        cl_Loop_Files loop = new cl_Loop_Files(country);
                        break;
                    }
                case "budget":
                    {
                        /*cl_Parser_MX Parser = new cl_Parser_MX();
                        Parser.StartParsing();*/
                        cl_Loop_Files loop = new cl_Loop_Files(country);
                        break;
                    }

            }

            if (country != "budget")
            {
                Console.WriteLine("Do you want to form PQR? Y - Yes, N - No");
                string reply = Console.ReadKey().Key.ToString();


                if (reply.Equals("Y"))
                {
                    cl_PQR_Forming pqr = new cl_PQR_Forming(country);
                }
            }

            Console.ReadKey();

        }

    }
}
