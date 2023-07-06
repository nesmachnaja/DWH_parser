using robot.Parsers;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;

namespace robot
{
    class cl_Loop_Files
    {
        string[] files;
        cl_Parser_SMS_test parse_sms;
        cl_Parser_KZ_test parse_kz;
        cl_Parser_MD_test parse_md;
        cl_Parser_MX_test parse_mx;
        cl_Parser_MKD_test parse_mkd;
        cl_Parser_budget parse_budget;
        int dca_num = 0;
        int snap_num = 0;
        int cess_num = 0;

        public cl_Loop_Files(string country)
        {
            Console.WriteLine("Appoint folder path:");
            string path = Console.ReadLine();
            if (Directory.Exists(path))
            {
                files = Directory.GetFiles(path, @"*.xls?", SearchOption.TopDirectoryOnly);
            }

            if (country.ToLower() == "sms") //|| country.ToLower() == "kz")
            {
                foreach (string file_path in files)
                {
                    string pattern = @"([^~\$]ces\S+$)|([^~\$]prosh\S+$)|([^~\$]portf\S+$)";
                    Match result = Regex.Match(file_path, pattern);
                    if (result.Value.ToString() != "")
                    {
                        parse_sms = new cl_Parser_SMS_test();
                        parse_sms.StartParsing(country, file_path);
                    }
                }

                if (country == "sms") parse_sms.CessPostProcessing();
                parse_sms.SnapPostProcessing();
            }

            if (country.ToLower() == "kz")
            {
                foreach (string file_path in files)
                {
                    string pattern = @"([^~\$]ces\S+$)|([^~\$]prosh\S+$)|([^~\$]port.+$)";
                    Match result = Regex.Match(file_path, pattern);
                    if (result.Value.ToString() != "")
                    {
                        parse_kz = new cl_Parser_KZ_test();
                        parse_kz.StartParsing(country, file_path);
                    }
                }

                //if (country == "sms") parse_sms.CessPostProcessing();
                parse_kz.SnapPostProcessing();
            }

            if (country.ToLower() == "md")
            {
                dca_num = 0;
                snap_num = 0;
                cess_num = 0;

                foreach (string file_path in files)
                {
                    string pattern = @"([^~\$]Plati.+$)|([^~\$]Moldova_SNAP.+$)|([^~\$]Moldova_WO.+$)";
                    Match result = Regex.Match(file_path, pattern);
                    if (result.Value.ToString() != "")
                    {
                        if (result.Value.ToString().Contains("Plati")) dca_num++;
                        if (result.Value.ToString().Contains("SNAP") || result.Value.ToString().Contains("WO")) snap_num++;
                        parse_md = new cl_Parser_MD_test();
                        parse_md.StartParsing(file_path);
                        //Console.WriteLine(result.Value.ToString());
                    }
                }

                if (dca_num != 0) parse_md.DcaPostProcessing();
                if (snap_num != 0) parse_md.SnapPostProcessing();
            }

            if (country.ToLower() == "mx")
            {
                //dca_num = 0;
                //snap_num = 0;
                //cess_num = 0;

                foreach (string file_path in files)
                {
                    string pattern = @"(.*\\[^~\$]*Data_exchance_format.+$)|(.*\\[^~\$]*cessions.+$)";
                    Match result = Regex.Match(file_path, pattern);
                    if (result.Value.ToString() != "")
                    {
                        //Console.WriteLine(result.Value.ToString());
                        if (result.Value.ToString().Contains("Data_exchance_format")) dca_num++;
                        if (result.Value.ToString().Contains("cessions")) cess_num++;
                        parse_mx = new cl_Parser_MX_test();
                        parse_mx.StartParsing(country, file_path);
                    }
                }
            }
            if (country.ToLower() == "mkd")
            {
                //dca_num = 0;
                //snap_num = 0;
                //cess_num = 0;

                foreach (string file_path in files)
                {
                    string pattern = @"(.*\\[^~\$]*Loan\+snapshot.+$)|(.*\\[^~\$]*DCA.+$)";
                    Match result = Regex.Match(file_path, pattern);
                    if (result.Value.ToString() != "")
                    {
                        parse_mkd = new cl_Parser_MKD_test();
                        parse_mkd.StartParsing(country, file_path);
                    }
                }
            }

            if (country.ToLower() == "budget")
            {
                foreach (string file_path in files)
                {
                    string pattern = @".*\\[^~\$]*\w*_budg.+$";
                    Match result = Regex.Match(file_path, pattern);
                    if (result.Value.ToString() != "")
                    {
                        parse_budget = new cl_Parser_budget();
                        parse_budget.StartParsing(file_path);
                    }
                }

            }
        }
    }
}
