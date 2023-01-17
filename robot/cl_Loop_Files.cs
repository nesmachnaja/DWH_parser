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
        cl_Parser_MD_test parse_md;

        public cl_Loop_Files(string country)
        {
            Console.WriteLine("Appoint folder path:");
            string path = Console.ReadLine();
            if (Directory.Exists(path))
            {
                files = Directory.GetFiles(path, @"*.xlsx", SearchOption.TopDirectoryOnly);
            }

            if (country.ToLower() == "sms")
            {
                foreach (string file_path in files)
                {
                    string pattern = @"(ces\S+$)|(prosh\S+$)|(portf\S+$)";
                    Match result = Regex.Match(file_path, pattern);
                    if (result.Value.ToString() != "")
                    {
                        parse_sms = new cl_Parser_SMS_test();
                        parse_sms.StartParsing(file_path);
                    }
                }

                parse_sms.CessPostProcessing();
                parse_sms.SnapPostProcessing();
            }
            
            if (country.ToLower() == "md")
            {
                int dca_num = 0;
                int snap_num = 0;

                foreach (string file_path in files)
                {
                    string pattern = @"(Plati.+$)|(Moldova_SNAP.+$)|(Moldova_WO.+$)";
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

                //Console.WriteLine(dca_num.ToString() + snap_num.ToString());
                if (dca_num != 0) parse_md.DcaPostProcessing();
                if (snap_num != 0) parse_md.SnapPostProcessing();
            }
        }
    }
}
