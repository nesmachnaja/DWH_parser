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
        }
    }
}
