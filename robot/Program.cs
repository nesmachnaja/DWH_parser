using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace robot
{
    class Program
    {
        static void Main(string[] args)
        {
            cl_Parser Parser = new cl_Parser();
            Parser.parse_MKD_DCA();
        }
    }
}
