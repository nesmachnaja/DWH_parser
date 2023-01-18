using Microsoft.Office.Interop.Excel;
using robot.DataSet1TableAdapters;
using robot.RiskTableAdapters;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace robot.Parsers
{
    class cl_Parser
    {
        public int lastUsedRow;
        public COUNTRY_LogTableAdapter logAdapter;
        public SP sp = new SP();
        public SPRisk sprisk = new SPRisk();
        public string report;
        public string pathFile;
        public DateTime reestr_date;
        public int success = 0;
        public string brand = "";
        public cl_Tasks task;
        public cl_Send_Report send_report;
        //public Application ex = new Application();
    }
}
