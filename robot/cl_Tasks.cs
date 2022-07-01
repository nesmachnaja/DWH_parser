using robot.DataSet1TableAdapters;
using robot.RiskTableAdapters;
using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using static robot.DataSet1;

namespace robot
{
    class cl_Tasks
    {
        public cl_Tasks()
        {
            Task task = new Task(() =>
            {
                SPRisk sprisk = new SPRisk();
                sprisk.sp_MX_TOTAL_CESS(DateTime.Parse("31.05.2022"));
                //COUNTRY_LogTableAdapter adpr = new COUNTRY_LogTableAdapter();
                //adpr.GetData();
                Console.WriteLine("Ok");
            },
            TaskCreationOptions.LongRunning);


            task.RunSynchronously();

            cl_Send_Report send_report = new cl_Send_Report("MX_CESS", 1);


            //task.Start();
            //task.Wait();
            //SPRisk sprisk = new SPRisk();
        }
    }
}
