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
                //SP sp = new SP();
                sprisk.sp_MD_TOTAL_SNAP(); //(DateTime.Parse("31.05.2022"));
                //sp.sp_SMS_TOTAL_SNAP_CFIELD(); // (DateTime.Parse("31.05.2022"));
                Console.WriteLine("Ok");
            },
            TaskCreationOptions.LongRunning);

            task.RunSynchronously();


            //cl_Send_Report send_report = new cl_Send_Report("LIGA_SNAP", 1);


            //task.Start();
            //task.Wait();
            //SPRisk sprisk = new SPRisk();
        }
    }
}
