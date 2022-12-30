using Microsoft.SqlServer.Management.Smo;
using Microsoft.SqlServer.Management.Smo.Agent;
using robot.RiskTableAdapters;
using System;
using System.Collections.Generic;
using System.Collections.Specialized;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using static robot.Risk;

namespace robot
{
    class cl_PQR_Forming
    {
        string _country;
        string _connection_string;
        WorkingJobsTableAdapter jobs = new WorkingJobsTableAdapter();
        cl_Tasks task;

        public cl_PQR_Forming (string country)
        {
            _country = country;

            if (_country == "mkd") 
            { 
                task = new cl_Tasks("exec Risk.dbo.sp_MKD2_LGD");
                return;
            }

            WorkingJobsDataTable working_jobs = jobs.GetData();
            if (working_jobs.Count == 0 || working_jobs.Select(row => row.job_name.Equals("Robot_PQR_LGD")).ElementAt(0) == false)
            {
                StartPQRJob();
            }
            else
            {
                Console.WriteLine("Job is busy");
            }

        }

        private void StartPQRJob()
        {
            cl_Connection_String connection_string = new cl_Connection_String("msdb");
            _connection_string = connection_string.connectionString;

            string pattern = @"Data Source=\S+;";
            Match result = Regex.Match(_connection_string, pattern);
            string server_name = result.Value.ToString().Replace("Data Source=", "").Replace(";", "");

            pattern = @"User ID=\S+;";
            result = Regex.Match(_connection_string, pattern);
            string login = result.Value.ToString().Replace("User ID=", "").Replace(";", "");

            pattern = @"Password=\S+$";
            result = Regex.Match(_connection_string, pattern);
            string password = result.Value.ToString().Replace("Password=", "");

            cl_Tasks task = new cl_Tasks("exec msdb.dbo.sp_update_jobstep @job_id = '35E49CD6-ABF2-40B2-BA1B-439EAA480D5D', @step_id = 1, @command = 'exec sp_Report_PQR_cube @country = ''" + _country + "'''");
            task = new cl_Tasks("exec msdb.dbo.sp_update_jobstep @job_id = '35E49CD6-ABF2-40B2-BA1B-439EAA480D5D', @step_id = 2, @command = 'exec sp_Report_LGD91_cube @country = ''" + _country + "'''");
            task = new cl_Tasks("exec msdb.dbo.sp_update_jobstep @job_id = '35E49CD6-ABF2-40B2-BA1B-439EAA480D5D', @step_id = 3, @command = 'exec sp_Report_LGD181_cube @country = ''" + _country + "'''");
            task = new cl_Tasks("exec msdb.dbo.sp_update_jobstep @job_id = '35E49CD6-ABF2-40B2-BA1B-439EAA480D5D', @step_id = 4, @command = 'exec sp_Report_LGD365_cube @country = ''" + _country + "'''");


            Server server = new Server(server_name);
            server.ConnectionContext.LoginSecure = false;
            server.ConnectionContext.Login = login;
            server.ConnectionContext.Password = password;

            server.JobServer.Jobs["Robot_PQR_LGD"]?.Start();
        }
    }
}
