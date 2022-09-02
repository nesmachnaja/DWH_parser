using Microsoft.SqlServer.Management.Smo;
using Microsoft.SqlServer.Management.Smo.Agent;
using System;
using System.Collections.Generic;
using System.Collections.Specialized;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;

namespace robot
{
    class cl_PQR_Forming
    {
        string _country;
        string _connection_string;

        public cl_PQR_Forming (string country)
        {
            //string asd = "EXEC msdb.dbo.sp_add_jobstep @job_id=N'1d6930b5-4903-4334-9bec-ddae4845a867', @step_name=N'step 2', \r\n\t\t@step_id=2, \r\n\t\t@cmdexec_success_code=0, \r\n\t\t@on_success_action=3, \r\n\t\t@on_success_step_id=0, \r\n\t\t@on_fail_action=2, \r\n\t\t@on_fail_step_id=0, \r\n\t\t@retry_attempts=0, \r\n\t\t@retry_interval=0, \r\n\t\t@os_run_priority=0, @subsystem=N'TSQL', \r\n\t\t@command=N'\r\ndeclare @country = ''DR''\r\nselect 2', \r\n\t\t@database_name=N'DWH_Risk', \r\n\t\t@flags=0";
            //Regex regex = new Regex(@"country = ''", RegexOptions.IgnoreCase);
            //asd.Contains(@"country = ''");

            //string text = "declare @country = ''CO''";
            //string pattern = @"country = ''\w+''";
            //string target = "country = ''" + country + "''";
            //Regex regex = new Regex(pattern);
            //Match result = Regex.Match(text, pattern);

            _country = country;
            
            cl_Connection_String connection_string = new cl_Connection_String("msdb");
            _connection_string = connection_string.connectionString;

            string pattern = @"Data Source=\S+;";
            Match result = Regex.Match(_connection_string, pattern);
            string server_name = result.Value.ToString().Replace("Data Source=","").Replace(";","");

            pattern = @"User ID=\S+;";
            result = Regex.Match(_connection_string, pattern);
            string login = result.Value.ToString().Replace("User ID=", "").Replace(";","");

            pattern = @"Password=\S+$";
            result = Regex.Match(_connection_string, pattern);
            string password = result.Value.ToString().Replace("Password=", "");

            cl_Tasks task = new cl_Tasks("exec msdb.dbo.sp_update_jobstep @job_id = '35E49CD6-ABF2-40B2-BA1B-439EAA480D5D', @step_id = 1, @command = 'exec sp_Report_PQR_cube @country = ''" + _country + "'''");
            task = new cl_Tasks("exec msdb.dbo.sp_update_jobstep @job_id = '35E49CD6-ABF2-40B2-BA1B-439EAA480D5D', @step_id = 2, @command = 'exec sp_Report_LGD91_cube @country = ''" + _country + "'''");
            task = new cl_Tasks("exec msdb.dbo.sp_update_jobstep @job_id = '35E49CD6-ABF2-40B2-BA1B-439EAA480D5D', @step_id = 3, @command = 'exec sp_Report_LGD181_cube @country = ''" + _country + "'''");


            Server server = new Server(server_name);
            server.ConnectionContext.LoginSecure = false;
            server.ConnectionContext.Login = login;
            server.ConnectionContext.Password = password;

            server.JobServer.Jobs["Robot_PQR_LGD"]?.Start();

        }

    }
}
