using Newtonsoft.Json.Linq;
using robot.DataSet1TableAdapters;
using robot.RiskTableAdapters;
using System;
using System.Collections.Generic;
using System.Data;
using System.Data.SqlClient;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading;
using System.Threading.Tasks;

namespace robot
{
    class cl_Tasks
    {
        static JObject connections;
        string database_name = "Risk";
        string connectionString = "";

        public cl_Tasks(string procedure_calling)
        {
            string _procedure_calling = procedure_calling;

            //SqlCommand command = new SqlCommand("exec dbo.sp_MD_TOTAL_SNAP_CFIELD");
            //command.CommandTimeout = 300;
            //command.ExecuteNonQuery();

            GetConnectionString();
            SqlConnection connection = new SqlConnection(connectionString);
            database_name = procedure_calling.Replace("exec ", "").Substring(0, procedure_calling.IndexOf(".") - 5);

            Task task = new Task(() =>
            {
                //SPRisk sprisk = new SPRisk();
                //SP sp = new SP();
                try
                {
                    SqlCommand command = new SqlCommand(_procedure_calling);
                    connection.Open();
                    command.Connection = connection;
                    command.CommandTimeout = 600;
                    command.ExecuteNonQuery();

                    connection.Close();

                    //sprisk.sp_MD_TOTAL_SNAP_CFIELD(); //(DateTime.Parse("31.07.2022"));
                    Console.WriteLine("Ok");
                }
                catch (SqlException exc)
                {
                    Console.WriteLine(exc.Message);
                    connection.Close();
                }
                //sp.sp_SMS_TOTAL_SNAP_CFIELD(); // (DateTime.Parse("31.05.2022"));
            },
            TaskCreationOptions.LongRunning);

            task.RunSynchronously();


            //cl_Send_Report send_report = new cl_Send_Report("LIGA_SNAP", 1);


            //task.Start();
            //task.Wait();
            //SPRisk sprisk = new SPRisk();
        }

        private void GetConnectionString()
        {
            try
            {
                connections = JObject.Parse(File.ReadAllText(@"js_Connections.json"));
                JToken connection_param;
                foreach (JObject connection in connections["connections"])
                    if (connection["name"].ToString().Equals(database_name))
                    {
                        connection_param = connection["parameters"];
                        connectionString = connection_param["connectionString"].ToString();

                        return;
                    }
            }
            catch (FileNotFoundException ex)
            {
                Console.WriteLine("Configuration file wasnt found.");
                Console.ReadLine();
                return;
            }
        }
    }
}
