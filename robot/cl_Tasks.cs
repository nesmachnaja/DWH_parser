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
using System.Text.RegularExpressions;
using System.Threading;
using System.Threading.Tasks;

namespace robot
{
    class cl_Tasks
    {
        //static JObject connections;
        string database_name;
        //string connectionString = "";
        
        public cl_Tasks(string procedure_calling)
        {
            string _procedure_calling = procedure_calling;

            //SqlCommand command = new SqlCommand("exec dbo.sp_MD_TOTAL_SNAP_CFIELD");
            //command.CommandTimeout = 300;
            //command.ExecuteNonQuery();

            database_name = procedure_calling.Replace("exec ", "").Substring(0, procedure_calling.IndexOf(".") - 5);
            cl_Connection_String connection_string = new cl_Connection_String(database_name);
            //GetConnectionString();
            SqlConnection connection = new SqlConnection(connection_string.connectionString);

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
        
        public cl_Tasks(string procedure_calling, DataTable dt)
        {
            string _procedure_calling = procedure_calling;

            database_name = procedure_calling.Replace("exec ", "").Substring(0, procedure_calling.IndexOf(".") - 5);

            string pattern = @"@\S+";
            Match result = Regex.Match(procedure_calling, pattern); 
            string param_name = result.ToString();

            pattern = @"exec \S+";
            result = Regex.Match(procedure_calling, pattern); 
            procedure_calling = result.ToString().Replace("exec ","");
            
            cl_Connection_String connection_string = new cl_Connection_String(database_name);
            //GetConnectionString();
            SqlConnection connection = new SqlConnection(connection_string.connectionString);


            Task task = new Task(() =>
            {
                try
                {
                    SqlCommand command = new SqlCommand(procedure_calling);
                    command.CommandType = CommandType.StoredProcedure;
                    connection.Open();
                    command.Connection = connection;
                    //command.Parameters[0].GetType();
                    DataTable dataTable = new DataTable();
                    dataTable = dt;
                    SqlParameter param = command.Parameters.AddWithValue(param_name, dataTable);
                    param.TypeName = "dbo.tp_" + param_name.Replace("@","");
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
            },
            TaskCreationOptions.LongRunning);

            task.RunSynchronously();


        }

        //private void GetConnectionString()
        //{
        //    try
        //    {
        //        connections = JObject.Parse(File.ReadAllText(@"js_Connections.json"));
        //        JToken connection_param;
        //        foreach (JObject connection in connections["connections"])
        //            if (connection["name"].ToString().Equals(database_name))
        //            {
        //                connection_param = connection["parameters"];
        //                connectionString = connection_param["connectionString"].ToString();

        //                return;
        //            }
        //    }
        //    catch (FileNotFoundException ex)
        //    {
        //        Console.WriteLine("Configuration file wasnt found.");
        //        Console.ReadLine();
        //        return;
        //    }
        //}
    }
}
