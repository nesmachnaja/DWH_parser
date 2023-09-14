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
        string pattern;
        Match result;
        public int query_result = 0;
        public int success = 0;

        public cl_Tasks(string procedure_calling)
        {
            Execute_SQL_Query(procedure_calling, string.Empty);

            /*string _procedure_calling = procedure_calling;

            //SqlCommand command = new SqlCommand("exec dbo.sp_MD_TOTAL_SNAP_CFIELD");
            //command.CommandTimeout = 300;
            //command.ExecuteNonQuery();

            pattern = @"\s\w+\.";
            result = Regex.Match(_procedure_calling, pattern);

            database_name = result.ToString().Replace(" ","").Replace(".",""); //procedure_calling.Replace("exec ", "").Substring(0, procedure_calling.IndexOf(".") - 5);
            cl_Connection_String connection_string = new cl_Connection_String(database_name);
            SqlConnection connection = new SqlConnection(connection_string.connectionString);

            Task task = new Task(() =>
            {
                try
                {
                    SqlCommand command = new SqlCommand(_procedure_calling);
                    connection.Open();
                    command.Connection = connection;
                    command.CommandTimeout = 600;
                    command.ExecuteNonQuery();

                    connection.Close();

                    Console.WriteLine("Ok");
                }
                catch (SqlException exc)
                {
                    Console.WriteLine(exc.Message);
                    connection.Close();
                }
            },
            TaskCreationOptions.LongRunning);

            task.RunSynchronously();*/
        }

        public cl_Tasks(string query, string result_field)
        {
            Execute_SQL_Query(query, result_field);
        }

        private void Execute_SQL_Query(string query, string result_field)
        {
            string _query = query;

            pattern = @"\s\w+\.";
            result = Regex.Match(_query, pattern);

            database_name = result.ToString().Replace(" ", "").Replace(".", ""); //procedure_calling.Replace("exec ", "").Substring(0, procedure_calling.IndexOf(".") - 5);
            cl_Connection_String connection_string = new cl_Connection_String(database_name);
            SqlConnection connection = new SqlConnection(connection_string.connectionString);

            Task task = new Task(() =>
            {
                try
                {
                    SqlCommand command = new SqlCommand(_query);
                    connection.Open();
                    command.Connection = connection;
                    command.CommandTimeout = 600;
                    if (result_field != string.Empty)
                    {
                        using (var reader = command.ExecuteReader())
                        {
                            if (reader.Read())
                            {
                                query_result = reader.GetInt32(reader.GetOrdinal(result_field));
                            }
                        }
                    }
                    else command.ExecuteNonQuery();

                    connection.Close();

                    Console.WriteLine("Ok");
                    success = 1;
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

        public cl_Tasks(string procedure_calling, DataTable dt)
        {
            string _procedure_calling = procedure_calling;

            database_name = procedure_calling.Replace("exec ", "").Substring(0, procedure_calling.IndexOf(".") - 5);

            pattern = @"@\S+";
            result = Regex.Match(procedure_calling, pattern); 
            string param_name = result.ToString();

            pattern = @"exec \S+";
            result = Regex.Match(procedure_calling, pattern); 
            procedure_calling = result.ToString().Replace("exec ","");
            
            cl_Connection_String connection_string = new cl_Connection_String(database_name);
            SqlConnection connection = new SqlConnection(connection_string.connectionString);


            Task task = new Task(() =>
            {
                try
                {
                    SqlCommand command = new SqlCommand(procedure_calling);
                    command.CommandType = CommandType.StoredProcedure;
                    connection.Open();
                    command.Connection = connection;
                    DataTable dataTable = new DataTable();
                    dataTable = dt;
                    SqlParameter param = command.Parameters.AddWithValue(param_name, dataTable);
                    param.TypeName = "dbo.tp_" + param_name.Replace("@","");
                    command.CommandTimeout = 600;
                    command.ExecuteNonQuery();

                    connection.Close();

                    Console.WriteLine("Ok");
                    success = 1;
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


    }
}
