using Newtonsoft.Json.Linq;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace robot
{
    class cl_Connection_String
    {
        static JObject connections;
        string _database_name;
        public string connectionString = "";

        public cl_Connection_String(string database_name)
        {
            _database_name = database_name;
            GetConnectionString();
        }

        private void GetConnectionString()
        {
            try
            {
                connections = JObject.Parse(File.ReadAllText(@"js_Connections.json"));
                JToken connection_param;
                foreach (JObject connection in connections["connections"])
                    if (connection["name"].ToString().Equals(_database_name))
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
