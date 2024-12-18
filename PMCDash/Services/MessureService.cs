using System;
using System.Collections.Generic;
using System.Data;
using System.Data.SqlClient;
using System.Linq;
using System.Threading.Tasks;

namespace PMCDash.Services
{
    public class MessureService
    {
        public string connectString { get; set; }

        public List<Tolerance> getToolTolerance(string SqlStr)
        {
            List<Tolerance> result = new List<Tolerance>();

            using (var conn = new SqlConnection(connectString))
            {
                if (conn.State != ConnectionState.Open)
                    conn.Open();
                using (SqlCommand comm = new SqlCommand(SqlStr, conn))
                {
                    using (SqlDataReader SqlData = comm.ExecuteReader())
                    {
                        if (SqlData.HasRows)
                        {
                            while (SqlData.Read())
                            {
                                result.Add(new Tolerance
                                {
                                    Id = int.Parse(SqlData["Id"].ToString().Trim()),
                                    toolId = SqlData["ToolID"].ToString().Trim(),
                                    position = int.Parse(SqlData["Position"].ToString().Trim()),
                                    value = double.Parse(SqlData["Value"].ToString().Trim()),
                                });
                            }
                        }
                    }
                }
            }

            return result;
        }
    }

    public class Tolerance
    {
        public int Id { get; set; }
        public string toolId { get; set; }
        public int position { get; set; }
        public double value { get; set; }

    }
}
