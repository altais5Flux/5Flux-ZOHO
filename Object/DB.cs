using System;
using System.Collections.Generic;
using System.Configuration;
using System.Data;
using System.Data.SqlClient;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace WebservicesSage.Object
{
    class DB
    {
        private static SqlConnection cnn;
        private static void Connect()
        {
            string connetionString;
            connetionString = @"Data Source=" + ConfigurationManager.AppSettings["SERVER"].ToString() + ";Initial Catalog=" + ConfigurationManager.AppSettings["DBNAME"].ToString() + ";User ID=" + ConfigurationManager.AppSettings["SQLUSER"].ToString() + ";Password=" + ConfigurationManager.AppSettings["SQLPWD"].ToString();
            cnn = new SqlConnection(connetionString);
            cnn.Open();
        }

        public static void Disconnect()
        {
            cnn.Close();
        }

        public static SqlDataReader Select(string sql)
        {
            Connect();

            SqlCommand command = new SqlCommand(sql, cnn);
            SqlDataReader dataReader = command.ExecuteReader();

            //Disconnect();
            return dataReader;
        }
        public static void Update(string ZohoEntityID, string CT_Num)
        {
            try
            {
                var sql = "UPDATE ODIAM.dbo.F_COMPTET SET ZohoEntityID = @ZohoEntityID WHERE CT_Num = @CT_Num ";
                
                Connect();
                using (SqlCommand command = new SqlCommand(sql, cnn))
                    {
                        command.Parameters.Add("@ZohoEntityID", SqlDbType.VarChar).Value = ZohoEntityID;
                        command.Parameters.Add("@CT_Num", SqlDbType.VarChar).Value = CT_Num;
                        
                        command.ExecuteNonQuery();
                        Disconnect();
                    }
                
            }
            catch (Exception s)
            {
                Disconnect();
                StringBuilder sb = new StringBuilder();
                sb.Append(DateTime.Now + "prospect :"+CT_Num + Environment.NewLine);
                sb.Append(DateTime.Now + "ZohoEntityID :" + ZohoEntityID + Environment.NewLine);
                sb.Append(DateTime.Now + s.Message + Environment.NewLine);
                sb.Append(DateTime.Now + s.StackTrace + Environment.NewLine);
                File.AppendAllText("Log\\ProspectZohoEntityId.txt", sb.ToString());
                sb.Clear();
            }
          
        }

    }
}
