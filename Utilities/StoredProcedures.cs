using Nskd;
using System;
using System.Data;
using System.Data.SqlClient;
using System.Text;
using System.Text.RegularExpressions;

namespace HttpDataServices.Utilities
{
    class StoredProcedures
    {
        private static void WriteToConsole(RequestPackage rqp)
        {
            StringBuilder msg = new StringBuilder();
            msg.AppendFormat("{0:yyyy-MM-dd hh:mm:ss}\n", DateTime.Now);
            msg.AppendFormat("SessionId: '{0}'\n", rqp.SessionId);
            msg.AppendFormat("Command: '{0}'\n", rqp.Command);
            if (rqp.Parameters != null)
            {
                foreach (RequestParameter p in rqp.Parameters)
                {
                    msg.AppendFormat("Parameter: name = '{0}', value = '{1}'\n", p.Name, p.Value);
                }
            }
            Console.WriteLine(msg.ToString());
        }
        private static ResponsePackage CutName(RequestPackage rqp)
        {
            ResponsePackage rsp = null;
            String name = rqp["Name"] as String;
            if (!String.IsNullOrWhiteSpace(name))
            {
                rsp = new ResponsePackage();
                rsp.Status = "HttpDataServices.Utilities.StoredProcedures.cutName()";

                DataTable dt = Db.GetCuts();

                foreach (DataRow dr in dt.Rows)
                {
                    name = (new Regex(dr["сокращаемое"] as String)).Replace(name, dr["сокращённое"] as String);
                }

            }
            rsp.Status = name;
            return rsp;
        }
        public static ResponsePackage Execute(RequestPackage rqp)
        {
            ResponsePackage rsp = null;

            if (rqp != null)
            {
                switch (rqp.Command)
                {
                    case "CutName":
                        rsp = CutName(rqp);
                        break;
                    default:
                        WriteToConsole(rqp);
                        break;
                }
            }
            return rsp;
        }
    }
    public class Db
    {
        private static String cnString = String.Format("Data Source={0};Integrated Security=True", Program.MainSqlServerDataSource);
        public static DataTable GetCuts()
        {
            DataTable dt = new DataTable();
            try
            {
                SqlCommand cmd = new SqlCommand();
                cmd.Connection = new SqlConnection(cnString);
                cmd.CommandType = CommandType.StoredProcedure;
                cmd.CommandText = "[Utilities].[dbo].[получить_сокращения]";
                SqlDataAdapter da = new SqlDataAdapter(cmd);
                da.Fill(dt);
            }
            catch (Exception e) { Console.WriteLine(e.Message); }
            return dt;
        }
    }
}
