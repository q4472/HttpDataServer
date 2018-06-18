using Nskd;
using System;
using System.Data;
using System.Data.SqlClient;

namespace Project12
{
    public class SqlServer
    {
        private static String cnString = String.Format("Data Source={0};Initial Catalog=Pharm-Sib;Integrated Security=True", Program.MainSqlServerDataSource);

        public static ResponsePackage Exec(RequestPackage rqp)
        {
            ResponsePackage rsp = new ResponsePackage
            {
                Status = String.Format("Запрос к Sql серверу '{0}'.\n", Program.MainSqlServerDataSource),
                Data = null
            };
            SqlCommand cmd = new SqlCommand()
            {
                CommandType = CommandType.StoredProcedure,
                CommandText = rqp.Command,
                CommandTimeout = 300
            };
            if ((rqp != null) && (rqp.Parameters != null))
            {
                foreach (var p in rqp.Parameters)
                {
                    if (p != null)
                    {
                        String key = p.Name;
                        if (!String.IsNullOrWhiteSpace(key))
                        {
                            String parameterName = ((key[0] == '@') ? key : "@" + key);
                            cmd.Parameters.AddWithValue(parameterName, p.Value);
                        }
                    }
                }
            }
            try
            {
                using (cmd.Connection = new SqlConnection(cnString))
                {
                    rsp.Data = new DataSet();
                    SqlDataAdapter da = new SqlDataAdapter(cmd);
                    int rCode = da.Fill(rsp.Data);
                    rsp.Status += String.Format("Код возврата: {0}.\n", rCode);
                }
            }
            catch (SqlException ex) 
            { 
                Log.Write(ex.Message); 
                rsp.Status += String.Format("{0}\n", ex.Message); 
                rsp.Data = null; 
            }
            return rsp;
        }
    }
}
