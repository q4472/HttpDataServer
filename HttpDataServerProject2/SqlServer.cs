using Nskd;
using System;
using System.Data;
using System.Data.SqlClient;

namespace HttpDataServerProject2
{
    public class SqlServer
    {
        private static String cnString = String.Format("Data Source={0};Initial Catalog=Pharm-Sib;Integrated Security=True", Program.MainSqlServerDataSource);

        public static ResponsePackage Exec(RequestPackage rqp)
        {
            ResponsePackage rsp = new ResponsePackage
            {
                Status = "Запрос к Sql серверу.",
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
                        Object value = p.Value;
                        if (!String.IsNullOrWhiteSpace(key) && (value != null) && (value != DBNull.Value))
                        {
                            String parameterName = ((key[0] == '@') ? key : "@" + key);
                            if (value.GetType().ToString() == "System.String")
                            {
                                value = ((String)value).Replace((Char)0xfffd, ' ');
                                if (String.IsNullOrWhiteSpace((String)value))
                                {
                                    value = DBNull.Value;
                                }
                            }
                            cmd.Parameters.AddWithValue(parameterName, value);
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
                    int rowCount = da.Fill(rsp.Data);
                    rsp.Status = String.Format("Обработано {0} строк.", rowCount);
                }
            }
            catch (SqlException ex) 
            { 
                Log.Write(ex.ToString()); 
                rsp.Status = ex.ToString(); 
                rsp.Data = null; 
            }
            return rsp;
        }
    }
}
