using System;
using System.Collections.Generic;
using System.Data;
using System.Data.SqlClient;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Manager
{
    class SqlServer
    {
        private static String cnStr = String.Format(
            "Data Source={0};" +
            "Initial Catalog=phs_oc8;" +
            "Integrated Security=True",
            "192.168.135.14"
            );
        public static DataTable DownloadPaymentList()
        {
            DataTable dt = new DataTable();
            SqlCommand cmd = new SqlCommand();
            cmd.Connection = new SqlConnection(cnStr);
            cmd.CommandType = CommandType.StoredProcedure;
            cmd.CommandText = "[dbo].[СписаниеСРасчетногоСчета Выбрать]";
            SqlDataAdapter da = new SqlDataAdapter(cmd);
            da.Fill(dt);
            return dt;
        }
        public static void UploadPartnerList(DataTable dt)
        {
            SqlCommand cmd = new SqlCommand();
            cmd.Connection = new SqlConnection(cnStr);
            cmd.CommandType = CommandType.StoredProcedure;
            cmd.CommandText = "[dbo].[Контрагенты Добавить]";
            cmd.Parameters.Add("@Наименование", SqlDbType.NVarChar, 100);
            cmd.Parameters.Add("@НаименованиеПолное", SqlDbType.NVarChar, 250);
            cmd.Parameters.Add("@ИНН", SqlDbType.NVarChar, 50);
            cmd.Parameters.Add("@КПП", SqlDbType.NVarChar, 9);
            if ((dt != null) && (dt.Rows.Count > 0))
            {
                try
                {
                    cmd.Connection.Open();
                    foreach (DataRow dr in dt.Rows)
                    {
                        cmd.Parameters["@Наименование"].Value = dr[0];
                        cmd.Parameters["@НаименованиеПолное"].Value = dr[1];
                        cmd.Parameters["@ИНН"].Value = dr[2];
                        cmd.Parameters["@КПП"].Value = dr[3];
                        cmd.ExecuteNonQuery();
                    }
                }
                catch (Exception e) { Console.WriteLine(e.ToString()); }
                finally { cmd.Connection.Close(); }
            }
        }
        public static void UploadPaymentList(DataTable dt)
        {
            SqlCommand cmd = new SqlCommand();
            cmd.Connection = new SqlConnection(cnStr);
            cmd.CommandType = CommandType.StoredProcedure;
            cmd.CommandText = "[dbo].[ХозрасчетныйДвиженияССубконто Добавить]";
            cmd.Parameters.Add("@Дата", SqlDbType.Date);
            cmd.Parameters.Add("@Сумма", SqlDbType.Float);
            cmd.Parameters.Add("@ДтСчетКод", SqlDbType.NVarChar, 32);
            cmd.Parameters.Add("@ДтКонтрагентИНН", SqlDbType.NVarChar, 50);
            cmd.Parameters.Add("@ДтКонтрагентКПП", SqlDbType.NVarChar, 9);
            cmd.Parameters.Add("@ДтКонтрагентНаименование", SqlDbType.NVarChar, 100);
            cmd.Parameters.Add("@ДтКонтрагентНаименованиеПолное", SqlDbType.NVarChar, 250);
            cmd.Parameters.Add("@КтСчетКод", SqlDbType.NVarChar, 32);
            cmd.Parameters.Add("@КтКонтрагентИНН", SqlDbType.NVarChar, 50);
            cmd.Parameters.Add("@КтКонтрагентКПП", SqlDbType.NVarChar, 9);
            cmd.Parameters.Add("@КтКонтрагентНаименование", SqlDbType.NVarChar, 100);
            cmd.Parameters.Add("@КтКонтрагентНаименованиеПолное", SqlDbType.NVarChar, 250);
            cmd.Parameters.Add("@РегистраторПредставление", SqlDbType.NVarChar, 100);
            cmd.Parameters.Add("@РегистраторНазначениеПлатежа", SqlDbType.NVarChar, 210);
            cmd.Parameters.Add("@НомерТоргов", SqlDbType.NVarChar, 19);
            if ((dt != null) && (dt.Rows.Count > 0))
            {
                try
                {
                    cmd.Connection.Open();
                    foreach (DataRow dr in dt.Rows)
                    {
                        int ci = 0;
                        cmd.Parameters["@Дата"].Value = dr[ci++];
                        cmd.Parameters["@Сумма"].Value = dr[ci++];
                        cmd.Parameters["@ДтСчетКод"].Value = dr[ci++];
                        cmd.Parameters["@ДтКонтрагентИНН"].Value = dr[ci++];
                        cmd.Parameters["@ДтКонтрагентКПП"].Value = dr[ci++];
                        cmd.Parameters["@ДтКонтрагентНаименование"].Value = dr[ci++];
                        cmd.Parameters["@ДтКонтрагентНаименованиеПолное"].Value = dr[ci++];
                        cmd.Parameters["@КтСчетКод"].Value = dr[ci++];
                        cmd.Parameters["@КтКонтрагентИНН"].Value = dr[ci++];
                        cmd.Parameters["@КтКонтрагентКПП"].Value = dr[ci++];
                        cmd.Parameters["@КтКонтрагентНаименование"].Value = dr[ci++];
                        cmd.Parameters["@КтКонтрагентНаименованиеПолное"].Value = dr[ci++];
                        cmd.Parameters["@РегистраторПредставление"].Value = dr[ci++];
                        cmd.Parameters["@РегистраторНазначениеПлатежа"].Value = dr[ci++];
                        cmd.Parameters["@НомерТоргов"].Value = dr[ci++];
                        cmd.ExecuteNonQuery();
                    }
                }
                catch (Exception e) { Console.WriteLine(e.ToString()); }
                finally { cmd.Connection.Close(); }
            }
        }
    }
}
