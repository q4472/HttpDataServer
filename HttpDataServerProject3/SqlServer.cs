using System;
using System.Collections;
using System.Collections.Generic;
using System.Data;
using System.Data.SqlClient;
using System.Globalization;
using System.Text.RegularExpressions;
using System.Threading;

namespace HttpDataServerProject3
{
    public class SqlServer
    {
        private static String mCnString = String.Format("Data Source={0};Initial Catalog=Pharm-Sib;Integrated Security=True", Program.MainSqlServerDataSource);

        public static String GetXlSetting(String name)
        {
            String value = null;
            SqlConnection cn = new SqlConnection(mCnString);
            cn.Open();
            try
            {
                String cmdText = "[dbo].[xl_settings_get_by_name]";
                SqlCommand cmd = new SqlCommand(cmdText, cn);
                cmd.CommandType = CommandType.StoredProcedure;
                cmd.Parameters.AddWithValue("@name", name);
                Object v = cmd.ExecuteScalar();
                if (v != DBNull.Value) value = (String)v;
            }
            catch (Exception) { }
            finally { cn.Close(); }
            return value;
        }
        public static void SetXlSetting(String name, String value)
        {
            SqlConnection cn = new SqlConnection(mCnString);
            cn.Open();
            try
            {
                String cmdText = "[dbo].[xl_settings_upsert]";
                SqlCommand cmd = new SqlCommand(cmdText, cn);
                cmd.CommandType = CommandType.StoredProcedure;
                cmd.Parameters.AddWithValue("@name", name);
                cmd.Parameters.AddWithValue("@value", value);
                cmd.ExecuteNonQuery();
            }
            catch (Exception) { }
            finally { cn.Close(); }
        }
        public static Int32 UpsertXlAgrTable(DataTable dt)
        {
            Int32 count = 0;
            if (dt.Rows.Count > 1)
            {
                SqlConnection cn = new SqlConnection(mCnString);
                SqlCommand cmd = new SqlCommand();
                cmd.CommandText = "[dbo].[xl_договоры_покупатели_2015_upsert]";
                cmd.CommandType = CommandType.StoredProcedure;
                cmd.Connection = cn;
                cn.Open();
                for (int ri = 1; ri < dt.Rows.Count; ri++)
                {
                    DataRow dr = dt.Rows[ri];
                    cmd.Parameters.Clear();
                    try
                    {
                        cmd.Parameters.AddWithValue("@id", Convert.ToGuidFromBase64(dr[0]));
                        cmd.Parameters.AddWithValue("@timestamp", Convert.ToDateTimeFromTimeSpan(dr[1]));
                        Object buff = Convert.ToBytesFromBase64(dr[2]);
                        if (buff != DBNull.Value) cmd.Parameters.AddWithValue("@hash", (Byte[])buff);
                        cmd.Parameters.AddWithValue("@f1", Convert.ToSqlChar(dr[3]));
                        cmd.Parameters.AddWithValue("@f2", Convert.ToSqlChar(dr[4]));
                        cmd.Parameters.AddWithValue("@f3", Convert.ToSqlChar(dr[5]));
                        cmd.Parameters.AddWithValue("@f4", Convert.ToSqlChar(dr[6]));
                        cmd.Parameters.AddWithValue("@f5", Convert.ToSqlChar(dr[7]));
                        cmd.Parameters.AddWithValue("@f6", Convert.ToSqlDatetime(dr[8]));
                        cmd.Parameters.AddWithValue("@f7", Convert.ToSqlDatetime(dr[9]));
                        cmd.Parameters.AddWithValue("@f8", Convert.ToSqlDatetime(dr[10]));
                        cmd.Parameters.AddWithValue("@f9", Convert.ToSqlDatetime(dr[11]));
                        cmd.Parameters.AddWithValue("@f10", Convert.ToSqlChar(dr[12]));
                        cmd.Parameters.AddWithValue("@f11", Convert.ToSqlChar(dr[13]));
                        cmd.Parameters.AddWithValue("@f12", Convert.ToSqlChar(dr[14]));
                        cmd.Parameters.AddWithValue("@f13", Convert.ToSqlFloat(dr[15]));
                        cmd.Parameters.AddWithValue("@f14", Convert.ToSqlInt(dr[16]));
                        cmd.Parameters.AddWithValue("@f15", Convert.ToSqlChar(dr[17]));
                        cmd.Parameters.AddWithValue("@f16", Convert.ToSqlChar(dr[18]));
                    }
                    catch (Exception exc) { Log.Write("UpsertAgrTable pars: " + exc.Message); }
                    try
                    {
                        count = cmd.ExecuteNonQuery();
                    }
                    catch (Exception exc) { Log.Write("UpsertAgrTable exec: " + exc.Message); }
                }
                cn.Close();
            }
            return count;
        }
        private static class Convert
        {
            private static CultureInfo ic = CultureInfo.InvariantCulture;
            public static Object ToGuidFromBase64(Object value)
            {
                Object result = DBNull.Value;
                if (value != DBNull.Value)
                {
                    String v = value as String;
                    if (!String.IsNullOrWhiteSpace(v))
                    {
                        Object buff = ToBytesFromBase64(value);
                        if (buff != DBNull.Value)
                        {
                            result = new Guid((Byte[])buff);
                        }
                    }
                }
                return result;
            }
            public static Object ToBytesFromBase64(Object value)
            {
                Object result = DBNull.Value;
                if (value != DBNull.Value)
                {
                    String v = value as String;
                    if (!String.IsNullOrWhiteSpace(v))
                    {
                        v += ((v.Length % 4) == 2) ? "==" : "=";
                        try
                        {
                            result = System.Convert.FromBase64String(v);
                        }
                        catch (Exception) { }
                    }
                }
                return result;
            }
            public static Object ToDateTimeFromTimeSpan(Object value)
            {
                Object result = DBNull.Value;
                if (value != DBNull.Value)
                {
                    String v = value as String;
                    if (!String.IsNullOrWhiteSpace(v))
                    {
                        Object seconds = ToSqlFloat(v);
                        if (seconds != DBNull.Value)
                        {
                            DateTime baseDate = new DateTime(2000, 1, 1);
                            result = baseDate.AddSeconds((Double)seconds);
                        }
                    }
                }
                return result;
            }
            public static Object ToSqlDatetime(Object value)
            {
                Object result = DBNull.Value;
                if (value != DBNull.Value)
                {
                    String v = value as String;
                    if (!String.IsNullOrWhiteSpace(v))
                    {
                        if ((new Regex("\\.").Matches(v).Count > 1) || v.Contains("-") || v.Contains("/"))
                        {
                            result = System.Convert.ToDateTime(v);
                        }
                        else // xl - value is a day count.
                        {
                            Object days = ToSqlFloat(v);
                            if (days != DBNull.Value)
                            {
                                DateTime baseDate = new DateTime(1899, 12, 30);
                                result = baseDate.AddDays((Double)days);
                            }
                        }
                    }
                }
                return result;
            }
            public static Object ToSqlFloat(Object value)
            {
                Object result = DBNull.Value;
                if (value != DBNull.Value)
                {
                    String v = value as String;
                    if (!String.IsNullOrWhiteSpace(v))
                    {
                        v = (new Regex(@"[^\d\.\,]")).Replace(v, "");
                        v = v.Replace(",", ".");
                        Double d;
                        if (Double.TryParse(v, NumberStyles.Float, ic, out d))
                        {
                            result = d;
                        }
                    }
                }
                return result;
            }
            public static Object ToSqlInt(Object value)
            {
                Object result = DBNull.Value;
                if (value != DBNull.Value)
                {
                    String v = value as String;
                    if (!String.IsNullOrWhiteSpace(v))
                    {
                        v = (new Regex(@"\D")).Replace(v, "");
                        Int32 i;
                        if (Int32.TryParse(v, NumberStyles.Integer, ic, out i))
                        {
                            result = i;
                        }
                    }
                }
                return result;
            }
            public static Object ToSqlChar(Object value)
            {
                Object result = DBNull.Value;
                if (value != DBNull.Value)
                {
                    String v = value as String;
                    if (!String.IsNullOrWhiteSpace(v))
                    {
                        result = v;
                    }
                }
                return result;
            }
        }
    }
}
