using System;
using System.Collections.Generic;
using System.Data;
using System.Data.OleDb;
using System.Data.SqlClient;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Nskd;

namespace HttpDataServerProject31
{
    class Program
    {
        // Если нет параметров при запуске программы, то работаем с Sql server на моём компъютере.
        private static String MainSqlServerDataSource = "192.168.135.77";

        static void Main(string[] args)
        {
            Int32 year = DateTime.Now.Year;
            if (args != null)
            {
                foreach (String arg in args)
                {
                    if (arg == "-d14" || arg == "/d14") { MainSqlServerDataSource = "192.168.135.14"; }
                    if( arg != null && arg.Length > 0 && arg[0] == '2')
                    {
                        Int32.TryParse(arg, out year);
                    }
                }
            }
            //Console.WriteLine(year);
            Export(year.ToString());
        }
        public static String Export(String year)
        {
            String status = "?";

            // Список полей, которые будем выгружать
            Object[] md = new Object[] { 
                //            0 field, 1 header
                new String[] {"f0", "id"},
                new String[] {"f1", "Вид контракта"},
                new String[] {"f2", "№ п/п (внутр)"},
                new String[] {"f3", "№ договора (внешн)"},
                new String[] {"f4", "Наименование клиента"},
                new String[] {"f5", "Город (населенный пункт)"},
                new String[] {"f6", "Дата договора"},
                new String[] {"f7", "Дата внесения в реестр"},
                new String[] {"f8", "Контрольная дата возврата"},
                new String[] {"f9", "Дата возврата"},
                new String[] {"f10", "Нал"},
                new String[] {"f11", "Лицензия"},
                new String[] {"f12", "Ответственный менеджер"},
                new String[] {"f13", "Сумма"},
                new String[] {"f14", "Код 1С"},
                new String[] {"f15", "№ торгов"},
                new String[] {"f16", "Комментарий"}
            };

            // Файл куда будем выгружать
            String now = DateTime.Now.ToString("yyyy-MM-dd HH-mm-ss");
            String fileName = @"\\SRV-TS2\work\Реестры договоров\Выгрузка\Выгрузка за " + year + " год " + now + ".xlsx";

            // Таблица которую будем выгружать
            DataTable dt = new DataTable();
            String connectionstring = String.Format("Data Source={0};Initial Catalog=Pharm-Sib;Integrated Security=True", MainSqlServerDataSource);
            SqlConnection cn = new SqlConnection(connectionstring);
            String sql = "" +
                "select " +
                " [f0], [f1], [f2], [f3], [f4], [f5], [f6], [f7], [f8], [f9], " +
                " [f10], [f11], [f12], [f13], [f14], [f15], [f16] " +
                "from [dbo].[договоры 15 16 17] " +
                "where (len([f2]) > 3) and ([f2] like N'%-" + year.Substring(2) + "') " +
                "order by [num];";
            SqlDataAdapter da = new SqlDataAdapter(sql, cn);
            da.Fill(dt);
            dt.TableName = "[договоры_покупатели_" + year + "]";
            for (int fi = 0; fi < md.Length; fi++)
            {
                String columnName = ((String[])md[fi])[0];
                String caption = ((String[])md[fi])[1];
                dt.Columns[columnName].Caption = caption;
            }

            // Выгружаем
            status = OleExcel.DataTableToExcelFile(dt, fileName);

            return status;
        }
    }
    class OleExcel
    {
        public static string Message { get; private set; }

        ///
        /// Преобразует Excel-файл в DataTable
        ///
        ///Таблица для загрузки данных
        ///Полный путь к Excel-файлу
        ///SQL-запрос. Используйте $SHEETS$ для выбоки по всем листам
        public static void ExcelFileToDataTable(out DataTable dtData, string sFile, string sRequest)
        {
            DataSet dsData = new DataSet();

            string sConnStr = String.Format("Provider=Microsoft.ACE.OLEDB.12.0;Data Source={0};Extended Properties=\"{1};HDR=YES\";", sFile, sFile.EndsWith(".xlsx") ? "Excel 12.0 Xml" : "Excel 8.0");

            using (OleDbConnection odcConnection = new OleDbConnection(sConnStr))
            {
                odcConnection.Open();
                if (sRequest.IndexOf("$SHEETS$") != -1)
                {
                    using (DataTable dtMetadata = odcConnection.GetOleDbSchemaTable(OleDbSchemaGuid.Tables, new object[4] { null, null, null, "TABLE" }))
                    {
                        for (int i = 0; i < dtMetadata.Rows.Count; i++)
                            if (dtMetadata.Rows[i]["TABLE_NAME"].ToString().IndexOf("$") == -1)
                                dtMetadata.Rows.Remove(dtMetadata.Rows[i]);

                        foreach (DataRow drRow in dtMetadata.Rows)
                        {
                            string sLocalRequest = sRequest.Replace("$SHEETS$", String.Format("[{0}]", drRow["TABLE_NAME"]));
                            OleDbCommand odcCommand = new OleDbCommand(sLocalRequest, odcConnection);
                            using (OleDbDataAdapter oddaAdapter = new OleDbDataAdapter(((OleDbCommand)odcCommand)))
                                oddaAdapter.Fill(dsData);
                        }
                    }
                }
                else
                {
                    OleDbCommand odcCommand = new OleDbCommand(sRequest, odcConnection);
                    using (OleDbDataAdapter oddaAdapter = new OleDbDataAdapter(odcCommand))
                        oddaAdapter.Fill(dsData);
                }
                odcConnection.Close();
            }

            dtData = dsData.Tables[0];
        }

        ///
        /// Преобразует dataTable в Excel-файл
        ///
        ///Данные
        ///Полный путь к файлу
        ///Excel 2007 или новее
        /// В случае успеха возвращает true
        public static String DataTableToExcelFile(DataTable data, String fileName, bool b2007 = true)
        {
            String status = null;
            try
            {
                string cnStr = String.Format(
                    "Provider=Microsoft.ACE.OLEDB.12.0;" +
                    "Data Source={0};" +
                    "Extended Properties=\"{1};HDR=YES\";", fileName, b2007 ? "Excel 12.0 Xml" : "Excel 8.0");

                using (OleDbConnection cn = new OleDbConnection(cnStr))
                {
                    cn.Open();
                    using (OleDbCommand cmd = new OleDbCommand())
                    {
                        cmd.Connection = cn;

                        // Создание таблицы
                        cmd.CommandText = GenerateSqlStatementCreateTable(data);
                        cmd.ExecuteNonQuery();

                        // Генерируем скрипт создания строк со значениями в качестве параметров.
                        string sColumns, sParameters;
                        GenerateColumnsString(data, out sColumns, out sParameters);
                        cmd.CommandText = string.Format("INSERT INTO {0} ({1}) VALUES ({2})", data.TableName, sColumns, sParameters);

                        // Создание строк с данными.
                        for (int ri = 0; ri < data.Rows.Count; ri++)
                        {
                            DataRow dr = data.Rows[ri];
                            // Устанавливаем параметры для INSERT.
                            cmd.Parameters.Clear();
                            for (int ci = 0; ci < data.Columns.Count; ci++)
                            {
                                if (dr.IsNull(ci))
                                {
                                    cmd.Parameters.AddWithValue("@p" + ci, DBNull.Value);
                                }
                                else
                                {
                                    switch (data.Columns[ci].DataType.ToString())
                                    {
                                        case "System.DateTime":
                                            cmd.Parameters.AddWithValue("@p" + ci, ((DateTime)dr[ci]).ToString("dd.MM.yyyy"));
                                            break;
                                        default:
                                            cmd.Parameters.AddWithValue("@p" + ci, dr[ci]);
                                            break;
                                    }
                                }
                            }
                            cmd.ExecuteNonQuery();
                        }
                    }
                    cn.Close();
                }
                status = "Ok";
            }
            catch (Exception ex)
            {
                status = ex.ToString();
            }
            return status;
        }

        /// <summary>
        ///     Создает список столбцов ([columnname0],[columnname1],[columnname2])
        ///     и соответствующих им параметров (@p0,@p1,@p2).
        ///     В качестве разделителя используется запятая
        /// </summary>
        /// <param name="dtData">Данные</param>
        /// <param name="sColumns">Список столбцов</param>
        /// <param name="sParams">Список параметров</param>
        private static void GenerateColumnsString(DataTable dtData, out string sColumns, out string sParams)
        {
            StringBuilder sbColumns = new StringBuilder();
            StringBuilder sbParams = new StringBuilder();
            for (int i = 0; i < dtData.Columns.Count; i++)
            {
                if (i != 0)
                {
                    sbColumns.Append(',');
                    sbParams.Append(',');
                }
                sbColumns.AppendFormat("[{0}]", dtData.Columns[i].Caption);
                sbParams.AppendFormat("@p{0}", i);
            }

            sColumns = sbColumns.ToString();
            sParams = sbParams.ToString();
        }

        /// <summary>
        ///     Создает SQL-скрипт для создания таблицы, в соответствии с DataTable
        /// </summary>
        /// <param name="data">Данные</param>
        /// <returns>Возвращает запрос 'CREATE TABLE...'</returns>
        private static string GenerateSqlStatementCreateTable(DataTable data)
        {
            StringBuilder sb = new StringBuilder();
            sb.AppendFormat("CREATE TABLE {0} (", data.TableName);
            for (int i = 0; i < data.Columns.Count; i++)
            {
                DataColumn dc = data.Columns[i];
                string dataType;
                switch (dc.DataType.ToString())
                {
                    case "System.DateTime":
                        dataType = "DATETIME";
                        break;
                    case "System.Double":
                    case "System.Single":
                        dataType = "DOUBLE";
                        break;
                    case "System.Int16":
                    case "System.Int32":
                        dataType = "INT";
                        break;
                    default:
                        dataType = "NVARCHAR";
                        break;
                }
                sb.AppendFormat("[{0}] {1},", dc.Caption, dataType);
            }
            if (data.Columns.Count > 0) sb.Length--;
            sb.Append(")");
            return sb.ToString();
        }
    }
}
