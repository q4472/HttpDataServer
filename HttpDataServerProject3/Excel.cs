using Nskd;
using System;
using System.Collections.Generic;
using System.Data;
using System.Data.OleDb;
using System.Data.SqlClient;
using System.Diagnostics;
using System.Globalization;
using System.IO;
using System.Runtime.InteropServices;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading;
using Excel = Microsoft.Office.Interop.Excel;

namespace HttpDataServerProject3
{
    public static class XlServer
    {
        private static Object thisLock = new Object();
        public static ResponsePackage Exec(RequestPackage rqp)
        {
            ResponsePackage rsp = new ResponsePackage();
            rsp.Status = "ok";
            lock (thisLock)
            {
                switch (rqp.Command)
                {
                    case "RefreshSqlFromXl":
                        XlStoredProcedure.GetFreshData();
                        break;
                    case "Добавить":
                        XlStoredProcedure.Append(rqp);
                        break;
                    case "Обновить":
                        XlStoredProcedure.Update(rqp);
                        break;
                    case "Prep.F4.GetXlsTables":
                        rsp.Data = XlStoredProcedure.GetXlsTables(rqp);
                        break;
                    case "Prep.F4.GetDocTables":
                        rsp.Data = XlStoredProcedure.GetDocTables(rqp);
                        break;
                    default:
                        break;
                }
            }
            return rsp;
        }
    }
    /// <summary>
    /// Этот класс хранит подключение к Excel. Оно одно на всех.
    /// Открывается и выдаётся (если ещё не открыто) по запросу Open().
    /// Закрывается по таймеру через 15 минут после последнего запроса Open() или по запросу Close().
    /// </summary>
    public static class XlConnection
    {
        //private static Boolean isConnected;
        //private static Boolean isConnectingProcessStarts;
        //private static Boolean isDisconnectingProcessStarts;
        //private static DateTime disconnectingProcessStartTime;

        //private static XlTimer xlTimer = new XlTimer();

        public static String FileName = SqlServer.GetXlSetting("file name");
        public static String SheetName = SqlServer.GetXlSetting("sheet name");
        public static Excel.Application App;
        public static Excel.Workbook Book;
        public static Excel.Worksheet Sheet;
        public static void Open()
        {
            Log.Write("XlConnection.Open()");
            // проверяем текущее подключение
            //if (App == null)
            {
                // если нет
                // Запускаем Excel и таймер на 15 мин.
                try
                {
                    // сначала пробуем подключиться к уже открытому экземпляру
                    App = Marshal.GetActiveObject("Excel.Application") as Excel.Application;
                }
                catch (Exception)
                {
                    App = null;
                }
                if (App == null)
                {
                    // или создаём новый экземпляр Exel
                    App = new Excel.Application();
                }

                App.Visible = false;
                App.DisplayAlerts = false;

                // подключаемся к файлу
                Book = App.Workbooks.Open(FileName, false, false);
                Sheet = Book.Sheets[SheetName];

                //isConnected = true;

                // тамер на 15 минут
                //xlTimer.Start(60000, 15); // 60 сек 15 раз
            }

            // Каждый запрос начинает отсчёт времени сначала.
            //xlTimer.Reset(60000, 15);
        }
        public static void Close()
        {
            // Отключаемся от файла с сохраненнием.
            if (Sheet != null)
            {
                Sheet = null;
                releaseComObject(Sheet);
            }
            if (Book != null)
            {
                Book.Close(true);
                releaseComObject(Book);
            }
            if (App != null)
            {
                App.Quit();
                releaseComObject(App);
                //isConnected = false;
            }
        }
        private static void releaseComObject(object obj)
        {
            if (obj != null)
            {
                if (obj.GetType().ToString() == "System.__ComObject")
                {
                    try
                    {
                        System.Runtime.InteropServices.Marshal.ReleaseComObject(obj);
                        obj = null;
                    }
                    catch (Exception ex)
                    {
                        obj = null;
                        Log.Write(ex.ToString());
                    }
                    finally
                    {
                        GC.Collect();
                    }
                }
            }
        }
    }
    public static class XlStoredProcedure
    {
        public static void GetFreshData()
        {
            DateTime xlUpdated = DateTime.Parse(SqlServer.GetXlSetting("last update date"));
            FileInfo fi = new FileInfo(XlConnection.FileName);
            if (xlUpdated > fi.LastWriteTime) { return; }

            TimeSpan ts = xlUpdated - new DateTime(2000, 1, 1);

            // Загрузка из Excel новых данных
            DataTable dt = new DataTable();
            String xlCnString =
                "Provider='Microsoft.ACE.OLEDB.12.0';" +
                "Data Source='" + XlConnection.FileName + "';" +
                "Extended Properties='Excel 12.0 Macro;HDR=No;IMEX=1;';";

            using (OleDbConnection xlCn = new OleDbConnection(xlCnString))
            {
                OleDbCommand cmd = new OleDbCommand();
                cmd.Connection = xlCn;
                cmd.CommandText =
                    "select * " +
                    "from [" + XlConnection.SheetName + "$A:S] " +
                    "where f2 > '" + ts.TotalSeconds.ToString() + "'";
                try
                {
                    xlCn.Open();
                    //SqlServer.LogWrite("OleDbConnection.State: " + xlCn.State.ToString());
                    (new OleDbDataAdapter(cmd)).Fill(dt);
                }
                catch (Exception ex) { Log.Write(ex.Message); }
                finally
                {
                    xlCn.Close();
                    //SqlServer.LogWrite("OleDbConnection.State: " + xlCn.State.ToString());
                }
            }
            // Обновить Sql Server
            if ((dt != null) && (dt.Rows.Count > 0)) // есть ответ и как минимум строка заголовка
            {
                SqlServer.UpsertXlAgrTable(dt);
                SqlServer.SetXlSetting("last update date", DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss"));
            }
            return;
        }
        public static void Append(RequestPackage rqp)
        {
            try
            {
                Guid guid = Guid.NewGuid();
                if (!String.IsNullOrWhiteSpace(rqp["f0"] as String))
                {
                    Guid.TryParse((rqp["f0"] as String), out guid);
                }

                String f0 = System.Convert.ToBase64String(guid.ToByteArray()).Substring(0, 22);

                XlConnection.Open();
                Excel.Worksheet sheet = XlConnection.Sheet;
                // ищем первую пустую строку начиная с третьей
                int ri;
                for (ri = 3; ri < 10000; ri++)
                {
                    if (sheet.Cells[ri, 1].Value == null) break;
                }
                if (ri < 10000)
                {
                    // добавляем ячейки
                    Excel.Range row = sheet.Range[sheet.Cells[ri, 1], sheet.Cells[ri, 19]];

                    sheet.Unprotect();
                    XlConnection.App.EnableEvents = false;
                    row[1, 1] = f0;
                    XlConnection.App.EnableEvents = true;
                    sheet.Protect();

                    row[1, 4] = rqp["f3"] as String;
                    //row[1, 5] = rqp["f4"] as String; // сама обновится через макрос Exel
                    row[1, 6] = rqp["f5"] as String;
                    row[1, 7] = rqp["f6"] as String;
                    row[1, 8] = rqp["f7"] as String;
                    row[1, 9] = XlConvert.ToXlDatetime(rqp["f8"] as String);
                    row[1, 10] = XlConvert.ToXlDatetime(rqp["f9"] as String);
                    row[1, 11] = XlConvert.ToXlDatetime(rqp["f10"] as String);
                    row[1, 12] = XlConvert.ToXlDatetime(rqp["f11"] as String);
                    row[1, 13] = rqp["f12"] as String;
                    row[1, 14] = rqp["f13"] as String;
                    row[1, 15] = rqp["f14"] as String;
                    row[1, 16] = XlConvert.ToXlFloat(rqp["f15"] as String);
                    row[1, 17] = XlConvert.ToXlInt(rqp["f16"] as String);
                    row[1, 18] = rqp["f17"] as String;
                    row[1, 19] = rqp["f18"] as String;
                    XlConnection.App.DisplayAlerts = false;
                    XlConnection.Book.Save();
                }
            }
            catch (Exception ex)
            {
                Log.Write("XlStoredProcedure.Append(): " + ex.ToString());
            }
            finally { XlConnection.Close(); }
        }
        public static void Update(RequestPackage rqp)
        {
            try
            {
                Guid guid = Guid.Parse(rqp["f0"] as String);
                String f0 = System.Convert.ToBase64String(guid.ToByteArray()).Substring(0, 22);

                XlConnection.Open();

                Excel.Worksheet sheet = XlConnection.Sheet;
                Excel.Range usedRange = sheet.UsedRange;
                Excel.Range cell = usedRange.Find(f0);
                if (cell != null) // нашли - исправляем уже существующую строку
                {
                    Excel.Range row = sheet.Range[sheet.Cells[cell.Row, 1], sheet.Cells[cell.Row, 19]];
                    row[1, 4] = rqp["f3"] as String;
                    row[1, 5] = rqp["f4"] as String;
                    row[1, 6] = rqp["f5"] as String;
                    row[1, 7] = rqp["f6"] as String;
                    row[1, 8] = rqp["f7"] as String;
                    row[1, 9] = XlConvert.ToXlDatetime(rqp["f8"] as String);
                    row[1, 10] = XlConvert.ToXlDatetime(rqp["f9"] as String);
                    row[1, 11] = XlConvert.ToXlDatetime(rqp["f10"] as String);
                    row[1, 12] = XlConvert.ToXlDatetime(rqp["f11"] as String);
                    row[1, 13] = rqp["f12"] as String;
                    row[1, 14] = rqp["f13"] as String;
                    row[1, 15] = rqp["f14"] as String;
                    row[1, 16] = XlConvert.ToXlFloat(rqp["f15"] as String);
                    row[1, 17] = XlConvert.ToXlInt(rqp["f16"] as String);
                    row[1, 18] = rqp["f17"] as String;
                    row[1, 19] = rqp["f18"] as String;
                    XlConnection.App.DisplayAlerts = false;
                    XlConnection.Book.Save();
                }
            }
            catch (Exception ex)
            {
                Log.Write("XlStoredProcedure.Update(): " + ex.ToString());
            }
            finally { XlConnection.Close(); }
        }
        public static DataSet GetXlsTables(RequestPackage rqp)
        {
            Console.WriteLine("ReadXlsTables");
            DataSet ds = null;
            if (rqp != null && rqp.Parameters != null && rqp.Parameters.Length > 0)
            {
                Byte[] bytes = null;
                foreach (RequestParameter p in rqp.Parameters)
                {
                    if (p.Name == "fileStream")
                    {
                        bytes = p.Value as Byte[];
                    }
                }
                if (bytes != null)
                {
                    Console.WriteLine(bytes.Length.ToString());
                    String fileName = "M:\\Temp\\" + Guid.NewGuid().ToString() + ".xls";
                    try
                    {
                        FileStream tempFileStream = File.Open(fileName, FileMode.Create);
                        tempFileStream.Write(bytes, 0, bytes.Length);
                        tempFileStream.Close();
                        ds = FromXlsFile(fileName);
                    }
                    catch (Exception e) { ds = null; throw new Exception("Ошибка при разборе потока xls.", e); }
                    finally { File.Delete(fileName); }
                }
            }
            return ds;
        }
        private static DataSet FromXlsFile(String fileName)
        {
            DataSet ds = new DataSet();
            try
            {
                String xlCnString =
                 "Provider='Microsoft.ACE.OLEDB.12.0';" +
                 "Data Source='" + fileName + "';" +
                 "Extended Properties='Excel 12.0 Macro;HDR=No;IMEX=1;';";
                using (OleDbConnection xlCn = new OleDbConnection(xlCnString))
                {
                    try
                    {
                        xlCn.Open();
                        DataTable schema = xlCn.GetOleDbSchemaTable(OleDbSchemaGuid.Tables, null);
                        foreach (DataRow schemaRow in schema.Rows)
                        {
                            String sheetName = schemaRow["TABLE_NAME"] as String;
                            DataTable dt = new DataTable(sheetName);
                            OleDbCommand cmd = new OleDbCommand();
                            cmd.Connection = xlCn;
                            cmd.CommandText = "select * from [" + sheetName + "] ";
                            (new OleDbDataAdapter(cmd)).Fill(dt);
                            ds.Tables.Add(dt);
                        }
                        xlCn.Close();
                    }
                    catch (Exception ex) { Console.WriteLine(ex.Message); }
                }
                /*
                Int32 ProcessId = 0;
                List<Int32> pids = new List<Int32>();
                foreach (Process p in Process.GetProcesses())
                {
                    pids.Add(p.Id);
                }
                Excel.Application excel = new Excel.Application();
                foreach (Process p in Process.GetProcesses())
                {
                    if (!pids.Contains(p.Id) && (p.ProcessName.ToUpper() == "EXCEL"))
                    {
                        ProcessId = p.Id;
                        break;
                    }
                }

                excel.Visible = false;
                Excel.Workbook book = excel.Workbooks.Open(Filename: fileName, ReadOnly: true);
                try
                {
                    foreach (Excel.Worksheet sheet in book.Worksheets)
                    {
                        // для каждого листа в Excel делаем свою таблицу в DataSet
                        DataTable dt = new DataTable();
                        ds.Tables.Add(dt);

                        // В DataTable длаем 20 колонок
                        for (int ci = 0; ci < 20; ci++)
                        {
                            dt.Columns.Add(String.Format("Column{0}", ci), typeof(String));
                        }

                        // В DataTable длаем 100 строк
                        for (int ri = 0; ri < 100; ri++)
                        {
                            dt.Rows.Add(dt.NewRow());
                        }

                        // Заполняем из Excel
                        for (int ri = 0; ri < 100; ri++)
                        {
                            for (int ci = 0; ci < 20; ci++)
                            {
                                dt.Rows[ri][ci] = Convert.ToString(((Excel.Range)sheet.Cells[ri + 1, ci + 1]).Value2);
                            }
                        }
                    }
                    
                }
                catch (Exception e)
                {
                    ds = null; throw new Exception("Ошибка при разборе документа xls." + ProcessId.ToString(), e);
                }
                
                finally
                {
                    book.Close(SaveChanges: false);
                    //excel.Workbooks.Close();
                    if (ProcessId > 0)
                    {
                        Process.GetProcessById(ProcessId).Kill();
                    }
                }
            */
            }

            catch (Exception e) { ds = null; throw new Exception("Ошибка при разборе файла xls.", e); }
            return ds;
        }
        public static DataSet GetDocTables(RequestPackage rqp)
        {
            Console.WriteLine("ReadDocTables");
            DataSet ds = null;
            if (rqp != null && rqp.Parameters != null && rqp.Parameters.Length > 0)
            {
                Byte[] bytes = null;
                foreach (RequestParameter p in rqp.Parameters)
                {
                    if (p.Name == "fileStream")
                    {
                        bytes = p.Value as Byte[];
                    }
                }
                if (bytes != null)
                {
                    Console.WriteLine(bytes.Length.ToString());
                    String fileName = "M:\\Temp\\" + Guid.NewGuid().ToString() + ".doc";
                    try
                    {
                        FileStream tempFileStream = File.Open(fileName, FileMode.Create);
                        tempFileStream.Write(bytes, 0, bytes.Length);
                        tempFileStream.Close();
                        ds = NskdInterop.WordTablesReader.FromDocFile(fileName);
                    }
                    catch (Exception e) { ds = null; throw new Exception("Ошибка при разборе потока doc.", e); }
                    finally { File.Delete(fileName); }
                }
            }
            return ds;
        }
    }

    static class XlConvert
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
                    Object seconds = 0;// ToSqlFloat(v);
                    if (seconds != DBNull.Value)
                    {
                        DateTime baseDate = new DateTime(2000, 1, 1);
                        result = baseDate.AddSeconds((Double)seconds);
                    }
                }
            }
            return result;
        }
        public static Object ToXlDatetime(Object value)
        {
            Object result = null;
            String v = value as String;
            if (!String.IsNullOrWhiteSpace(v))
            {
                v = v.Replace("\uFFFD", "");
                DateTime dt;
                if (DateTime.TryParse(v, out dt))
                {
                    result = dt;
                };
            }
            return result;
        }
        public static Object ToXlFloat(Object value)
        {
            Object result = null;
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
            return result;
        }
        public static Object ToXlInt(Object value)
        {
            Object result = null;
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

    class XlTimer
    {
        private static Timer timer;
        private int invokeCount;
        private int maxCount;

        public XlTimer()
        {
            timer = new Timer(CheckStatus, null, Timeout.Infinite, Timeout.Infinite);
            invokeCount = 0;
            maxCount = 0;
            Log.Write("Creating timer.");
        }
        public void Reset(int period, int count)
        {
            timer.Change(period, period);
            invokeCount = 0;
            maxCount = count;
            Log.Write("Reseting timer.");
        }
        public void Start(int period, int count)
        {
            timer.Change(period, period);
            invokeCount = 0;
            maxCount = count;
            Log.Write("Starting timer.");
        }
        public void CheckStatus(Object stateInfo)
        {
            invokeCount++;
            Log.Write("Checking status " + invokeCount.ToString() + ".");
            if (invokeCount >= maxCount)
            {
                // Останавливаем таймер и выгружаем Excel.
                timer.Change(Timeout.Infinite, Timeout.Infinite);
                invokeCount = 0;
                maxCount = 0;
                XlConnection.Close();
            }
        }
    }

}
