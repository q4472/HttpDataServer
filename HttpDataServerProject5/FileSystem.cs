using Nskd;
using System;
using System.Collections;
using System.Data;
using System.Data.SqlClient;
using System.Globalization;
using System.IO;
using System.Net;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading;

namespace HttpDataServerProject5
{
    public static class Fs
    {
        private static String docsDirectory = @"\\SRV-TS2\work\Тендерный\2015\Декабрь";
        private static String ctDirectory = @"\\SHD\st_sertificat";
        private static String rdDirectory = @"\\SHD\reg_doc";
        private static String contractsDirectoty = @"\\SHD\Kontract";
        private static String PathCombine(String home, String path = null, String name = null)
        {
            String dest = home;
            if (!String.IsNullOrEmpty(path))
            {
                String p = path.Replace('/', '\\');
                if (p[0] == '\\') p = p.Substring(1);
                if (p.Length > 0)
                {
                    if (p[p.Length - 1] == '\\') p = p.Substring(0, p.Length - 1);
                    if (p.Length > 0)
                    {
                        dest += "\\" + p;
                    }
                }
            }
            if (!String.IsNullOrEmpty(name))
            {
                dest += "\\" + name;
            }
            return dest;
        }
        public static String BasePathCombine(String dir = null, String path = null, String name = null)
        {
            String dest = PathCombine(dir, path, name);
            return dest;
        }
        public static String GetAbsolutePath(String alias, String path)
        {
            String absolutePath = null;
            String baseDir = null;
            switch (alias ?? "")
            {
                case "prep_f0":
                    baseDir = docsDirectory;
                    break;
                case "docs_ct":
                    baseDir = ctDirectory;
                    break;
                case "docs_rd":
                    baseDir = rdDirectory;
                    break;
                case "contracts":
                    baseDir = contractsDirectoty;
                    break;
                default:
                    break;
            }
            if (baseDir != null)
            {
                path = path ?? "";
                try
                {
                    absolutePath = BasePathCombine(baseDir, path);
                }
                catch { }
            }
            return absolutePath;
        }
        public static String GetRelativePath(String alias, String absolutePath)
        {
            String path = "";
            String baseDir = null;
            switch (alias ?? "")
            {
                case "prep_f0":
                    baseDir = docsDirectory;
                    break;
                case "docs_ct":
                    baseDir = ctDirectory;
                    break;
                case "docs_rd":
                    baseDir = rdDirectory;
                    break;
                case "contracts":
                    baseDir = contractsDirectoty;
                    break;
                default:
                    break;
            }
            if (baseDir != null)
            {
                path = absolutePath.Replace(baseDir, "");
            }
            return path;
        }


        private static String zakUrlC = "http://zakupki.gov.ru/epz/order/notice/ea44/view/common-info.html";
        private static String zakUrlD = "http://zakupki.gov.ru/epz/order/notice/ea44/view/documents.html";
        private static String zak44fzFilestore = "http://zakupki.gov.ru/44fz/filestore";
        private static String CutName(String name)
        {
            String result = name;
            RequestPackage rqp = new RequestPackage();
            rqp.SessionId = new Guid();
            rqp.Command = "CutName";
            rqp.Parameters = new RequestParameter[] {
                new RequestParameter {Name = "Name", Value = name}
            };
            ResponsePackage rsp = rqp.GetResponse("http://192.168.135.14:11009");
            if (rsp == null)
            {
                Console.WriteLine("HttpDataServerProject8.StoredProcedures.cutName(): rsp is null.\n");
            }
            else
            {
                result = rsp.Status;
            }
            return result;
        }
        private static String CreateCustomerDirs(DirectoryInfo di, DataTable dt)
        {
            String firstCustomerName = "x";
            if (dt != null && dt.Rows.Count > 0 && dt.Columns.Count > 0)
            {
                foreach (DataRow dr in dt.Rows)
                {
                    String cName = dr[0] as String;
                    cName = (new Regex(@"\W")).Replace(cName, " ");
                    while (cName.Contains("  ")) cName = cName.Replace("  ", " ");
                    cName = cName.Trim();
                    cName = CutName(cName);
                    cName = cName.Substring(0, Math.Min(cName.Length, 128));
                    if (!String.IsNullOrWhiteSpace(cName))
                    {
                        if (firstCustomerName == "x") firstCustomerName = cName;
                        di.CreateSubdirectory(cName);
                    }
                }
            }
            return firstCustomerName;
        }
        private static void CreateRefFile(DirectoryInfo di, String auctionNumber)
        {
            using (StreamWriter writer = File.CreateText(di.FullName + @"\Ссылка на аукцион.url"))
            {
                writer.WriteLine("[InternetShortcut]");
                writer.WriteLine("URL=" + zakUrlC + "?regNumber=" + auctionNumber);
                writer.Flush();
            }
        }
        private static String DecodeContentDispositionHttpHeader(String codedString)
        {
            String decodedString = String.Empty;

            if (!String.IsNullOrWhiteSpace(codedString))
            {
                StringBuilder sb = new StringBuilder();
                Int32 cIndex = 0;
                Boolean cont = true;
                Byte[] buff = new Byte[(codedString.Length / 3) + 1];
                while (cont && cIndex < codedString.Length)
                {
                    Char c = codedString[cIndex++];
                    if (c != '%')
                    {
                        sb.Append(c);
                    }
                    else
                    {
                        // начинаем собирать массив байт
                        Int32 bIndex = 0;
                        while (c == '%')
                        {
                            if (cIndex + 1 < codedString.Length)
                            {
                                Byte b = 0;
                                Byte.TryParse(codedString.Substring(cIndex, 2), NumberStyles.HexNumber, null as IFormatProvider, out b);
                                buff[bIndex++] = b;
                            }
                            cIndex += 2;
                            if (cIndex < codedString.Length)
                            {
                                c = codedString[cIndex++];
                            }
                            else { cont = false; break; }
                        }
                        sb.Append(Encoding.UTF8.GetString(buff, 0, bIndex));
                        --cIndex;
                    }
                }
                decodedString = sb.ToString();
            }
            return decodedString;
        }
        private static String GetFileNameFromContentDispositionHttpHeader(String contentDispositionValue)
        {
            String fileName = null;

            //Console.WriteLine(contentDispositionValue);
            String temp = DecodeContentDispositionHttpHeader(contentDispositionValue);
            // attachment; filename="Извещение 741.rar"; filename*=UTF-8''Извещение 741.rar
            //Console.WriteLine(temp);

            Int32 q1i = temp.IndexOf("filename=\"", 0);
            if (q1i >= 0 && q1i + 10 < temp.Length)
            {
                Int32 q2i = temp.IndexOf('"', q1i + 10);
                if (q2i > q1i)
                {
                    fileName = temp.Substring(q1i + 10, q2i - (q1i + 10));
                }
            }
            return fileName;
        }
        private static void DownloadDocFiles(Guid sessinId, DirectoryInfo di, String auctionNumber)
        {
            String uri = zakUrlD + "?regNumber=" + auctionNumber;

            //Console.WriteLine(uri);

            HttpWebRequest rq = WebRequest.CreateHttp(uri);
            rq.UseDefaultCredentials = true;
            rq.UserAgent = "Mozilla/5.0"; // сайт не отвечает на автоматические запросы. поэтому притворяемся браузером.
            rq.Timeout = 10000; // 10 sec.

            Thread.Sleep(1000);

            WebResponse rs = null;
            String body = null;
            using (rs = rq.GetResponse())
            {
                using (StreamReader reader = new StreamReader(rs.GetResponseStream()))
                {
                    body = reader.ReadToEnd();
                }
            }

            if (!String.IsNullOrEmpty(body))
            {
                // Пропустить неактивные ссылки
                Int32 dr = body.IndexOf("Действующая редакция");
                Int32 sIndex = (dr >= 0) ? dr : 0;
                Int32 bIndex = 0; // индекс начала ссылки
                Int32 eIndex = 0; // индекс окончания ссылки
                Int32 cnt = 0;
                while (sIndex < body.Length)
                {
                    // индекс начала ссылки
                    bIndex = body.IndexOf(zak44fzFilestore, sIndex);
                    if (bIndex >= 0)
                    {
                        // индекс окончания ссылки (")
                        eIndex = body.IndexOf("\"", bIndex);
                        if (eIndex > bIndex)
                        {
                            // сама ссылка
                            uri = body.Substring(bIndex, eIndex - bIndex);

                            Db.SessionLogWriteLine(sessinId,
                                "HttpDataServerProject8.StoredProcedures.downloadDocFiles()",
                                String.Format("Начало загрузки документа по ссылке: {0} - '{1}'", cnt, uri));

                            rq = WebRequest.CreateHttp(uri);
                            rq.UseDefaultCredentials = true;
                            rq.UserAgent = "Mozilla/5.0";
                            rq.Timeout = 20000;

                            Thread.Sleep(1000);

                            try
                            {
                                using (rs = rq.GetResponse())
                                {
                                    String contentDispositionHttpHeader = rs.Headers["Content-Disposition"];
                                    String fileName = GetFileNameFromContentDispositionHttpHeader(contentDispositionHttpHeader);

                                    String path = di.FullName + @"\" + fileName;
                                    if (!File.Exists(path))
                                    {
                                        using (FileStream fs = File.Create(path))
                                        {
                                            rs.GetResponseStream().CopyTo(fs);
                                        }
                                    }
                                }
                            }
                            catch (Exception e)
                            {
                                Db.SessionLogWriteLine(sessinId,
                                    "HttpDataServerProject8.StoredProcedures.downloadDocFiles()",
                                    String.Format("Возникло исключение: {0} в {1}", e.Message, e.Source));
                            }

                            Db.SessionLogWriteLine(sessinId,
                                "HttpDataServerProject8.StoredProcedures.downloadDocFiles()",
                                String.Format("Конец загрузки документа по ссылке: {0} - '{1}'", cnt, uri));

                            cnt++;

                            sIndex = eIndex + 1;
                            Thread.Sleep(1000);
                            continue;
                        }
                        else break;
                    }
                    else break;
                }
            }
        }
        private static void CreatePersonalRef(String targetPath, String dirName, String alias, String path)
        {
            try
            {
                DirectoryInfo di = FsBase.AddDirectory(alias, path);
                if (di.Exists)
                {
                    String shortcutLocation = Path.Combine(di.FullName, dirName + ".lnk");
                    IWshRuntimeLibrary.WshShell shell = new IWshRuntimeLibrary.WshShell();
                    IWshRuntimeLibrary.IWshShortcut shortcut = (IWshRuntimeLibrary.IWshShortcut)shell.CreateShortcut(shortcutLocation);
                    shortcut.TargetPath = targetPath;
                    shortcut.Save();
                }
            }
            catch (Exception e) { Console.WriteLine("createPersonalRef: " + e.Message); }
        }
        public static void AddContractDirectory(RequestPackage rqp)
        {
            Guid sessionId = rqp.SessionId;
            Object temp = rqp["auction_uid"];
            if (temp != null && temp.GetType() == typeof(Guid))
            {
                Guid auctionUid = (Guid)temp;
                DataSet ds = Db.Prep.GetDirInf(sessionId, auctionUid);
                if (ds != null && ds.Tables.Count > 1)
                {
                    DataTable dt0 = ds.Tables[0];
                    if (dt0 != null && dt0.Rows.Count > 0 && dt0.Columns.Count > 0)
                    {
                        String auctionNumber = (dt0.Rows[0]["an"] as String) ?? "x";
                        String sd = (dt0.Rows[0]["sd"] as String) ?? "x";
                        String sdDate = "x"; if (sd.Length >= 2) sdDate = sd.Substring(0, 2);
                        String sdMonth = "x"; if (sd.Length >= 5) sdMonth = sd.Substring(3, 2);
                        String sdYear = "x"; if (sd.Length >= 10) sdYear = sd.Substring(6, 4);
                        String ed = (dt0.Rows[0]["ed"] as String) ?? "x";
                        String edDate = "x"; if (ed.Length > 2) edDate = ed.Substring(0, 2);
                        String edMonth = "x"; if (ed.Length >= 5) edMonth = ed.Substring(3, 2);
                        String userName = (dt0.Rows[0]["un"] as String) ?? "x";
                        String distrName = (dt0.Rows[0]["dn"] as String) ?? "x";
                        String alias = "contracts";
                        String path = String.Format(@"\Заявки\{0:d4}\{1:d2}\{2:d2}.{1:d2}-{4:d2}.{3:d2} {5} {6} {7}", sdYear, sdMonth, sdDate, edMonth, edDate, auctionNumber, userName, distrName);

                        DirectoryInfo di = FsBase.AddDirectory(alias, path);
                        if (di.Exists)
                        {
                            Db.SessionLogWriteLine(sessionId, "HttpDataServerProject8.StoredProcedures.addContractDirectory()", String.Format("Создан каталог: '{0}'", di.FullName));

                            String firstCustomerName = CreateCustomerDirs(di, ds.Tables[1]);

                            CreateRefFile(di, auctionNumber);

                            DownloadDocFiles(sessionId, di, auctionNumber);

                            alias = "persons";
                            String custName = firstCustomerName.Substring(0, Math.Min(firstCustomerName.Length, 100));
                            String dirName = String.Format(@"{0:d2}.{1:d2}-{2:d2}.{3:d2} {4} {5}", sdDate, sdMonth, edDate, edMonth, auctionNumber, custName);
                            switch (userName)
                            {
                                case "Коледова":
                                    path = String.Format(@"\Коледова Юлия Ивановна\АУКЦИОНЫ\Заявки\{0:d4}\{1:d2}", sdYear, sdMonth);
                                    CreatePersonalRef(di.FullName, dirName, alias, path);
                                    break;
                                case "Сущева":
                                    path = String.Format(@"\Сущева Ольга Николаевна\АУКЦИОНЫ\Заявки\{0:d4}\{1:d2}", sdYear, sdMonth);
                                    CreatePersonalRef(di.FullName, dirName, alias, path);
                                    break;
                                case "Шанина":
                                    path = String.Format(@"\Шанина Елена Николаевна\АУКЦИОНЫ\Заявки\{0:d4}\{1:d2}", sdYear, sdMonth);
                                    CreatePersonalRef(di.FullName, dirName, alias, path);
                                    break;
                                case "Углова":
                                    path = String.Format(@"\Углова Алёна Александровна\АУКЦИОНЫ\Заявки\{0:d4}\{1:d2}", sdYear, sdMonth);
                                    CreatePersonalRef(di.FullName, dirName, alias, path);
                                    break;
                                case "Соколов":
                                    path = String.Format(@"\Соколов Евгений Анатольевич\АУКЦИОНЫ\Заявки\{0:d4}\{1:d2}", sdYear, sdMonth);
                                    CreatePersonalRef(di.FullName, dirName, alias, path);
                                    break;
                                case "Магергут":
                                case "Скворцова":
                                case "Лобанова":
                                    alias = "mag_dep";
                                    path = String.Format(@"\новые аукционы");
                                    try
                                    {
                                        String status;
                                        // пробуем взять папку в Контрактах
                                        String sAalias = "contracts";
                                        String sPath = String.Format(@"\Заявки\{0:d4}\{1:d2}", sdYear, sdMonth);
                                        DirectoryInfo di1 = FsBase.GetDirectoryInfo(sAalias, sPath, auctionNumber);
                                        if (di1 != null)
                                        {
                                            // копируем в отдел Магергут
                                            status = FsBase.CopyFiles(alias, path, di1);
                                        }
                                        else
                                        {
                                            status = String.Format("Ошибка копирования: Папка с аукционом № {0} не найдена.", auctionNumber);
                                        }
                                        Console.WriteLine(status);
                                    }
                                    catch (Exception e) { Console.WriteLine(e.Message); }
                                    break;
                                case "Завалова":
                                case "Борисова":
                                    alias = "zav_dep";
                                    path = String.Format(@"\новые аукционы");
                                    try
                                    {
                                        String status;
                                        // пробуем взять папку в Контрактах
                                        String sAalias = "contracts";
                                        String sPath = String.Format(@"\Заявки\{0:d4}\{1:d2}", sdYear, sdMonth);
                                        DirectoryInfo di1 = FsBase.GetDirectoryInfo(sAalias, sPath, auctionNumber);
                                        if (di1 != null)
                                        {
                                            // копируем в отдел Заваловой
                                            status = FsBase.CopyFiles(alias, path, di1);
                                        }
                                        else
                                        {
                                            status = String.Format("Ошибка копирования: Папка с аукционом № {0} не найдена.", auctionNumber);
                                        }
                                        Console.WriteLine(status);
                                    }
                                    catch (Exception e) { Console.WriteLine(e.Message); }
                                    break;
                                default:
                                    break;
                            }
                        }
                    }
                }
            }
        }
    }
    class CommandSwitcher
    {
        public static ResponsePackage Exec(RequestPackage rqp)
        {
            ResponsePackage rsp = null;
            if (rqp != null)
            {
                switch (rqp.Command)
                {
                    case "GetDirectoryInfo":
                        rsp = FsStoredProcedure.GetDirectoryInfo(rqp);
                        break;
                    case "GetFileContents":
                        rsp = FsStoredProcedure.GetFileContents(rqp);
                        break;
                    case "AddDirectory":
                        rsp = FsStoredProcedure.AddDirectory(rqp);
                        break;
                    default:
                        break;
                }
            }
            return rsp;
        }
    }
    class FsStoredProcedure
    {
        private static Byte[] readFile(String path)
        {
            Byte[] buff = null;
            try
            {
                String root = Path.GetPathRoot(path);
                String dir = Path.GetDirectoryName(path).Replace(root, "");
                String file = Path.GetFileName(path);

                DirectoryInfo diRoot = new DirectoryInfo(root);
                DirectoryInfo diDir;
                if (String.IsNullOrWhiteSpace(dir) || ((dir.Length == 1) && (dir[0] == '\\')))
                {
                    diDir = diRoot;
                }
                else
                {
                    if (dir[0] == '\\')
                    {
                        dir = dir.Substring(1);
                    }
                    diDir = diRoot.GetDirectories(dir)[0];
                }
                FileInfo[] fis = diDir.GetFiles(file);
                if (fis.Length > 0)
                {
                    FileInfo fi = fis[0];
                    using (FileStream fs = fi.OpenRead())
                    {
                        Int32 len = (Int32)fs.Length;
                        buff = new Byte[len];
                        fs.Read(buff, 0, len);
                    }
                }
            }
            catch (Exception e) { Log.Write(e.ToString()); }
            return buff;
        }
        public static ResponsePackage GetDirectoryInfo(RequestPackage rqp)
        {
            ResponsePackage rsp = new ResponsePackage();
            rsp.Status = "Trying to get a directory information.";

            // Пока нигде не используется.
            // Это задел на будущее.
            Guid sessionId = rqp.SessionId;

            // псевдоним для базового каталога
            String alias = rqp["alias"] as String;
            
            // путь от базового каталога
            String path = rqp["path"] as String;
            
            // полный путь
            String absolutePath = Fs.GetAbsolutePath(alias, path);
            
            //Console.WriteLine(absolutePath);

            if (!String.IsNullOrWhiteSpace(absolutePath))
            {
                try
                {
                    DirectoryInfo pdi = new DirectoryInfo(absolutePath);
                    if (pdi.Exists)
                    {
                        DataSet ds = new DataSet();
                        rsp.Data = ds;

                        // таблица для каталогов
                        DataTable ddt = new DataTable("Directories");
                        ds.Tables.Add(ddt);
                        ddt.Columns.Add("name", typeof(String));
                        {
                            // первая строка это запись о каталоге который был запрошен (parent)
                            DataRow dr = ddt.NewRow();
                            ddt.Rows.Add(dr);
                            dr[0] = Fs.GetRelativePath(alias, pdi.FullName); //((name == "") ? "\\" : name);
                            //Console.WriteLine("   +'" + (String)dr[0] + "'");
                            // каждая следующая строка это запись о вложенном каталоге (child)
                            foreach (DirectoryInfo cdi in pdi.GetDirectories())
                            {
                                dr = ddt.NewRow();
                                ddt.Rows.Add(dr);
                                dr[0] = Fs.GetRelativePath(alias, cdi.FullName);
                                //Console.WriteLine("   +'" + (String)dr[0] + "'");
                            }
                        }

                        // таблица для вложенных файлов
                        DataTable fdt = new DataTable("Files");
                        ds.Tables.Add(fdt);
                        fdt.Columns.Add("name", typeof(String));
                        {
                            // каждая строка это запись о вложенном файле (child)
                            foreach (FileInfo cfi in pdi.GetFiles())
                            {
                                DataRow dr = fdt.NewRow();
                                fdt.Rows.Add(dr);
                                dr[0] = Fs.GetRelativePath(alias, cfi.FullName);
                                //Console.WriteLine("   -'" + (String)dr[0] + "'");
                            }
                        }
                    }
                }
                catch (Exception e) { rsp.Status = "Error."; Log.Write(e.ToString()); }
            }
            return rsp;
        }
        public static ResponsePackage GetFileContents(RequestPackage rqp)
        {
            ResponsePackage rsp = new ResponsePackage();
            rsp.Status = "Trying to get a file information.";

            // Пока нигде не используется.
            // Это задел на будущее.
            Guid sessionId = rqp.SessionId;

            // псевдоним для базового каталога
            String alias = rqp["alias"] as String;

            // путь от базового каталога
            String path = rqp["path"] as String;
            
            // полный путь
            String absolutePath = Fs.GetAbsolutePath(alias, path);
            
            //Console.WriteLine("absolutePath: '" + absolutePath + "'");

            if (!String.IsNullOrWhiteSpace(absolutePath))
            {
                try
                {
                    DataSet ds = new DataSet();
                    rsp.Data = ds;
                    DataTable dt = new DataTable();
                    ds.Tables.Add(dt);
                    dt.Columns.Add("contents", typeof(String));
                    DataRow dr = dt.NewRow();
                    dt.Rows.Add(dr);

                    Byte[] buff = readFile(absolutePath);
                    if (buff != null)
                    {
                        dr[0] = Convert.ToBase64String(buff);
                    }
                }
                catch (Exception e) { rsp.Status = "Error."; Log.Write(e.ToString()); }
            }
            return rsp;
        }
        public static ResponsePackage AddDirectory(RequestPackage rqp)
        {
            ResponsePackage rsp = new ResponsePackage();
            rsp.Status = "Trying to add a directory.";

            // Пока нигде не используется.
            // Это задел на будущее.
            Guid sessionId = rqp.SessionId;

            // псевдоним для базового каталога
            String alias = rqp["alias"] as String;

            // путь от базового каталога
            String path = rqp["path"] as String;
            
            // полный путь
            String absolutePath = Fs.GetAbsolutePath(alias, path);
            
            //Console.WriteLine(absolutePath);

            if (!String.IsNullOrWhiteSpace(absolutePath))
            {
                try
                {
                    Directory.CreateDirectory(absolutePath);
                }
                catch (Exception e) { rsp.Status = e.Message; }
            }
            return rsp;
        }
    }
    class Db
    {
        private static String cnString = String.Format("Data Source={0};Integrated Security=True", Program.MainSqlServerDataSource);
        public class Prep
        {
            public static DataSet GetDirInf(Guid sessionId, Guid auctionUid)
            {
                DataSet ds = null;
                try
                {
                    SqlCommand cmd = new SqlCommand();
                    cmd.Connection = new SqlConnection(cnString);
                    cmd.CommandType = CommandType.StoredProcedure;
                    cmd.CommandText = "[Pharm-Sib].[dbo].[prep_get_dir_inf]";
                    cmd.Parameters.AddWithValue("session_id", sessionId);
                    cmd.Parameters.AddWithValue("auction_uid", auctionUid);
                    SqlDataAdapter da = new SqlDataAdapter(cmd);
                    ds = new DataSet();
                    da.Fill(ds);
                }
                catch (Exception e) { Console.WriteLine(e.Message); }
                return ds;
            }
            public static DataSet GetDirInf(Guid sessionId, String auctionNumber)
            {
                DataSet ds = null;
                try
                {
                    SqlCommand cmd = new SqlCommand();
                    cmd.Connection = new SqlConnection(cnString);
                    cmd.CommandType = CommandType.StoredProcedure;
                    cmd.CommandText = "[Pharm-Sib].[dbo].[prep_get_dir_inf]";
                    cmd.Parameters.AddWithValue("session_id", sessionId);
                    cmd.Parameters.AddWithValue("auction_number", auctionNumber);
                    SqlDataAdapter da = new SqlDataAdapter(cmd);
                    ds = new DataSet();
                    da.Fill(ds);
                }
                catch (Exception e) { Console.WriteLine(e.Message); }
                return ds;
            }
        }
        public static void SessionLogWriteLine(Guid sessionUid, String src, String msg)
        {
            try
            {
                SqlCommand cmd = new SqlCommand();
                cmd.Connection = new SqlConnection(cnString);
                cmd.CommandType = CommandType.StoredProcedure;
                cmd.CommandText = "[phs_s].[dbo].[session_logs_insert]";
                cmd.Parameters.AddWithValue("session_uid", sessionUid);
                cmd.Parameters.AddWithValue("src", src);
                cmd.Parameters.AddWithValue("message", msg);
                using (cmd.Connection)
                {
                    cmd.Connection.Open();
                    cmd.ExecuteNonQuery();
                }
            }
            catch (Exception e) { Console.WriteLine(e.Message); }
        }
        public static ResponsePackage Exec(RequestPackage rqp)
        {
            ResponsePackage rsp = new ResponsePackage();
            rsp.Status = "Запрос к Sql серверу.";
            rsp.Data = new DataSet();
            SqlCommand cmd = new SqlCommand();
            cmd.CommandType = CommandType.StoredProcedure;
            cmd.CommandText = rqp.Command;
            cmd.CommandTimeout = 300;
            if ((rqp != null) && (rqp.Parameters != null))
            {
                foreach (var p in rqp.Parameters)
                {
                    if (p != null)
                    {
                        String key = p.Name;
                        Object value = p.Value;
                        if (!String.IsNullOrWhiteSpace(key))
                        {
                            cmd.Parameters.AddWithValue("@" + key, value);
                        }
                    }
                }
            }
            try
            {
                using (cmd.Connection = new SqlConnection(cnString))
                {
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
    class FsBase
    {
        private static String docsDirectory = @"\\SRV-TS2\work";        // H:\Документы
        private static String personsDirectory = @"\\TS\sotrudniki";    // H:\Сотрудники (Exchange)
        private static String ctDirectory = @"\\SHD\st_sertificat";     // H:\Сертификаты
        private static String rdDirectory = @"\\SHD\reg_doc";           // H:\Рег. документы
        private static String contractsDirectory = @"\\SHD\Kontract";   // H:\Контракты
        private static String magDirectory = @"\\SHD\magergut";         // Отдел Магергут
        private static String zavDirectory = @"\\SHD\Zavalova";         // Отдел Заваловой
        private static String pathCombine(String home, String path = null, String name = null)
        {
            String dest = home;
            if (!String.IsNullOrEmpty(path))
            {
                String p = path.Replace('/', '\\');
                if (p[0] == '\\') p = p.Substring(1);
                if (p.Length > 0)
                {
                    if (p[p.Length - 1] == '\\') p = p.Substring(0, p.Length - 1);
                    if (p.Length > 0)
                    {
                        dest += "\\" + p;
                    }
                }
            }
            if (!String.IsNullOrEmpty(name))
            {
                dest += "\\" + name;
            }
            return dest;
        }
        private static String basePathCombine(String dir = null, String path = null, String name = null)
        {
            String dest = pathCombine(dir, path, name);
            return dest;
        }
        private static String getAbsolutePath(String alias, String path)
        {
            String absolutePath = null;
            String baseDir = null;
            switch (alias ?? "")
            {
                case "prep_f0":
                    baseDir = docsDirectory;
                    break;
                case "docs_ct":
                    baseDir = ctDirectory;
                    break;
                case "docs_rd":
                    baseDir = rdDirectory;
                    break;
                case "contracts":
                    baseDir = contractsDirectory;
                    break;
                case "persons":
                    baseDir = personsDirectory;
                    break;
                case "mag_dep":
                    baseDir = magDirectory;
                    break;
                case "zav_dep":
                    baseDir = zavDirectory;
                    break;
                default:
                    break;
            }
            if (baseDir != null)
            {
                path = path ?? "";
                try
                {
                    absolutePath = basePathCombine(baseDir, path);
                }
                catch { }
            }
            return absolutePath;
        }
        public static DirectoryInfo AddDirectory(String alias, String path)
        {
            DirectoryInfo di = null;
            // полный путь
            String absolutePath = getAbsolutePath(alias, path);

            //Console.WriteLine(absolutePath);

            if (!String.IsNullOrWhiteSpace(absolutePath))
            {
                try
                {
                    di = Directory.CreateDirectory(absolutePath);
                }
                catch (Exception e) { Console.WriteLine(e.Message); }
            }
            return di;
        }
        public static Boolean Exists(String alias, String path, String auctionNumber)
        {
            Boolean r = false;
            String absolutePath = getAbsolutePath(alias, path);
            DirectoryInfo di = new DirectoryInfo(absolutePath);
            if (di.Exists)
            {
                DirectoryInfo[] dis = di.GetDirectories(String.Format("*{0}*", auctionNumber));
                r = (dis.Length > 0);
            }
            return r;
        }
        public static DirectoryInfo GetDirectoryInfo(String alias, String path, String auctionNumber)
        {
            DirectoryInfo di = null;
            String absolutePath = getAbsolutePath(alias, path);
            DirectoryInfo pdi = new DirectoryInfo(absolutePath);
            if (pdi.Exists)
            {
                DirectoryInfo[] dis = pdi.GetDirectories(String.Format("*{0}*", auctionNumber));
                if (dis.Length > 0)
                {
                    di = dis[0];
                }
            }
            return di;
        }
        public static String CopyFiles(String alias, String path, DirectoryInfo sDi)
        {
            String r = String.Format("{0}, {1}, {2}", alias, path, sDi.FullName);
            String absolutePath = getAbsolutePath(alias, path);
            DirectoryInfo newDir = Directory.CreateDirectory(absolutePath + @"\" + sDi.Name);

            DirectoryCopy(sDi, newDir, true);

            return r;
        }
        private static void DirectoryCopy(DirectoryInfo sDir, DirectoryInfo dDir, bool copySubDirs)
        {
            if (!sDir.Exists)
            {
                throw new DirectoryNotFoundException(
                    "Source directory does not exist or could not be found.");
            }

            FileInfo[] files = sDir.GetFiles();
            foreach (FileInfo file in files)
            {
                String temppath = Path.Combine(dDir.FullName, file.Name);
                file.CopyTo(temppath, true);
            }

            if (copySubDirs)
            {
                DirectoryInfo[] dirs = sDir.GetDirectories();
                foreach (DirectoryInfo subdir in dirs)
                {
                    String tempPath = Path.Combine(dDir.FullName, subdir.Name);
                    DirectoryInfo tempDir = Directory.CreateDirectory(tempPath);
                    DirectoryCopy(subdir, tempDir, copySubDirs);
                }
            }
        }
    }
}
