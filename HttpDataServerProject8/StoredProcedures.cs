using Nskd;
using System;
using System.Data;
using System.Data.SqlClient;
using System.Globalization;
using System.IO;
using System.Net;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading;

namespace HttpDataServerProject8
{
    class StoredProcedures
    {
        public static void WriteRequestPackageToConsole(RequestPackage rqp)
        {
            String msg = String.Format("{0:yyyy-MM-dd hh:mm:ss}\n", DateTime.Now);
            msg += String.Format("SessionId: '{0}'\n", rqp.SessionId);
            msg += String.Format("Command: '{0}'\n", rqp.Command);
            if (rqp.Parameters != null)
            {
                foreach (RequestParameter p in rqp.Parameters)
                {
                    msg += String.Format("Parameter: name = '{0}', value = '{1}'\n", p.Name, p.Value);
                }
            }
            Log.Write(String.Format(msg));
        }
    }
    public class Prep
    {
        private static String zakUrl = "http://zakupki.gov.ru/";
        private static String zak44CUrl = "http://zakupki.gov.ru/epz/order/notice/ea44/view/common-info.html";
        private static String zak223CUrl = "http://zakupki.gov.ru/223/purchase/public/purchase/info/common-info.html";
        private static String zak44DUrl = "http://zakupki.gov.ru/epz/order/notice/ea44/view/documents.html";
        private static String zak44FilestoreUrl = "http://zakupki.gov.ru/44fz/filestore";
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
                Log.Write(String.Format("HttpDataServerProject8.StoredProcedures.cutName(): rsp is null.\n"));
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
                    if (!String.IsNullOrWhiteSpace(cName))
                    {
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
            }
            return firstCustomerName;
        }
        private static void CreateRefFile(DirectoryInfo di, String auctionNumber)
        {
            using (StreamWriter writer = File.CreateText(di.FullName + @"\Ссылка на аукцион.url"))
            {
                writer.WriteLine("[InternetShortcut]");
                switch (auctionNumber.Length)
                {
                    case 11:
                        writer.WriteLine("URL=" + zak223CUrl + "?regNumber=" + auctionNumber);
                        break;
                    case 19:
                        writer.WriteLine("URL=" + zak44CUrl + "?regNumber=" + auctionNumber);
                        break;
                    default:
                        writer.WriteLine("URL=" + zakUrl);
                        break;
                }

                writer.Flush();
            }
        }
        private static String GetFileNameFromContentDispositionHttpHeader(String contentDispositionValue)
        {
            String fileName = null;

            //Log.Write(String.Format(contentDispositionValue));
            String temp = DecodeContentDispositionHttpHeader(contentDispositionValue);
            // attachment; filename="Извещение 741.rar"; filename*=UTF-8''Извещение 741.rar
            //Log.Write(String.Format(temp));

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
        private static void DownloadDocFiles(Guid sessinId, DirectoryInfo di, String auctionNumber)
        {
            if (String.IsNullOrWhiteSpace(auctionNumber) || auctionNumber.Length != 19) { return; }

            String uri = zak44DUrl + "?regNumber=" + auctionNumber;

            //Log.Write(String.Format(uri));

            HttpWebRequest rq = WebRequest.CreateHttp(uri);
            rq.UseDefaultCredentials = true;
            rq.UserAgent = "Mozilla/5.0"; // сайт не отвечает на автоматические запросы. поэтому притворяемся браузером.
            rq.Timeout = 10000; // 10 sec.

            Thread.Sleep(1000);

            WebResponse rs = null;
            String body = null;
            try
            {
                using (rs = rq.GetResponse())
                {
                    using (StreamReader reader = new StreamReader(rs.GetResponseStream()))
                    {
                        body = reader.ReadToEnd();
                    }
                }
            }
            catch(Exception e) { Log.Write(String.Format("{0}\n{1}", rq.RequestUri, e.Message)); }

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
                    bIndex = body.IndexOf(zak44FilestoreUrl, sIndex);
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
                DirectoryInfo di = Fs.AddDirectory(alias, path);
                if (di.Exists)
                {
                    String shortcutLocation = Path.Combine(di.FullName, dirName + ".lnk");
                    IWshRuntimeLibrary.WshShell shell = new IWshRuntimeLibrary.WshShell();
                    IWshRuntimeLibrary.IWshShortcut shortcut = (IWshRuntimeLibrary.IWshShortcut)shell.CreateShortcut(shortcutLocation);
                    shortcut.TargetPath = targetPath;
                    shortcut.Save();
                }
            }
            catch (Exception e) { Log.Write(String.Format("createPersonalRef: " + e.Message)); }
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
                        String distrName = (auctionNumber.Length == 19) ? (dt0.Rows[0]["dn"] as String) ?? "x" : "x";
                        String dAlias = "contracts";
                        String dPath = String.Format(
                            @"\Заявки\{0:d4}\{1:d2}\{2:d2}.{1:d2}-{4:d2}.{3:d2} {5} {6} {7}", 
                            sdYear, sdMonth, sdDate, 
                            edMonth, edDate, 
                            auctionNumber, userName, distrName);

                        DirectoryInfo dDi = Fs.AddDirectory(dAlias, dPath);
                        if (dDi.Exists)
                        {
                            Db.SessionLogWriteLine(sessionId, "HttpDataServerProject8.StoredProcedures.addContractDirectory()", String.Format("Создан каталог: '{0}'", dDi.FullName));

                            String firstCustomerName = CreateCustomerDirs(dDi, ds.Tables[1]);

                            CreateRefFile(dDi, auctionNumber);

                            DownloadDocFiles(sessionId, dDi, auctionNumber);

                            dAlias = "persons";
                            String custName = firstCustomerName.Substring(0, Math.Min(firstCustomerName.Length, 100));
                            String dirName = String.Format(@"{0:d2}.{1:d2}-{2:d2}.{3:d2} {4} {5}", sdDate, sdMonth, edDate, edMonth, auctionNumber, custName);
                            switch (userName)
                            {
                                case "Перевалова":
                                    dPath = String.Format(@"\Перевалова Юлия Викторовна\АУКЦИОНЫ\Заявки\{0:d4}\{1:d2}", sdYear, sdMonth);
                                    CreatePersonalRef(dDi.FullName, dirName, dAlias, dPath);
                                    break;
                                case "Коледова":
                                    dPath = String.Format(@"\Коледова Юлия Ивановна\АУКЦИОНЫ\Заявки\{0:d4}\{1:d2}", sdYear, sdMonth);
                                    CreatePersonalRef(dDi.FullName, dirName, dAlias, dPath);
                                    break;
                                case "Сущева":
                                    dPath = String.Format(@"\Сущева Ольга Николаевна\АУКЦИОНЫ\Заявки\{0:d4}\{1:d2}", sdYear, sdMonth);
                                    CreatePersonalRef(dDi.FullName, dirName, dAlias, dPath);
                                    break;
                                case "Шанина":
                                    dPath = String.Format(@"\Шанина Елена Николаевна\АУКЦИОНЫ\Заявки\{0:d4}\{1:d2}", sdYear, sdMonth);
                                    CreatePersonalRef(dDi.FullName, dirName, dAlias, dPath);
                                    break;
                                case "Морева":
                                    dPath = String.Format(@"\Морева Марина Олеговна\АУКЦИОНЫ\Заявки\{0:d4}\{1:d2}", sdYear, sdMonth);
                                    CreatePersonalRef(dDi.FullName, dirName, dAlias, dPath);
                                    break;
                                case "Углова":
                                    dPath = String.Format(@"\Углова Алёна Александровна\АУКЦИОНЫ\Заявки\{0:d4}\{1:d2}", sdYear, sdMonth);
                                    CreatePersonalRef(dDi.FullName, dirName, dAlias, dPath);
                                    break;
                                case "Панафидина":
                                    dPath = String.Format(@"\Панафидина Екатерина Николаевна\АУКЦИОНЫ\Заявки\{0:d4}\{1:d2}", sdYear, sdMonth);
                                    CreatePersonalRef(dDi.FullName, dirName, dAlias, dPath);
                                    break;
                                case "Соколов":
                                    dPath = String.Format(@"\Соколов Евгений Анатольевич\АУКЦИОНЫ\Заявки\{0:d4}\{1:d2}", sdYear, sdMonth);
                                    CreatePersonalRef(dDi.FullName, dirName, dAlias, dPath);
                                    break;
                                case "Магергут":
                                case "Скворцова":
                                case "Лобанова":
                                    dAlias = "mag_dep";
                                    dPath = String.Format(@"\новые аукционы");
                                    try
                                    {
                                        String status;
                                        // пробуем взять папку в Контрактах
                                        String sAalias = "contracts";
                                        String sPath = String.Format(@"\Заявки\{0:d4}\{1:d2}", sdYear, sdMonth);
                                        DirectoryInfo sDi = Fs.GetDirectoryInfo(sAalias, sPath, auctionNumber);
                                        if (sDi != null)
                                        {
                                            // копируем в отдел Магергут
                                            status = Fs.CopyDirectory(sDi, dAlias, dPath);
                                        }
                                        else
                                        {
                                            status = String.Format("Ошибка копирования: Папка с аукционом № {0} не найдена.", auctionNumber);
                                        }
                                        Log.Write(String.Format(status));
                                    }
                                    catch (Exception e) { Log.Write(String.Format(e.Message)); }
                                    break;
                                case "Завалова":
                                case "Борисова":
                                case "Подрезова":
                                case "Пирожкова":
                                    dAlias = "zav_dep";
                                    dPath = String.Format(@"\новые аукционы");
                                    try
                                    {
                                        String status;
                                        // пробуем взять папку в Контрактах
                                        String sAalias = "contracts";
                                        String sPath = String.Format(@"\Заявки\{0:d4}\{1:d2}", sdYear, sdMonth);
                                        DirectoryInfo sDi = Fs.GetDirectoryInfo(sAalias, sPath, auctionNumber);
                                        if (sDi != null)
                                        {
                                            // копируем в отдел Заваловой
                                            status = Fs.CopyDirectory(sDi, dAlias, dPath);
                                        }
                                        else
                                        {
                                            status = String.Format("Ошибка копирования: Папка с аукционом № {0} не найдена.", auctionNumber);
                                        }
                                        Log.Write(String.Format(status));
                                    }
                                    catch (Exception e) { Log.Write(String.Format(e.Message)); }
                                    break;
                                default:
                                    break;
                            }
                        }
                    }
                }
            }
        }
        public static ResponsePackage PassToTender(RequestPackage rqp)
        {
            ResponsePackage rsp = new ResponsePackage();
            rsp.Status = "HttpDataServerProject8.StoredProcedures.passToTender()";
            if (rqp != null)
            {
                Guid sessionId = rqp.SessionId;
                String auctionNumber = rqp["auction_number"] as String;
                if (!String.IsNullOrWhiteSpace(auctionNumber))
                {
                    DataSet ds = Db.Prep.GetDirInf(sessionId, auctionNumber);
                    if (ds != null && ds.Tables.Count > 1)
                    {
                        DataTable dt0 = ds.Tables[0];
                        if (dt0 != null && dt0.Rows.Count > 0 && dt0.Columns.Count > 0)
                        {
                            //String auctionNumber = (dt0.Rows[0]["an"] as String) ?? "x";
                            String sd = (dt0.Rows[0]["sd"] as String) ?? "x";
                            String sdDate = "x"; if (sd.Length >= 2) sdDate = sd.Substring(0, 2);
                            String sdMonth = "x"; if (sd.Length >= 5) sdMonth = sd.Substring(3, 2);
                            String sdYear = "x"; if (sd.Length >= 10) sdYear = sd.Substring(6, 4);
                            String[] monthNames = new String[] {
                                "Январь", "Февраль", "Март", "Апрель", "Май", "Июнь", "Июль", "Август", "Сентябрь", "Октябрь", "Ноябрь", "Декабрь" };
                            String mm = sdMonth;
                            Int32 m;
                            if (mm.Length == 2 && mm[0] == '0') { mm = mm.Substring(1); }
                            if (Int32.TryParse(mm, out m))
                            {
                                String dAlias = "prep_f0"; // H:\Документы
                                String dPath = String.Format(@"\Тендерный\{0:d4}\{1}", sdYear, monthNames[m - 1]);
                                // проверка существования папки в тендерном отделе
                                if (Fs.Exists(dAlias, dPath, auctionNumber))
                                {
                                    rsp.Status = String.Format("Ошибка копирования: Папка с аукционом № {0} уже есть в Тендерном отделе.", auctionNumber);
                                }
                                else
                                {
                                    // пробуем взять папку в Контрактах
                                    String sAalias = "contracts"; // @"\\SHD\Kontract";
                                    String sPath = String.Format(@"\Заявки\{0:d4}\{1:d2}", sdYear, sdMonth);
                                    DirectoryInfo sDi = Fs.GetDirectoryInfo(sAalias, sPath, auctionNumber);
                                    if (sDi != null)
                                    {
                                        try
                                        {
                                            // копируем в Тендерный
                                            rsp.Status = Fs.CopyDirectory(sDi, dAlias, dPath);
                                        }
                                        catch(Exception e)
                                        {
                                            Log.Write(String.Format(e.Message));
                                            rsp.Status = String.Format("Ошибка копирования: Папка с аукционом № {0}: {1}", auctionNumber, e.Message);
                                        }
                                    }
                                    else
                                    {
                                        rsp.Status = String.Format("Ошибка копирования: Папка с аукционом № {0} не найдена.", auctionNumber);
                                    }
                                }
                            }
                        }
                    }
                }
            }
            return rsp;
        }
    }
    public class Fs
    {
        private static String docsDirectory = @"\\SRV-TS2\work";        // H:\Документы
        private static String personsDirectory = @"\\TS\sotrudniki";    // H:\Сотрудники (Exchange)
        private static String ctDirectory = @"\\SHD\st_sertificat";     // H:\Сертификаты
        private static String rdDirectory = @"\\SHD\reg_doc";           // H:\Рег. документы
        private static String contractsDirectory = @"\\SHD\Kontract";   // H:\Контракты
        private static String magDirectory = @"\\SHD\magergut";         // Отдел Магергут
        private static String zavDirectory = @"\\SHD\Zavalova";         // Отдел Заваловой
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
        private static String BasePathCombine(String dir = null, String path = null, String name = null)
        {
            String dest = PathCombine(dir, path, name);
            return dest;
        }
        private static String GetAbsolutePath(String alias, String path)
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
                    absolutePath = BasePathCombine(baseDir, path);
                }
                catch { }
            }
            return absolutePath;
        }
        public static DirectoryInfo AddDirectory(String alias, String path)
        {
            DirectoryInfo di = null;
            // полный путь
            String absolutePath = GetAbsolutePath(alias, path);

            //Log.Write(String.Format(absolutePath));

            if (!String.IsNullOrWhiteSpace(absolutePath))
            {
                try
                {
                    di = Directory.CreateDirectory(absolutePath);
                }
                catch (Exception e) { Log.Write(String.Format(e.Message)); }
            }
            return di;
        }
        public static Boolean Exists(String alias, String path, String auctionNumber)
        {
            Boolean r = false;
            String absolutePath = GetAbsolutePath(alias, path);
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
            String absolutePath = GetAbsolutePath(alias, path);
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
        public static String CopyDirectory(DirectoryInfo sDi, String alias, String path)
        {
            String r = String.Format("{0}, {1}, {2}", alias, path, sDi.FullName);
            String absolutePath = GetAbsolutePath(alias, path);
            DirectoryInfo dDi = Directory.CreateDirectory(absolutePath + @"\" + sDi.Name);

            CopyDirectory(sDi, dDi, true);

            return r;
        }
        private static void CopyDirectory(DirectoryInfo sDi, DirectoryInfo dDi, bool copySubDirs)
        {
            if (!sDi.Exists)
            {
                throw new DirectoryNotFoundException("Source directory does not exist or could not be found.");
            }

            FileInfo[] files = sDi.GetFiles();
            foreach (FileInfo file in files)
            {
                String temppath = Path.Combine(dDi.FullName, file.Name);
                file.CopyTo(temppath, true);
            }

            if (copySubDirs)
            {
                DirectoryInfo[] subDirs = sDi.GetDirectories();
                foreach (DirectoryInfo subDir in subDirs)
                {
                    String tempPath = Path.Combine(dDi.FullName, subDir.Name);
                    DirectoryInfo tempDi = Directory.CreateDirectory(tempPath);
                    CopyDirectory(subDir, tempDi, copySubDirs);
                }
            }
        }
    }
    public class Db
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
                catch (Exception e) { Log.Write(String.Format(e.Message)); }
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
                catch (Exception e) { Log.Write(String.Format(e.Message)); }
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
            catch (Exception e) { Log.Write(String.Format(e.Message)); }
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
}
