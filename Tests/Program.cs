using Nskd;
using Nskd.V83;
using System;
using System.ComponentModel;
using System.Data;
using System.Data.SqlClient;
using System.Diagnostics;
using System.Globalization;
using System.IO;
using System.Net;
using System.Net.Mail;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Text;
using System.Text.RegularExpressions;
using System.Xml.Serialization;
//using Word = Microsoft.Office.Interop.Word;

namespace Tests
{
    internal class Native
    {
        public delegate void SignalHandler(uint consoleSignal);

        [DllImport("Kernel32", EntryPoint = "SetConsoleCtrlHandler")]
        public static extern bool SetSignalHandler(SignalHandler handler, bool add);
        /// <summary>
        ///     При закрытии программы удаляет все процессы 1cv8*
        /// </summary>
        /// <param name="consoleSignal">перечисление передаваемое операционной системой</param>
        public static void HandlerRoutine(uint consoleSignal)
        {
            switch (consoleSignal)
            {
                case 0: // CtrlC
                case 1: // CtrlBreak
                case 2: // Close
                case 5: // LogOff
                case 6: // Shutdown
                    Process[] ps = Process.GetProcesses();
                    foreach (Process p in ps)
                    {
                        if ((p.ProcessName.Length >= 4) && (p.ProcessName.Substring(0, 4).ToUpper() == "1CV8"))
                        {
                            p.Kill();
                        }
                    }
                    break;
                default:
                    break;
            }
        }
    }
    class Program
    {
        private static Native.SignalHandler signalHangler = null;
        static bool mailSent = false;
        private static void SendCompletedCallback(object sender, AsyncCompletedEventArgs e)
        {
            // Get the unique identifier for this asynchronous operation.
            String userToken = e.UserState.ToString();

            if (e.Cancelled)
            {
                Console.WriteLine("[{0}] Send canceled.", userToken);
            }
            if (e.Error != null)
            {
                Console.WriteLine("[{0}] {1}", userToken, e.Error.ToString());
            }
            else
            {
                Console.WriteLine("Message sent.");
            }
            mailSent = true;
        }
        static void Main(string[] args)
        {
            // при закрытии программы будет вызвана процедура 'handlerRoutine'
            //signalHangler += Native.HandlerRoutine;
            //Native.SetSignalHandler(signalHangler, true);

            try
            {
                //TestRegex();
                TestLinks();
                //TestInterop();
                //TestXmlSerializer();
                //Test1c();
                //Console.ReadKey();
                //MailServer.Exec();
                //TestHttpDataServer();
                //Test1Cv8UT11();
            }
            catch (Exception e) { Console.WriteLine(e); }
            Console.ReadKey();
        }

        private static Object thisLock = new Object();
        private static void Test1Cv8UT11()
        {
            //DataSet ds = null;
            lock (thisLock)
            {
                // Подключаемся к 1с83.
                String connectString = @"File=""C:\Users\sokolov\Documents\1C\Trade"";Usr=""СоколовЕА"";Pwd=""й"";";
                var globalContext = new GlobalContext(connectString);
                if (globalContext != null)
                {
                    String cnString = String.Format("Data Source={0};Initial Catalog=Pharm-Sib;Integrated Security=True", "192.168.135.77");
                    SqlCommand cmd = new SqlCommand()
                    {
                        CommandText = @"
                                select  
	                                ID, DESCR, v_igs.*
                                from (
	                                select top (100) ID, DESCR 
	                                from phs_oc_mirror.dbo.sc33 
	                                where ISFOLDER <> 1 and ISMARK = 0 
	                                order by ID desc
	                                ) as sc33
                                join Goods.dbo.спр_1с_keys as k3 on k3.sc33_id = sc33.id
                                join Goods.dbo.items on items.src_id = 3 and items.src_uid = k3.uid
                                join Goods.dbo.v_igs on v_igs.uid = items.group_uid
                                where v_igs.наименование is not null
                                order by v_igs.наименование, descr
                            ",
                        CommandType = CommandType.Text,
                        Connection = new SqlConnection(cnString)
                    };
                    SqlDataAdapter da = new SqlDataAdapter(cmd);
                    DataSet ds = new DataSet();
                    da.Fill(ds);
                    DataTable dt = ds.Tables[0];

                    Console.WriteLine(dt.Rows.Count);

                    Console.WriteLine((globalContext == null) ? "null" : globalContext.GetType().ToString());
                    try
                    {
                        Console.WriteLine("Подсоединились.");

                        // эти дополнительные реквизиты должны быть уже добавлены в 1с ПланыВидовХарактеристик.ДополнительныеРеквизитыИСведения
                        var ДопРеквизиты = globalContext.ПланыВидовХарактеристик.ДополнительныеРеквизитыИСведения;
                        var ДопРеквизитНаименование = ДопРеквизиты.НайтиПоНаименованию("Торговое наименование");
                        var ДопРеквизитФорма = ДопРеквизиты.НайтиПоНаименованию("Форма выпуска");
                        var ДопРеквизитДозировка = ДопРеквизиты.НайтиПоНаименованию("Дозировка");
                        var ДопРеквизитУпаковка = ДопРеквизиты.НайтиПоНаименованию("Упаковка");
                        var ДопРеквизитКоличество = ДопРеквизиты.НайтиПоНаименованию("Кол-во в потр. упаковке");
                        var ДопРеквизитПроизводитель = ДопРеквизиты.НайтиПоНаименованию("Производитель");
                        var ДопРеквизитСтрана = ДопРеквизиты.НайтиПоНаименованию("Страна");

                        var Запрос = globalContext.Запрос;
                        Запрос.Текст = @"
                            ВЫБРАТЬ
	                            УпаковкиЕдиницыИзмерения.Ссылка КАК Ссылка
                            ИЗ
	                            Справочник.УпаковкиЕдиницыИзмерения КАК УпаковкиЕдиницыИзмерения
                            ГДЕ
                                НЕ УпаковкиЕдиницыИзмерения.ПометкаУдаления
                                И УпаковкиЕдиницыИзмерения.Наименование = ""шт""
                            ";
                        var РезультатЗапроса = Запрос.Выполнить();
                        var ВыборкаИзРезультатаЗапроса = РезультатЗапроса.Выбрать();
                        if (ВыборкаИзРезультатаЗапроса.Количество() != 1) { throw new Exception("В справочнике 'УпаковкиЕдиницыИзмерения' не найден элемент 'шт'."); }
                        ВыборкаИзРезультатаЗапроса.Следующий();
                        var УпаковкиЕдиницыИзмерения_Шт_Ссылка = ВыборкаИзРезультатаЗапроса.GetProperty("Ссылка");
                        Console.WriteLine("В справочнике 'УпаковкиЕдиницыИзмерения' найден элемент 'шт'.");

                        var Справочники = globalContext.Справочники;

                        var СправочникВидыНоменклатуры = Справочники.ВидыНоменклатуры;
                        var ВидНоменклатурыТовар = СправочникВидыНоменклатуры.НайтиПоНаименованию("Товар", true);
                        while (!ВидНоменклатурыТовар.Пустая() && ВидНоменклатурыТовар.ЭтоГруппа)
                        {
                            Console.WriteLine("В справочнике 'ВидыНоменклатуры' найдена группа 'Товар'.");
                            ВидНоменклатурыТовар = СправочникВидыНоменклатуры.НайтиПоНаименованию("Товар", true, ВидНоменклатурыТовар);
                        }
                        if (ВидНоменклатурыТовар.Пустая()) { throw new Exception("В справочнике 'ВидыНоменклатуры' не найден элемент 'Товар'."); }
                        Console.WriteLine("В справочнике 'ВидыНоменклатуры' найден элемент '{0}'", ВидНоменклатурыТовар.Наименование);

                        
                        var Перечисления = globalContext.Перечисления;
                        /*
                        var ПеречислениеТипыНоменклатуры = Перечисления.ТипыНоменклатуры;
                        var ТипНоменклатурыТовар = new Перечисления.ПеречисленияСсылка(ПеречислениеТипыНоменклатуры.GetProperty("Товар"));
                        if (ТипНоменклатурыТовар.Пустая()) { throw new Exception("В перечислении 'ТипыНоменклатуры' не найдено значение 'Товар'."); }
                        Console.WriteLine("В перечислении 'ТипыНоменклатуры' найдено значение 'Товар'.");

                        var ПеречисленияСтавкиНДС = Перечисления.СтавкиНДС;
                        var СтавкиНДС10 = new Перечисления.ПеречисленияСсылка(ПеречисленияСтавкиНДС.GetProperty("НДС10"));
                        if (СтавкиНДС10.Пустая()) { throw new Exception("В перечислении 'СтавкиНДС' не найдено значение 'НДС10'."); }
                        Console.WriteLine("В перечислении 'СтавкиНДС' найдено значение 'НДС10'.");

                        var ВариантыОформленияПродажи = Перечисления.ВариантыОформленияПродажи;
                        var ВариантыОформленияПродажиРеализацияТоваровУслуг = new Перечисления.ПеречисленияСсылка(ВариантыОформленияПродажи.GetProperty("РеализацияТоваровУслуг"));
                        if (ВариантыОформленияПродажиРеализацияТоваровУслуг.Пустая()) { throw new Exception("В перечислении 'ВариантыОформленияПродажи' не найдено значение 'РеализацияТоваровУслуг'."); }
                        Console.WriteLine("В перечислении 'ВариантыОформленияПродажи' найдено значение 'РеализацияТоваровУслуг'.");

                        var ВариантыИспользованияХарактеристикНоменклатуры = Перечисления.ВариантыИспользованияХарактеристикНоменклатуры;

                        var ВариантыИспользованияХарактеристикНоменклатуры_НеИспользовать = new Перечисления.ПеречисленияСсылка(ВариантыИспользованияХарактеристикНоменклатуры.GetProperty("НеИспользовать"));
                        if (ВариантыИспользованияХарактеристикНоменклатуры_НеИспользовать.Пустая()) { throw new Exception("В перечислении 'ВариантыИспользованияХарактеристикНоменклатуры' не найдено значение 'НеИспользовать'."); }
                        Console.WriteLine("В перечислении 'ВариантыИспользованияХарактеристикНоменклатуры' найдено значение 'НеИспользовать'.");

                        var ВариантыИспользованияХарактеристикНоменклатуры_ОбщиеДляВидаНоменклатуры = new Перечисления.ПеречисленияСсылка(ВариантыИспользованияХарактеристикНоменклатуры.GetProperty("ОбщиеДляВидаНоменклатуры"));
                        if (ВариантыИспользованияХарактеристикНоменклатуры_ОбщиеДляВидаНоменклатуры.Пустая()) { throw new Exception("В перечислении 'ВариантыИспользованияХарактеристикНоменклатуры' не найдено значение 'ОбщиеДляВидаНоменклатуры'."); }
                        Console.WriteLine("В перечислении 'ВариантыИспользованияХарактеристикНоменклатуры' найдено значение 'ОбщиеДляВидаНоменклатуры'.");
                        */
                        var СправочникНоменклатура = Справочники.Номенклатура;

                        // удаляем всё 
                        Console.WriteLine("Удаляем все корневые элементы. Все остальные удалятся автоматически.");

                        var Выборка = СправочникНоменклатура.Выбрать();
                        while (Выборка.Следующий())
                        {
                            var СсылкаНаРодителя = Выборка.Родитель;
                            if (СсылкаНаРодителя.Пустая())
                            {
                                Выборка.ПолучитьОбъект().Удалить();
                            }
                        }

                        Console.WriteLine("Удалили.");

                        //var f0 = СправочникНоменклатура.СоздатьГруппу();
                        //f0.Наименование = "Товары";
                        //f0.Записать();

                        var f1 = СправочникНоменклатура.СоздатьГруппу();
                        f1.Наименование = "Лекарственные средства";
                        //f1.Родитель = f0.Ссылка;
                        f1.Записать();


                        // перебираем таблицу из sql
                        Int32 ri = 0;
                        DataRow dr = dt.Rows[ri];
                        Boolean theEnd = false;
                        while (ri < dt.Rows.Count && !theEnd)
                        {
                            String f2Descr = dr["наименование"] as String;

                            Console.WriteLine(f2Descr);
                            /*
                            Справочники.СправочникОбъект f2 = СправочникНоменклатура.СоздатьГруппу();
                            f2.Наименование = f2Descr;
                            f2.Родитель = f1.Ссылка;
                            f2.Записать();

                            while (f2Descr == dr["наименование"] as String && !theEnd)
                            {
                                String i1Descr = dr["descr"] as String;
                                Справочники.СправочникОбъект i1 = СправочникНоменклатура.СоздатьЭлемент();
                                i1.Наименование = i1Descr;
                                i1.Родитель = f2.Ссылка;
                                i1.SetProperty("ВидНоменклатуры", ВидНоменклатурыТовар);
                                i1.SetProperty("ТипНоменклатуры", ТипНоменклатурыТовар);
                                i1.SetProperty("ЕдиницаИзмерения", УпаковкиЕдиницыИзмерения_Шт_Ссылка);
                                i1.SetProperty("СтавкаНДС", СтавкиНДС10);
                                i1.SetProperty("ВариантОформленияПродажи", ВариантыОформленияПродажиРеализацияТоваровУслуг);
                                i1.SetProperty("ИспользованиеХарактеристик", ВариантыИспользованияХарактеристикНоменклатуры_НеИспользовать);

                                var q2 = new ТабличнаяЧасть(i1.GetProperty("ДополнительныеРеквизиты"));

                                var q3 = q2.Добавить();
                                q3.SetProperty("Свойство", ДопРеквизитНаименование);
                                q3.SetProperty("Значение", f2Descr);

                                q3 = q2.Добавить();
                                q3.SetProperty("Свойство", ДопРеквизитФорма);
                                q3.SetProperty("Значение", dr["форма_выпуска"] as String);

                                q3 = q2.Добавить();
                                q3.SetProperty("Свойство", ДопРеквизитДозировка);
                                q3.SetProperty("Значение", dr["дозировка"] as String);

                                q3 = q2.Добавить();
                                q3.SetProperty("Свойство", ДопРеквизитУпаковка);
                                q3.SetProperty("Значение", dr["упаковка"] as String);

                                q3 = q2.Добавить();
                                q3.SetProperty("Свойство", ДопРеквизитКоличество);
                                q3.SetProperty("Значение", Decimal.Parse(dr["количество"] as String));

                                q3 = q2.Добавить();
                                q3.SetProperty("Свойство", ДопРеквизитПроизводитель);
                                q3.SetProperty("Значение", dr["производитель"] as String);

                                q3 = q2.Добавить();
                                q3.SetProperty("Свойство", ДопРеквизитСтрана);
                                q3.SetProperty("Значение", dr["страна"] as String);
                                

                                i1.Записать();

                                while (i1Descr == dr["descr"] as String)
                                {
                                    ri++;
                                    if (ri >= dt.Rows.Count) { theEnd = true; break; }
                                    dr = dt.Rows[ri];
                                }
                            }
                            */
                        }


                    }
                    catch (Exception e) { Console.WriteLine(e); }
                    finally
                    {
                        //if (globalContext != null) { globalContext.Release(); }
                        GC.Collect();
                    }
                }
            }
        }
        private static void TestRegex()
        {
            String src = "12.3 vr/45тыс.6rt+231 тыс sfdg/ 1 105,46";
            Console.WriteLine("'{0}'", src);

            String[] ps = src.Split('+');
            Regex re = new Regex(@""
                + @"^\s*" // начальные пробелы пропускаем
                          // три группы
                + @"(\d+[\d\., ]*)" // обязательное число с фиксированной точкой без разделителя групп (g1)
                + @"\s*"
                + @"(тыс|млн)?\.?" // множитель (g2)
                + @"\s*"
                + @"(%|\w*)?\.?" // единица измерения (g3)
                );
            for (int pi = 0; pi < ps.Length; pi++)
            {
                String p = ps[pi];
                Console.WriteLine("    {0}: '{1}'", pi, p);
                String[] ud = p.Split(new Char[] { '/', '|', '\\' });
                for (int i = 0; i < 2; i++)
                {
                    String t = ud.Length > i ? ud[i] : String.Empty;
                    Console.WriteLine("        {0}: '{1}'", i, t);
                    Match m = re.Match(t);
                    GroupCollection gs = m.Groups;
                    // три группы
                    Double g1 = 0;
                    if (gs.Count > 1)
                    {
                        Double.TryParse(gs[1].Value.Replace(" ", String.Empty).Replace(',', '.'), NumberStyles.Any, CultureInfo.InvariantCulture, out g1);
                    }
                    String g2 = gs.Count > 2 ? gs[2].Value : String.Empty;
                    switch (g2)
                    {
                        case "тыс":
                            g1 *= 1000;
                            break;
                        case "млн":
                            g1 *= 1000000;
                            break;
                        default:
                            break;
                    }
                    String g3 = gs.Count > 3 ? gs[3].Value : String.Empty;

                    Console.WriteLine("            '{0}', '{1}'", g1, g3);
                }
            }
        }
        private static void TestLinks()
        {
            // 1. читать все ссылки на файлы РУ из [Pharm-Sib].[dbo].[files]
            DataTable links = new DataTable();

            String cnString = "Data Source=192.168.135.14;Integrated Security=True";
            SqlCommand cmd = new SqlCommand
            {
                Connection = new SqlConnection(cnString),
                CommandType = CommandType.StoredProcedure,
                CommandText = "[Pharm-Sib].[dbo].[рег_уд__файлы__get]"
            };
            cmd.Parameters.AddWithValue("max_rows_count", 0);
            (new SqlDataAdapter(cmd)).Fill(links);

            // 2. каждую ссылку проверить
            int dCount = 0;
            int fCount = 0;
            int errCount = 0;
            foreach (DataRow row in links.Rows)
            {
                String path = row["path"] as String;
                path = @"\\SHD\reg_doc" + path.Replace(@"/", @"\");
                DirectoryInfo di = new DirectoryInfo(path);
                if (di.Exists)
                {
                    // это каталог - всё хорошо - переходим на следующую ссылку
                    dCount++;
                    continue;
                }
                // это не каталог - может быть это файл
                FileInfo fi = new FileInfo(path);
                if (fi.Exists)
                {
                    // это файл - всё хорошо - переходим на следующую ссылку
                    fCount++;
                    continue;
                }
                // это не каталог и не файл - сообщаем об ошибке
                errCount++;
                Console.WriteLine($"{errCount} {path}");

                /*
                cmd.CommandText =
                    @"update [dbo].[files] set [deleted] = 1 where [file_id] = N'" + file_id + "';";
                try
                {
                    cmd.Connection.Open();
                    cmd.ExecuteNonQuery();
                }
                catch (Exception e) { Console.WriteLine(e.ToString()); }
                finally { cmd.Connection.Close(); }
                */

            }
            Console.WriteLine($"всего ссылок: {links.Rows.Count}, каталогов: {dCount}, файлов: {fCount}, ошибок: {errCount}");
        }
        private static void TestHttpDataServer()
        {
            RequestPackage rqp = new RequestPackage
            {
                Command = "ПолучитьСписокПриходныхНакладных",
                SessionId = Guid.NewGuid(),
                /*
                Parameters = new RequestParameter[] {
                    new RequestParameter { Name = "auction_number", Value = "31806726909" },
                    new RequestParameter { Name = "overwrite", Value = true },
                }
                */
            };
            Console.WriteLine("Request to 'http://127.0.0.1:11015/'.\n");
            ResponsePackage rsp = rqp.GetResponse("http://127.0.0.1:11015/");
            if (rsp == null)
            {
                Console.WriteLine("rsp is null.\n");
            }
            else
            {
                Console.WriteLine(String.Format("rsp.Status: '{0}'\n", rsp.Status));
                if(rsp.Data != null)
                {
                    Console.WriteLine(String.Format("Получен набор данных."));
                    DataSet ds = rsp.Data;
                    if (ds.Tables.Count > 0)
                    {
                        Console.WriteLine($"В нём таблиц: {ds.Tables.Count}.");
                        foreach(DataTable dt in ds.Tables)
                        {
                            Console.WriteLine($"Таблица: '{dt.TableName}'");
                        }
                    }
                    else { Console.WriteLine("Таблиц в нём нет."); }
                }
            }
            return;
        }
        /*
            string sLogin = "sibdomain/sokolov";
            string sPassword = "1234548";
            string sComputer = "192.168.135.14";

            //создание процесса на удаленной машине
            ManagementScope ms;
            ConnectionOptions co = new ConnectionOptions();
            co.Username = sLogin;
            co.Password = sPassword;
            co.EnablePrivileges = true;
            co.Impersonation = ImpersonationLevel.Impersonate;

            ms = new ManagementScope(string.Format(@"\\{0}\root\CIMV2", sComputer), co);

            ms.Connect();

            ManagementPath path = new ManagementPath("Win32_Process");
            System.Management.ManagementClass classObj = new System.Management.ManagementClass(ms, path, null);
            System.Management.ManagementBaseObject inParams = null;
            inParams = classObj.GetMethodParameters("Create");
            inParams["CommandLine"] = "notepad.exe";
            inParams["CurrentDirectory"] = "C:\\WINDOWS\\system32\\";
            ManagementBaseObject outParams = classObj.InvokeMethod("Create", inParams, null); 
        */
        /*
            // Загрузка из Excel новых данных
            DataTable dt = new DataTable();
            String xlCnString =
                "Provider='Microsoft.ACE.OLEDB.12.0';" +
                "Data Source='C:\\Users\\sokolov\\Desktop\\6703_16_аэ_приложение_№1.xls';" +
                "Extended Properties='Excel 12.0 Macro;HDR=No;IMEX=1;';";
            using (OleDbConnection xlCn = new OleDbConnection(xlCnString))
            {
                try
                {
                    xlCn.Open();
                    DataTable schema = xlCn.GetOleDbSchemaTable(OleDbSchemaGuid.Tables, null);
                    String sheetName = schema.Rows[0]["TABLE_NAME"] as String;
                    OleDbCommand cmd = new OleDbCommand();
                    cmd.Connection = xlCn;
                    cmd.CommandText =
                        "select * " +
                        "from [" + sheetName + "] ";
                    (new OleDbDataAdapter(cmd)).Fill(dt);
                    xlCn.Close();
                }
                catch (Exception ex) { Console.WriteLine(ex.Message); }
            }
        */
        /*
        MailAddress from = new MailAddress("sokolov_ea@farmsib.ru", "Автоматическая рассылка", System.Text.Encoding.UTF8);
        MailAddress to = new MailAddress("sokolov_ea@farmsib.ru");
        MailMessage message = new MailMessage(from, to);
        message.SubjectEncoding = System.Text.Encoding.UTF8;
        message.Subject = "Сообщение о изменении судебного статуса.";
        message.BodyEncoding = System.Text.Encoding.UTF8;
        message.Body = "This is a test e-mail message sent by an application.";
        SmtpClient client = new SmtpClient("nicmail.ru", 25);
        client.UseDefaultCredentials = false;
        client.Credentials = new NetworkCredential("sokolov_ea@farmsib.ru", "69Le5PfLQCpQY");
        client.SendCompleted += new SendCompletedEventHandler(SendCompletedCallback);

        object userToken = Guid.NewGuid();
        client.SendAsync(message, userToken);

        Console.WriteLine("Sending message... press c to cancel mail. Press any other key to exit.");
        string answer = Console.ReadLine();
        // If the user canceled the send, and mail hasn't been sent yet,
        // then cancel the pending operation.
        if (answer.StartsWith("c") && mailSent == false)
        {
            client.SendAsyncCancel();
        }
        // Clean up.
        message.Dispose();
        Console.WriteLine("Goodbye.");
        */
        /*
        try
        {
                
            Type v8ComConnector = Type.GetTypeFromProgID("V83.ComConnector");
            Object v8 = Activator.CreateInstance(v8ComConnector);
            Object[] arguments = { @"Srvr=""srv-82:1741""; Ref=""BUH""; Usr=""Соколов Евгений COM1""; Pwd=""yNFxfrvqxP"";" };
            var x = v8ComConnector.InvokeMember("Connect", BindingFlags.Public | BindingFlags.InvokeMethod | BindingFlags.Static, null, v8, arguments);

            var query = InvokeObjectMethod(v8, x, "NewObject", new Object[] { "Запрос" });
            SetObjectProperty(v8, query, "Текст", new Object[] { "ВЫБРАТЬ Контрагенты.Наименование ИЗ Справочник.Контрагенты КАК Контрагенты" });
            var result = InvokeObjectMethod(v8, query, "Выполнить", new Object[] { });

            //var refs = GetObjectProperty(v8, x, "Справочники");
            //var result = GetObjectProperty(v8, refs, "Контрагенты");

            var selection = InvokeObjectMethod(v8, result, "Выбрать", new Object[] { });
            while ((bool)InvokeObjectMethod(v8, selection, "Следующий", new Object[] { }))
            {
                var name = GetObjectProperty(v8, selection, "Наименование");
                Console.WriteLine("Наименование: " + name);
            }

            //v8.GetType().InvokeMember("Exit", BindingFlags.Public | BindingFlags.InvokeMethod | BindingFlags.Static, null, x, new Object[] { false });
            //InvokeObjectMethod(v8, x, "Exit", new Object[] { false });

            //System.Runtime.InteropServices.Marshal.FinalReleaseComObject(x);
            //System.Runtime.InteropServices.Marshal.FinalReleaseComObject(v8);
            //GC.Collect();
                
        }
        catch (Exception e) { Console.WriteLine(e.ToString()); }
        */
        private static void TestInterop()
        {
            /*
            DataSet ds = new DataSet();

            String fileName = @"C:\Users\sokolov\Desktop\Заказы\ФОРМА ПРЕДЛОЖЕНИЯ УЧАСТНИКА ЗАКУПКИ- тест .docx";
            //String fileName = @"C:\Users\sokolov\Desktop\Заказы\ФОРМА ПРЕДЛОЖЕНИЯ УЧАСТНИКА ЗАКУПКИ- тест .doc";
            //String fileName = @"C:\Users\sokolov\Desktop\Заказы\Р­Р”_Р»РµРє.СЃСЂРµРґСЃС‚РІР°_(РѕС‚С…Р°СЂРєРёРІР°СЋС‰РёРµ).docx";
            Word.Application word = new Word.Application();
            word.Visible = false;
            Word.Document doc = word.Documents.Open(FileName: fileName, ReadOnly: true);
            try
            {
                foreach (Word.Table table in doc.Tables)
                {
                    DataTable dt = new DataTable(table.ID);
                    ds.Tables.Add(dt);

                    for (int ci = 0; ci < table.Columns.Count; ci++)
                    {
                        DataColumn dc = new DataColumn("Column" + ci.ToString(), typeof(String));
                        dt.Columns.Add(dc);
                    }
                    for (int ri = 0; ri < table.Rows.Count; ri++)
                    {
                        DataRow dr = dt.NewRow();
                        dt.Rows.Add(dr);
                    }
                    foreach (Word.Cell cell in table.Range.Cells)
                    {
                        String text = cell.Range.Text;
                        text = (new Regex(@"[\x00-\x1f]")).Replace(text, " ");
                        text = text.Trim();
                        dt.Rows[cell.RowIndex - 1][cell.ColumnIndex - 1] = text;
                    }
                }
            }
            catch (Exception e)
            {
                Console.Write(e.ToString());
            }
            finally
            {
                doc.Close(SaveChanges: false);
                word.Quit(SaveChanges: false);
                KillProcesses("WinWord");
            }

            foreach (DataTable table in ds.Tables)
            {
                Boolean isFinded = false;
                Mnn mnn = Mnn.Find(table);
                if ((mnn != null) && (table.Rows.Count > 1))
                {
                    //DataRow dr = table.Rows[table.Rows.IndexOf(mnn.Row) + 1];
                    //if (!String.IsNullOrWhiteSpace(dr[mnn.Column] as String))
                    {
                        isFinded = true;
                    }
                }
                if (isFinded)
                {
                    Console.WriteLine("Таблица " + table.TableName);
                    foreach (DataRow row in table.Rows)
                    {
                        Console.WriteLine("____Строка");
                        foreach (DataColumn col in table.Columns)
                        {
                            Console.Write("________");
                            String text = row[col] as String;
                            text = text ?? "null";
                            Console.WriteLine("'" + text + "'");
                        }
                    }
                }
            }
            */
        }
        class Mnn
        {
            private static Regex re = new Regex("Международное непатентованное наименование|Мнн|МНН");
            private static Int32 maxCnt = 5; // количество первых строчек для поиска

            public DataRow Row { get; private set; }
            public DataColumn Column { get; private set; }
            private Mnn(DataRow row, DataColumn column) { Row = row; Column = column; }
            public static Mnn Find(DataTable dt)
            {
                Mnn mnn = null;
                if (dt != null)
                {
                    // ищем только в первых строчках пока не найдём.
                    for (int ri = 0; (ri < Math.Min(dt.Rows.Count, maxCnt)) && (mnn == null); ri++)
                    {
                        DataRow dr = dt.Rows[ri];
                        foreach (DataColumn dc in dt.Columns)
                        {
                            if ((dr[dc] != DBNull.Value) && (re.IsMatch((String)dr[dc])))
                            {
                                mnn = new Mnn(dr, dc);
                                break;
                            }
                        }
                    }
                }
                return mnn;
            }
        }
        private static void KillProcesses(String name)
        {
            Process[] ps = Process.GetProcesses();
            foreach (Process p in ps)
            {
                if (p.ProcessName.ToUpper() == name.ToUpper())
                {
                    p.Kill();
                    Console.WriteLine("Process {0,5:#####} {1:s} has been killed.", p.Id, p.ProcessName);
                }
            }
        }
        private static void TestXmlSerializer()
        {
            RequestPackage rqp = new RequestPackage();
            rqp.Command = "Проверка";
            rqp.Parameters = new RequestParameter[] {
                new RequestParameter {Name="фываы0", Value=1},
                new RequestParameter {Name="фываы1", Value=1.2},
                new RequestParameter {Name="фываы2", Value=DateTime.Now},
                new RequestParameter {Name="фываы2", Value="q"},
                new RequestParameter {Name="фываы2", Value=Guid.NewGuid()},
                new RequestParameter {Name="фываы2", Value= new Byte[]{1,2,3}},
                new RequestParameter {Name="фываы2", Value= true}
            };

            ResponsePackage rsp = new ResponsePackage();
            rsp.Status = "asfлывоар";
            rsp.Data = new DataSet();
            rsp.Data.Tables.Add(new DataTable("qqq"));
            rsp.Data.Tables[0].Columns.Add("c1", typeof(Boolean));
            DataRow dr = rsp.Data.Tables[0].NewRow();
            dr[0] = true;
            rsp.Data.Tables[0].Rows.Add(dr);
            dr = rsp.Data.Tables[0].NewRow();
            dr[0] = false;
            rsp.Data.Tables[0].Rows.Add(dr);
            dr = rsp.Data.Tables[0].NewRow();
            dr[0] = DBNull.Value;
            rsp.Data.Tables[0].Rows.Add(dr);

            Byte[] buff = new Byte[8192];
            MemoryStream stream = new MemoryStream(buff);
            // кодировка без (ef bb bf) - заголовок Utf8 порядка байт
            TextWriter tr = new StreamWriter(stream, new UTF8Encoding(false));

            XmlSerializer sr = new XmlSerializer(typeof(RequestPackage));
            sr.Serialize(tr, rqp);

            Int32 len = (Int32)stream.Position;

            tr.Close();

            Console.WriteLine("---------------------------------------------------");
            Console.WriteLine(Encoding.UTF8.GetString(buff, 0, len));
            Console.WriteLine("---------------------------------------------------");
        }
        private static void Test1c()
        {
            Report1 r = new Report1("47449");
            DataTable dt = r.RegMoney;
            Console.WriteLine(dt.Rows.Count.ToString());
            foreach (DataRow dr in dt.Rows)
            {
                foreach (DataColumn dc in dt.Columns)
                {
                    Console.Write(dr[dc].ToString());
                }
                Console.WriteLine();
            }
        }
        public static object GetObjectProperty(object v8, object refObject, string propertyName)
        {
            return v8.GetType().InvokeMember(propertyName, BindingFlags.GetProperty, null, refObject, null);
        }
        public static object SetObjectProperty(object v8, object refObject, string propertyName, Object[] value)
        {
            return v8.GetType().InvokeMember(propertyName, BindingFlags.SetProperty, null, refObject, value);
        }
        public static object InvokeObjectMethod(object v8, object refObject, string methodName, Object[] parameters)
        {
            return v8.GetType().InvokeMember(methodName, BindingFlags.Public | BindingFlags.InvokeMethod | BindingFlags.Static, null, refObject, parameters);
        }
    }
    public static class Fs
    {
        //private static String home = @"C:\inetpub\Ank1_Data";
        private static String uploadsDirectory = @"C:\inetpub\Ank1_Data\Uploads";
        public static String docsDirectory = @"\\SRV-TS2\work\Тендерный\2015\Декабрь";
        public static String ctDirectory = @"\\SHD\st_sertificat";
        public static String rdDirectory = @"\\SHD\reg_doc";
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
        public static String UploadsPathCombine(String path = null, String name = null)
        {
            String dest = PathCombine(uploadsDirectory, path, name);
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
                default:
                    break;
            }
            if (baseDir != null)
            {
                path = absolutePath.Replace(baseDir, "");
            }
            return path;
        }
        /*
        public static void SavePostedFile(HttpPostedFileBase pf, String path)
        {
            String[] fnps = pf.FileName.Split('\\');
            if (fnps.Length > 0)
            {
                String fn = fnps[fnps.Length - 1];
                String guid = Guid.NewGuid().ToString();
                String dest = UploadsPathCombine(path, fn + "." + guid);
                try
                {
                    pf.SaveAs(dest);
                }
                catch (Exception ex)
                {
                    throw new Exception("Ank1.Data.Fs.SavePostedFile(): " + ex.Message);
                }
            }
        }
        */
        public static String ReadToEndFromUploads(String path)
        {
            String str = null;
            String src = UploadsPathCombine(path);
            using (StreamReader sr = new StreamReader(src))
            {
                str = sr.ReadToEnd();
            }
            return str;
        }
        public static String ReadToEnd(String path)
        {
            String str = null;
            using (StreamReader sr = new StreamReader(path))
            {
                str = sr.ReadToEnd();
            }
            return str;
        }
        /*
        public static String WriteStringAsUtf8FileToUploads(String str, String path)
        {
            Byte[] buff = System.Text.Encoding.UTF8.GetBytes(str);
            String dest = UploadsPathCombine(path);
            using (System.IO.FileStream fs = System.IO.File.Create(dest))
            {
                fs.Write(buff, 0, buff.Length);
                fs.Close();
            }
            Guid pointer = Db.GetFileIdByName(dest);
            return pointer.ToString();
        }
        */
    }
    class FsStoredProcedure
    {
        public static ResponsePackage GetDirectoryInfo(RequestPackage rqp)
        {
            ResponsePackage rsp = new ResponsePackage();
            rsp.Status = "Trying to get a directory information.";

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
                catch (Exception e) { rsp.Status = "Error."; Console.WriteLine(e.ToString()); }
            }
            return rsp;
        }
        public static ResponsePackage GetFileContents(RequestPackage rqp)
        {
            ResponsePackage rsp = new ResponsePackage();
            rsp.Status = "Trying to get a file information.";

            // псевдоним для базового каталога
            String alias = rqp["alias"] as String;
            // путь от базового каталога
            String path = rqp["path"] as String;
            // полный путь
            String absolutePath = Fs.GetAbsolutePath(alias, path);
            Console.WriteLine("absolutePath: '" + absolutePath + "'");

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

                    Byte[] buff = ReadFile(absolutePath);
                    if (buff != null)
                    {
                        dr[0] = Convert.ToBase64String(buff);
                    }
                }
                catch (Exception e) { rsp.Status = "Error."; Console.WriteLine(e.ToString()); }
            }
            return rsp;
        }
        private static Byte[] ReadFile(String path)
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
            catch (Exception e) { Console.WriteLine(e.ToString()); }
            return buff;
        }
    }
    class MailServer
    {
        private static bool mailSent = false;
        private static void SendCompletedCallback(object sender, AsyncCompletedEventArgs e)
        {
            // Get the unique identifier for this asynchronous operation.
            String userToken = e.UserState.ToString();

            if (e.Cancelled)
            {
                Console.WriteLine("[{0}] Send canceled.", userToken);
            }
            if (e.Error != null)
            {
                Console.WriteLine("[{0}] {1}", userToken, e.Error.ToString());
            }
            else
            {
                Console.WriteLine("Message sent.");
            }
            mailSent = true;
        }
        private static void Send(String address, String subject, String body, String attachment = null)
        {
            Console.WriteLine(String.Format("address: {0}, subject: {1}, body: {2}", address, subject, body));
            MailAddress from = new MailAddress("sokolov_ea@farmsib.ru", "Автоматическая рассылка", System.Text.Encoding.UTF8);
            MailAddress to = new MailAddress(address);
            MailMessage message = new MailMessage(from, to);
            message.SubjectEncoding = System.Text.Encoding.UTF8;
            message.Subject = subject;
            message.BodyEncoding = System.Text.Encoding.UTF8;
            message.IsBodyHtml = true;
            message.Body = body;
            if (attachment != null)
            {
                message.Attachments.Add(new Attachment(new MemoryStream(Encoding.UTF8.GetBytes(attachment)), "Запрос.html"));
            }
            SmtpClient client = new SmtpClient("nicmail.ru", 25);
            client.UseDefaultCredentials = false;
            client.Credentials = new NetworkCredential("sokolov_ea@farmsib.ru", "69Le5PfLQCpQY");
            client.SendCompleted += new SendCompletedEventHandler(SendCompletedCallback);

            object userToken = Guid.NewGuid();
            //client.SendAsync(message, userToken);
            try
            {
                client.Send(message);
            }
            catch (Exception e) { Console.WriteLine(e.ToString()); }
        }

        public static void Exec()
        {
            String address = "sokolov_ea@farmsib.ru";
            String subject = "Сообщение о изменении судебного статуса.";
            String body = "<html><head><title></title></head><body><font color='red'>test</font></body></html>";
            String attachment = null;

            try
            {
                Send(address, subject, body);
            }
            catch (Exception ex) { Console.WriteLine(ex.ToString()); }

            return;
        }
    }
}
