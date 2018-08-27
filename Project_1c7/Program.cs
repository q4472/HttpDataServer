using Nskd;
using System;
using System.Diagnostics;
using System.IO;
using System.Net;
using System.Net.Sockets;
using System.Runtime.InteropServices;
using System.Text;

namespace Project_1c7
{
    // Обязательно перед первым запуском в cmd.exe выполнить
    // netsh http add urlacl url=http://+:{port}/ user={DOMAIN}\{user}
    internal class Native
    {
        public delegate void SignalHandler(uint consoleSignal);

        [DllImport("Kernel32", EntryPoint = "SetConsoleCtrlHandler")]
        public static extern bool SetSignalHandler(SignalHandler handler, bool add);
        /// <summary>
        ///     При закрытии программы удаляет все процессы 1cv7s
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
                    KillProcessesByName("1CV7S");
                    break;
                default:
                    break;
            }
        }
        // call from Main
        public static void HandlerRoutine()
        {
            KillProcessesByName("1CV7S");
        }
        private static void KillProcessesByName(String name)
        {
            Process[] ps = Process.GetProcesses(); //ByName("1cv7s");
            foreach (Process p in ps)
            {
                if (p.ProcessName.ToUpper() == "1CV7S")
                {
                    p.Kill();
                }
            }
        }
    }
    class Program
    {
        public static String V77CnString = null;
        public static String MainSqlServer = null;
        static void Main(string[] args)
        {
            // при закрытии программы будут закрыты ВСЕ процессы 1cv7s
            Native.SetSignalHandler(Native.HandlerRoutine, true);

            String localIp = null;
            IPAddress[] ips = Dns.GetHostEntry(Dns.GetHostName()).AddressList;
            for (int i = 0; i < ips.Length; i++)
            {
                if (ips[i].AddressFamily == AddressFamily.InterNetwork)
                {
                    localIp = ips[i].ToString();
                    break;
                }
            }

            // работаем с локальным Sql server.
            MainSqlServer = localIp;

            String v77CnStringD = null;
            String v77CnStringN = null;
            String v77CnStringP = null;
            String port = null;
            foreach (String arg in args)
            {
                String[] nv = arg.Split('=');
                String n = nv[0];
                if (nv.Length > 1)
                {
                    String v = nv[1];
                    if (n == "1cd" || n == "/1cd" || n == "-1cd") { v77CnStringD = v; }
                    if (n == "1cn" || n == "/1cn" || n == "-1cn") { v77CnStringN = v; }
                    if (n == "1cp" || n == "/1cp" || n == "-1cp") { v77CnStringP = v; }
                    if (n == "port" || n == "/port" || n == "-port") { port = v; }
                }
            }

            if (v77CnStringD != null && v77CnStringN != null && v77CnStringP != null)
            {
                V77CnString = String.Format($"/d\"{v77CnStringD}\" /n\"{v77CnStringN}\" /p\"{v77CnStringP}\"");
            }

            if (V77CnString != null)
            {
                if (port != null)
                {
                    Log.Write(String.Format($"Start accept on 'http://{localIp}:{port}/' with main SqlServer on: '{MainSqlServer}' and 1C connection string '{V77CnString}'."));

                    HttpServer server = new HttpServer();
                    server.OnIncomingRequest += new HttpServer.RequestDelegate(OnIncomingRequest);
                    server.Start($"http://+:{port}/");
                }
                else { Log.Write("Не задан порт прослушивания."); }
            }
            else { Log.Write("Не заданы параметры подключения к 1С."); }

            Console.ReadKey();
            // при выходе из программы будут закрыты ВСЕ процессы 1cv7s
            Native.HandlerRoutine();
        }
        public static void OnIncomingRequest(HttpListenerContext context)
        {
            //Log.Write("IncomingRequest");
            try
            {
                HttpListenerRequest incomingRequest = context.Request;
                if (Http.RequestIsAcceptable(incomingRequest))
                {
                    // разбор запроса
                    Encoding enc = Encoding.UTF8;
                    String body = null;
                    using (StreamReader sr = new StreamReader(incomingRequest.InputStream, enc))
                    {
                        body = sr.ReadToEnd();
                    }
                    //Log.Write(body);
                    if (!String.IsNullOrWhiteSpace(body))
                    {
                        if (body[0] == '<')
                        {
                            RequestPackage rqp = RequestPackage.ParseXml(body);
                            if (rqp != null)
                            {
                                ResponsePackage rsp = null;
                                // исполнение запроса
                                rsp = OcStoredProcedure.Exec1(rqp, V77CnString);
                                // запись ответа
                                Byte[] buff = (rsp == null) ? new Byte[0] : rsp.ToXml(enc);
                                context.Response.OutputStream.Write(buff, 0, buff.Length);
                            }
                        }
                        else { Log.Write("Непонятен формат запроса: " + body); }
                    }
                    else { Log.Write("Запрос пуст."); }
                }
                else { Log.Write("Запрос не прошел проверку."); }
            }
            catch (Exception ex) { Log.Write(ex.ToString()); }
            finally { context.Response.Close(); }
        }
    }
}
