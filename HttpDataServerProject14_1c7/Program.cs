﻿using System;
using System.IO;
using System.Net;
using System.Text;
using System.Runtime.InteropServices;
using System.Diagnostics;
using Nskd;
using System.Net.Sockets;

namespace HttpDataServerProject14
{
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
                    Process[] ps = Process.GetProcesses(); //ByName("1cv7s");
                    foreach (Process p in ps)
                    {
                        if (p.ProcessName.ToUpper() == "1CV7S")
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
        private static HttpServer server;
        private static Native.SignalHandler signalHangler = null;

        private static String V77GarzaCnString1 = @"/d""\\SRV-TS2\dbase_1c$\ФК_Гарза"" /n""Соколов_Евгений_клиент_1"" /p""yNFxfrvqxP""";
        private static String V77GarzaCnString2 = @"/d""\\SRV-TS2\dbase_1c$\ФК_Гарза"" /n""Соколов_Евгений_клиент_2"" /p""yNFxfrvqxP""";
        private static String V77CnString = null;

        // Если нет параметров при запуске программы, то работаем с Sql server на моём компъютере.
        public static String MainSqlServerDataSource = "192.168.135.77";

        static void Main(string[] args)
        {
            // при закрытии программы будет вызвана процедура 'handlerRoutine'
            signalHangler += Native.HandlerRoutine;
            Native.SetSignalHandler(signalHangler, true);

            // Обязательно перед первым запуском в cmd.exe выполнить
            // netsh http add urlacl url=http://+:11014/ user=DOMAIN\user
            if (args != null)
            {
                foreach (String arg in args)
                {
                    if (arg == "-d14" || arg == "/d14") { MainSqlServerDataSource = "192.168.135.14"; }
                }
            }

            IPAddress ip = null;
            IPAddress[] ips = Dns.GetHostEntry(Dns.GetHostName()).AddressList;
            for (int i = 0; i < ips.Length; i++)
            {
                if (ips[i].AddressFamily == AddressFamily.InterNetwork)
                {
                    ip = ips[i];
                    break;
                }
            }

            V77CnString = (ip.ToString() == "192.168.135.14") ? V77GarzaCnString1 : V77GarzaCnString2;

            Log.Write(String.Format("Start accept on 'http://{0}:11014/' with MainSqlServerDataSource: '{1}'.", ip, MainSqlServerDataSource));

            server = new HttpServer();
            server.OnIncomingRequest += new HttpServer.RequestDelegate(OnIncomingRequest);
            server.Start("http://+:11014/");
        }
        public static void OnIncomingRequest(HttpListenerContext context)
        {
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
                    if (!String.IsNullOrWhiteSpace(body))
                    {
                        if (body[0] == '<')
                        {
                            RequestPackage rqp = RequestPackage.ParseXml(body);
                            if (rqp != null)
                            {
                                // исполнение запроса
                                ResponsePackage rsp = OcStoredProcedure.Exec1(rqp, V77CnString);
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
