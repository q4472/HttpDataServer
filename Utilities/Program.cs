using Nskd;
using System;
using System.IO;
using System.Net;
using System.Net.Sockets;
using System.Text;

namespace HttpDataServices.Utilities
{
    class Program
    {
        private static HttpServer server;
        private static IPAddress HostIPv4
        {
            get
            {
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
                return ip;
            }
        }

        // Если нет параметров при запуске программы, то работаем с Sql server на моём компъютере.
        public static String MainSqlServerDataSource = "192.168.135.77";

        static void Main(string[] args)
        {
            // Обязательно перед первым запуском в cmd.exe выполнить
            // netsh http add urlacl url=http://+:11009/ user=sokolov
            if (args != null)
            {
                foreach (String arg in args)
                {
                    if (arg == "-d14" || arg == "/d14") { MainSqlServerDataSource = "192.168.135.14"; }
                }
            }
            Log.Write(String.Format("Start accept on 'http://{0}:11009/' with MainSqlServerDataSource: '{1}'.", HostIPv4, MainSqlServerDataSource));
            server = new HttpServer();
            server.OnIncomingRequest += new HttpServer.RequestDelegate(OnIncomingRequest);
            server.Start("http://+:11009/");
        }
        public static void OnIncomingRequest(HttpListenerContext context)
        {
            try
            {
                HttpListenerRequest request = context.Request;
                if (Http.RequestIsAcceptable(request))
                {
                    // разбор запроса
                    String body = null;
                    using (StreamReader sr = new StreamReader(request.InputStream, request.ContentEncoding))
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
                                ResponsePackage rsp = StoredProcedures.Execute(rqp);
                                // запись ответа
                                Byte[] buff = (rsp == null) ? new Byte[0] : rsp.ToXml(Encoding.UTF8);
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
