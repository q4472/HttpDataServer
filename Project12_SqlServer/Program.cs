using Nskd;
using System;
using System.IO;
using System.Net;
using System.Net.Sockets;
using System.Text;
using System.Threading;

namespace Project12
{
    class Program
    {
        private static HttpServer Server = null;
        private static Boolean IsDebugginMode = false;
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

        public static void Main(String[] args)
        {
            // Обязательно перед первым запуском в cmd.exe выполнить
            // netsh http add urlacl url=http://+:11012/ user=SIBDOMAIN\Sokolov
            if (args != null)
            {
                foreach (String arg in args)
                {
                    if (arg == "-d14" || arg == "/d14") { MainSqlServerDataSource = "192.168.135.14"; }
                    if (arg == "-debug" || arg == "/debug") { IsDebugginMode = true; }
                }
            }
            Log.Write(String.Format("Start accept on 'http://{0}:11012/' with MainSqlServerDataSource: '{1}'.", HostIPv4, MainSqlServerDataSource));
            Server = new HttpServer();
            Server.OnIncomingRequest += new HttpServer.RequestDelegate(OnIncomingRequest);
            Server.Start("http://+:11012/");
        }
        public static void OnIncomingRequest(HttpListenerContext context)
        {
            if (context == null) { return; }
            HttpListenerRequest request = context.Request;
            HttpListenerResponse response = context.Response;
            if (IsDebugginMode)
            {
                String address = request.RemoteEndPoint.Address.ToString();
                String method = request.HttpMethod;
                String host = request.Url.Host;
                String path = request.Url.AbsolutePath;
                Log.Write(String.Format("address: {0}, method: {1}, host: {2}, path: {3}", address, method, host, path));
            }
            if (Http.RequestIsAcceptable(request))
            {
                // разбор запроса
                Encoding enc = Encoding.UTF8;
                String body = null;
                using (StreamReader sr = new StreamReader(request.InputStream, enc))
                {
                    body = sr.ReadToEnd();
                }
                if (!String.IsNullOrWhiteSpace(body))
                {
                    if (IsDebugginMode)
                    {
                        Log.Write("Request.InputStream: " + body);
                    }
                    RequestPackage rqp = null;
                    switch (body[0])
                    {
                        case '<':
                            rqp = RequestPackage.ParseXml(body);
                            break;
                        case '{':
                            rqp = RequestPackage.ParseJson(body);
                            break;
                        default:
                            break;
                    }
                    if (rqp != null)
                    {
                        // исполнение запроса
                        ResponsePackage rsp = SqlServer.Exec(rqp);
                        // запись ответа
                        Byte[] buff = (rsp == null) ? new Byte[0] : rsp.ToXml(enc);
                        if (IsDebugginMode)
                        {
                            Log.Write("Response.OutputStream: " + enc.GetString(buff));
                        }
                        response.OutputStream.Write(buff, 0, buff.Length);
                    }
                    else { Log.Write("Непонятен формат запроса: " + body); }
                }
                else { Log.Write("Запрос пуст."); }
            }
            else { Log.Write("Запрос не прошел проверку."); }
        }
    }
}
