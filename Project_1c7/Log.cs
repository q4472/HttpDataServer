using System;
using System.Diagnostics;
using System.Threading;

namespace Project_1c7
{
    class Log
    {
        public static void Write(String msg)
        {
            DateTime now = DateTime.Now;
            Int32 processId = Process.GetCurrentProcess().Id;
            Int32 threadId = Thread.CurrentThread.ManagedThreadId;
            String msg1 = String.Format($"{now:yyyy.MM.dd HH:mm:ss} {processId} {threadId}> {msg}");
            Console.WriteLine(msg1);
            Nskd.SqlServer.LogWrite(msg, processId.ToString(), "Message", Program.MainSqlServer);
        }
    }
}
