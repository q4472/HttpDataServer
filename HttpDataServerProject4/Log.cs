using System;
using System.Threading;

namespace HttpDataServerProject4
{
    class Log
    {
        public static void Write(String msg)
        {
            DateTime now = DateTime.Now;
            Int32 threadId = Thread.CurrentThread.ManagedThreadId;
            String msg1 = String.Format("11004 {0:yyyy.MM.dd HH:mm:ss} {1}> {2}", now, threadId, msg);
            Console.WriteLine(msg1);
            Nskd.SqlServer.LogWrite(msg, "11004", "Message", Program.MainSqlServerDataSource);
        }
    }
}
