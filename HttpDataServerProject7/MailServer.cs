using Nskd;
using System;
using System.ComponentModel;
using System.IO;
using System.Net;
using System.Net.Mail;
using System.Text;

namespace HttpDataServerProject7
{
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
            Log.Write(String.Format("address: {0}, subject: {1}, body: {2}", address, subject, body));
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
            client.SendAsync(message, userToken);
        }

        public static ResponsePackage Exec(RequestPackage rqp)
        {
            ResponsePackage rsp = new ResponsePackage();
            rsp.Status = "Запрос к e-mail серверу nimail.ru на отправку почтового сообщения по smtp.";
            rsp.Data = null;
            String address = null;
            String subject = null;
            String body = null;
            String attachment = null;
            switch (rqp.Command)
            {
                case "Prep.F4.SendEmail":
                    if ((rqp != null) && (rqp.Parameters != null))
                    {
                        foreach (var p in rqp.Parameters)
                        {
                            if (p != null)
                            {
                                String key = p.Name;
                                Object value = p.Value;
                                if (!String.IsNullOrWhiteSpace(key) && (value != null) && (value != DBNull.Value))
                                {
                                    if (value.GetType() == typeof(String))
                                    {
                                        if (key == "address") { address = (String)value; }
                                        if (key == "subject") { subject = (String)value; }
                                        if (key == "body") { body = (String)value; }
                                        if (key == "attachment") { attachment = (String)value; }
                                    }
                                }
                            }
                        }
                    }
                    try
                    {
                        Send(address, subject, body, attachment);
                    }
                    catch (Exception ex) { Log.Write(ex.ToString()); }
                    break;
                default:
                    subject = "Сообщение о изменении судебного статуса.";
                    if ((rqp != null) && (rqp.Parameters != null))
                    {
                        foreach (var p in rqp.Parameters)
                        {
                            if (p != null)
                            {
                                String key = p.Name;
                                Object value = p.Value;
                                if (!String.IsNullOrWhiteSpace(key) && (value != null) && (value != DBNull.Value))
                                {
                                    if (value.GetType().ToString() == "System.String")
                                    {
                                        if (key == "address") { address = (String)value; }
                                        if (key == "msg") { body = (String)value; }
                                    }
                                }
                            }
                        }
                    }
                    try
                    {
                        Send(address, subject, body);
                    }
                    catch (Exception ex) { Log.Write(ex.ToString()); }
                    break;
            }
            return rsp;
        }
    }
}
