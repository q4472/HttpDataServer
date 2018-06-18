using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Reflection;
using System.Text;
using System.Threading.Tasks;

namespace Tests
{
    class Report1
    {
        public DataTable RegMoney = null;
        public Report1(String clientCode)
        {
            try
            {
                OneCv77 oc = new OneCv77();
                using (oc)
                {
                    oc.ConnectTo1cV77S();
                    if (oc.IsConnected)
                    {
                        RegMoney = oc.GetData(clientCode);
                    }
                }
            }
            catch (Exception exp)
            {
                //HttpDataServer.SqlServer.LogWrite(exp.ToString());
            }
        }

    }
    public class OneCv77 : IDisposable
    {
        private Type v77SApplication;
        private Object app;
        public Boolean IsConnected;
        public OneCv77()
        {
            // ищем приложение 1С77
            v77SApplication = Type.GetTypeFromProgID("V77S.Application");
            app = null;
            IsConnected = false;
        }
        public void Dispose()
        {
            if (app != null)
            {
                System.Runtime.InteropServices.Marshal.ReleaseComObject(app);
                app = null;
            }
        }
        public void ConnectTo1cV77S()
        {
            IsConnected = false;
            if (v77SApplication != null)
            {
                for (int tryNumber = 1; tryNumber <= 3; tryNumber++)
                {
                    // получаем экземпляр COM объекта 1с
                    app = Activator.CreateInstance(v77SApplication);
                    if (app != null)
                    {
                        Int32 code = (int)InvokeMethod(app, "RMTrade"); // код для режима 'предприятие'
                        String folder = EnvVars.App1c77Folder;
                        String userName = EnvVars.App1c77UserName + tryNumber.ToString();
                        String userPassword = EnvVars.App1c77UserPassword;

                        // запускаем в режиме предприятие
                        IsConnected = (Boolean)InvokeMethod(app, "Initialize", new Object[] { 
                            code, 
                            "/d\"" + folder + "\" /n" + userName + " /p" + userPassword, 
                            "NO_SPLASH_SHOW" });
                        if (IsConnected) break;
                    }
                }
            }
        }
        public DataTable GetData(String clientCode)
        {
            DataTable dt = new DataTable();
            dt.Columns.Add("Автор", typeof(String));
            dt.Columns.Add("Клиент", typeof(String));
            dt.Columns.Add("Документ", typeof(String));
            dt.Columns.Add("ДатаПолучения", typeof(DateTime));
            dt.Columns.Add("СрокОплаты", typeof(DateTime));
            dt.Columns.Add("Просрочка", typeof(Double));
            dt.Columns.Add("Сумма", typeof(Double));
            dt.Columns.Add("Счёт", typeof(String));
            dt.Columns.Add("Договор", typeof(String));
            dt.Columns.Add("ЭкземплярВернули", typeof(String));

            Object rM = InvokeMethod(app, "CreateObject", "Регистр.Деньги");

            DateTime vbdate = DateTime.Now;

            Object sC = InvokeMethod(app, "CreateObject", "Справочник.Клиенты");
            Object c = null;
            if ((Double)InvokeMethod(sC, "НайтиПоКоду", Int32.Parse(clientCode)) == 1) // 47449
            {
                c = InvokeMethod(sC, "ТекущийЭлемент");
            }
            if (c != null)
            {
                InvokeMethod(rM, "УстановитьЗначениеФильтра", new Object[] { "Клиент", c });
            }

            InvokeMethod(rM, "ВыбратьИтоги");
            while ((Double)InvokeMethod(rM, "ПолучитьИтог") == 1)
            {
                Double summ = (Double)GetAttrib(rM, "Сумма");
                if (summ > 0) continue;

                DateTime nullDateTime = new DateTime(1899, 12, 30);
                String vid = String.Empty;
                Object doc = GetAttrib(rM, "КрДок");
                DateTime dpgk = nullDateTime;
                DateTime srok = nullDateTime;
                Double prosr = 0;
                if ((Double)InvokeMethod(app, "ПустоеЗначение", doc) == 1)
                {
                }
                else
                {
                    vid = (String)InvokeMethod(doc, "Вид");
                    if ((vid != "Расходная") && (vid != "Акт"))
                    {
                    }
                    else
                    {
                        //if(Автор.Выбран() = 1) && (Автор <> Док.Автор) продолжить
                        if (vid == "Расходная")
                        {
                            dpgk = (DateTime)GetAttrib(doc, "ДатаПолученияГрузаКлиентом");
                            Double otsr = (Double)GetAttrib(doc, "Отсрочка");
                            //LogWrite(otsr.ToString());
                            //break;
                            //if ((Double)InvokeMethod(app, "ПустоеЗначение", dpgk) == 1)
                            if (dpgk == nullDateTime)
                            {
                                srok = (DateTime)GetAttrib(doc, "ДатаДок");
                            }
                            else
                            {
                                srok = (DateTime)dpgk;
                            }
                            srok = srok.AddDays(otsr);
                            prosr = (vbdate - srok).Days;
                        }
                    }
                }
                DataRow dr = dt.NewRow();
                dt.Rows.Add(dr);
                {
                    dr["Автор"] = GetAttrib(GetAttrib(doc, "Автор"), "Наименование");
                    dr["Клиент"] = GetAttrib(c, "Наименование");
                    dr["Документ"] = GetAttrib(doc, "НомерДок");
                    dr["ДатаПолучения"] = (vid == "Расходная") ? dpgk : nullDateTime;
                    dr["СрокОплаты"] = srok;
                    dr["Просрочка"] = (prosr < 0) ? 0 : prosr;
                    dr["Сумма"] = (-1) * summ;
                    Object doc1 = GetAttrib(doc, "ДокументОснование");
                    dr["Счёт"] = GetAttrib(doc1, "НомерДок");
                    Object doc2 = GetAttrib(doc, "Договор");
                    dr["Договор"] = GetAttrib(doc2, "Наименование");
                    Object enu = GetAttrib(doc, "ЭкземплярВернули");
                    dr["ЭкземплярВернули"] = InvokeMethod(enu, "Идентификатор");
                }
                /*
                LogWrite(GetAttrib(dr["Автор"], "Наименование"));
                LogWrite(" / ");
                LogWrite(GetAttrib(dr["Клиент"], "Наименование"));
                LogWrite(
                    GetAttrib(dr["Документ"], "НомерДок") + ": " +
                    ((Double)dr["Сумма"]).ToString() + " " +
                    ((DateTime)dr["ДатаПолучения"]).ToString("yyyy-MM-dd") + " " +
                    ((DateTime)dr["СрокОплаты"]).ToString("yyyy-MM-dd") + " " +
                    dr["Просрочка"].ToString()
                    );
                LogWrite();
                 */
            }

            return dt;
        }


        private static BindingFlags INVOKE_METHOD = BindingFlags.Public | BindingFlags.InvokeMethod | BindingFlags.Static;

        public static Object InvokeMethod(Object obj, String name)
        {
            return obj.GetType().InvokeMember(name, INVOKE_METHOD, null, obj, null);
        }
        public static Object InvokeMethod(Object obj, String name, Object value)
        {
            return obj.GetType().InvokeMember(name, INVOKE_METHOD, null, obj, new Object[] { value });
        }
        public static Object InvokeMethod(Object obj, String name, Object[] args)
        {
            return obj.GetType().InvokeMember(name, INVOKE_METHOD, null, obj, args);
        }
        public static Object GetAttrib(Object obj, String name)
        {
            return InvokeMethod(obj, "ПолучитьАтрибут", name);
        }
        public static Object SetAttrib(Object obj, String name, Object value)
        {
            return InvokeMethod(obj, "УстановитьАтрибут", new Object[] { name, value });
        }
    }

    public static class EnvVars
    {
        public static String SqlServerPharmSibConnectionString = String.Format(
            "Data Source={0};" +
            "Initial Catalog=Pharm-Sib;" +
            "Integrated Security=False;" +
            "Persist Security Info=False;" +
            "User=sa;" +
            "Password=sasa;" +
            "Connect Timeout=10"
            , "192.168.135.77"
            );

        public static String App1c77Folder = @"\\SRV-TS2\dbase_1c$\Фармацея Фарм-Сиб";
        public static String App1c77UserName = @"Соколов_Евгений_клиент_"; // имя для пула подключений. номер будкт добавлен потом.
        public static String App1c77UserPassword = @"yNFxfrvqxP";
    }
}
