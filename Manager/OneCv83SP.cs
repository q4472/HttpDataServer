using Nskd.V83;
using System;
using System.Collections.Generic;
using System.Data;
using System.Text.RegularExpressions;

namespace Manager
{
    public class OcStoredProcedures
    {
        private static String ocCnString = @"Srvr=""srv-82:1741"";Ref=""BUH"";Usr=""Соколов Евгений"";Pwd=""yNFxfrvqxP"";";
        private static Object thisLock = new Object();
        private static object sys;

        public static DataSet Exec0(String command, Dictionary<String, Object> pars)
        {
            DataSet ds = null;
            lock (thisLock)
            {
                GlobalContext globalContext = new GlobalContext(ocCnString);
                if (globalContext != null)
                {
                    try
                    {
                        DataTable dt = null;
                        switch (command)
                        {
                            case "GetPartnerList":
                                dt = GetPartnerList(globalContext);
                                if (dt != null)
                                {
                                    ds = new DataSet();
                                    ds.Tables.Add(dt);
                                }
                                break;
                            case "GetPaymentList":
                                DateTime sd = (DateTime)pars["НачалоПериода"];
                                DateTime ed = (DateTime)pars["КонецПериода"];
                                dt = GetPaymentList(globalContext, sd, ed);
                                if (dt != null)
                                {
                                    ds = new DataSet();
                                    ds.Tables.Add(dt);
                                }
                                break;
                            default:
                                break;
                        }
                    }
                    catch (Exception e) { Nskd.SqlServer.LogWrite(e.ToString(), "OneCv83", "Exception", "192.168.135.14"); }
                }
            }
            return ds;
        }
        public static DataTable GetPartnerList(GlobalContext globalContext)
        {
            DataTable partners = new DataTable();
            partners.Columns.Add("descr", typeof(String));
            partners.Columns.Add("descrF", typeof(String));
            partners.Columns.Add("inn", typeof(String));
            partners.Columns.Add("kpp", typeof(String));

            try
            {
                var Запрос = globalContext.Запрос;
                // ПЕРВЫЕ 50
                Запрос.SetProperty("Текст", @"
                    ВЫБРАТЬ 
                        Контрагент.Наименование КАК Наименование,
                        Контрагент.НаименованиеПолное КАК НаименованиеПолное,
                        Контрагент.ИНН КАК ИНН,
                        Контрагент.КПП КАК КПП
                    ИЗ 
                        Справочник.Контрагенты КАК Контрагент
                    ");
                //ГДЕ
                //Контрагент.ПометкаУдаления = Ложь
                var РезультатЗапроса = Запрос.Выполнить();
                var ВыборкаИзРезультатаЗапроса = РезультатЗапроса.Выбрать();
                while (ВыборкаИзРезультатаЗапроса.Следующий())
                {
                    String descr = (String)(ВыборкаИзРезультатаЗапроса.GetProperty("Наименование"));
                    String descrF = (String)(ВыборкаИзРезультатаЗапроса.GetProperty("НаименованиеПолное"));
                    String inn = (String)(ВыборкаИзРезультатаЗапроса.GetProperty("ИНН"));
                    String kpp = (String)(ВыборкаИзРезультатаЗапроса.GetProperty("КПП"));
                    DataRow dr = partners.NewRow();
                    partners.Rows.Add(dr);
                    dr[0] = (descr ?? "").Trim();
                    dr[1] = (descrF ?? "").Trim();
                    dr[2] = (inn ?? "").Trim();
                    dr[3] = (kpp ?? "").Trim();
                    //Console.WriteLine((String)dr[0] + " | " + (String)dr[1]);
                }
                SqlServer.UploadPartnerList(partners);
                Console.WriteLine(partners.Rows.Count.ToString());
            }
            catch (Exception e) { Console.WriteLine(e.ToString()); }
            return partners;
        }
        public static DataTable GetPaymentList(GlobalContext globalContext, DateTime sd, DateTime ed)
        {
            DataTable payments = new DataTable();
            payments.Columns.Add("date", typeof(DateTime));
            payments.Columns.Add("amount", typeof(Double));
            payments.Columns.Add("d_a_code", typeof(String));
            payments.Columns.Add("d_p_inn", typeof(String));
            payments.Columns.Add("d_p_kpp", typeof(String));
            payments.Columns.Add("d_p_descr", typeof(String));
            payments.Columns.Add("d_p_descr_f", typeof(String));
            payments.Columns.Add("c_a_code", typeof(String));
            payments.Columns.Add("c_p_inn", typeof(String));
            payments.Columns.Add("c_p_kpp", typeof(String));
            payments.Columns.Add("c_p_descr", typeof(String));
            payments.Columns.Add("c_p_descr_f", typeof(String));
            payments.Columns.Add("r_descr", typeof(String));
            payments.Columns.Add("r_note", typeof(String));
            payments.Columns.Add("trade_num", typeof(String));

            String ssd = sd.ToString("yyyy, MM, dd, 0, 0, 0");
            String sed = ed.ToString("yyyy, MM, dd, 23, 59, 59");

            var Запрос = globalContext.Запрос;
            try
            {
                Запрос.Текст = @"
                    ВЫБРАТЬ
                    	ХозрасчетныйДвиженияССубконто.Период КАК Период,
                    	ХозрасчетныйДвиженияССубконто.Сумма КАК Сумма,
	                    ХозрасчетныйДвиженияССубконто.СчетДт.Код КАК ДтСчетКод,
	                    ВЫБОР
                            КОГДА (Выразить(СчетДт.Код КАК СТРОКА(5)) = ""76.06"")
                                ТОГДА ХозрасчетныйДвиженияССубконто.СубконтоДт1.ИНН
                            ИНАЧЕ """"
                        КОНЕЦ КАК ДтКонтрагентИНН,
                    	ВЫБОР
                            КОГДА (Выразить(СчетДт.Код КАК СТРОКА(5)) = ""76.06"")
                                ТОГДА ХозрасчетныйДвиженияССубконто.СубконтоДт1.КПП
                            ИНАЧЕ """"
                        КОНЕЦ КАК ДтКонтрагентКПП,
	                    ВЫБОР
                            КОГДА (Выразить(СчетДт.Код КАК СТРОКА(5)) = ""76.06"")
                                ТОГДА ХозрасчетныйДвиженияССубконто.СубконтоДт1.Наименование
                            ИНАЧЕ """"
                        КОНЕЦ КАК ДтКонтрагентНаименование,
	                    ВЫБОР
                            КОГДА (Выразить(СчетДт.Код КАК СТРОКА(5)) = ""76.06"")
                                ТОГДА ХозрасчетныйДвиженияССубконто.СубконтоДт1.НаименованиеПолное
                            ИНАЧЕ """"
                        КОНЕЦ КАК ДтКонтрагентНаименованиеПолное,
	                    ХозрасчетныйДвиженияССубконто.СчетКт.Код КАК КтСчетКод,
	                    ВЫБОР
                            КОГДА (Выразить(СчетКт.Код КАК СТРОКА(5)) = ""76.06"")
                                ТОГДА ХозрасчетныйДвиженияССубконто.СубконтоКт1.ИНН
                            ИНАЧЕ """"
                        КОНЕЦ КАК КтКонтрагентИНН,
	                    ВЫБОР
                            КОГДА (Выразить(СчетКт.Код КАК СТРОКА(5)) = ""76.06"")
                                ТОГДА ХозрасчетныйДвиженияССубконто.СубконтоКт1.КПП
                            ИНАЧЕ """"
                        КОНЕЦ КАК КтКонтрагентКПП,
	                    ВЫБОР
                            КОГДА (Выразить(СчетКт.Код КАК СТРОКА(5)) = ""76.06"")
                                ТОГДА ХозрасчетныйДвиженияССубконто.СубконтоКт1.Наименование
                            ИНАЧЕ """"
                        КОНЕЦ КАК КтКонтрагентНаименование,
                    	ВЫБОР
                            КОГДА (Выразить(СчетКт.Код КАК СТРОКА(5)) = ""76.06"")
                                ТОГДА ХозрасчетныйДвиженияССубконто.СубконтоКт1.НаименованиеПолное
                            ИНАЧЕ """"
                        КОНЕЦ КАК КтКонтрагентНаименованиеПолное,
	                    ХозрасчетныйДвиженияССубконто.Регистратор.Представление КАК РегистраторПредставление,
	                    ХозрасчетныйДвиженияССубконто.Регистратор.НазначениеПлатежа КАК РегистраторНазначениеПлатежа
                    ИЗ
	                    РегистрБухгалтерии.Хозрасчетный.ДвиженияССубконто(
		                    ДАТАВРЕМЯ(" + ssd + @"), 
		                    ДАТАВРЕМЯ(" + sed + @"), 
		                    ((Выразить(СчетДт.Код КАК СТРОКА(5)) = ""76.06"")
                    		    ИЛИ (Выразить(СчетКт.Код КАК СТРОКА(5)) = ""76.06""))) КАК ХозрасчетныйДвиженияССубконто
                    ";
                var РезультатЗапроса = Запрос.Выполнить();
                var ВыборкаИзРезультатаЗапроса = РезультатЗапроса.Выбрать();
                while (ВыборкаИзРезультатаЗапроса.Следующий())
                {
                    DateTime date = (DateTime)(ВыборкаИзРезультатаЗапроса.GetProperty("Период"));
                    Double amount = Convert.ToDouble(ВыборкаИзРезультатаЗапроса.GetProperty("Сумма"));
                    String daCode = ВыборкаИзРезультатаЗапроса.GetProperty("ДтСчетКод") as String;
                    String dpInn = ВыборкаИзРезультатаЗапроса.GetProperty("ДтКонтрагентИНН") as String;
                    String dpKpp = ВыборкаИзРезультатаЗапроса.GetProperty("ДтКонтрагентКПП") as String;
                    String dpDescr = ВыборкаИзРезультатаЗапроса.GetProperty("ДтКонтрагентНаименование") as String;
                    String dpDescrF = ВыборкаИзРезультатаЗапроса.GetProperty("ДтКонтрагентНаименованиеПолное") as String;
                    String caCode = ВыборкаИзРезультатаЗапроса.GetProperty("КтСчетКод") as String;
                    String cpInn = ВыборкаИзРезультатаЗапроса.GetProperty("КтКонтрагентИНН") as String;
                    String cpKpp = ВыборкаИзРезультатаЗапроса.GetProperty("КтКонтрагентКПП") as String;
                    String cpDescr = ВыборкаИзРезультатаЗапроса.GetProperty("КтКонтрагентНаименование") as String;
                    String cpDescrF = ВыборкаИзРезультатаЗапроса.GetProperty("КтКонтрагентНаименованиеПолное") as String;
                    String rDescr = ВыборкаИзРезультатаЗапроса.GetProperty("РегистраторПредставление") as String;
                    String rNote = ВыборкаИзРезультатаЗапроса.GetProperty("РегистраторНазначениеПлатежа") as String;

                    //String aCode = (new Regex(@"[^КкДда][№: ] ?(\d{18,20})")).Match((note ?? "").Trim()).Groups[1].Value; // КБК, КД, Код дохода
                    String trNum = (new Regex(@"(0\d{4}0\d{5}1\d0\d{5})")).Match((rNote ?? "").Trim()).Groups[1].Value;
                    
                    DataRow dr = payments.NewRow();
                    payments.Rows.Add(dr);
                    int ci = 0;
                    dr[ci++] = date;
                    dr[ci++] = amount;
                    dr[ci++] = (daCode ?? "").Trim();
                    dr[ci++] = (dpInn ?? "").Trim();
                    dr[ci++] = (dpKpp ?? "").Trim();
                    dr[ci++] = (dpDescr ?? "").Trim();
                    dr[ci++] = (dpDescrF ?? "").Trim();
                    dr[ci++] = (caCode ?? "").Trim();
                    dr[ci++] = (cpInn ?? "").Trim();
                    dr[ci++] = (cpKpp ?? "").Trim();
                    dr[ci++] = (cpDescr ?? "").Trim();
                    dr[ci++] = (cpDescrF ?? "").Trim();
                    dr[ci++] = (rDescr ?? "").Trim();
                    dr[ci++] = (rNote ?? "").Trim();
                    dr[ci++] = (trNum ?? "");
                    /*
                    Console.WriteLine(
                        ((DateTime)dr[0]).ToString("yyyy-MM-dd") + " | " +
                        String.Format("{0,10:n2}", (Double)dr[2]) + " | " +
                        (String)dr[1] + " | " +
                        (String)dr[7]
                        );
                    */
                }
            }
            catch (Exception e) { Console.WriteLine(e.ToString()); }
            return payments;
        }
    }
}
