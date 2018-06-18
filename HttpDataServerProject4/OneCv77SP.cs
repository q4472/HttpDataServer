using Nskd;
using Nskd.V77;
using System;
using System.Data;
using System.Text.RegularExpressions;

namespace HttpDataServerProject4
{
    public class OcStoredProcedure
    {
        private static GlobalContext V77gc;

        private static Object thisLock = new Object();

        public static ResponsePackage Exec1(RequestPackage rqp, String V77CnString)
        {
            ResponsePackage rsp = new ResponsePackage();
            lock (thisLock)
            {
                if (V77gc == null) { V77gc = new GlobalContext(V77CnString); }
                if (V77gc != null && V77gc.ComObject != null)
                {
                    switch (rqp.Command)
                    {
                        case "Добавить":
                            rsp = Agrs.F0Add(rqp);
                            break;
                        case "Обновить":
                            rsp = Agrs.F0Update(rqp);
                            break;
                        case "Удалить":
                            rsp = Agrs.F0Delete(rqp);
                            break;
                        case "Docs1c/F0/Save":
                            rsp = Docs1c.Save(rqp);
                            break;
                        case "[dbo].[oc_клиенты_select_1]":
                            rsp = F1GetCustTable(rqp);
                            break;
                        case "[dbo].[oc_сотрудники_select_1]":
                            rsp = F1GetStuffTable(rqp);
                            break;
                        case "ПолучитьСписокРасходныхНакладных":
                            rsp = ПолучитьСписокРасходныхНакладных();
                            break;
                        case "ПолучитьИз1СФармСибРасходнуюНакладную":
                            rsp = ПолучитьИз1СФармСибРасходнуюНакладную(rqp);
                            break;
                        default:
                            rsp.Status = String.Format($"Не найдена команда '{rqp.Command}'.");
                            break;
                    }
                }
            }
            return rsp;
        }
        public class Agrs
        {
            public static ResponsePackage F0Add(RequestPackage rqp)
            {
                ResponsePackage rsp = new ResponsePackage();
                Int32 code = -1;
                // Договора нет. Надо создать новую запись.
                // Вставка нового элемента в справочник Договоры.
                // сначала ищем Клиента так как Договоры - подчинённый справочник 
                /*
                try
                {
                    Console.WriteLine((rqp["ВладелецКод"] as String) + ": " + (rqp["Владелец"] as String));
                    GetByCode(root, "Клиенты", rqp["ВладелецКод"] as String);
                }
                catch (Exception e) { Console.WriteLine(e.ToString()); }
                */
                //V77Object cust = GetByDescr(root, "Клиенты", rqp["Владелец"] as String);
                var cust = GetByCode("Клиенты", rqp["ВладелецКод"] as String);
                if (cust != null)
                {
                    // Клиента нашли. Открываем справочник Договоры
                    var Договора = new Справочник(V77gc.СоздатьОбъект("Справочник.Договора"));
                    // в нём работаем только с записями по Клиенту
                    Договора.ИспользоватьВладельца(cust);

                    // Ищем папку с нужным годом
                    String year = DateTime.Now.Year.ToString();
                    String dd = OcConvert.ToD(rqp["ДатаДоговора"] as String);
                    if (dd.Length == 10)
                    {
                        year = dd.Substring(6, 4);
                    }
                    // 1 - поиск внутри установленного подчинения.
                    if (Договора.НайтиПоНаименованию(year, 1) == 0)
                    {
                        // Папку с нужным годом не нашли. Создаём новую.
                        Договора.НоваяГруппа();
                        Договора.Наименование = year;
                        Договора.Записать();
                    }

                    // Папка или найдена или только что создана. Берём её как родителя.
                    var parent = Договора.ТекущийЭлемент();
                    Договора.ИспользоватьРодителя(parent);

                    // Создаём новый элемент справочника Договоры
                    Договора.Новый();

                    // Изменяем содержимое полей.
                    var empl = GetByDescr("Сотрудники", rqp["ОтветЛицо"] as String);
                    var pres = GetByDescr("Клиенты", rqp["Представитель"] as String);
                    var adda = GetByDescr("Договора", rqp["ДопСоглашение"] as String);

                    Договора.УстановитьАтрибут("Владелец", cust);
                    Договора.УстановитьАтрибут("Наименование", rqp["Наименование"] as String);
                    Договора.УстановитьАтрибут("ДатаДоговора", OcConvert.ToD(rqp["ДатаДоговора"] as String));
                    Договора.УстановитьАтрибут("ДатаОкончания", OcConvert.ToD(rqp["ДатаОкончания"] as String));
                    Договора.УстановитьАтрибут("Пролонгация", rqp["Пролонгация"] as String);
                    Договора.УстановитьАтрибут("ОтветЛицо", ((empl == null) ? null : empl));
                    Договора.УстановитьАтрибут("СуммаДоговора", OcConvert.ToN(rqp["СуммаДоговора"] as String));
                    Договора.УстановитьАтрибут("ОтсрочкаПлатежа", OcConvert.ToN(rqp["ОтсрочкаПлатежа"] as String));
                    Договора.УстановитьАтрибут("Представитель", ((pres == null) ? null : pres));
                    Договора.УстановитьАтрибут("ДопСоглашение", ((adda == null) ? null : adda));
                    Договора.УстановитьАтрибут("Примечание", rqp["Примечание"] as String);
                    Договора.УстановитьАтрибут("НомерТоргов", rqp["НомерТоргов"] as String);
                    Договора.УстановитьАтрибут("ГосударственныйИдентификатор", rqp["ГосударственныйИдентификатор"] as String);
                    Договора.УстановитьАтрибут("НовыйДоговор", 1);

                    // Сохранить.
                    Договора.Записать();

                    code = Int32.Parse(Договора.Код);
                }
                else
                {
                    Log.Write("В 1С.Сравочник.Клиенты не наден элемент: " + (rqp["Владелец"] as String));
                }
                rsp.Data = new DataSet();
                rsp.Data.Tables.Add(new DataTable());
                rsp.Data.Tables[0].Columns.Add("code", typeof(Int32));
                rsp.Data.Tables[0].Rows.Add(rsp.Data.Tables[0].NewRow());
                rsp.Data.Tables[0].Rows[0][0] = code;
                return rsp;
            }
            public static ResponsePackage F0Update(RequestPackage rqp)
            {
                ResponsePackage rsp = new ResponsePackage();
                Int32 code = -1;
                if (Int32.TryParse(rqp["Код"] as String, out code))
                {
                    // Пробуем найти Договор.
                    var Договора = new Справочник(V77gc.СоздатьОбъект("Справочник.Договора"));
                    if (Договора.НайтиПоКоду(code.ToString()) == 1)
                    {
                        if (Договора.Выбран() == 1) // Договор найден.
                        {
                            var empl = GetByDescr("Сотрудники", rqp["ОтветЛицо"] as String);
                            var pres = GetByDescr("Клиенты", rqp["Представитель"] as String);
                            var adda = GetByDescr("Договора", rqp["ДопСоглашение"] as String);

                            Договора.УстановитьАтрибут("Наименование", rqp["Наименование"] as String);
                            Договора.УстановитьАтрибут("ДатаДоговора", OcConvert.ToD(rqp["ДатаДоговора"] as String));
                            Договора.УстановитьАтрибут("ДатаОкончания", OcConvert.ToD(rqp["ДатаОкончания"] as String));
                            Договора.УстановитьАтрибут("Пролонгация", rqp["Пролонгация"] as String);
                            Договора.УстановитьАтрибут("ОтветЛицо", ((empl == null) ? null : empl));
                            Договора.УстановитьАтрибут("СуммаДоговора", OcConvert.ToN(rqp["СуммаДоговора"] as String));
                            Договора.УстановитьАтрибут("ОтсрочкаПлатежа", OcConvert.ToN(rqp["ОтсрочкаПлатежа"] as String));
                            Договора.УстановитьАтрибут("Представитель", ((pres == null) ? null : pres));
                            Договора.УстановитьАтрибут("ДопСоглашение", ((adda == null) ? null : adda));
                            Договора.УстановитьАтрибут("Примечание", rqp["Примечание"] as String);
                            Договора.УстановитьАтрибут("НомерТоргов", rqp["НомерТоргов"] as String);
                            Договора.УстановитьАтрибут("ГосударственныйИдентификатор", rqp["ГосударственныйИдентификатор"] as String);

                            Договора.Записать();

                            code = Int32.Parse(Договора.Код);
                        }
                    }
                }
                else
                {
                    Log.Write("Ошибка: не правильный код для 1С.Справочник.Договора: " + String.Format("{0:s}", rqp["Код"] as String));
                }
                rsp.Data = new DataSet();
                rsp.Data.Tables.Add(new DataTable());
                rsp.Data.Tables[0].Columns.Add("code", typeof(Int32));
                rsp.Data.Tables[0].Rows.Add(rsp.Data.Tables[0].NewRow());
                rsp.Data.Tables[0].Rows[0][0] = code;
                return rsp;
            }
            public static ResponsePackage F0Delete(RequestPackage rqp)
            {
                ResponsePackage rsp = new ResponsePackage();
                Int32 code = -1;
                if (Int32.TryParse(rqp["Код"] as String, out code))
                {
                    var Договора = new Справочник(V77gc.СоздатьОбъект("Справочник.Договора"));
                    if (Договора.НайтиПоКоду(code.ToString()) == 1)
                    {
                        Договора.Удалить(0); // 0 - пометка на удаление
                    }
                }
                else
                {
                    Log.Write("Ошибка: не правильный код для 1С.Справочник.Договора: " + String.Format("{0:s}", rqp["Код"] as String));
                }
                rsp.Data = new DataSet();
                rsp.Data.Tables.Add(new DataTable());
                rsp.Data.Tables[0].Columns.Add("code", typeof(Int32));
                rsp.Data.Tables[0].Rows.Add(rsp.Data.Tables[0].NewRow());
                rsp.Data.Tables[0].Rows[0][0] = code;
                return rsp;
            }
        }
        public class Docs1c
        {
            public static ResponsePackage Save(RequestPackage rqp)
            {
                ResponsePackage rsp = new ResponsePackage();
                Int32 code = -1;
                String docDate = null;
                String docNo = null;
                String track = null;
                String sDate = null;
                String rDate = null;
                if ((rqp != null) && (rqp.Parameters != null))
                {
                    foreach (var p in rqp.Parameters)
                    {
                        if (p != null)
                        {
                            if (p.Name == "f1") { docDate = p.Value as String; }
                            if (p.Name == "f2") { docNo = p.Value as String; }
                            if (p.Name == "f4") { track = p.Value as String; }
                            if (p.Name == "f5") { sDate = p.Value as String; }
                            if (p.Name == "f6") { rDate = p.Value as String; }
                        }
                    }
                }
                Log.Write(String.Format("'{0}', '{1}', '{2}', '{3}', '{4}'", docDate, docNo, track, sDate, rDate));
                // Пробуем найти Расходную.
                Документ Расходная = null; // new Документ(V77gc.СоздатьОбъект("Документ.Расходная"));
                var Расходные = new Документ(V77gc.СоздатьОбъект("Документ.Расходная"));
                if (Расходные.ВыбратьДокументы(new DateTime(2018, 4, 1), new DateTime(2018, 12, 31)) == 1)
                {
                    while (Расходные.ПолучитьДокумент() == 1)
                    {
                        if (Расходные.НомерДок.Trim() == docNo.Trim().Replace("\"", ""))
                        {
                            Расходная = Расходные.ТекущийДокумент();

                        }
                    }
                }
                if (Расходная != null) //.НайтиПоНомеру(docNo, Convert.ToDateTime(docDate)) == 1)
                {
                    // Расходная найдена.
                    Console.WriteLine("Расходная найдена.");
                    // Справочник Рыссылки
                    var Рассылки = new Справочник(V77gc.СоздатьОбъект("Справочник.Рассылки"));
                    if (Рассылки.НайтиПоРеквизиту("РасходнаяНакладная", Расходная, 1) == 0)
                    {
                        // Рассылка не найдена.
                        Console.WriteLine("Рассылка не найдена.");
                        Рассылки.Новый();
                        var rncd = Расходная.ТекущийДокумент();
                        Рассылки.УстановитьАтрибут("РасходнаяНакладная", rncd);
                    }
                    // Рассылка найдена или только-что создана.
                    Console.WriteLine("Рассылка найдена или только-что создана.");
                    Рассылки.УстановитьАтрибут("Наименование", track);
                    DateTime d;
                    if ((!String.IsNullOrWhiteSpace(sDate)) && (DateTime.TryParse(sDate, out d)))
                    {
                        Рассылки.УстановитьАтрибут("ДатаОтправкиДокументов", d);
                    }
                    else
                    {
                        Рассылки.УстановитьАтрибут("ДатаОтправкиДокументов", null);
                    }
                    if ((!String.IsNullOrWhiteSpace(rDate)) && (DateTime.TryParse(rDate, out d)))
                    {
                        Рассылки.УстановитьАтрибут("ДатаПолученияДокументов", d);
                    }
                    else
                    {
                        Рассылки.УстановитьАтрибут("ДатаПолученияДокументов", null);
                    }
                    Рассылки.Записать();
                }
                rsp.Data = new DataSet();
                rsp.Data.Tables.Add(new DataTable());
                rsp.Data.Tables[0].Columns.Add("code", typeof(Int32));
                rsp.Data.Tables[0].Rows.Add(rsp.Data.Tables[0].NewRow());
                rsp.Data.Tables[0].Rows[0][0] = code;
                return rsp;
            }
        }
        private static Справочник GetByCode(String name, String code)
        {
            Справочник obj = null;
            if (!String.IsNullOrWhiteSpace(name))
            {
                var справочник = new Справочник(V77gc.СоздатьОбъект("Справочник." + name));
                // режим 0 - поиск во всём справочнике
                if (справочник.НайтиПоКоду(code, 0) == 1)
                {
                    obj = справочник.ТекущийЭлемент();
                }
                справочник.Dispose(); справочник = null;
            }
            return obj;
        }
        private static Справочник GetByDescr(String name, String descr)
        {
            Справочник obj = null;
            if (!String.IsNullOrWhiteSpace(name))
            {
                var справочник = new Справочник(V77gc.СоздатьОбъект("Справочник." + name));
                // режим 0 - поиск во всём справочнике
                if (справочник.НайтиПоНаименованию(descr, 0) == 1)
                {
                    obj = справочник.ТекущийЭлемент();
                }
                справочник.Dispose(); справочник = null;
            }
            return obj;
        }
        private static ResponsePackage F1GetCustTable(RequestPackage rqp)
        {
            ResponsePackage rsp = new ResponsePackage();
            DataTable dt = new DataTable();
            dt.Columns.Add("Код", typeof(String));
            dt.Columns.Add("Наименование", typeof(String));
            dt.Columns.Add("ИНН", typeof(String));
            dt.Columns.Add("Город", typeof(String));
            var filter = rqp["DESCR"] as String;
            if (!String.IsNullOrWhiteSpace(filter))
            {
                Regex re = new Regex(filter, RegexOptions.IgnoreCase);
                try
                {
                    var Клиенты = new Справочник(V77gc.СоздатьОбъект("Справочник.Клиенты"));
                    try
                    {
                        if (Клиенты.ВыбратьЭлементы() == 1)
                        {
                            Int32 cnt = 0;
                            while (Клиенты.ПолучитьЭлемент() == 1 && cnt++ < 10000)
                            {
                                var Наименование = Клиенты.Наименование.Trim();
                                var ИНН = ((String)Клиенты.ПолучитьАтрибут("ИНН")).Trim();
                                if (re.IsMatch(Наименование) || re.IsMatch(ИНН))
                                {
                                    DataRow dr = dt.NewRow();
                                    dt.Rows.Add(dr);
                                    dr["Код"] = Клиенты.Код.Trim();
                                    dr["Наименование"] = Наименование;
                                    dr["ИНН"] = ((String)Клиенты.GetProperty("ИНН")).Trim();
                                    dr["Город"] = ((String)Клиенты.GetProperty("Город")).Trim();
                                }
                            }
                        }
                    }
                    catch (Exception) { }
                    finally { if (Клиенты != null) { Клиенты.Dispose(); Клиенты = null; } }
                }
                catch (Exception) { }
                //finally { if (V77Garza != null) { V77Garza.Dispose(); V77Garza = null; } }
            }
            rsp.Data = new DataSet();
            rsp.Data.Tables.Add(dt);
            return rsp;
        }
        private static ResponsePackage F1GetStuffTable(RequestPackage rqp)
        {
            ResponsePackage rsp = new ResponsePackage();
            DataTable dt = new DataTable();
            dt.Columns.Add("Код", typeof(String));
            dt.Columns.Add("Наименование", typeof(String));
            var filter = rqp["DESCR"] as String;
            if (!String.IsNullOrWhiteSpace(filter))
            {
                Regex re = new Regex(filter, RegexOptions.IgnoreCase);
                try
                {
                    var Сотрудники = new Справочник(V77gc.СоздатьОбъект("Справочник.Сотрудники"));
                    try
                    {
                        if (Сотрудники.ВыбратьЭлементы() == 1)
                        {
                            Int32 cnt = 0;
                            while (Сотрудники.ПолучитьЭлемент() == 1 && cnt++ < 10000)
                            {
                                var Наименование = Сотрудники.Наименование.Trim();
                                if (re.IsMatch(Наименование))
                                {
                                    DataRow dr = dt.NewRow();
                                    dt.Rows.Add(dr);
                                    dr["Код"] = Сотрудники.Код.Trim();
                                    dr["Наименование"] = Наименование;
                                }
                            }
                        }
                    }
                    catch (Exception) { }
                    finally { if (Сотрудники != null) { Сотрудники.Dispose(); Сотрудники = null; } }
                }
                catch (Exception) { }
                //finally { if (V77Garza != null) { V77Garza.Dispose(); V77Garza = null; } }
            }
            rsp.Data = new DataSet();
            rsp.Data.Tables.Add(dt);
            return rsp;
        }
        private static ResponsePackage ПолучитьСписокРасходныхНакладных()
        {
            ResponsePackage rsp = new ResponsePackage();
            DataTable dt = new DataTable("СписокРасходныхНакладных");
            dt.Columns.Add("ДатаДок", typeof(DateTime));
            dt.Columns.Add("НомерДок", typeof(String));
            dt.Columns.Add("КлиентНаименование", typeof(String));
            try
            {
                var ТекстЗапроса = String.Format(@"
                    Без итогов;
                    Период с '{0}';
                    НомерДок = Документ.Расходная.НомерДок;
                    ДатаДок = Документ.Расходная.ДатаДок;
                    КлиентНаименование = Документ.Расходная.Клиент.Наименование;
                    Группировка НомерДок Без групп;
                    Условие(Найти(КлиентНаименование, ""{1}"") > 0);
                ", "01.05.2018", "ФК ГАРЗА");
                if (V77gc != null && V77gc.Запрос.Выполнить(ТекстЗапроса) == 1)
                {
                    //Console.WriteLine("Запрос выполнен.");
                    while (V77gc.Запрос.Группировка() == 1)
                    {
                        DataRow dr = dt.NewRow();
                        dt.Rows.Add(dr);
                        dr["ДатаДок"] = V77gc.Запрос.ПолучитьАтрибут("ДатаДок");
                        dr["НомерДок"] = ((String)V77gc.Запрос.ПолучитьАтрибут("НомерДок")).Trim();
                        dr["КлиентНаименование"] = ((String)V77gc.Запрос.ПолучитьАтрибут("КлиентНаименование")).Trim();
                    }
                }
            }
            catch (Exception e) { Console.WriteLine(e); }
            rsp.Data = new DataSet();
            rsp.Data.Tables.Add(dt);
            return rsp;
        }
        private static ResponsePackage ПолучитьИз1СФармСибРасходнуюНакладную(RequestPackage rqp)
        {
            ResponsePackage rsp = new ResponsePackage
            {
                Status = "HttpDataServer.HttpDataServerProject04_1c7.OneCv77SP.HttpDataServerProject4.OcStoredProcedure.ПолучитьИз1СФармСибРасходнуюНакладную()"
            };

            String fsN = rqp["fsN"] as String;
            if (!String.IsNullOrWhiteSpace(fsN))
            {
                rsp.Status += "\n" + fsN;

                String НомерНакладной = fsN.Replace("\"", "");

                V77Расходная v77Расходная = null;

                var Расходная = НайтиРасходнуюНакладнуюПоНомеру(НомерНакладной);
                if (Расходная != null)
                {
                    Документ СчетФактура = null;
                    var Документы = new Документ(V77gc.СоздатьОбъект("Документ"));
                    if (Документы.ВыбратьПодчиненныеДокументы(null, null, Расходная) == 1)
                    {
                        while (Документы.ПолучитьДокумент() == 1)
                        {
                            СчетФактура = Документы.ТекущийДокумент();
                            if (СчетФактура.Вид() == "Счет_фактура")
                            {
                                break;
                            }
                            else { СчетФактура = null; }
                        }
                    }
                    Документы.Dispose(); Документы = null;

                    v77Расходная = new V77Расходная
                    {
                        НомерДок = Расходная.НомерДок.Trim(),
                        ДатаДок = Расходная.ДатаДок,
                    };

                    if (СчетФактура != null)
                    {
                        v77Расходная.НомерСФ = СчетФактура.НомерДок.Trim();
                        v77Расходная.ДатаСФ = СчетФактура.ДатаДок;
                        СчетФактура.Dispose(); СчетФактура = null;
                    }

                    var РеестрЦенНаЖНЛВП = new Справочник(V77gc.СоздатьОбъект("Справочник.РеестрЦенНаЖНЛВП"));
                    var КоличествоСтрок = Convert.ToInt32(Расходная.КоличествоСтрок());
                    Console.WriteLine("Количество строк табличной части: '{0}'", КоличествоСтрок);
                    if (КоличествоСтрок > 0)
                    {
                        for (int Номер = 1; Номер <= КоличествоСтрок; Номер++)
                        {
                            Расходная.ПолучитьСтрокуПоНомеру(Номер);
                            var ТоварСсылка = new Справочник(Расходная.ПолучитьАтрибут("Товар"));
                            var СерияСсылка = new Справочник(Расходная.ПолучитьАтрибут("Серия"));
                            {
                                var dr = v77Расходная.ТабличнаяЧасть.NewRow();
                                dr["Товар"] = ТоварСсылка.Наименование.Trim();
                                if (РеестрЦенНаЖНЛВП.НайтиПоРеквизиту("Товар", ТоварСсылка, 1) == 1)
                                {
                                    var ЦенаРуб = new Справочник(РеестрЦенНаЖНЛВП.ПолучитьАтрибут("ЦенаРуб"));
                                    dr["ЦенаРеестраРуб"] = (Double)ЦенаРуб.Получить();
                                    ЦенаРуб.Dispose(); ЦенаРуб = null;
                                }
                                dr["Серия"] = new V77Серия(СерияСсылка);
                                dr["Количество"] = (Double)Расходная.ПолучитьАтрибут("Количество");
                                dr["Цена"] = (Double)Расходная.ПолучитьАтрибут("Цена");
                                dr["Сумма"] = (Double)Расходная.ПолучитьАтрибут("Сумма");
                                dr["НДС"] = (Double)Расходная.ПолучитьАтрибут("НДС");
                                dr["СуммаЧистая"] = (Double)Расходная.ПолучитьАтрибут("СуммаЧистая");
                                dr["СуммаНП"] = (Double)Расходная.ПолучитьАтрибут("СуммаНП");
                                dr["Всего"] = (Double)Расходная.ПолучитьАтрибут("Всего");
                                v77Расходная.ТабличнаяЧасть.Rows.Add(dr);
                            }
                            СерияСсылка.Dispose(); СерияСсылка = null;
                            ТоварСсылка.Dispose(); ТоварСсылка = null;
                        }
                    }
                    РеестрЦенНаЖНЛВП.Dispose(); РеестрЦенНаЖНЛВП = null;
                    Расходная.Dispose(); Расходная = null;
                }

                rsp.Data = new DataSet();
                DataTable dt = rsp.Data.Tables.Add();
                dt.Columns.Add("json", typeof(String));
                dt.Rows.Add(JsonV2.ToString(v77Расходная));
                Console.WriteLine(dt.Rows[0][0]);
            }
            else { rsp.Status += "\nНе задан номер расходной накладной."; }
            return rsp;
        }
        /// <summary>Ищет расходную накладную по номеру за последние 3 месяца.</summary>
        private static Документ НайтиРасходнуюНакладнуюПоНомеру(String Номер)
        {
            Документ Расходная = null;
            String ТекстЗапроса = String.Format(@"
                        Без итогов;
                        Период с '{0}';
                        ДокументРасходная = Документ.Расходная.ТекущийДокумент;
                        НомерДок = Документ.Расходная.НомерДок;
                        Условие(НомерДок = ""{1}"");
                        Группировка ДокументРасходная без групп;
                        ", DateTime.Now.AddMonths(-3).ToString("dd.MM.yyyy"), Номер);
            try
            {
                if (V77gc.Запрос.Выполнить(ТекстЗапроса) == 1 && V77gc.Запрос.Группировка() == 1)
                {
                    Расходная = new Документ(V77gc.Запрос.ПолучитьАтрибут("ДокументРасходная"));
                }
            }
            catch (Exception e) { Log.Write(e.Message); }
            return Расходная;
        }
    }
    class V77Серия
    {
        public String Код;
        public String Наименование;
        public DateTime ГоденДо;
        public Double КолЦУ;
        //public Double Розн_Цена;
        //public Double Прих_Цена;
        //public Double Прих_ЦенаСНалогами;
        public Double ЦенаИзг;
        //public Справочник Валюта; // Валюты
        public String Лаборатория; // Справочник.Клиенты.Наименование
        public DateTime СрокДействия;
        public String НомерСерт;
        public String ГТД;
        //public Double Коммисия;
        //public Перечисление ЗаВалюту;
        //public Справочник Поставщик; // Клиенты
        //public Справочник ДоговорПоставки; // Договора
        //public Документ Поставка;
        public String Адрес;
        public Double Масса;
        public Double Объем;
        public Double СрокГодности;
        public Double Фасовка;
        public String Декларант;
        //public Справочник Покупатель; // Сотрудники
        //public Double КрМаркЦена; // Периодический
        //public Double Прих_ЦенаСоСкидкой;
        //public Double ЦенаПрайс; // Периодический
        //public String ТЦСКЛС;
        //public DateTime ДатаРегистрации;
        public String Закупщик; // Справочник.Сотрудники.Наименование
        //public Справочник ДоговорПокупателя; // Договора
        //public DateTime ДатаИзмененияПокупателя;
        //public String ВремяИзмененияПокупателя;
        //public Справочник АвторИзмененияПокупателя; // Сотрудники
        //public String ИсторияИзмененияАдресов;
        public String Закупщик2; // Справочник.Сотрудники.Наименование
        //public Справочник ТипДоставки; // ТипыДоставок
        public V77Серия(Справочник СерияСсылка)
        {
            Справочник sc = null;
            Код = СерияСсылка.Код.Trim();
            Наименование = СерияСсылка.Наименование.Trim();
            ГоденДо = (DateTime)СерияСсылка.ПолучитьАтрибут("ГоденДо");
            КолЦУ = (Double)СерияСсылка.ПолучитьАтрибут("КолЦУ");
            ЦенаИзг = (Double)СерияСсылка.ПолучитьАтрибут("ЦенаИзг");
            sc = new Справочник(СерияСсылка.ПолучитьАтрибут("Лаборатория")); Лаборатория = sc.Наименование.Trim(); sc.Dispose(); sc = null;
            СрокДействия = (DateTime)СерияСсылка.ПолучитьАтрибут("СрокДействия");
            НомерСерт = ((String)СерияСсылка.ПолучитьАтрибут("НомерСерт")).Trim();
            ГТД = ((String)СерияСсылка.ПолучитьАтрибут("ГТД")).Trim();
            Адрес = ((String)СерияСсылка.ПолучитьАтрибут("Адрес")).Trim();
            Фасовка = (Double)СерияСсылка.ПолучитьАтрибут("Фасовка");
            Масса = (Double)СерияСсылка.ПолучитьАтрибут("Масса");
            Объем = (Double)СерияСсылка.ПолучитьАтрибут("Объем");
            СрокГодности = (Double)СерияСсылка.ПолучитьАтрибут("СрокГодности");
            Декларант = ((String)СерияСсылка.ПолучитьАтрибут("Декларант")).Trim();
            sc = new Справочник(СерияСсылка.ПолучитьАтрибут("Закупщик")); Закупщик = sc.Наименование.Trim(); sc.Dispose(); sc = null;
            sc = new Справочник(СерияСсылка.ПолучитьАтрибут("Закупщик2")); Закупщик2 = sc.Наименование.Trim(); sc.Dispose(); sc = null;
        }
    }
    class V77Расходная
    {
        public String НомерДок;
        public DateTime ДатаДок;
        public String НомерСФ;
        public DateTime ДатаСФ;
        public DataTable ТабличнаяЧасть;
        public V77Расходная()
        {
            ТабличнаяЧасть = new DataTable();
            ТабличнаяЧасть.Columns.Add("Товар", typeof(String)); // Справочник.Товары.Наименование
            ТабличнаяЧасть.Columns.Add("ЦенаРеестраРуб", typeof(Double));
            ТабличнаяЧасть.Columns.Add("Серия", typeof(V77Серия));
            ТабличнаяЧасть.Columns.Add("Количество", typeof(Double));
            ТабличнаяЧасть.Columns.Add("Цена", typeof(Double));
            ТабличнаяЧасть.Columns.Add("Сумма", typeof(Double));
            ТабличнаяЧасть.Columns.Add("НДС", typeof(Double));
            ТабличнаяЧасть.Columns.Add("СуммаЧистая", typeof(Double));
            ТабличнаяЧасть.Columns.Add("СуммаНП", typeof(Double));
            ТабличнаяЧасть.Columns.Add("Всего", typeof(Double));
        }
    }
}
