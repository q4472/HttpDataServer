﻿using Nskd;
using Nskd.V77;
using System;
using System.Collections;
using System.Data;
using System.Text.RegularExpressions;

namespace HttpDataServerProject14
{
    public class OcStoredProcedure
    {
        private static GlobalContext V77gc;

        private static Object thisLock = new Object();

        public static ResponsePackage Exec1(RequestPackage rqp, String V77CnString)
        {
            ResponsePackage rsp = null;
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
                        case "ПолучитьСписокПриходныхНакладных":
                            rsp = ПолучитьСписокПриходныхНакладных();
                            break;
                        case "ДобавитьВ1СФКГарзаПриходнуюНакладную":
                            rsp = ДобавитьВ1СФКГарзаПриходнуюНакладную(rqp);
                            break;
                        case "ПолучитьДоговорПоКоду":
                            rsp = ПолучитьДоговорПоКоду(rqp);
                            break;
                        default:
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
                //var cust = GetByDescr("Клиенты", rqp["Владелец"] as String);
                var cust = GetByCode("Клиенты", rqp["ВладелецКод"] as String);
                if (cust != null)
                {
                    // Клиента нашли. Открываем справочник Договоры
                    var Договоры = new Справочник(V77gc.СоздатьОбъект("Справочник.Договора"));
                    // в нём работаем только с записями по Клиенту
                    Договоры.ИспользоватьВладельца(cust);

                    // Ищем папку с нужным годом
                    String year = DateTime.Now.Year.ToString();
                    String dd = OcConvert.ToD(rqp["ДатаДоговора"] as String);
                    if (dd.Length == 10)
                    {
                        year = dd.Substring(6, 4);
                    }
                    // 1 - поиск внутри установленного подчинения.
                    if (Договоры.НайтиПоНаименованию(year, 1) == 0)
                    {
                        // Папку с нужным годом не нашли. Создаём новую.
                        Договоры.НоваяГруппа();
                        Договоры.Наименование = year;
                        Договоры.Записать();
                    }

                    // Папка или найдена или только что создана. Берём её как родителя.
                    var parent = Договоры.ТекущийЭлемент();
                    Договоры.ИспользоватьРодителя(parent);

                    // Создаём новый элемент справочника Договоры
                    Договоры.Новый();

                    // Изменяем содержимое полей.
                    var empl = GetByDescr("Сотрудники", rqp["ОтветЛицо"] as String);
                    //var pres = GetByDescr("Клиенты", rqp["Представитель"] as String);
                    //var adda = GetByDescr("Договора", rqp["ДопСоглашение"] as String);

                    Договоры.УстановитьАтрибут("Владелец", cust);
                    Договоры.УстановитьАтрибут("Наименование", rqp["Наименование"] as String);
                    Договоры.УстановитьАтрибут("ДатаДоговора", OcConvert.ToD(rqp["ДатаДоговора"] as String));
                    Договоры.УстановитьАтрибут("ДатаОкончания", OcConvert.ToD(rqp["ДатаОкончания"] as String));
                    Договоры.УстановитьАтрибут("Пролонгация", rqp["Пролонгация"] as String);
                    Договоры.УстановитьАтрибут("ОтветЛицо", ((empl == null) ? null : empl));
                    Договоры.УстановитьАтрибут("СуммаДоговора", OcConvert.ToN(rqp["СуммаДоговора"] as String));
                    Договоры.УстановитьАтрибут("ОтсрочкаПлатежа", OcConvert.ToN(rqp["ОтсрочкаПлатежа"] as String));
                    //Договора.УстановитьАтрибут("Представитель", ((pres == null) ? null : pres));
                    //Договора.УстановитьАтрибут("ДопСоглашение", ((adda == null) ? null : adda));
                    Договоры.УстановитьАтрибут("Примечание", rqp["Примечание"] as String);
                    Договоры.УстановитьАтрибут("НомерТоргов", rqp["НомерТоргов"] as String);
                    Договоры.УстановитьАтрибут("ГосударственныйИдентификатор", rqp["ГосударственныйИдентификатор"] as String);
                    Договоры.УстановитьАтрибут("НовыйДоговор", 1);

                    // Сохранить.
                    Договоры.Записать();

                    code = Int32.Parse(Договоры.Код);
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
                var Расходная = new Документ(V77gc.СоздатьОбъект("Документ.Расходная"));
                if (Расходная.НайтиПоНомеру(docNo, Convert.ToDateTime(docDate)) == 1)
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
        private static ResponsePackage ПолучитьСписокПриходныхНакладных()
        {
            ResponsePackage rsp = new ResponsePackage();
            DataTable dt = new DataTable("СписокПриходныхНакладных");
            dt.Columns.Add("ДатаДок", typeof(DateTime));
            dt.Columns.Add("НомерДок", typeof(String));
            dt.Columns.Add("КлиентНаименование", typeof(String));
            dt.Columns.Add("НомерТН", typeof(String));
            try
            {
                var ТекстЗапроса = String.Format(@"
                    Без итогов;
                    Период с '{0}';
                    ОбрабатыватьДокументы Все; 
                    НомерДок = Документ.Приходная.НомерДок;
                    ДатаДок = Документ.Приходная.ДатаДок;
                    КлиентНаименование = Документ.Приходная.Клиент.Наименование;
                    НомерТН = Документ.Приходная.НомерТН;
                    Группировка НомерДок Без групп;
                    Условие(Найти(КлиентНаименование, ""{1}"") > 0);
                ", "01.05.2018", "Фарм-Сиб");
                if (V77gc.Запрос.Выполнить(ТекстЗапроса) == 1)
                {
                    Console.WriteLine("Запрос выполнен.");
                    while (V77gc.Запрос.Группировка() == 1)
                    {
                        DataRow dr = dt.NewRow();
                        dt.Rows.Add(dr);
                        dr["ДатаДок"] = V77gc.Запрос.ПолучитьАтрибут("ДатаДок");
                        dr["НомерДок"] = ((String)V77gc.Запрос.ПолучитьАтрибут("НомерДок")).Trim();
                        dr["КлиентНаименование"] = ((String)V77gc.Запрос.ПолучитьАтрибут("КлиентНаименование")).Trim();
                        dr["НомерТН"] = ((String)V77gc.Запрос.ПолучитьАтрибут("НомерТН")).Trim();
                    }
                    if (dt.Rows.Count > 0)
                    {
                        rsp.Data = new DataSet();
                        rsp.Data.Tables.Add(dt);
                    }
                }
            }
            catch (Exception e) { Console.WriteLine(e); }
            return rsp;
        }
        private static ResponsePackage ДобавитьВ1СФКГарзаПриходнуюНакладную(RequestPackage rqp)
        {
            ResponsePackage rsp = new ResponsePackage
            {
                Status = "HttpDataServer.HttpDataServerProject14_1v7.OneCv77SP.HttpDataServerProject14.OcStoredProcedure.ДобавитьВ1СФКГарзаПриходнуюНакладную()"
            };
            if (rqp != null)
            {
                var json = rqp["РасходнаяНакладная"] as String;
                if (!String.IsNullOrWhiteSpace(json))
                {
                    //rsp.Status += "\n" + json;
                    Hashtable РасходнаяНакладная = null;
                    try
                    {
                        РасходнаяНакладная = Nskd.JsonV3.Parse(json) as Hashtable;
                    }
                    catch (Exception) { }
                    if (РасходнаяНакладная != null)
                    {
                        var ДокументПриходная = new Документ(V77gc.СоздатьОбъект("Документ.Приходная"));
                        ДокументПриходная.Новый();
                        V77Object ТипДоставки;
                        V77Object Маркетолог;
                        // шапка
                        {
                            Маркетолог = null;
                            {
                                if (РасходнаяНакладная.ContainsKey("Маркетолог"))
                                {
                                    var Сотрудники = new Справочник(V77gc.СоздатьОбъект("Справочник.Сотрудники"));
                                    if (Сотрудники.НайтиПоНаименованию(РасходнаяНакладная["НомерСФ"] as String, 0, 1) == 1)
                                    {
                                        Маркетолог = Сотрудники.ТекущийЭлемент();
                                    }
                                    Сотрудники.Dispose(); Сотрудники = null;
                                }
                            }
                            V77Object Фирма = null;
                            {
                                var Фирмы = new Справочник(V77gc.СоздатьОбъект("Справочник.Фирмы"));
                                if (Фирмы.НайтиПоНаименованию("ООО \"ФК ГАРЗА\"", 0, 1) == 1) Фирма = Фирмы.ТекущийЭлемент();
                                Фирмы.Dispose(); Фирмы = null;
                                if (Фирма != null)
                                {
                                    ДокументПриходная.УстановитьАтрибут("Фирма", Фирма);
                                    Фирма.Dispose(); Фирма = null;
                                }
                            }
                            V77Object Склад = null;
                            {
                                var Склады = new Справочник(V77gc.СоздатьОбъект("Справочник.Склады"));
                                if (Склады.НайтиПоНаименованию("Главный склад", 0, 1) == 1) Склад = Склады.ТекущийЭлемент();
                                Склады.Dispose(); Склады = null;
                                if (Склад != null)
                                {
                                    ДокументПриходная.УстановитьАтрибут("Склад", Склад);
                                    Склад.Dispose(); Склад = null;
                                }
                            }
                            V77Object Клиент = null;
                            {
                                var Клиенты = new Справочник(V77gc.СоздатьОбъект("Справочник.Клиенты"));
                                if (Клиенты.НайтиПоНаименованию("ООО \"Фарм-Сиб\"", 0, 1) == 1) Клиент = Клиенты.ТекущийЭлемент();
                                Клиенты.Dispose(); Клиенты = null;
                                if (Клиент != null)
                                {
                                    ДокументПриходная.УстановитьАтрибут("Клиент", Клиент);
                                    Клиент.Dispose(); Клиент = null;
                                }
                            }
                            V77Object Валюта = null;
                            {
                                var Валюты = new Справочник(V77gc.СоздатьОбъект("Справочник.Валюты"));
                                if (Валюты.НайтиПоНаименованию("Рубль", 0, 1) == 1) Валюта = Валюты.ТекущийЭлемент();
                                Валюты.Dispose(); Валюты = null;
                                if (Валюта != null)
                                {
                                    ДокументПриходная.УстановитьАтрибут("Валюта", Валюта);
                                    Валюта.Dispose(); Валюта = null;
                                }
                            }
                            ДокументПриходная.УстановитьАтрибут("Курс", 1.0);

                            var ПризнПрихНакл = V77gc.Перечисление.ПолучитьАтрибут("ПризнПрихНакл");
                            {
                                var Закупка = ПризнПрихНакл.ЗначениеПоИдентификатору("Закупка");
                                ДокументПриходная.УстановитьАтрибут("ПризнакНакладной", Закупка);
                                Закупка.Dispose(); Закупка = null;
                                ПризнПрихНакл.Dispose(); ПризнПрихНакл = null;
                            }
                            var Налоги = V77gc.Перечисление.ПолучитьАтрибут("Налоги");
                            {
                                var НДССверхуБезНП = Налоги.ЗначениеПоИдентификатору("НДССверхуБезНП");
                                ДокументПриходная.УстановитьАтрибут("Налоги", НДССверхуБезНП);
                                НДССверхуБезНП.Dispose(); НДССверхуБезНП = null;
                                Налоги.Dispose(); Налоги = null;
                            }
                            ТипДоставки = null;
                            {
                                var ТипыДоставок = new Справочник(V77gc.СоздатьОбъект("Справочник.ТипыДоставок"));
                                if (ТипыДоставок.НайтиПоНаименованию("транзит", 0, 1) == 1) ТипДоставки = ТипыДоставок.ТекущийЭлемент();
                                ТипыДоставок.Dispose(); ТипыДоставок = null;
                                if (ТипДоставки != null)
                                {
                                    ДокументПриходная.УстановитьАтрибут("ТипДоставкиШ", ТипДоставки);
                                    //ТипДоставки.Dispose(); ТипДоставки = null;
                                }
                            }
                            if (РасходнаяНакладная.ContainsKey("НомерДок"))
                            {
                                ДокументПриходная.УстановитьАтрибут("НомерТН", РасходнаяНакладная["НомерДок"]);
                            }

                            if (РасходнаяНакладная.ContainsKey("НомерСФ"))
                            {
                                ДокументПриходная.УстановитьАтрибут("НомерСФ", РасходнаяНакладная["НомерСФ"]);
                            }

                            if (РасходнаяНакладная.ContainsKey("ДатаДок"))
                            {
                                ДокументПриходная.УстановитьАтрибут("ДатаНакладной", РасходнаяНакладная["ДатаДок"]);
                            }

                            if (РасходнаяНакладная.ContainsKey("ДатаСФ"))
                            {
                                ДокументПриходная.УстановитьАтрибут("ДатаСФ", РасходнаяНакладная["ДатаСФ"]);
                            }
                        }
                        // табличная часть
                        if (РасходнаяНакладная.Contains("ТабличнаяЧасть"))
                        {
                            if (РасходнаяНакладная["ТабличнаяЧасть"] is DataTable ТабличнаяЧасть && ТабличнаяЧасть.Rows.Count > 0)
                            {
                                var Товары = new Справочник(V77gc.СоздатьОбъект("Справочник.Товары"));
                                var Лаборатория = new Справочник(V77gc.СоздатьОбъект("Справочник.Клиенты"));
                                var Сотрудники = new Справочник(V77gc.СоздатьОбъект("Справочник.Сотрудники"));
                                foreach (DataRow dr in ТабличнаяЧасть.Rows)
                                {
                                    if (Товары.НайтиПоНаименованию((String)dr["Товар"], 0) == 1)
                                    {
                                        if (ТабличнаяЧасть.Columns.Contains("Серия") && dr["Серия"] is Hashtable Серия)
                                        {
                                            try
                                            {
                                                System.Globalization.CultureInfo ic = System.Globalization.CultureInfo.InvariantCulture;
                                                ДокументПриходная.НоваяСтрока();
                                                ДокументПриходная.УстановитьАтрибут("Товар", Товары.ТекущийЭлемент());
                                                if (ТабличнаяЧасть.Columns.Contains("Количество") && dr["Количество"] != DBNull.Value)
                                                {
                                                    ДокументПриходная.УстановитьАтрибут("Количество", Convert.ToDouble(dr["Количество"], ic));
                                                }
                                                if (Серия.ContainsKey("Наименование"))
                                                {
                                                    ДокументПриходная.УстановитьАтрибут("СерияСтр", Серия["Наименование"]);
                                                }
                                                if (ТабличнаяЧасть.Columns.Contains("Цена") && dr["Цена"] != DBNull.Value)
                                                {
                                                    ДокументПриходная.УстановитьАтрибут("Цена", Convert.ToDouble(dr["Цена"], ic));
                                                }
                                                if (ТабличнаяЧасть.Columns.Contains("Сумма") && dr["Сумма"] != DBNull.Value)
                                                {
                                                    ДокументПриходная.УстановитьАтрибут("Сумма", Convert.ToDouble(dr["Сумма"], ic));
                                                }
                                                if (ТабличнаяЧасть.Columns.Contains("СуммаНП") && dr["СуммаНП"] != DBNull.Value)
                                                {
                                                    ДокументПриходная.УстановитьАтрибут("СуммаНП", Convert.ToDouble(dr["СуммаНП"], ic));
                                                }
                                                if (ТабличнаяЧасть.Columns.Contains("Всего") && dr["Всего"] != DBNull.Value)
                                                {
                                                    ДокументПриходная.УстановитьАтрибут("Всего", Convert.ToDouble(dr["Всего"], ic));
                                                }

                                                if (Серия.ContainsKey("ГоденДо"))
                                                {
                                                    ДокументПриходная.УстановитьАтрибут("ГоденДо", Серия["ГоденДо"]);
                                                }
                                                if (Серия.ContainsKey("КолЦУ"))
                                                {
                                                    ДокументПриходная.УстановитьАтрибут("КолЦУ", Convert.ToDouble(Серия["КолЦУ"], ic));
                                                }
                                                if (Серия.ContainsKey("ЦенаИзг"))
                                                {
                                                    ДокументПриходная.УстановитьАтрибут("ЦенаИзг", Convert.ToDouble(Серия["ЦенаИзг"], ic));
                                                }
                                                if (Серия.ContainsKey("Лаборатория"))
                                                {
                                                    if (Лаборатория.НайтиПоНаименованию(Серия["Лаборатория"] as String, 0) == 1)
                                                    {
                                                        ДокументПриходная.УстановитьАтрибут("Лаборатория", Лаборатория.ТекущийЭлемент());
                                                    }
                                                }
                                                if (Серия.ContainsKey("СрокДействия"))
                                                {
                                                    ДокументПриходная.УстановитьАтрибут("СрокДействия", Серия["СрокДействия"]);
                                                }
                                                if (Серия.ContainsKey("НомерСерт"))
                                                {
                                                    ДокументПриходная.УстановитьАтрибут("НомерСерт", Серия["НомерСерт"]);
                                                }
                                                if (Серия.ContainsKey("ГТД"))
                                                {
                                                    ДокументПриходная.УстановитьАтрибут("ГТД", Серия["ГТД"]);
                                                }
                                                if (Серия.ContainsKey("Адрес"))
                                                {
                                                    ДокументПриходная.УстановитьАтрибут("Адрес", Серия["Адрес"]);
                                                }
                                                if (Серия.ContainsKey("Фасовка"))
                                                {
                                                    ДокументПриходная.УстановитьАтрибут("Фасовка", Convert.ToDouble(Серия["Фасовка"], ic));
                                                }
                                                if (Серия.ContainsKey("Масса"))
                                                {
                                                    ДокументПриходная.УстановитьАтрибут("Масса", Convert.ToDouble(Серия["Масса"], ic));
                                                }
                                                if (Серия.ContainsKey("Объем"))
                                                {
                                                    ДокументПриходная.УстановитьАтрибут("Объем", Convert.ToDouble(Серия["Объем"], ic));
                                                }
                                                if (Серия.ContainsKey("СрокГодности"))
                                                {
                                                    ДокументПриходная.УстановитьАтрибут("ГоденДоМес", Convert.ToDouble(Серия["СрокГодности"], ic));
                                                }
                                                if (Серия.ContainsKey("Декларант"))
                                                {
                                                    ДокументПриходная.УстановитьАтрибут("Декларант", Серия["Декларант"]);
                                                }
                                                ДокументПриходная.УстановитьАтрибут("ТипДоставки", ТипДоставки);
                                                if (ТабличнаяЧасть.Columns.Contains("ЦенаРеестраРуб") && dr["ЦенаРеестраРуб"] != DBNull.Value)
                                                {
                                                    ДокументПриходная.УстановитьАтрибут("ЦенаРеестраРуб", Convert.ToDouble(dr["ЦенаРеестраРуб"], ic));
                                                }
                                                if (Серия.ContainsKey("Закупщик"))
                                                {
                                                    if (Сотрудники.НайтиПоНаименованию(Серия["Закупщик"] as String, 0) == 1)
                                                    {
                                                        ДокументПриходная.УстановитьАтрибут("Закупщик", Сотрудники.ТекущийЭлемент());
                                                    }
                                                }
                                                if (Серия.ContainsKey("Закупщик2"))
                                                {
                                                    if (Сотрудники.НайтиПоНаименованию(Серия["Закупщик2"] as String, 0) == 1)
                                                    {
                                                        ДокументПриходная.УстановитьАтрибут("Закупщик2", Сотрудники.ТекущийЭлемент());
                                                    }
                                                }
                                                if (Маркетолог != null)
                                                {
                                                    ДокументПриходная.УстановитьАтрибут("КтоЗаказал", Маркетолог);
                                                }
                                            }
                                            catch (Exception e) { Console.WriteLine(e); }
                                        }
                                    }
                                    else { rsp.Status += String.Format("Не нашли в Гарза.Справочник.Товары: '{0}'", Товары.Наименование); }
                                }
                                Сотрудники.Dispose(); Сотрудники = null;
                                Лаборатория.Dispose(); Лаборатория = null;
                                Товары.Dispose(); Товары = null;
                            }
                        }
                        Маркетолог?.Dispose(); Маркетолог = null;
                        ТипДоставки?.Dispose(); ТипДоставки = null;

                        ДокументПриходная.Записать();
                        ДокументПриходная.Dispose(); ДокументПриходная = null;
                    }
                }
                else
                {
                    rsp.Status += "\nРасходнаяНакладная не получена.";
                }
            }
            return rsp;
        }
        private static ResponsePackage ПолучитьДоговорПоКоду(RequestPackage rqp)
        {
            ResponsePackage rsp = new ResponsePackage();
            DataTable dt = new DataTable("Договоры");
            dt.Columns.Add("ДатаОкончания", typeof(DateTime));
            dt.Columns.Add("Пролонгация", typeof(String));
            dt.Columns.Add("ОтсрочкаПлатежа", typeof(Double));
            try
            {
                var Код = rqp["Код"] as String;
                if (!String.IsNullOrWhiteSpace(Код))
                {
                    var Договоры = new Справочник(V77gc.СоздатьОбъект("Справочник.Договора"));
                    if (Договоры.НайтиПоКоду(Код) == 1)
                    {
                        DataRow dr = dt.NewRow();
                        dt.Rows.Add(dr);
                        dr["ДатаОкончания"] = Договоры.ПолучитьАтрибут("ДатаОкончания");
                        dr["Пролонгация"] = Договоры.ПолучитьАтрибут("Пролонгация");
                        dr["ОтсрочкаПлатежа"] = Договоры.ПолучитьАтрибут("ОтсрочкаПлатежа");
                    }
                    Договоры.Dispose(); Договоры = null;
                    if (dt.Rows.Count > 0)
                    {
                        rsp.Data = new DataSet();
                        rsp.Data.Tables.Add(dt);
                    }
                }
            }
            catch (Exception e) { Console.WriteLine(e); }
            return rsp;
        }
    }
}
