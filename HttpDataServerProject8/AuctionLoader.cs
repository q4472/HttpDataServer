using Nskd;
using System;
using System.Collections.Generic;
using System.Data;
using System.Data.SqlClient;
using System.IO;
using System.Net;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading;
using System.Web;

namespace HttpDataServerProject8
{
    class AuctionLoader
    {
        public static ResponsePackage Load(RequestPackage rqp)
        {
            ResponsePackage rsp = new ResponsePackage();

            Boolean overwrite = false;
            Object temp = rqp["overwrite"];
            if (temp != null)
            {
                if (temp.GetType() == typeof(Boolean)) { overwrite = (Boolean)temp; }
                if (temp.GetType() == typeof(String)) { Boolean.TryParse((String)temp, out overwrite); }
            }

            String auctionNumber = rqp["auction_number"] as String;

            Guid? auctionUid = null;
            if (!String.IsNullOrWhiteSpace(auctionNumber))
            {
                auctionUid = GetAuctionUid(auctionNumber);

                if (auctionUid == null || overwrite)
                {
                    auctionUid = LoadAndSave(auctionNumber);
                }
            }
            else { rsp.Status = "Не задан номер аукциона"; }

            if (auctionUid != null)
            {
                rsp.Status = "ОК";
                rsp.Data = new DataSet();
                rsp.Data.Tables.Add();
                rsp.Data.Tables[0].Columns.Add(new DataColumn("auction_uid", typeof(Guid)));
                rsp.Data.Tables[0].Rows.Add(new object[] { auctionUid });
            }
            else { rsp.Status = "Не удалось найти или загрузить информацию об аукционе."; }

            return rsp;
        }
        private static Guid? GetAuctionUid(String auctionNumber)
        {
            Guid? auctionUid = null;
            SqlCommand cmd = new SqlCommand
            {
                Connection = new SqlConnection(Env.cnString),
                CommandText = "[Auctions].[dbo].[get_auction_uid]",
                CommandType = CommandType.StoredProcedure
            };
            cmd.Parameters.AddWithValue("@auction_number", auctionNumber);
            using (cmd.Connection)
            {
                Object o = null;
                cmd.Connection.Open();
                o = cmd.ExecuteScalar();
                if (o != null && o.GetType() == typeof(Guid))
                {
                    auctionUid = (Guid)o;
                }
            }
            return auctionUid;
        }
        private static Guid? LoadAndSave(String auctionNumber)
        {
            Guid? auctionUid = null;
            try
            {
                if (!String.IsNullOrWhiteSpace(auctionNumber) && (auctionNumber.Length == 11 || auctionNumber.Length == 19))
                {
                    String html = ZakupkiGovRu.GetAuctionInf(auctionNumber);
                    if (!String.IsNullOrWhiteSpace(html))
                    {
                        if (auctionNumber.Length == 11) // 223 фз
                        {
                            auctionUid = Fz223.SaveAuctionInf(html);
                        }
                        if (auctionNumber.Length == 19) // 44 фз
                        {
                            auctionUid = Fz44.SaveAuctionInf(html);
                        }
                    }
                    else { Log.Write(String.Format($"Информация о заявке '{auctionNumber}' не загрузилась.")); }
                }
                else { Log.Write(String.Format($"Не указан номер заявки.")); }
            }
            catch (Exception e) { Log.Write(String.Format(e.Message)); }
            return auctionUid;
        }
    }
    class ZakupkiGovRu
    {
        public static String GetAuctionInf(String auctionNumber)
        {
            String html = null;
            if (!String.IsNullOrWhiteSpace(auctionNumber) && auctionNumber.Length == 19)
            {
                UriBuilder ub = new UriBuilder
                {
                    Scheme = "http",
                    Host = "zakupki.gov.ru",
                    Path = "/epz/order/notice/ea44/view/common-info.html",
                    Query = String.Format("regNumber={0}", auctionNumber)
                };
                html = Utilities.GetResponse(ub.Uri);
            }
            if (!String.IsNullOrWhiteSpace(auctionNumber) && auctionNumber.Length == 11)
            {
                UriBuilder ub = new UriBuilder
                {
                    Scheme = "http",
                    Host = "zakupki.gov.ru",
                    Path = "/223/purchase/public/purchase/info/common-info.html",
                    Query = String.Format("regNumber={0}", auctionNumber)
                };
                html = Utilities.GetResponse(ub.Uri);
            }
            return html;
        }
        private static Int32 GetHtmlDivFinishIndex(String inputString, Int32 htmlDivStartIndex)
        {
            Int32 htmlDivFinishIndex = -1;
            if (!String.IsNullOrWhiteSpace(inputString) && inputString.Length > htmlDivStartIndex)
            {
                Int32 currentIndex = htmlDivStartIndex + 4;
                Boolean br = false;
                while (!br && currentIndex < inputString.Length)
                {
                    Int32 nextDivIndex = inputString.IndexOf("div", currentIndex);
                    if (nextDivIndex > htmlDivStartIndex)
                    {
                        switch (inputString[nextDivIndex - 1])
                        {
                            case '<':
                                currentIndex = GetHtmlDivFinishIndex(inputString, nextDivIndex - 1);
                                break;
                            case '/':
                                if (inputString[nextDivIndex - 2] == '<')
                                {
                                    htmlDivFinishIndex = Math.Min(nextDivIndex + 4, inputString.Length);
                                    br = true;
                                }
                                break;
                            default:
                                break;
                        }
                    }
                }
            }
            return htmlDivFinishIndex;
        }
        private static String BuildQueryStringForRegion(String regionNumber, String publishDate, String searchString, Regex re, Int32 recordsPerPage = 50, Int32 pageNumber = 1)
        {
            publishDate = String.Format("{0}.{1}.20{2}", publishDate.Substring(0, 2), publishDate.Substring(3, 2), publishDate.Substring(6, 2));
            StringBuilder sb = new StringBuilder();
            sb.AppendFormat("searchString={0}", HttpUtility.UrlEncode(searchString));
            sb.AppendFormat("&pageNumber={0}", pageNumber);
            sb.Append("&sortDirection=false");
            sb.AppendFormat("&recordsPerPage=_{0}", recordsPerPage);
            sb.Append("&showLotsInfoHidden=false");
            if (re.ToString().Contains("19")) { sb.Append("&fz44=on"); }
            if (re.ToString().Contains("11")) { sb.Append("&fz223=on"); }
            sb.Append("&af=on");
            sb.Append("&ca=on");
            //sb.Append("&pc=on"); // закупка завершена
            sb.Append("&priceFrom=");
            sb.Append("&priceTo=");
            sb.Append("&currencyId=1");
            sb.Append("&agencyTitle=");
            sb.Append("&agencyCode=");
            sb.Append("&agencyFz94id=");
            sb.Append("&agencyFz223id=");
            sb.Append("&agencyInn=");
            sb.AppendFormat("&region_regions_{0}=region_regions_{0}", regionNumber);
            sb.AppendFormat("&regions={0}", regionNumber);
            sb.AppendFormat("&publishDateFrom={0}", publishDate);
            sb.AppendFormat("&publishDateTo={0}", publishDate);
            sb.Append("&sortBy=UPDATE_DATE");
            sb.Append("&updateDateFrom=");
            sb.Append("&updateDateTo=");
            return sb.ToString();
        }
        public static ResponsePackage LoadAuctionNumbers(RequestPackage rqp)
        {
            ResponsePackage rsp = new ResponsePackage();
            rsp.Status = "Start HttpDataServerProject8.ZakupkiGovRu.LoadAuctionNumbers().";
            if (rqp != null && rqp.Parameters != null)
            {
                String regionNumber = rqp["region_number"] as String;
                String publishDate = rqp["publish_date"] as String;
                rsp.Status += String.Format(" Region_number: '{0}' publish_date: '{1}'", regionNumber, publishDate);
                if (!String.IsNullOrWhiteSpace(regionNumber) && !String.IsNullOrWhiteSpace(publishDate))
                {
                    rsp.Status += " Creating new DataSet.";
                    rsp.Data = new DataSet();
                    rsp.Data.Tables.Add();
                    rsp.Data.Tables[0].Columns.Add("auction_number", typeof(String));

                    List<String> auctionNumbers = new List<string>();

                    Regex re44 = new Regex(@"№ (\d{19})\D?");
                    Regex re223 = new Regex(@"№ (\d{11})\D?");
                    UriBuilder ub = new UriBuilder
                    {
                        Scheme = "http",
                        Host = "zakupki.gov.ru",
                        Path = "/epz/order/quicksearch/search_eis.html"
                    };
                    String[] ss = new String[] { "лекарств", "препарат", "медикамент" };
                    Int32 recordsPerPage = 50;
                    Int32 pageCount = 1;
                    foreach (Regex re in new Regex[] { re44, re223 })
                    {
                        foreach (String s in ss)
                        {
                            Int32 pageNumber = 0;
                            while (pageNumber++ < pageCount)
                            {
                                ub.Query = BuildQueryStringForRegion(regionNumber, publishDate, s, re, recordsPerPage, pageNumber);
                                String receivedString = Utilities.GetResponse(ub.Uri);

                                if (!String.IsNullOrWhiteSpace(receivedString))
                                {
                                    // найти количество записей в ответе
                                    Int32 recordCount = GetRecordCount(receivedString);
                                    if (pageNumber == 1)
                                    {
                                        pageCount = (recordCount / recordsPerPage) + 1;
                                    }

                                    MatchCollection ms = re.Matches(receivedString);
                                    foreach (Match m in ms)
                                    {
                                        if (m.Groups.Count > 1)
                                        {
                                            String value = m.Groups[1].Value;
                                            if (!String.IsNullOrWhiteSpace(value))
                                            {
                                                if (!auctionNumbers.Contains(value))
                                                {
                                                    auctionNumbers.Add(value);
                                                    rsp.Data.Tables[0].Rows.Add(new object[] { value });
                                                }
                                            }
                                        }
                                    }
                                }
                            }
                        }
                    }
                }
            }
            return rsp;
        }
        private static Int32 GetRecordCount(String receivedString)
        {
            Int32 recordCount = 0;
            Int32 sIndex = receivedString.IndexOf("Всего записей");
            if (sIndex >= 0)
            {
                String temp = receivedString.Substring(sIndex, 100);
                Regex re = new Regex(@"\d+");
                Match match = re.Match(temp);
                Int32.TryParse(match.Value, out recordCount);
            }
            return recordCount;
        }
    }
    class Fz44
    {
        public static Guid SaveAuctionInf(String html)
        {
            Guid aUid = new Guid();

            // записываем на SQL сервер в базу Auctions общую информацию о заявке
            aUid = ParseAndSaveAuction44fzCommonInf(html);

            // записываем на SQL сервер в базу Auctions информацию о заказчиках (их может быть несколько)
            ParseAndSaveAuction44FzCustomerRequirement(aUid, html);

            return aUid;
        }
        private static Guid ParseAndSaveAuction44fzCommonInf(String html)
        {
            Guid aUid = new Guid();

            // берём основные блоки

            // noticeTabBox содержит пары <h2 class="noticeBoxH2"> и <div class="noticeTabBoxWrapper"> с общей информацией для всех заказчиков
            // <h2> заголовк общей информации
            // <div> содержимое общей информации

            // заявка
            Int32 h0 = new SectionIndexes(html, new String[] { "cardHeader", ">" }).Index1;
            // общая_информация_о_закупке
            Int32 contentTabBoxBlock = new SectionIndexes(html, new String[] { "contentTabBoxBlock", ">" }, h0).Index1;
            // общая_информация_о_закупке
            Int32 t0 = new SectionIndexes(html, new String[] { "Общая информация", "<table", ">" }, contentTabBoxBlock).Index1;
            // информация_об_организации_осуществляющей_определение_поставщика
            Int32 t1 = new SectionIndexes(html, new String[] { "Информация об организации", "<table", ">" }, contentTabBoxBlock).Index1;
            // информация_о_процедуре_закупки
            Int32 t2 = new SectionIndexes(html, new String[] { "Информация о процедуре", "<table", ">" }, contentTabBoxBlock).Index1;
            // начальная_максимальная_цена_контракта
            Int32 t3 = new SectionIndexes(html, new String[] { "Начальная (максимальная)", "<table", ">" }, contentTabBoxBlock).Index1;
            // информация_об_объекте_закупки
            Int32 t4 = new SectionIndexes(html, new String[] { "Информация об объекте", "<table", ">" }, contentTabBoxBlock).Index1;
            // преимущества_требования_к_участникам
            Int32 t5 = new SectionIndexes(html, new String[] { "Преимущества, требования", "<table", ">" }, contentTabBoxBlock).Index1;

            Object[][] md44 = new Object[][] { 
                // закупка
                new Object[] { "номер",                                         h0, new String[] { "Закупка", "№" },                                        @"(\d*)" },
                new Object[] { "дата_размещения",                               h0, new String[] { "Размещено:" },                                          @"(.*?)</div>" },
                new Object[] { "кооператив",                                    h0, new String[] { "cooperative" },                                         null, "1" },
                // общая_информация_о_закупке
                new Object[] { "способ_определения_поставщика",                 t0, new String[] { "Способ определения поставщика", "<td", ">" },           @"(.*?)</td>" },
                new Object[] { "наименование_электронной_площадки_в_интернете", t0, new String[] { "Наименование электронной площадки", "<td", ">" },       @"(.*?)</td>" },
                new Object[] { "адрес_электронной_площадки_в_интернете",        t0, new String[] { "Адрес электронной площадки", "<td", ">", "<a", ">" },   @"(.*?)</a>" },
                new Object[] { "размещение_осуществляет",                       t0, new String[] { "Размещение осуществляет", "<td", ">" },                 @"(.*?)</td>" },
                new Object[] { "объект_закупки",                                t0, new String[] { "Наименование объекта закупки", "<td", ">" },            @"(.*?)</td>" },
                new Object[] { "этап_закупки",                                  t0, new String[] { "Этап закупки", "<td", ">" },                            @"(.*?)</td>" },
                new Object[] { "сведения_о_связи_с_позицией_плана_графика",     t0, new String[] { "Сведения о связи", "<td", ">" },                        @"(.*?)</td>" },
                new Object[] { "номер_типового_контракта",                      t0, new String[] { "Номер типового контракта", "<td", ">" },                @"(.*?)</td>" },
                // информация_об_организации_осуществляющей_определение_поставщика
                new Object[] { "организация_осуществляющая_размещение",         t1, new String[] { "Организация, осуществляющая размещение", "<td", ">" },  @"(.*?)</td>" },
                new Object[] { "почтовый_адрес",                                t1, new String[] { "Почтовый адрес", "<td", ">" },                          @"(.*?)</td>" },
                new Object[] { "место_нахождения",                              t1, new String[] { "Место нахождения", "<td", ">" },                        @"(.*?)</td>" },
                new Object[] { "ответственное_должностное_лицо",                t1, new String[] { "Ответственное должностное лицо", "<td", ">" },          @"(.*?)</td>" },
                new Object[] { "адрес_электронной_почты",                       t1, new String[] { "Адрес электронной почты", "<td", ">" },                 @"(.*?)</td>" },
                new Object[] { "номер_контактного_телефона",                    t1, new String[] { "Номер контактного телефона", "<td", ">" },              @"(.*?)</td>" },
                new Object[] { "факс",                                          t1, new String[] { "Факс", "<td", ">" },                                    @"(.*?)</td>" },
                new Object[] { "дополнительная_информация",                     t1, new String[] { "Дополнительная информация", "<td", ">" },               @"(.*?)</td>" },
                // информация_о_процедуре_закупки
                new Object[] { "дата_и_время_начала_подачи_заявок",             t2, new String[] { "Дата и время начала подачи", "<td", ">" },              @"(.*?)</td>" },
                new Object[] { "дата_и_время_окончания_подачи_заявок",          t2, new String[] { "Дата и время окончания подачи", "<td", ">" },           @"(.*?)</td>" },
                new Object[] { "место_подачи_заявок",                           t2, new String[] { "Место подачи заявок", "<td", ">" },                     @"(.*?)</td>" },
                new Object[] { "порядок_подачи_заявок",                         t2, new String[] { "Порядок подачи заявок", "<td", ">" },                   @"(.*?)</td>" },
                new Object[] { "дата_окончания_срока_рассмотрения_первых_частей", t2, new String[] { "Дата окончания срока", "<td", ">" },                  @"(.*?)</td>" },
                new Object[] { "дата_проведения_аукциона_в_электронной_форме",  t2, new String[] { "Дата проведения", "<td", ">" },                         @"(.*?)</td>" },
                new Object[] { "время_проведения_аукциона",                     t2, new String[] { "Время проведения", "<td", ">" },                        @"(.*?)</td>" },
                new Object[] { "дополнительная_информация2",                    t2, new String[] { "Дополнительная информация", "<td", ">" },               @"(.*?)</td>" },
                // начальная_максимальная_цена_контракта
                new Object[] { "начальная_максимальная_цена_контракта",         t3, new String[] { "Начальная (максимальная) цена", "<td", ">" },           @"(.*?)</td>" },
                new Object[] { "валюта",                                        t3, new String[] { "Валюта", "<td", ">" },                                  @"(.*?)</td>" },
                new Object[] { "источник_финансирования",                       t3, new String[] { "Источник финансирования", "<td", ">" },                 @"(.*?)</td>" },
                new Object[] { "идентификационный_код_закупки",                 t3, new String[] { "Идентификационный код закупки", "<td", ">" },           @"(.*?)</td>" },
                new Object[] { "оплата_исполнения_контракта_по_годам",          t3, new String[] { "Оплата исполнения контракта", "<td", ">" },             @"(.*?)</td>" }, // @"../div[contains(@class, 'addingTbl')]/table//tr[1]/td[2]"
                // информация_об_объекте_закупки
                new Object[] { "описание_объекта_закупки",                      t4, new String[] { "Описание объекта закупки" },                            @"(.*?)</td>" },
                new Object[] { "условия_запреты_и_ограничения_допуска_товаров", t4, new String[] { "Условия", "запреты", "ограничения", "<td", ">" },       @"(.*?)</td>" },
                new Object[] { "табличная_часть_в_формате_html",                t4, new String[] { "Табличная часть", "html", "<td", ">" },                 @"(.*?)</td>" },
                // преимущества_требования_к_участникам
                new Object[] { "преимущества",                                  t5, new String[] { "Преимущества", "<td", ">" },                            @"(.*?)</td>" },
                new Object[] { "требования",                                    t5, new String[] { "Требования", "<td", ">" },                              @"(.*?)</td>" },
                new Object[] { "ограничения",                                   t5, new String[] { "Ограничения", "<td", ">" },                             @"(.*?)</td>" }
            };

            SqlCommand cmd = new SqlCommand
            {
                Connection = new SqlConnection(Env.cnString),
                CommandText = "[Auctions].[dbo].[save_auction_inf]",
                CommandType = CommandType.StoredProcedure
            };
            foreach (Object[] p in md44)
            {
                Object v;
                Int32 index = new SectionIndexes(html, (String[])p[2], (Int32)p[1]).Index1;
                if (index >= 0)
                {
                    if (p[3] is String r) v = Utilities.GetValueByRegex(html, r, index); else v = p[4];
                    if (v is String) v = Utilities.NormString(v as String);
                    if (v != null) cmd.Parameters.AddWithValue((String)p[0], v);
                }
            }
            using (cmd.Connection)
            {
                Object o = null;
                cmd.Connection.Open();
                o = cmd.ExecuteScalar();
                if (o != null && o.GetType() == typeof(Guid)) { aUid = (Guid)o; }
            }
            return aUid;
        }
        private static void ParseAndSaveAuction44FzCustomerRequirement(Guid aUid, String html)
        {
            // здесь может быть два варианта
            // первый - когда списки expand пусты - сохраняем из основного блока
            // второй - сохраняем из списков
            // кажется теперь только второй вариант 2018-06-20

            // второй вариант
            // baseIndex

            Object[][] mdCust = new Object[][] {
                // требования_заказчика
                new Object[] { "наименование_заказчика",                    0, new String[] { "&nbsp;" },  "(.*?)</h3>" },
                // условия_контракта
                new Object[] { "место_доставки_товара",                     0, new String[] { "Место доставки", "<td", ">" },                       "(.*?)</td>" },
                new Object[] { "сроки_поставки_товара",                     0, new String[] { "Сроки поставки", "<td", ">" },                       "(.*?)</td>" },
                new Object[] { "сведения_о_связи_с_позицией_плана_графика", 0, new String[] { "Сведения о связи", "Сведения о связи", "<td", ">" }, "(.*?)</td>" },
                new Object[] { "оплата_исполнения_контракта_по_годам",      0, new String[] { "Оплата исполнения", "<td", ">" },                    "(.*?)</td>" },
                // обеспечение_заявок
                new Object[] { "размер_обеспечения",                        0, new String[] { "Обеспечение заявок", "Размер обеспечения", "<td", ">" },  "(.*?)</td>" },
                new Object[] { "порядок_внесения_денежных_средств",         0, new String[] { "Обеспечение заявок", "Порядок внесения", "<td", ">" },    "(.*?)</td>" },
                new Object[] { "платежные_реквизиты",                       0, new String[] { "Обеспечение заявок", "Платежные реквизиты", "<td", ">" }, "(.*?)</td>" },
                // обеспечение_исполнения_контракта
                new Object[] { "размер_обеспечения_2",                      0, new String[] { "Обеспечение исполнения", "Размер обеспечения", "<td", ">" },     "(.*?)</td>" },
                new Object[] { "порядок_предоставления_обеспечения",        0, new String[] { "Обеспечение исполнения", "Порядок предоставления", "<td", ">" }, "(.*?)</td>" },
                new Object[] { "платежные_реквизиты_2",                     0, new String[] { "Обеспечение исполнения", "Платежные реквизиты", "<td", ">" },    "(.*?)</td>" },
            };

            Int32 bi = new SectionIndexes(html, new String[] { "Требования заказчиков" }).Index1;

            bi = new SectionIndexes(html, new String[] { "Требования заказчика" }, bi).Index1;
            while (bi >= 0)
            {
                SqlCommand cmd = new SqlCommand
                {
                    Connection = new SqlConnection(Env.cnString),
                    CommandText = "[Auctions].[dbo].[save_customer_requirement]",
                    CommandType = CommandType.StoredProcedure
                };
                cmd.Parameters.AddWithValue("закупка_uid", aUid);
                foreach (Object[] p in mdCust)
                {
                    Object v;
                    Int32 index = new SectionIndexes(html, (String[])p[2], bi).Index1;
                    if (index >= 0)
                    {
                        if (p[3] is String r) v = Utilities.GetValueByRegex(html, r, index); else v = p[4];
                        if (v is String) v = Utilities.NormString(v as String);
                        if (v != null) cmd.Parameters.AddWithValue((String)p[0], v);
                    }
                }
                using (cmd.Connection)
                {
                    Object o = null;
                    cmd.Connection.Open();
                    o = cmd.ExecuteScalar();
                    if (o != null && o.GetType() == typeof(Guid)) { aUid = (Guid)o; }
                }
                bi = new SectionIndexes(html, new String[] { "Требования заказчика" }, bi).Index1;
            }
        }
    }
    class Fz223
    {
        public static Guid SaveAuctionInf(String html)
        {
            Guid aUid = new Guid();
            Regex re = new Regex("<td[^>]*?>(.*?)</td>", RegexOptions.Singleline);
            Object[][] md223 = new Object[][] { 
                // Закупка
                new Object[] { "номер", "Общие сведения о закупке", "Номер извещения" },
                new Object[] { "дата_размещения", "Общие сведения о закупке", "Дата размещения извещения" },
                //new Object[] { "кооператив", null, "0"}, // это значение по умолчанию
                // Общие сведения о закупке
                new Object[] { "номер_извещения", "Общие сведения о закупке", "Номер извещения" },
                new Object[] { "способ_размещения_закупки", "Общие сведения о закупке", "Способ размещения закупки" },
                new Object[] { "наименование_закупки", "Общие сведения о закупке", "Наименование закупки" },
                new Object[] { "закупка_осуществляется_вследствие_аварии", "Общие сведения о закупке", "Закупка осуществляется вследствие аварии" },
                new Object[] { "редакция", "Общие сведения о закупке", "Редакция" },
                new Object[] { "наименование_электронной_площадки", "Общие сведения о закупке", "Наименование электронной площадки" },
                new Object[] { "адрес_электронной_площадки", "Общие сведения о закупке", "Адрес электронной площадки" },
                new Object[] { "дата_размещения_извещения", "Общие сведения о закупке", "Дата размещения извещения" },
                new Object[] { "дата_размещения_текущей_редакции_извещения", "Общие сведения о закупке", "Дата размещения текущей редакции" },
                // Заказчик
                new Object[] { "наименование_организации", "Заказчик", "Наименование организации" },
                new Object[] { "инн_кпп", "Заказчик", "ИНН" },
                new Object[] { "огрн", "Заказчик", "ОГРН" },
                new Object[] { "адрес_места_нахождения", "Заказчик", "Адрес места нахождения" },
                new Object[] { "почтовый_адрес", "Заказчик", "Почтовый адрес" },
                // Контактное лицо
                new Object[] { "организация", "Контактное лицо", "Организация" },
                new Object[] { "контактное_лицо", "Контактное лицо", "Контактное лицо" },
                new Object[] { "электронная_почта", "Контактное лицо", "Электронная почта" },
                new Object[] { "телефон", "Контактное лицо", "Телефон" },
                new Object[] { "факс", "Контактное лицо", "Факс" },
                // Требования к участникам закупки
                new Object[] { "требование_к_отсутствию_участников_закупки_в_реестре_недобросовестных_поставщиков", "Требования к участникам закупки", "Требование к отсутствию" },
                // Порядок размещения закупки
                new Object[] { "дата_и_время_окончания_подачи_заявок", "Порядок размещения закупки", "Дата и время окончания подачи" },
                new Object[] { "дата_окончания_срока_рассмотрения_заявок", "Порядок размещения закупки", "Дата окончания срока рассмотрения" },
                new Object[] { "дата_и_время_проведение_аукциона", "Порядок размещения закупки", "Дата и время проведения" },
                // Предоставление документации
                new Object[] { "срок_предоставления", "Предоставление документации", "Срок" },
                new Object[] { "место_предоставления", "Предоставление документации", "Место" },
                new Object[] { "порядок_предоставления", "Предоставление документации", "Порядок" },
                new Object[] { "официальный_сайт", "Предоставление документации", "Официальный сайт" },
                new Object[] { "внесение_платы_за_предоставление_конкурсной_документации", "Предоставление документации", "Внесение платы" }
                };
            SqlCommand cmd = new SqlCommand
            {
                Connection = new SqlConnection(Env.cnString),
                CommandText = "[Auctions].[dbo].[save_auction223_inf]",
                CommandType = CommandType.StoredProcedure
            };
            foreach (Object[] mdRow in md223)
            {
                String dbName = mdRow[0] as String;
                String blockName = mdRow[1] as String;
                String keyText = mdRow[2] as String;

                Int32 blockIndex = new SectionIndexes(html, new String[] { blockName }).Index1;
                if (blockIndex >= 0)
                {
                    Int32 keyIndex = new SectionIndexes(html, new String[] { keyText }).Index1;
                    if (keyIndex >= 0)
                    {
                        Match match = re.Match(html, keyIndex);
                        String value = null;
                        if (match.Success && match.Groups.Count > 1)
                        {
                            value = match.Groups[1].Value.Trim();
                            value = Utilities.NormString(value);
                            cmd.Parameters.AddWithValue(dbName, value);
                        }
                    }
                }
            }
            
            using (cmd.Connection)
            {
                Object o = null;
                cmd.Connection.Open();
                o = cmd.ExecuteScalar();
                if (o != null && o.GetType() == typeof(Guid)) { aUid = (Guid)o; }
            }
            
            return aUid;
        }
    }
    class Env
    {
        public static String cnString = String.Format(@"Data Source={0};Initial Catalog=master;Integrated Security=True", Program.MainSqlServerDataSource);
    }
    class Utilities
    {
        public static String NormString(String s)
        {
            String r = s;
            if (!String.IsNullOrWhiteSpace(r))
            {
                r = r.Replace("&nbsp;", " ").Replace("&laquo;", "«").Replace("&raquo;", "»").Replace("&#034;", "\"").Replace("&#34;", "\"");
                r = r.Replace("\t", " ").Replace("\n", " ").Replace("\r", " ");
                while (r.Contains("  ")) r = r.Replace("  ", " ");
                r = r.Trim();
            }
            return r;
        }
        public static String GetResponse(Uri uri)
        {
            String html = null;

            Thread.Sleep(1000);

            HttpWebRequest request = (HttpWebRequest)WebRequest.Create(uri);
            request.Accept = "text/html";
            request.UserAgent = "Mozilla/5.0";
            request.UseDefaultCredentials = true;
            request.Timeout = 10000; // 10 sec.
            //request.Proxy = new WebProxy("127.0.0.1", 9050);

            HttpWebResponse response = null;

            try
            {
                response = (HttpWebResponse)request.GetResponse();
            }
            catch (Exception e) { Log.Write(String.Format("{0}\n{1}", request.RequestUri, e.Message)); }

            if (response != null)
            {
                if (response.StatusCode == HttpStatusCode.OK)
                {
                    Stream receivedStream = response.GetResponseStream();
                    Encoding encoding = Encoding.GetEncoding(response.CharacterSet);
                    using (StreamReader readStream = new StreamReader(receivedStream, encoding))
                    {
                        html = readStream.ReadToEnd();
                    }
                }
                else { Log.Write(response.StatusDescription); }
            }
            return html;
        }
        public static String GetValueByRegex(String src, String re, Int32 startAt = 0)
        {
            String value = null;
            Regex regex = new Regex(re, RegexOptions.Singleline);
            Match match = regex.Match(src, startAt);
            if (match.Success && match.Groups.Count > 1)
            {
                value = match.Groups[1].Value.Trim();
            }
            return value;
        }
    }
    class SectionIndexes
    {
        private String src;
        private Int32 i0 = -1;
        private Int32 i1 = -1;
        private Int32 i2 = -1;
        private Int32 i3 = -1;

        public Int32 Index0
        {
            get { return i0; }
            set
            {
                if (value < -1 || value > src.Length) throw new ArgumentOutOfRangeException();
                else { i0 = i1 = i2 = i3 = value; }
            }
        }
        public Int32 Index1
        {
            get { return i1; }
            set
            {
                if (value < i0 || value > src.Length) throw new ArgumentOutOfRangeException();
                else { i1 = i2 = i3 = value; }
            }
        }
        public Int32 Index2
        {
            get { return i2; }
            set
            {
                if (value < i1 || value > src.Length) throw new ArgumentOutOfRangeException();
                else { i2 = i3 = value; }
            }
        }
        public Int32 Index3
        {
            get { return i3; }
            set
            {
                if (value < i2 || value > src.Length) throw new ArgumentOutOfRangeException();
                else i3 = value;
            }
        }

        public SectionIndexes(String src, String[] bf = null, Int32 startIndex = 0, String[] ef = null, Int32 endIndex = 0)
        {
            if (src is null) throw new ArgumentNullException();
            this.src = src;

            if (src is null || src.Length == 0) return;
            if (endIndex <= 0 || endIndex > src.Length) endIndex = src.Length;

            if (bf != null && bf.Length > 0)
            {
                bool bfIsFound = false;
                while (!bfIsFound && startIndex < endIndex)
                {
                    foreach (String value in bf)
                    {
                        if (!String.IsNullOrEmpty(value))
                        {
                            Int32 index = src.IndexOf(value, startIndex);
                            if (index >= 0)
                            {
                                bfIsFound = true;
                                if (Index0 == -1)
                                {
                                    Index0 = index;
                                    Index1 = index + value.Length;
                                    Index2 = Index1;
                                    Index3 = Index2;
                                }
                                else
                                {
                                    Index1 = index + value.Length;
                                    Index2 = Index1;
                                    Index3 = Index2;
                                }
                                startIndex = Index1;
                            }
                            else
                            {
                                bfIsFound = false;
                                Index0 = -1;
                                Index1 = -1;
                                Index2 = -1;
                                Index3 = -1;
                                startIndex = endIndex;
                                break;
                            }
                        }
                    }
                }
                if (bfIsFound && ef != null && ef.Length > 0)
                {
                    startIndex = Index1;
                    bool efIsFound = false;
                    while (!efIsFound && startIndex < endIndex)
                    {
                        foreach (String value in bf)
                        {
                            if (!String.IsNullOrEmpty(value))
                            {
                                Int32 index = src.IndexOf(value, startIndex);
                                if (index >= 0)
                                {
                                    efIsFound = true;
                                    if (Index2 == Index1)
                                    {
                                        Index2 = index;
                                        Index3 = index + value.Length;
                                    }
                                    else
                                    {
                                        Index3 = index + value.Length;
                                    }
                                    startIndex = Index3;
                                }
                                else
                                {
                                    efIsFound = false;
                                    Index2 = Index1;
                                    Index3 = Index1;
                                    startIndex = endIndex;
                                    break;
                                }
                            }
                        }
                    }
                }
            }
        }
    }
}
