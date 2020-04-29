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
            String auctionUrl = rqp["auction_href"] as String;

            Guid? auctionUid = null;
            if (!String.IsNullOrWhiteSpace(auctionNumber))
            {
                auctionUid = GetAuctionUid(auctionNumber);
            }

            if (auctionUid == null || overwrite)
            {
                if (!String.IsNullOrWhiteSpace(auctionUrl) )
                {
                    auctionUid = LoadAndSave(auctionNumber, auctionUrl);
                }
                else { rsp.Status = "Не задан номер аукциона"; }
            }

            if (auctionUid != null)
            {
                rsp.Status = "ОК";
                rsp.Data = new DataSet();
                rsp.Data.Tables.Add();
                rsp.Data.Tables[0].Columns.Add(new DataColumn("auction_uid", typeof(Guid)));
                rsp.Data.Tables[0].Columns.Add(new DataColumn("auction_number", typeof(Guid)));
                rsp.Data.Tables[0].Columns.Add(new DataColumn("auction_url", typeof(Guid)));
                rsp.Data.Tables[0].Rows.Add(new object[] { auctionUid, auctionNumber, auctionUrl });
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
        private static Guid? LoadAndSave(String auctionNumber, String auctionUrl)
        {
            Guid? auctionUid = null;
            try
            {
                //if (!String.IsNullOrWhiteSpace(auctionNumber) && (auctionNumber.Length == 11 || auctionNumber.Length == 19))
                {
                    String html = ZakupkiGovRu.GetAuctionInf(auctionNumber, auctionUrl);
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
                //else { Log.Write(String.Format($"Не указан номер заявки.")); }
            }
            catch (Exception e) { Log.Write(String.Format(e.Message)); }
            return auctionUid;
        }
    }
    class ZakupkiGovRu
    {
        public static String GetAuctionInf(String auctionNumber, String auctionUrl)
        {
            String html = null;
            
            if (!String.IsNullOrWhiteSpace(auctionUrl))
            {
                Uri uri = new Uri("http://zakupki.gov.ru" + auctionUrl);
                html = Utilities.GetResponse(uri);
            }
            else
            
            {
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
        private static String BuildQueryStringForRegion(String regionNumber, String publishDate, String searchString, Int32 recordsPerPage = 50, Int32 pageNumber = 1)
        {
            publishDate = String.Format("{0}.{1}.20{2}", publishDate.Substring(0, 2), publishDate.Substring(3, 2), publishDate.Substring(6, 2));
            StringBuilder sb = new StringBuilder();
            sb.AppendFormat("searchString={0}", HttpUtility.UrlEncode(searchString));
            sb.Append("&morphology=on");
            sb.AppendFormat("&pageNumber={0}", pageNumber);
            //sb.Append("&sortDirection=false");
            sb.AppendFormat("&recordsPerPage=_{0}", recordsPerPage);
            sb.Append("&showLotsInfoHidden=false");
            sb.Append("&fz44=on");
            sb.Append("&fz223=on");
            sb.Append("&sortBy=UPDATE_DATE");
            sb.Append("&af=on");
            sb.Append("&ca=on");
            //sb.Append("&pc=on"); // закупка завершена
            //sb.Append("&priceFrom=");
            //sb.Append("&priceTo=");
            //sb.Append("&currencyId=1");
            //sb.Append("&agencyTitle=");
            //sb.Append("&agencyCode=");
            //sb.Append("&agencyFz94id=");
            //sb.Append("&agencyFz223id=");
            //sb.Append("&agencyInn=");
            //sb.AppendFormat("&region_regions_{0}=region_regions_{0}", regionNumber);
            //sb.AppendFormat("&regions={0}", regionNumber);
            sb.AppendFormat("&publishDateFrom={0}", publishDate);
            sb.AppendFormat("&publishDateTo={0}", publishDate);
            sb.Append("&currencyIdGeneral=-1");
            //sb.Append("&updateDateFrom=");
            //sb.Append("&updateDateTo=");
            sb.Append("&customerPlaceWithNested=on");
            sb.AppendFormat("&customerPlace={0}", regionNumber);
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
                    rsp.Data.Tables[0].Columns.Add("auction_href", typeof(String));

                    List<String> auctionNumbers = new List<string>();

                    UriBuilder ub = new UriBuilder
                    {
                        Scheme = "http",
                        Host = "zakupki.gov.ru",
                        Path = "/epz/order/extendedsearch/results.html"
                    };
                    String[] ss = new String[] { "лекарств", "препарат", "медикамент" };
                    foreach (String s in ss)
                    {
                        Int32 recordsPerPage = 50;
                        Int32 pageCount = 1;
                        Int32 pageNumber = 0;
                        while (pageNumber++ < pageCount)
                        {
                            ub.Query = BuildQueryStringForRegion(regionNumber, publishDate, s, recordsPerPage, pageNumber);
                            String receivedString = Utilities.GetResponse(ub.Uri);
                            if (!String.IsNullOrWhiteSpace(receivedString))
                            {
                                // найти количество записей в ответе
                                Int32 recordCount = GetRecordCount(receivedString);
                                if (pageNumber == 1)
                                {
                                    pageCount = (recordCount / recordsPerPage) + 1;
                                }
                                List<String> hrefs = GetRecordHrefList(receivedString);
                                foreach (String href in hrefs)
                                {
                                    Int32 si = href.IndexOf("regNumber=") + 10;
                                    String num = href.Substring(si);
                                    rsp.Data.Tables[0].Rows.Add(new Object[] { num, href });
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
            Int32 sIndex = receivedString.IndexOf("search-results__total");
            if (sIndex >= 0)
            {
                String temp = receivedString.Substring(sIndex, 100);
                Regex re = new Regex(@"\d+");
                Match match = re.Match(temp);
                Int32.TryParse(match.Value, out recordCount);
            }
            return recordCount;
        }
        private static List<String> GetRecordHrefList(String s)
        {
            List<String> hrefs = new List<string>();
            if (!String.IsNullOrWhiteSpace(s))
            {
                Int32 sIndex = 0;
                while (true)
                {
                    SectionIndexes sIs = new SectionIndexes(s, new String[] { "registry-entry__header-mid__number", "<a", "href=\"" }, sIndex, new string[] { "\"" });
                    String v = sIs.InnerText;
                    if (!String.IsNullOrWhiteSpace(v))
                    {
                        hrefs.Add(v);
                        sIndex = sIs.TEIndex;
                        continue;
                    }
                    break;
                }
            }
            return hrefs;
        }
    }
    class Fz44
    {
        public static Guid SaveAuctionInf(String html)
        {
            // записываем на SQL сервер в базу Auctions общую информацию о заявке
            Guid aUid = ParseAndSaveAuction44fzCommonInf(html);

            // записываем на SQL сервер в базу Auctions информацию о заказчиках (их может быть несколько)
            //ParseAndSaveAuction44FzCustomerRequirement(aUid, html);

            return aUid;
        }
        private static Guid ParseAndSaveAuction44fzCommonInf(String html)
        {
            Guid aUid = new Guid();

            // берём основные блоки
            Int32 c0 = new SectionIndexes(html, new String[] { "cardWrapper" }).HEIndex;
            // заявка
            Int32 h0 = new SectionIndexes(html, new String[] { "cardMainInfo" }, c0).HEIndex;
            // общая_информация_о_закупке
            Int32 t0 = new SectionIndexes(html, new String[] { "blockInfo__title", "span", "Общая информация о закупке", "</span>" }, c0).HEIndex;
            // информация_об_организации_осуществляющей_определение_поставщика
            Int32 t1 = new SectionIndexes(html, new String[] { "blockInfo__title", "span", "Контактная информация", "</span>" }, c0).HEIndex;
            // информация_о_процедуре_закупки
            Int32 t2 = new SectionIndexes(html, new String[] { "blockInfo__title", "span", "Информация о процедуре электронного аукциона", "</span>" }, c0).HEIndex;
            // начальная_максимальная_цена_контракта
            Int32 t3 = new SectionIndexes(html, new String[] { "blockInfo__title", "span", "Начальная (максимальная) цена контракта", "</span>" }, c0).HEIndex;
            // информация_об_объекте_закупки
            Int32 t4 = new SectionIndexes(html, new String[] { "blockInfo__title", "span", "Информация об объекте закупки", "</span>" }, c0).HEIndex;
            // преимущества_требования_к_участникам
            Int32 t5 = new SectionIndexes(html, new String[] { "blockInfo__title", "span", "Преимущества, требования к участникам", "</span>" }, c0).HEIndex;
            //
            Int32 t6 = new SectionIndexes(html, new String[] { "blockInfo__title", "span", "Документация об электронном аукционе", "</span>" }, c0).HEIndex;
            //
            Int32 t7 = new SectionIndexes(html, new String[] { "blockInfo__title", "span", "Обеспечение заявки", "</span>" }, c0).HEIndex;
            //
            Int32 t8 = new SectionIndexes(html, new String[] { "blockInfo__title", "span", "Условия контракта", "</span>" }, c0).HEIndex;
            //
            Int32 t9 = new SectionIndexes(html, new String[] { "blockInfo__title", "span", "Обеспечение исполнения контракта", "</span>" }, c0).HEIndex;
            //
            Int32 t10 = new SectionIndexes(html, new String[] { "blockInfo__title", "span", "Обеспечение гарантийных обязательств", "</span>" }, c0).HEIndex;
            //
            Int32 t11 = new SectionIndexes(html, new String[] { "blockInfo__title", "span", "Информация о банковском и (или) казначейском сопровождении контракта", "</span>" }, c0).HEIndex;

            // записываем на SQL сервер в базу Auctions общую информацию о заявке

            Object[][] md44 = new Object[][] { 
                // закупка
                new Object[] { "номер",                                         h0, new String[] { ">№ " },                                                                 new String[] { "<" } },
                new Object[] { "дата_размещения",                               h0, new String[] { "Размещено" },                                                           new String[] { "</div>" } },
                new Object[] { "кооператив",                                    h0, new String[] { "cooperative" },                                                         null, "1" },
                // общая_информация_о_закупке
                new Object[] { "способ_определения_поставщика",                 t0, new String[] { "section", "span", "Способ определения поставщика", "</span>" },         new String[] { "</section>" } },
                new Object[] { "наименование_электронной_площадки_в_интернете", t0, new String[] { "section", "span", "Наименование электронной площадки", "</span>" },     new String[] { "</section>" } },
                new Object[] { "адрес_электронной_площадки_в_интернете",        t0, new String[] { "section", "span", "Адрес электронной площадки", "</span>", "<a", ">" }, new String[] { "</section>" } },
                new Object[] { "размещение_осуществляет",                       t0, new String[] { "section", "span", "Размещение осуществляет", "</span>", "<a", ">" },    new String[] { "</section>" } },
                new Object[] { "объект_закупки",                                t0, new String[] { "section", "span", "Наименование объекта закупки", "</span>" },          new String[] { "</section>" } },
                new Object[] { "этап_закупки",                                  t0, new String[] { "section", "span", "Этап закупки", "</span>" },                          new String[] { "</section>" } },
                new Object[] { "сведения_о_связи_с_позицией_плана_графика",     t0, new String[] { "section", "span", "Сведения о связи", "</span>" },                      new String[] { "</section>" } },
                new Object[] { "номер_типового_контракта",                      t0, new String[] { "section", "span", "Номер типового контракта", "</span>" },              new String[] { "</section>" } },
                new Object[] { "дата_и_время_окончания_подачи_заявок",          t0, new String[] { "section", "span", "Дата и время окончания срока подачи", "</span>" },   new String[] { "</section>" } },
                // информация_об_организации_осуществляющей_определение_поставщика
                new Object[] { "организация_осуществляющая_размещение",         t1, new String[] { "section", "span", "Организация, осуществляющая размещение", "</span>" },new String[] { "</section>" } },
                new Object[] { "почтовый_адрес",                                t1, new String[] { "section", "span", "Почтовый адрес", "</span>" },                        new String[] { "</section>" } },
                new Object[] { "место_нахождения",                              t1, new String[] { "section", "span", "Место нахождения", "</span>" },                      new String[] { "</section>" } },
                new Object[] { "ответственное_должностное_лицо",                t1, new String[] { "section", "span", "Ответственное должностное лицо", "</span>" },        new String[] { "</section>" } },
                new Object[] { "адрес_электронной_почты",                       t1, new String[] { "section", "span", "Адрес электронной почты", "</span>" },               new String[] { "</section>" } },
                new Object[] { "номер_контактного_телефона",                    t1, new String[] { "section", "span", "Номер контактного телефона", "</span>" },            new String[] { "</section>" } },
                new Object[] { "факс",                                          t1, new String[] { "section", "span", "Факс", "</span>" },                                  new String[] { "</section>" } },
                new Object[] { "дополнительная_информация",                     t1, new String[] { "section", "span", "Дополнительная информация", "</span>" },             new String[] { "</section>" } },
                // информация_о_процедуре_закупки
                new Object[] { "дата_и_время_начала_подачи_заявок",             t2, new String[] { "section", "span", "Дата и время начала подачи", "</span>" },            new String[] { "</section>" } },
                new Object[] { "дата_и_время_окончания_подачи_заявок",          t2, new String[] { "section", "span", "Дата и время окончания подачи", "</span>" },         new String[] { "</section>" } },
                new Object[] { "место_подачи_заявок",                           t2, new String[] { "section", "span", "Место подачи заявок", "</span>" },                   new String[] { "</section>" } },
                new Object[] { "порядок_подачи_заявок",                         t2, new String[] { "section", "span", "Порядок подачи заявок", "</span>" },                 new String[] { "</section>" } },
                new Object[] { "дата_окончания_срока_рассмотрения_первых_частей",t2,new String[] { "section", "span", "Дата окончания срока", "</span>" },                  new String[] { "</section>" } },
                new Object[] { "дата_проведения_аукциона_в_электронной_форме",  t2, new String[] { "section", "span", "Дата проведения", "</span>" },                       new String[] { "</section>" } },
                new Object[] { "время_проведения_аукциона",                     t2, new String[] { "section", "span", "Время проведения", "</span>" },                      new String[] { "</section>" } },
                new Object[] { "дополнительная_информация2",                    t2, new String[] { "section", "span", "Дополнительная информация", "</span>" },             new String[] { "</section>" } },
                // начальная_максимальная_цена_контракта
                new Object[] { "начальная_максимальная_цена_контракта",         t3, new String[] { "section", "span", "Начальная (максимальная) цена", "</span>" },         new String[] { "</section>" } },
                new Object[] { "валюта",                                        t3, new String[] { "section", "span", "Валюта", "</span>" },                                new String[] { "</section>" } },
                new Object[] { "источник_финансирования",                       t3, new String[] { "section", "span", "Источник финансирования", "</span>" },               new String[] { "</section>" } },
                new Object[] { "идентификационный_код_закупки",                 t3, new String[] { "section", "span", "Идентификационный код закупки", "</span>" },         new String[] { "</section>" } },
                new Object[] { "оплата_исполнения_контракта_по_годам",          t3, new String[] { "section", "span", "Оплата исполнения контракта", "</span>" },           new String[] { "</section>" } },
                // информация_об_объекте_закупки
                new Object[] { "описание_объекта_закупки",                      t4, new String[] { "section", "span", "Описание объекта закупки", "</span>" },              new String[] { "</section>" } },
                new Object[] { "условия_запреты_и_ограничения_допуска_товаров", t4, new String[] { "section", "span", "Условия", "запреты", "ограничения", "</span>" },     new String[] { "</section>" } },
                new Object[] { "табличная_часть_в_формате_html",                t4, new String[] { "section", "span", "Табличная часть", "</span>" },                       new String[] { "</section>" } },
                // преимущества_требования_к_участникам
                new Object[] { "преимущества",                                  t5, new String[] { "section", "span", "Преимущества", "</span>" },                          new String[] { "</section>" } },
                new Object[] { "требования",                                    t5, new String[] { "section", "span", "Требования", "</span>" },                            new String[] { "</section>" } },
                new Object[] { "ограничения",                                   t5, new String[] { "section", "span", "Ограничения", "</span>" },                           new String[] { "</section>" } }
            };

            SqlCommand cmd = new SqlCommand
            {
                Connection = new SqlConnection(Env.cnString),
                CommandText = "[Auctions].[dbo].[save_auction_inf]",
                CommandType = CommandType.StoredProcedure
            };
            foreach (Object[] p in md44)
            {
                String v = null;
                SectionIndexes sect;
                if (p[3] != null)
                {
                    sect = new SectionIndexes(html, (String[])p[2], (Int32)p[1], (String[])p[3]);
                    v = sect.InnerText;
                }
                else // cooperative
                {
                    sect = new SectionIndexes(html, (String[])p[2]);
                    if (sect.HeadIsFound) v = (String)p[4];
                }
                if (v != null)
                {
                    v = Utilities.NormString(v);
                    String dbName = (String)p[0];
                    if (cmd.Parameters.Contains(dbName)) cmd.Parameters[dbName].Value = v;
                    else cmd.Parameters.AddWithValue(dbName, v);
                }
            }

            using (cmd.Connection)
            {
                Object o = null;
                cmd.Connection.Open();
                o = cmd.ExecuteScalar();
                if (o != null && o.GetType() == typeof(Guid)) { aUid = (Guid)o; }
            }



            // записываем на SQL сервер в базу Auctions информацию о заказчиках (их может быть несколько)

            // здесь может быть два варианта
            // первый - когда списки expand пусты - сохраняем из основного блока
            // второй - сохраняем из списков

            // первый вариант
            SqlCommand cmd1 = new SqlCommand
            {
                Connection = new SqlConnection(Env.cnString),
                CommandText = "[Auctions].[dbo].[save_customer_requirement]",
                CommandType = CommandType.StoredProcedure
            };
            cmd1.Parameters.AddWithValue("закупка_uid", aUid);
            cmd1.Parameters.AddWithValue("наименование_заказчика", cmd.Parameters["размещение_осуществляет"].Value);
            using (cmd1.Connection)
            {
                cmd1.Connection.Open();
                cmd1.ExecuteScalar();
            }


            /*
            Object[][] mdCust = new Object[][] {
                // требования_заказчика
                new Object[] { "наименование_заказчика",                    0, new String[] { "&nbsp;" },                                                       new String[] { "</h3>" } },
                // условия_контракта
                new Object[] { "место_доставки_товара",                     0, new String[] { "Место доставки", "<td", ">" },                                   new String[] { "</td>" } },
                new Object[] { "сроки_поставки_товара",                     0, new String[] { "Сроки поставки", "<td", ">" },                                   new String[] { "</td>" } },
                new Object[] { "сведения_о_связи_с_позицией_плана_графика", 0, new String[] { "Сведения о связи", "Сведения о связи", "<td", ">" },             new String[] { "</td>" } },
                new Object[] { "оплата_исполнения_контракта_по_годам",      0, new String[] { "Оплата исполнения", "<td", ">" },                                new String[] { "</td>" } },
                // обеспечение_заявок
                new Object[] { "размер_обеспечения",                        0, new String[] { "Обеспечение заявок", "Размер обеспечения", "<td", ">" },         new String[] { "</td>" } },
                new Object[] { "порядок_внесения_денежных_средств",         0, new String[] { "Обеспечение заявок", "Порядок внесения", "<td", ">" },           new String[] { "</td>" } },
                new Object[] { "платежные_реквизиты",                       0, new String[] { "Обеспечение заявок", "Платежные реквизиты", "<td", ">" },        new String[] { "</td>" } },
                // обеспечение_исполнения_контракта
                new Object[] { "размер_обеспечения_2",                      0, new String[] { "Обеспечение исполнения", "Размер обеспечения", "<td", ">" },     new String[] { "</td>" } },
                new Object[] { "порядок_предоставления_обеспечения",        0, new String[] { "Обеспечение исполнения", "Порядок предоставления", "<td", ">" }, new String[] { "</td>" } },
                new Object[] { "платежные_реквизиты_2",                     0, new String[] { "Обеспечение исполнения", "Платежные реквизиты", "<td", ">" },    new String[] { "</td>" } },
            };
            
            Int32 bi = new SectionIndexes(html, new String[] { "Требования заказчиков" }).Index1;

            var sect = new SectionIndexes(html, new String[] { "Требования заказчика" }, bi);
            if (!sect.HeadIsFound)
            {
                String v;
                foreach (Object[] p in mdCust)
                {
                    v = null;
                    sect = new SectionIndexes(html, (String[])p[2], (Int32)p[1], (String[])p[3]);
                    v = sect.InnerText;
                    if (v != null)
                    {
                        v = Utilities.NormString(v);
                        cmd.Parameters.AddWithValue((String)p[0], v);
                    }
                }
                sect = new SectionIndexes(html, new String[] { "Общая информация", "Размещение осуществляет", "<td", ">", "<a", ">" }, 0, new String[] { "</a>" });
                v = sect.InnerText;
                if (v != null)
                {
                    v = Utilities.NormString(v);
                    if (cmd.Parameters.Contains("наименование_заказчика")) cmd.Parameters["наименование_заказчика"].Value = v;
                    else cmd.Parameters.AddWithValue("наименование_заказчика", v);
                }
            }
            else
            {
                // второй вариант
                while (sect.HeadIsFound)
                {
                    Int32 ci = sect.Index1;
                    SqlCommand cmd = new SqlCommand
                    {
                        Connection = new SqlConnection(Env.cnString),
                        CommandText = "[Auctions].[dbo].[save_customer_requirement]",
                        CommandType = CommandType.StoredProcedure
                    };
                    cmd.Parameters.AddWithValue("закупка_uid", aUid);
                    foreach (Object[] p in mdCust)
                    {
                        String v = null;
                        sect = new SectionIndexes(html, (String[])p[2], ci, (String[])p[3]);
                        v = sect.InnerText;
                        if (v != null)
                        {
                            v = Utilities.NormString(v);
                            cmd.Parameters.AddWithValue((String)p[0], v);
                        }
                    }
                    using (cmd.Connection)
                    {
                        Object o = null;
                        cmd.Connection.Open();
                        o = cmd.ExecuteScalar();
                        if (o != null && o.GetType() == typeof(Guid)) { aUid = (Guid)o; }
                    }
                    sect = new SectionIndexes(html, new String[] { "Требования заказчика" }, ci);
                }
            }
            */
            return aUid;
        }
    }
    class Fz223
    {
        public static Guid SaveAuctionInf(String html)
        {
            Guid aUid = new Guid();
            Object[][] md223 = new Object[][] { 
                // Закупка
                new Object[] { "номер",                                     new String[] { "Общие сведения", "омер извещения",                              "<td", ">" },           new String[] { "</td>" } },
                new Object[] { "дата_размещения",                           new String[] { "Общие сведения", "Дата размещения извещения",                   "<td", ">" },           new String[] { "</td>" } },
                // Общие сведения о закупке 
                new Object[] { "номер_извещения",                           new String[] { "Общие сведения", "омер извещения",                              "<td", ">" },           new String[] { "</td>" } },
                new Object[] { "способ_размещения_закупки",                 new String[] { "Общие сведения", "Способ размещения закупки",                   "<td", ">" },           new String[] { "</td>" } },
                new Object[] { "наименование_закупки",                      new String[] { "Общие сведения", "Наименование закупки",                        "<td", ">" },           new String[] { "</td>" } },
                new Object[] { "закупка_осуществляется_вследствие_аварии",  new String[] { "Общие сведения", "Закупка осуществляется вследствие аварии",    "<td", ">" },           new String[] { "</td>" } },
                new Object[] { "редакция",                                  new String[] { "Общие сведения", "Редакция",                                    "<td", ">" },           new String[] { "</td>" } },
                new Object[] { "наименование_электронной_площадки",         new String[] { "Общие сведения", "Наименование электронной площадки",           "<td", ">" },           new String[] { "</td>" } },
                new Object[] { "адрес_электронной_площадки",                new String[] { "Общие сведения", "Адрес электронной площадки",                  "<td", ">" },           new String[] { "</td>" } },
                new Object[] { "дата_размещения_извещения",                 new String[] { "Общие сведения", "Дата размещения извещения",                   "<td", ">" },           new String[] { "</td>" } },
                new Object[] { "дата_размещения_текущей_редакции_извещения",new String[] { "Общие сведения", "Дата размещения текущей редакции",            "<td", ">" },           new String[] { "</td>" } },
                // Заказчик
                new Object[] { "наименование_организации",                  new String[] { "Заказчик", "Наименование организации",                          "<td", ">", "<a", ">" },new String[] { "</a>" } },
                new Object[] { "инн_кпп",                                   new String[] { "Заказчик", "ИНН",                                               "<td", ">" },           new String[] { "</td>" } },
                new Object[] { "огрн",                                      new String[] { "Заказчик", "ОГРН",                                              "<td", ">" },           new String[] { "</td>" } },
                new Object[] { "адрес_места_нахождения",                    new String[] { "Заказчик", "есто нахождения",                                   "<td", ">" },           new String[] { "</td>" } },
                new Object[] { "почтовый_адрес",                            new String[] { "Заказчик", "Почтовый адрес",                                    "<td", ">" },           new String[] { "</td>" } },
                // Контактное лицо
                new Object[] { "организация",                               new String[] { "Контактное лицо", "Организация",                                "<td", ">" },           new String[] { "</td>" } },
                new Object[] { "контактное_лицо",                           new String[] { "Контактное лицо", "Контактное лицо",                            "<td", ">" },           new String[] { "</td>" } },
                new Object[] { "электронная_почта",                         new String[] { "Контактное лицо", "Электронная почта",                          "<td", ">" },           new String[] { "</td>" } },
                new Object[] { "телефон",                                   new String[] { "Контактное лицо", "Телефон",                                    "<td", ">" },           new String[] { "</td>" } },
                new Object[] { "факс",                                      new String[] { "Контактное лицо", "Факс",                                       "<td", ">" },           new String[] { "</td>" } },
                // Требования к участникам закупки
                new Object[] { "требование_к_отсутствию_участников_закупки_в_реестре_недобросовестных_поставщиков",
                                                                            new String[] { "Требования к участникам закупки", "Требование к отсутствию",    "<td", ">" },           new String[] { "</td>" } },
                // Порядок размещения закупки
                new Object[] { "дата_и_время_окончания_подачи_заявок",      new String[] { "Порядок проведения процедуры", "Дата и время окончания подачи", "<td", ">" },           new String[] { "</td>" } },
                new Object[] { "дата_окончания_срока_рассмотрения_заявок",  new String[] { "Порядок проведения процедуры", "Дата рассмотрения",             "<td", ">" },           new String[] { "</td>" } },
                new Object[] { "дата_и_время_проведение_аукциона",          new String[] { "Порядок проведения процедуры", "Дата подведения итогов",        "<td", ">" },           new String[] { "</td>" } },
                // Предоставление документации
                new Object[] { "срок_предоставления",                       new String[] { "Предоставление документации", "Срок",                           "<td", ">" },           new String[] { "</td>" } },
                new Object[] { "место_предоставления",                      new String[] { "Предоставление документации", "Место",                          "<td", ">" },           new String[] { "</td>" } },
                new Object[] { "порядок_предоставления",                    new String[] { "Предоставление документации", "Порядок",                        "<td", ">" },           new String[] { "</td>" } },
                new Object[] { "официальный_сайт",                          new String[] { "Предоставление документации", "Официальный сайт",               "<td", ">" },           new String[] { "</td>" } },
                new Object[] { "внесение_платы_за_предоставление_конкурсной_документации",
                                                                            new String[] { "Предоставление документации", "Внесение платы",                 "<td", ">" },           new String[] { "</td>" } }
                };
            SqlCommand cmd = new SqlCommand
            {
                Connection = new SqlConnection(Env.cnString),
                CommandText = "[Auctions].[dbo].[save_auction223_inf]",
                CommandType = CommandType.StoredProcedure
            };
            foreach (Object[] p in md223)
            {
                var sect = new SectionIndexes(html, (String[])p[1], 0, (String[])p[2]);
                String v = sect.InnerText;
                if (v != null)
                {
                    String dbName = p[0] as String;
                    v = Utilities.NormString(v);
                    if (cmd.Parameters.Contains(dbName)) cmd.Parameters[dbName].Value = v;
                    else cmd.Parameters.AddWithValue(dbName, v);
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
                r = new Regex("<span[^>]*>").Replace(r, " ");
                r = r.Replace("</span>", " ");
                r = new Regex("<a[^>]*>").Replace(r, " ");
                r = r.Replace("</a>", " ");
                r = new Regex("<br[^>]*>").Replace(r, " ");
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
    }
    class SectionIndexes
    {
        private String src;
        private Int32 hbi; // индекс первого символа заголовка
        private Int32 hei; // индекс последнего символа заголовка
        private Int32 tbi; // индекс первого символа хвоста
        private Int32 tei; // индекс последнего символа хвоста

        /// <summary>
        /// индекс первого символа заголовка
        /// </summary>
        public Int32 HBIndex
        {
            get => hbi;
            set
            {
                if (value < 0 || value > src.Length) throw new ArgumentOutOfRangeException();
                else { hbi = hei = tbi = tei = value; }
            }
        }
        /// <summary>
        /// индекс последнего символа заголовка
        /// </summary>
        public Int32 HEIndex
        {
            get => hei;
            set
            {
                if (value < hbi || value > src.Length) throw new ArgumentOutOfRangeException();
                else { hei = tbi = tei = value; }
            }
        }
        /// <summary>
        /// индекс первого символа хвоста
        /// </summary>
        public Int32 TBIndex
        {
            get => tbi;
            set
            {
                if (value < hei || value > src.Length) throw new ArgumentOutOfRangeException();
                else { tbi = tei = value; }
            }
        }
        /// <summary>
        /// индекс последнего символа хвоста
        /// </summary>
        public Int32 TEIndex
        {
            get => tei;
            set
            {
                if (value < tbi || value > src.Length) throw new ArgumentOutOfRangeException();
                else tei = value;
            }
        }
        public Boolean HeadIsFound { get => HEIndex > HBIndex; }
        public Boolean TailIsFound { get => TEIndex > TBIndex; }
        public String InnerText { get => (HeadIsFound && TailIsFound) ? src.Substring(HEIndex, TBIndex - HEIndex) : null; }
        public SectionIndexes(String src, String[] hf = null, Int32 startIndex = 0, String[] tf = null)
        {
            if (src == null || startIndex < 0) throw new ArgumentException();
            this.src = src;

            HBIndex = startIndex;
            if (hf != null && hf.Length > 0)
            {
                bool hfIsFound = false;
                foreach (String value in hf)
                {
                    if (!String.IsNullOrEmpty(value))
                    {
                        Int32 index = src.IndexOf(value, HEIndex);
                        if (index >= HEIndex)
                        {
                            if (!hfIsFound) HBIndex = index;
                            HEIndex = index + value.Length;
                            hfIsFound = true;
                        }
                        else
                        {
                            HEIndex = HBIndex;
                            hfIsFound = false;
                            break;
                        }
                    }
                }
            }

            if (tf != null && tf.Length > 0)
            {
                bool efIsFound = false;
                foreach (String value in tf)
                {
                    if (!String.IsNullOrEmpty(value))
                    {
                        Int32 index = src.IndexOf(value, TEIndex);
                        if (index >= TEIndex)
                        {
                            if (!efIsFound) TBIndex = index;
                            TEIndex = index + value.Length;
                            efIsFound = true;
                        }
                        else
                        {
                            TEIndex = TBIndex;
                            efIsFound = false;
                            break;
                        }
                    }
                }
            }
        }
    }
}
