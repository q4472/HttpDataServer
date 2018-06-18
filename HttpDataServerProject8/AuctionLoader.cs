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
using System.Xml;

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
                if (auctionNumber != null && auctionNumber.Length == 11) // 223 фз
                {
                    // получаем 6 блоков Auction223FzInfBlockNames в виде xml документов с html разметкой <div... </div>
                    List<AuctionInfBlock> aibs = ZakupkiGovRu.GetAuction223FzInf(auctionNumber);
                    if (aibs != null)
                    {
                        if (aibs.Count < 6) { Log.Write(String.Format("Информация загрузилась не полностью. Всего блоков: {0}.", aibs.Count)); }
                        // сохраняем данные об аукционе
                        auctionUid = Fz223.SaveAuctionInf(aibs);
                    }
                    else { Log.Write(String.Format("Информация не загрузилась.")); }
                }
                if (auctionNumber != null && auctionNumber.Length == 19) // 44 фз
                {

                    XmlDocument d = ZakupkiGovRu.GetAuction44FzInf(auctionNumber);
                    if (d != null)
                    {
                        XmlNode mainBox = d.SelectSingleNode("/html/body/div[@class='cardWrapper']/div[@class='wrapper']/div[@class='mainBox']");
                        if (mainBox != null)
                        {
                            // сохраняем данные об аукционе
                            auctionUid = Fz44.SaveAuctionInf(auctionNumber, mainBox);
                        }
                    }
                    else { Log.Write(String.Format("Информация не загрузилась.")); }
                }
            }
            catch (Exception e) { Log.Write(String.Format(e.Message)); }
            return auctionUid;
        }
    }
    class ZakupkiGovRu
    {
        public static XmlDocument GetAuction44FzInf(String auctionNumber)
        {
            XmlDocument d = null;
            List<AuctionInfBlock> aibs = null;
            if (!String.IsNullOrWhiteSpace(auctionNumber) && auctionNumber.Length == 19)
            {
                UriBuilder ub = new UriBuilder
                {
                    Scheme = "http",
                    Host = "zakupki.gov.ru",
                    Path = "/epz/order/notice/ea44/view/common-info.html",
                    Query = String.Format("regNumber={0}", auctionNumber)
                };
                String receivedString = Utilities.GetResponse(ub.Uri);
                if (!String.IsNullOrWhiteSpace(receivedString))
                {
                    String temp = NormXmlString(receivedString);
                    aibs = GetAuctionInfBlockData(temp, Auction44FzInfBlockNames);
                    d = new XmlDocument();
                    d.LoadXml(temp);
                }
            }
            return d;
        }
        private static String[] Auction44FzInfBlockNames = new String[]
        {
            "Общая информация о закупке",
            "Информация об организации, осуществляющей определение поставщика (подрядчика, исполнителя)",
            "Информация о процедуре закупки",
            "Начальная (максимальная) цена контракта",
            "Информация об объекте закупки",
            "Преимущества, требования к участникам",

            "Требования заказчика", // этот блок появляется несколько раз по числу заказчиков (кооператив)
                                    // он содержит вложенные блоки описанные в Customer44FzInfBlockNames
                                    // порядок важен так как "Обеспечение заявок" входит в Customer44FzInfBlockNames

            "Условия контракта", // этот блок есть только если заказчик один (не кооператив)
            "Обеспечение заявок", // этот блок есть только если заказчик один (не кооператив)
            "Обеспечение исполнения контракта" // этот блок есть только если заказчик один (не кооператив)
        };
        private static String[] Customer44FzInfBlockNames = new String[]
        {
            "Сведения о связи с позицией плана-графика",
            "Начальная (максимальная) цена контракта",
            "Обеспечение заявок",
            "Обеспечение исполнения контракта"
        };
        private static String[] Auction223FzInfBlockNames = new String[]
        {
            "Общие сведения о закупке",
            "Заказчик",
            "Контактное лицо",
            "Требования к участникам закупки",
            "Порядок размещения закупки",
            "Предоставление документации"
        };
        public static List<AuctionInfBlock> GetAuction223FzInf(String auctionNumber)
        {
            List<AuctionInfBlock> aibs = null;
            if (!String.IsNullOrWhiteSpace(auctionNumber) && auctionNumber.Length == 11)
            {
                UriBuilder ub = new UriBuilder
                {
                    Scheme = "http",
                    Host = "zakupki.gov.ru",
                    Path = "/223/purchase/public/purchase/info/common-info.html",
                    Query = String.Format("regNumber={0}", auctionNumber)
                };
                String receivedString = Utilities.GetResponse(ub.Uri);
                if (!String.IsNullOrWhiteSpace(receivedString))
                {
                    String temp = NormXmlString(receivedString);
                    aibs = GetAuctionInfBlockData(temp, Auction223FzInfBlockNames);
                }
            }
            return aibs;
        }
        private static List<AuctionInfBlock> GetAuctionInfBlockData(String inputString, String[] blockNames)
        {
            List<AuctionInfBlock> aibs = new List<AuctionInfBlock>();
            Int32 inputStringCurrentIndex = 0;
            foreach (String blockName in blockNames)
            {
                while (true)
                {
                    XmlDocument d = null;
                    Int32 blockStartIndex = inputString.IndexOf(String.Format("{0}", blockName), inputStringCurrentIndex);
                    if (blockStartIndex >= 0)
                    {
                        // очередной блок найден
                        inputStringCurrentIndex = blockStartIndex + blockName.Length;
                        Int32 divStartIndex = GetHtmlDivStartIndex(inputString, blockStartIndex);
                        if (divStartIndex >= 0)
                        {
                            inputStringCurrentIndex = divStartIndex + 1;
                            Int32 divFinishIndex = GetHtmlDivFinishIndex(inputString, divStartIndex);
                            if (divFinishIndex >= 0)
                            {
                                inputStringCurrentIndex = divFinishIndex;
                                String htmlDiv = inputString.Substring(divStartIndex, divFinishIndex - divStartIndex);
                                d = new XmlDocument();
                                try
                                {
                                    d.LoadXml(htmlDiv);
                                }
                                catch (Exception e) { d = null; Log.Write(String.Format(e.Message)); }
                            }
                        }
                        aibs.Add(new AuctionInfBlock { Name = blockName, Data = d });

                    }
                    else { break; }
                    if (blockName != "Требования заказчика") { break; }
                }
            }
            return aibs;
        }
        private static Int32 GetHtmlDivStartIndex(String inputString, Int32 searchStartIndex)
        {
            Int32 divStartIndex = -1;
            if (!String.IsNullOrWhiteSpace(inputString) && inputString.Length > searchStartIndex)
            {
                divStartIndex = inputString.IndexOf("<div", searchStartIndex);
            }
            return divStartIndex;
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
        private static String NormXmlString(String s)
        {
            String temp = (new Regex(@"(?s)<script[^>]*>.*?<\/script>")).Replace(s, " ");
            temp = (new Regex(@"(?s)<style[^>]*>.*?<\/style>")).Replace(temp, " ");
            temp = (new Regex(@"(?s)<link[^>]*>")).Replace(temp, "");
            temp = (new Regex(@"(?s)<br[^>]*>")).Replace(temp, "");
            temp = (new Regex(@"(?s)<tr>\s*?<td colspan=""6"" class=""delimTr""></td>\s*?<tr>")).Replace(temp, @"<tr><td colspan=""6"" class=""delimTr""></td></tr>");
            temp = temp.Replace("&", "&amp;");
            temp = temp.Replace("<div class=\"td-overflow\">", " ");
            temp = temp.Replace(" class>", ">");
            temp = temp.Replace("\t", " ").Replace("\n", " ").Replace("\r", " ");
            temp = temp.Replace("&nbsp;", " ");
            while (temp.Contains("  ")) temp = temp.Replace("  ", " ");
            temp = temp.Replace("Общая информация</span></div> </td>", "Общая информация</span> </td>");
            temp = temp.Replace("Документы</span></div> </td>", "Документы</span> </td>");
            temp = temp.Replace("Журнал событий</span></div> </td>", "Журнал событий</span> </td>");
            temp = temp.Replace("</span ", "</span>");
            temp = temp.Replace("\"\">", "\">");
            temp = temp.Replace("=\">", "=\"\">");
            temp = temp.Replace("Результаты определения поставщика</span></div>", "Результаты определения поставщика</span>");
            return temp;
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
    class AuctionInfBlock
    {
        public String Name;
        public XmlDocument Data;
    }
    class Fz44
    {
        public static Guid SaveAuctionInf(List<AuctionInfBlock> aibs)
        {
            Guid aUid = new Guid();
            Object[][] md44 = new Object[][] {
                // закупка
                new Object[] { "номер", "Общая информация о закупке", "Способ определения поставщика" },
                    /*
                    new Object[] { "дата_размещения",                               cardHeader, "./div[@class='public']" },
                    new Object[] { "кооператив",                                    cardHeader, "./div[@class='public']/span[@class='cooperative']" },
                    // общая_информация_о_закупке
                    new Object[] { "способ_определения_поставщика",                 t0, ".//tr[1]/td[2]" },
                    new Object[] { "наименование_электронной_площадки_в_интернете", t0, ".//tr[2]/td[2]" },
                    new Object[] { "адрес_электронной_площадки_в_интернете",        t0, ".//tr[3]/td[2]" },
                    new Object[] { "размещение_осуществляет",                       t0, ".//tr[4]/td[2]" },
                    new Object[] { "объект_закупки",                                t0, ".//tr[5]/td[2]" },
                    new Object[] { "этап_закупки",                                  t0, ".//tr[6]/td[2]" },
                    new Object[] { "сведения_о_связи_с_позицией_плана_графика",     t0, ".//tr[7]/td[2]" },
                    new Object[] { "номер_типового_контракта",                      t0, ".//tr[8]/td[2]" },
                    // информация_об_организации_осуществляющей_определение_поставщика
                    new Object[] { "организация_осуществляющая_размещение",         t1, ".//tr[1]/td[2]" },
                    new Object[] { "почтовый_адрес",                                t1, ".//tr[2]/td[2]" },
                    new Object[] { "место_нахождения",                              t1, ".//tr[3]/td[2]" },
                    new Object[] { "ответственное_должностное_лицо",                t1, ".//tr[4]/td[2]" },
                    new Object[] { "адрес_электронной_почты",                       t1, ".//tr[5]/td[2]" },
                    new Object[] { "номер_контактного_телефона",                    t1, ".//tr[6]/td[2]" },
                    new Object[] { "факс",                                          t1, ".//tr[7]/td[2]" },
                    new Object[] { "дополнительная_информация",                     t1, ".//tr[8]/td[2]" },
                    // информация_о_процедуре_закупки
                    new Object[] { "дата_и_время_начала_подачи_заявок",             t2, ".//tr[1]/td[2]" },
                    new Object[] { "дата_и_время_окончания_подачи_заявок",          t2, ".//tr[2]/td[2]" },
                    new Object[] { "место_подачи_заявок",                           t2, ".//tr[3]/td[2]" },
                    new Object[] { "порядок_подачи_заявок",                         t2, ".//tr[4]/td[2]" },
                    new Object[] { "дата_окончания_срока_рассмотрения_первых_частей", t2, ".//tr[5]/td[2]" },
                    new Object[] { "дата_проведения_аукциона_в_электронной_форме",  t2, ".//tr[6]/td[2]" },
                    new Object[] { "время_проведения_аукциона",                     t2, ".//tr[7]/td[2]" },
                    new Object[] { "дополнительная_информация2",                    t2, ".//tr[8]/td[2]" },
                    // начальная_максимальная_цена_контракта
                    new Object[] { "начальная_максимальная_цена_контракта",         t3, ".//tr[1]/td[2]" },
                    new Object[] { "валюта",                                        t3, ".//tr[2]/td[2]" },
                    new Object[] { "источник_финансирования",                       t3, ".//tr[3]/td[2]" },
                    new Object[] { "идентификационный_код_закупки",                 t3, ".//tr[4]/td[2]" },
                    new Object[] { "оплата_исполнения_контракта_по_годам",          t3, "../div[contains(@class, 'addingTbl')]/table//tr[1]/td[2]" },
                    // информация_об_объекте_закупки
                    new Object[] { "условия_запреты_и_ограничения_допуска_товаров", t4, ".//tr[1]/td[2]" },
                    new Object[] { "табличная_часть_в_формате_html",                t4, ".//tr[2]/td[1]" },
                    // преимущества_требования_к_участникам
                    new Object[] { "преимущества",                                  t5, ".//tr[1]/td[2]" },
                    new Object[] { "требования",                                    t5, ".//tr[2]/td[2]" },
                    new Object[] { "ограничения",                                   t5, ".//tr[3]/td[2]" }
                    */
            };
            return aUid;
        }
        public static Guid SaveAuctionInf(String manager, XmlNode mainBox)
        {
            Guid aUid = new Guid();

            // берём основные блоки

            // заявка
            XmlNode cardHeader = mainBox.SelectSingleNode("./div[@class='cardHeader']");

            // общая_информация_о_закупке
            XmlNode contentTabBoxBlock = mainBox.SelectSingleNode("./div[contains(@class, 'contentTabBoxBlock')]");
            XmlNode noticeTabBox = contentTabBoxBlock.SelectSingleNode("./div[contains(@class, 'noticeTabBox')]");

            // noticeTabBox содержит пары <h2 class="noticeBoxH2"> и <div class="noticeTabBoxWrapper"> с общей информацией для всех заказчиков
            // делаем из пар два параллельных списка

            // список <h2> с заголовками общей информации
            XmlNodeList noticeBoxH2List = noticeTabBox.SelectNodes("./h2");
            XmlNode[] noticeBoxH2s = new XmlNode[9];
            if (noticeBoxH2List != null)
            {
                for (int i = 0; i < Math.Min(9, noticeBoxH2List.Count); i++)
                {
                    noticeBoxH2s[i] = noticeBoxH2List[i];
                }
            }

            // список <div> с содержимым общей информации
            XmlNodeList noticeTabBoxWrapperList = noticeTabBox.SelectNodes("./div[contains(@class, 'noticeTabBoxWrapper')]");
            XmlNode[] noticeTabBoxWrappers = new XmlNode[9];
            if (noticeTabBoxWrapperList != null)
            {
                for (int i = 0; i < Math.Min(9, noticeTabBoxWrapperList.Count); i++)
                {
                    noticeTabBoxWrappers[i] = noticeTabBoxWrapperList[i];
                }
            }

            // берём элементы с информацией о требованиях заказчиков
            // их два параллельных списка

            // первый список
            XmlNodeList noticeBoxExpandList = noticeTabBox.SelectNodes("./div[contains(@class, 'noticeBoxExpand')]");
            XmlNode[] noticeBoxExpands = new XmlNode[0];
            if (noticeBoxExpandList != null)
            {
                for (int i = 0; i < noticeBoxExpandList.Count; i++)
                {
                    Int32 size = noticeBoxExpands.Length;
                    Array.Resize<XmlNode>(ref noticeBoxExpands, size + 1);
                    noticeBoxExpands[size] = noticeBoxExpandList[i];
                }
            }

            // второй список
            XmlNodeList expandRowList = noticeTabBox.SelectNodes("./div[contains(@class, 'expandRow')]");
            XmlNode[] expandRows = new XmlNode[0];
            if (expandRowList != null)
            {
                for (int i = 0; i < expandRowList.Count; i++)
                {
                    Int32 size = expandRows.Length;
                    Array.Resize<XmlNode>(ref expandRows, size + 1);
                    expandRows[size] = expandRowList[i];
                }
            }

            // записываем на SQL сервер в базу Auctions общую информацию об аукционе
            aUid = ParseAndSaveAuctionCommonInf(cardHeader, noticeBoxH2s, noticeTabBoxWrappers);

            // записываем на SQL сервер в базу Auctions информацию о заказчиках (их может быть несколько)
            ParseAndSaveAuction44FzCustomerRequirement(manager, aUid, noticeTabBoxWrappers, noticeBoxExpands, expandRows);

            return aUid;
        }
        private static Guid ParseAndSaveAuctionCommonInf(XmlNode cardHeader, XmlNode[] noticeBoxH2s, XmlNode[] noticeTabBoxWrappers)
        {
            // ищем таблицы информационных блоков по заголовку
            XmlNode t0 = null; // общая_информация_о_закупке
            XmlNode t1 = null; // информация_об_организации_осуществляющей_определение_поставщика
            XmlNode t2 = null; // информация_о_процедуре_закупки
            XmlNode t3 = null; // начальная_максимальная_цена_контракта
            XmlNode t4 = null; // информация_об_объекте_закупки
            XmlNode t5 = null; // преимущества_требования_к_участникам
            for (int i = 0; i < noticeBoxH2s.Length; i++)
            {
                XmlNode n = noticeBoxH2s[i];
                if (n != null)
                {
                    String h = n.InnerText.Trim();
                    if (h.Length >= 17)
                    {
                        XmlNode d = noticeTabBoxWrappers[i];
                        if (d != null)
                        {
                            XmlNode t = d.SelectSingleNode("./table");
                            if (t != null)
                            {
                                switch (h.Substring(0, 17))
                                {
                                    case "Общая информация ": t0 = t; break;
                                    case "Информация об орг": t1 = t; break;
                                    case "Информация о проц": t2 = t; break;
                                    case "Начальная (максим": t3 = t; break;
                                    case "Информация об объ": t4 = t; break;
                                    case "Преимущества, тре": t5 = t; break;
                                    default: break;
                                }
                            }
                        }
                    }
                }
            }

            Object[][] md44 = new Object[][] { 
                    // закупка
                    new Object[] { "номер",                                         cardHeader, "./h1" },
                    new Object[] { "дата_размещения",                               cardHeader, "./div[@class='public']" },
                    new Object[] { "кооператив",                                    cardHeader, "./div[@class='public']/span[@class='cooperative']" },
                    // общая_информация_о_закупке
                    new Object[] { "способ_определения_поставщика",                 t0, ".//tr[1]/td[2]" },
                    new Object[] { "наименование_электронной_площадки_в_интернете", t0, ".//tr[2]/td[2]" },
                    new Object[] { "адрес_электронной_площадки_в_интернете",        t0, ".//tr[3]/td[2]" },
                    new Object[] { "размещение_осуществляет",                       t0, ".//tr[4]/td[2]" },
                    new Object[] { "объект_закупки",                                t0, ".//tr[5]/td[2]" },
                    new Object[] { "этап_закупки",                                  t0, ".//tr[6]/td[2]" },
                    new Object[] { "сведения_о_связи_с_позицией_плана_графика",     t0, ".//tr[7]/td[2]" },
                    new Object[] { "номер_типового_контракта",                      t0, ".//tr[8]/td[2]" },
                    // информация_об_организации_осуществляющей_определение_поставщика
                    new Object[] { "организация_осуществляющая_размещение",         t1, ".//tr[1]/td[2]" },
                    new Object[] { "почтовый_адрес",                                t1, ".//tr[2]/td[2]" },
                    new Object[] { "место_нахождения",                              t1, ".//tr[3]/td[2]" },
                    new Object[] { "ответственное_должностное_лицо",                t1, ".//tr[4]/td[2]" },
                    new Object[] { "адрес_электронной_почты",                       t1, ".//tr[5]/td[2]" },
                    new Object[] { "номер_контактного_телефона",                    t1, ".//tr[6]/td[2]" },
                    new Object[] { "факс",                                          t1, ".//tr[7]/td[2]" },
                    new Object[] { "дополнительная_информация",                     t1, ".//tr[8]/td[2]" },
                    // информация_о_процедуре_закупки
                    new Object[] { "дата_и_время_начала_подачи_заявок",             t2, ".//tr[1]/td[2]" },
                    new Object[] { "дата_и_время_окончания_подачи_заявок",          t2, ".//tr[2]/td[2]" },
                    new Object[] { "место_подачи_заявок",                           t2, ".//tr[3]/td[2]" },
                    new Object[] { "порядок_подачи_заявок",                         t2, ".//tr[4]/td[2]" },
                    new Object[] { "дата_окончания_срока_рассмотрения_первых_частей", t2, ".//tr[5]/td[2]" },
                    new Object[] { "дата_проведения_аукциона_в_электронной_форме",  t2, ".//tr[6]/td[2]" },
                    new Object[] { "время_проведения_аукциона",                     t2, ".//tr[7]/td[2]" },
                    new Object[] { "дополнительная_информация2",                    t2, ".//tr[8]/td[2]" },
                    // начальная_максимальная_цена_контракта
                    new Object[] { "начальная_максимальная_цена_контракта",         t3, ".//tr[1]/td[2]" },
                    new Object[] { "валюта",                                        t3, ".//tr[2]/td[2]" },
                    new Object[] { "источник_финансирования",                       t3, ".//tr[3]/td[2]" },
                    new Object[] { "идентификационный_код_закупки",                 t3, ".//tr[4]/td[2]" },
                    new Object[] { "оплата_исполнения_контракта_по_годам",          t3, "../div[contains(@class, 'addingTbl')]/table//tr[1]/td[2]" },
                    // информация_об_объекте_закупки
                    new Object[] { "условия_запреты_и_ограничения_допуска_товаров", t4, ".//tr[1]/td[2]" },
                    new Object[] { "табличная_часть_в_формате_html",                t4, ".//tr[2]/td[1]" },
                    // преимущества_требования_к_участникам
                    new Object[] { "преимущества",                                  t5, ".//tr[1]/td[2]" },
                    new Object[] { "требования",                                    t5, ".//tr[2]/td[2]" },
                    new Object[] { "ограничения",                                   t5, ".//tr[3]/td[2]" }
                };

            Guid aUid = SaveAuctionCommonInf(md44);

            return aUid;
        }
        private static Guid SaveAuctionCommonInf(Object[][] md44)
        {
            Guid aUid = new Guid();
            SqlCommand cmd = new SqlCommand
            {
                Connection = new SqlConnection(Env.cnString),
                CommandText = "[Auctions].[dbo].[save_auction_inf]",
                CommandType = CommandType.StoredProcedure
            };
            foreach (Object[] nx in md44)
            {
                String name = (String)nx[0];
                XmlNode node = (XmlNode)nx[1];
                if (node != null)
                {
                    node = node.SelectSingleNode((String)nx[2]);
                    if (node != null)
                    {
                        String value = Utilities.NormString(node.InnerText);
                        if (name == "номер") value = (new Regex(@"\d{19}")).Match(value).Value;
                        if (name == "кооператив") value = "1"; // иначе сюда вообще не дойдёт и value окажется равным 0 по умолчанию
                        if (name == "дата_размещения") value = value.Replace("Размещено: ", String.Empty);
                        if (!String.IsNullOrWhiteSpace(name))
                        {
                            if (name[0] != '@') { name = "@" + name; }
                            cmd.Parameters.AddWithValue(name, value);
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
        private static void ParseAndSaveAuction44FzCustomerRequirement(String manager, Guid aUid, XmlNode[] noticeTabBoxWrappers, XmlNode[] noticeBoxExpands, XmlNode[] expandRows)
        {
            // здесь может быть два варианта
            // первый - когда списки expand пусты - сохраняем из основного блока
            // второй - сохраняем из списков

            Object[][] mdCust = null;
            if (expandRows.Length == 0 || noticeBoxExpands.Length == 0)
            {
                mdCust = new Object[][] {
                        // требования_заказчика
                        new Object[] { "наименование_заказчика",                        noticeTabBoxWrappers[1], "./table//tr[1]/td[2]" },
                        // условия_контракта
                        new Object[] { "место_доставки_товара",                         noticeTabBoxWrappers[6], "./table//tr[1]/td[2]" },
                        new Object[] { "сроки_поставки_товара",                         noticeTabBoxWrappers[6], "./table//tr[2]/td[2]" },
                        new Object[] { "сведения_о_связи_с_позицией_плана_графика",     noticeTabBoxWrappers[0], "./table//tr[7]/td[2]" },
                        new Object[] { "оплата_исполнения_контракта_по_годам",          noticeTabBoxWrappers[3], "./div[contains(@class, 'addingTbl')]/table//tr[1]/td[2]" },
                        // обеспечение_заявок
                        new Object[] { "размер_обеспечения",                            noticeTabBoxWrappers[7], "./table//tr[2]/td[2]" },
                        new Object[] { "порядок_внесения_денежных_средств",             noticeTabBoxWrappers[7], "./table//tr[3]/td[2]" },
                        new Object[] { "платежные_реквизиты",                           noticeTabBoxWrappers[7], "./table//tr[4]/td[2]" },
                        // обеспечение_исполнения_контракта
                        new Object[] { "размер_обеспечения_2",                          noticeTabBoxWrappers[8], "./table//tr[2]/td[2]" },
                        new Object[] { "порядок_предоставления_обеспечения",            noticeTabBoxWrappers[8], "./table//tr[3]/td[2]" },
                        new Object[] { "платежные_реквизиты_2",                         noticeTabBoxWrappers[8], "./table//tr[4]/td[2]" }
                    };

                SaveAuctionCustomerRequirement(manager, aUid, mdCust);
            }
            else // второй вариант - со списком поставщиков
            {
                for (int i = 0; i < noticeBoxExpands.Length; i++)
                {
                    mdCust = new Object[][] {
                            // требования_заказчика
                            new Object[] { "наименование_заказчика",                        noticeBoxExpands[i], ".//h3" },
                            // условия_контракта
                            //new Object[] { "место_доставки_товара",                         noticeTabBoxWrappers[6], "table//tr[1]/td[2]" },
                            //new Object[] { "сроки_поставки_товара",                         noticeTabBoxWrappers[6], "table//tr[2]/td[2]" },
                            new Object[] { "сведения_о_связи_с_позицией_плана_графика",     expandRows[i], ".//div[contains(@class, 'noticeTabBoxWrapper')][1]/table//tr[1]/td[2]" },
                            new Object[] { "оплата_исполнения_контракта_по_годам",          expandRows[i], ".//div[contains(@class, 'noticeTabBoxWrapper')][2]/table//tr[1]/td[2]" },
                            // обеспечение_заявок
                            new Object[] { "размер_обеспечения",                            expandRows[i], ".//div[contains(@class, 'noticeTabBoxWrapper')][3]/table//tr[2]/td[2]" },
                            new Object[] { "порядок_внесения_денежных_средств",             expandRows[i], ".//div[contains(@class, 'noticeTabBoxWrapper')][3]/table//tr[3]/td[2]" },
                            new Object[] { "платежные_реквизиты",                           expandRows[i], ".//div[contains(@class, 'noticeTabBoxWrapper')][3]/table//tr[4]/td[2]" },
                            // обеспечение_исполнения_контракта
                            new Object[] { "размер_обеспечения_2",                          expandRows[i], ".//div[contains(@class, 'noticeTabBoxWrapper')][4]/table//tr[2]/td[2]" },
                            new Object[] { "порядок_предоставления_обеспечения",            expandRows[i], ".//div[contains(@class, 'noticeTabBoxWrapper')][4]/table//tr[3]/td[2]" },
                            new Object[] { "платежные_реквизиты_2",                         expandRows[i], ".//div[contains(@class, 'noticeTabBoxWrapper')][4]/table//tr[4]/td[2]" }
                        };
                    SaveAuctionCustomerRequirement(manager, aUid, mdCust);
                }
            }
        }
        private static void SaveAuctionCustomerRequirement(String manager, Guid aUid, Object[][] mdCus)
        {
            SqlCommand cmd = new SqlCommand
            {
                Connection = new SqlConnection(Env.cnString),
                CommandText = "[Auctions].[dbo].[save_customer_requirement]",
                CommandType = CommandType.StoredProcedure
            };
            cmd.Parameters.AddWithValue("@закупка_uid", aUid);
            foreach (Object[] nx in mdCus)
            {
                String name = (String)nx[0];
                XmlNode node = (XmlNode)nx[1];
                if (node != null)
                {
                    node = node.SelectSingleNode((String)nx[2]);
                    if (node != null)
                    {
                        String value = Utilities.NormString(node.InnerText);
                        if (name == "наименование_заказчика") value = value.Replace("Требования заказчика ", String.Empty);
                        if (name[0] != '@') { name = "@" + name; }
                        cmd.Parameters.AddWithValue(name, value);
                    }
                }
            }
            using (cmd.Connection)
            {
                cmd.Connection.Open();
                cmd.ExecuteNonQuery();
            }
            return;
        }
    }
    class Fz223
    {
        public static Guid SaveAuctionInf(List<AuctionInfBlock> aibs)
        {
            Guid aUid = new Guid();
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
                String value = String.Empty;

                foreach (AuctionInfBlock block in aibs)
                {
                    if (block.Name == blockName)
                    {
                        XmlDocument doc = block.Data;
                        XmlNode node = doc.SelectSingleNode(String.Format(".//tr/td[contains(.,\"{0}\")]/following-sibling::td", keyText));
                        if (node != null)
                        {
                            value = Utilities.NormString(node.InnerText);
                            if (dbName[0] != '@') { dbName = "@" + dbName; }
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
            String result = null;

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
                        result = readStream.ReadToEnd();
                    }
                }
                else { Log.Write(response.StatusDescription); }
            }
            return result;
        }
    }
}
