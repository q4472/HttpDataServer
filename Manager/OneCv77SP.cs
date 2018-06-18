using Nskd.Oc77;
using System;
using System.Collections.Generic;
using System.Data;
using System.Diagnostics;
using System.Globalization;
using System.Reflection;
using System.Text.RegularExpressions;
using System.Threading;

namespace Manager
{
    public class OcStoredProcedure
    {
        private static Object thisLock = new Object();
        public static String Exec0(DataTable paymens, Boolean glFlag)
        {
            String msg = "q";
            lock (thisLock)
            {
                // Запускаем процесс локальной 1с или подключаемся к уже запущенному.
                // Root COM object 1c 'System'.
                V77Connection v77cn = OcConnection.Open(); // (new OcConnection(glFlag)).Open();
                if ((v77cn != null) && (v77cn.Root != null) && (v77cn.Root.ComObject != null) && (v77cn.IsConnected))
                {
                    msg += " подключились";
                    V77System root = v77cn.Root;
                    //DataRow p = paymens.Rows[0];

                    // Вставка новых элементов в справочник ПереводыДляОбеспеченияКонтрактов.
                    V77Reference ps = (V77Reference)root.CreateObject("Справочник.ПереводыДляОбеспеченияКонтрактов");
                    foreach (DataRow p in paymens.Rows)
                    {
                        // Проверяем может там уже есть такой
                        //ps.FindByAttribute()

                        // Создаём новый элемент справочника ПереводыДляОбеспеченияКонтрактов
                        ps.New();
                        // Изменяем содержимое полей.
                        ps.SetAttrib("Наименование", p["НомерТоргов"]);
                        ps.SetAttrib("ПереводДокументДата", p["Дата"]);
                        ps.SetAttrib("ПереводДокументНомер", p["Номер"]);
                        ps.SetAttrib("ПереводДокументСумма", p["Сумма"]);
                        ps.SetAttrib("ПереводКонтрагентНаименование", p["КонтрагентНаименование"]);
                        ps.SetAttrib("ПереводКонтрагентНаименованиеПолное", p["КонтрагентНаименованиеПолное"]);
                        ps.SetAttrib("ПереводКонтрагентИнн", p["КонтрагентИНН"]);
                        ps.SetAttrib("ПереводКонтрагентКпп", p["КонтрагентКПП"]);
                        // Сохранить.
                        ps.Write();
                    }
                }
            }
            return msg;
        }
    }
}
