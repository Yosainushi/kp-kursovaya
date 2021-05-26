using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace KursovayaKP
{
    class Queries
    {
        public static string selectBank = "SELECT КодБанка, Название, Адрес, НомерТелефона FROM Банки WHERE Deleted = 0";
        public static string selectPochta = "SELECT КодОтделенияПочты, Название, Адрес, НомерТелефона FROM Почты WHERE Deleted = 0";
        public static string selectNalog = "SELECT КодНалога, Наименование, Сумма1Платежа  FROM Налоги WHERE Deleted = 0";
        public static string selectClient = "SELECT Налогоплательщики.КодНалогоплательщика, Налогоплательщики.Фамилия, Налогоплательщики.Имя, Налогоплательщики.Отчество, Налогоплательщики.НомерТелефона, Область.Название FROM Область INNER JOIN Налогоплательщики ON Область.КодОбласти = Налогоплательщики.КодОбласти WHERE Налогоплательщики.Deleted = 0";
        public static string selectOperacii = "SELECT Операции.КодОперации, Налогоплательщики.Фамилия, Налогоплательщики.Имя, Налогоплательщики.Отчество, Налоги.Наименование, Операции.ДатаОперации, Операции.Оплачено, ВидОплаты.ВидОплаты, Сотрудники.Фамилия FROM Налогоплательщики INNER JOIN(Налоги INNER JOIN (Сотрудники INNER JOIN (ВидОплаты INNER JOIN Операции ON ВидОплаты.КодВидаОплаты = Операции.КодВидаОплаты) ON Сотрудники.КодСотрудника = Операции.КодСотрудника) ON Налоги.КодНалога = Операции.КодНалога) ON Налогоплательщики.КодНалогоплательщика = Операции.КодНалогоплательщика WHERE Операции.Deleted = 0";
        public static string selectVidOplati = "SELECT КодВидаОплаты, ВидОплаты FROM ВидОплаты WHERE Deleted = 0";
        public static string selectOblast = "SELECT КодОбласти, Название FROM Область WHERE Deleted = 0";
    public static string selectSotrudnik = "SELECT КодСотрудника, Фамилия, Имя, Отчество, НомерТелефона FROM Сотрудники WHERE Deleted = 0";
        public static string selectPolz = "SELECT Пользователи.КодПользователя, Пользователи.Логин, Пользователи.Пароль, Пользователи.ПравоАдмина FROM Пользователи WHERE Deleted =0";
        public static string selectDolzh = "SELECT Налогоплательщики.Фамилия, Налогоплательщики.Имя,  Count(*) AS КолвоДолгов FROM Налогоплательщики INNER JOIN Операции ON Налогоплательщики.КодНалогоплательщика = Операции.КодНалогоплательщика WHERE(((Операции.КодНалогоплательщика)=[Налогоплательщики].[КодНалогоплательщика]) AND((Операции.Оплачено)= 'Нет')) GROUP BY Налогоплательщики.Фамилия, Налогоплательщики.Имя";
    }
}
