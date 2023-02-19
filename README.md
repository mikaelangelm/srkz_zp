# srkz_zp
Формирование Акта и Контракта с задачами из bitrix24.ru

👋 я - бот заполняющий word-шаблоны Акта и Контракта на основании переданных в JSON параметров вида \_<имя параметра>\_. 
Данные нигде не хранятся, кроме чата telegram (и оперативной памяти серверов хостинга в момент заполнения).
Для заполнения шаблона отправьте параметры в виде: 

 {"\_Итог\_": 50000.00, "\_Адрес\_": "адрес регистрации", "\_Паспорт\_": "паспортные данные", "\_Банк\_": "АО «Тинькоф-банк»", "\_ИНН\_": "инн", "\_СНИЛС\_": "снилс", "\_Счет\_": "банк.счет", "\_КС\_": "корр-счет", "\_БИК\_": "БИК банка"}
 
 - при каждом запуске происходит проверка наличия на портале Б24 у Сотрудника значения поля telegram, соответствующего логину telegram. Если не найдено 1 совпадение, то заполнение невозможно
 - все параметры являются необязательными, но незаполненные будут отображаться в виде \_<имя параметра>\_
 - (применительно к текущему веб-серверу) бот уходит в сон при долгом (~15 мин) отсутствии запросов. Выход из сна ~1 мин
 - параметр \_ТЗДата\_ - дата формирования актов. По умолчанию - сегодняшняя. Т.е. будут подбираться первые 8 завершенных до конца прошлого месяца задач из Б24, отсортированных по Убыванию разности времени старта и закрытия Задачи. Дату можно переопределить.
