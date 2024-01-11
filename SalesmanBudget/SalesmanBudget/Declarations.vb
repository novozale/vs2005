Module Declarations
    Public CompanyID As String                              'строка - номер компании в Scala
    Public Year As String                                   'cтрока - год в Scala
    Public UserCode As String                               'короткое имя пользователя Scala

    Public MyConnStr As String                              'строка соединения с БД
    Public MyNETConnStr As String                           '.NET строка соединения с БД
    Public MyConn As ADODB.Connection                       'соединение с БД
    Public MyRec As ADODB.Recordset                         'Рекордсет он в Африке рекордсет

    Public StrTotalStart As Integer                         'в какой строке начинается тело таблицы итоговых данных
    Public IndustryQTY As Integer                           'кол-во строк отраслей
    Public TypeQTY As Integer                               'кол-во строк типов
    Public MarketQTY As Integer                             'кол-во строк рынков
    Public IKAQTY As Integer                                'кол-во строк типов IKA
    Public ActiveAreaStart As Integer                       'строка - начало списка активных клиентов
    Public ActiveAreaFinish As Integer                      'строка - конец списка активных клиентов
    Public PassiveAreaStart As Integer                      'строка - начало списка неактивных клиентов
    Public PassiveAreaFinish As Integer                     'строка - Конец списка неактивных клиентов
    Public NewAreaStart As Integer                          'строка - начало списка новых клиентов
    Public NewAreaFinish As Integer                         'строка - Конец списка новых клиентов
End Module
