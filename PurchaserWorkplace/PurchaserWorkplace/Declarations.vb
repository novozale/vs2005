Module Declarations
    Public CompanyID As String                            'строка - номер компании в Scala
    Public Year As String                                 'cтрока - год в Scala
    Public UserCode As String                             'короткое имя пользователя Scala
    Public UserID As Integer                              'код пользователя Scala
    Public UserName As String                             'логин пользователя Scala
    Public FullName As String                             'ФИО пользователя Scala

    Public MyConnStr As String                            'строка соединения с БД
    Public MyNETConnStr As String                         '.NET строка соединения с БД
    Public MyConn As ADODB.Connection                     'соединение с БД
    Public MyRec As ADODB.Recordset                       'Рекордсет он в Африке рекордсет

    Public MySupplierCode As String                       'код поставщика
    Public MyWH As String                                 'код склада
    Public MyOrderID As String                            'Номер обобщенного заказа
    Public MyPurchOrderSum As Double                      'Сумма обобщенного заказа

    Public MySupplierSelectList As SupplierSelectList               'Реализация окна выбора поставщиков
    Public MyConsolidatedOrders As ConsolidatedOrders               'реализация окна формирования обобщенных заказов
    Public MyEditConsolidatedOrder As EditConsolidatedOrder         'Реализация окна создания / редактирования обобщенного заказа
    Public MyImportCommonConfirmation As ImportCommonConfirmation   'реализация окна импорта стандартного подтверждения о поставке
    Public MyErrorForm As ErrorForm                                 'реализация окна вывода сообщения об ошибке
    Public MyCreateSPTasks As CreateSPTasks                         'реализация окна создания задач в шарепойнте
    Public MySupplierInfo As SupplierInfo                           'реализация окна информации по поставщику по всем складам

End Module
