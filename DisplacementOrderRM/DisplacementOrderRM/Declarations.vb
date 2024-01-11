Module Declarations
    Public CompanyID As String                                'строка - номер компании в Scala
    Public Year As String                                     'cтрока - год в Scala
    Public UserCode As String                                 'короткое имя пользователя Scala
    Public UserID As Integer                                  'код пользователя Scala
    Public FullName As String                                 'ФИО пользователя Scala

    Public MyConnStr As String                                'строка соединения с БД
    Public MyNETConnStr As String                             '.NET строка соединения с БД
    Public MyConn As ADODB.Connection                         'соединение с БД
    Public MyRec As ADODB.Recordset                           'Рекордсет он в Африке рекордсет

    Public WHFrom As String                                   'склад отгрузки
    Public WHFromCode As String                               'склад отгрузки - код
    Public WHTo As String                                     'склад приемки
    Public WHToCode As String                                 'склад приемки - код
    Public MyOrderID As Integer                               'Номер обобщенного заказа

    Public MyPermission As Boolean                            'Принадлежит или нет к администраторской группе

    Public MyConsolidatedOrders As ConsolidatedOrders         'реализация окна формирования обобщенных заказов (отгрузок)
    Public MyEditConsolidatedOrder As EditConsolidatedOrder   'Реализация окна создания / редактирования обобщенного заказа
End Module
