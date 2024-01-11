Module Declarations
    Public CompanyID As String                            'строка - номер компании в Scala
    Public Year As String                                 'cтрока - год в Scala
    Public UserCode As String                             'короткое имя пользователя Scala
    Public UserID As Integer                              'код пользователя Scala
    Public FullName As String                             'ФИО пользователя Scala

    Public MyConnStr As String                            'строка соединения с БД
    Public MyNETConnStr As String                         '.NET строка соединения с БД
    Public MyConn As ADODB.Connection                     'соединение с БД
    Public MyRec As ADODB.Recordset                       'Рекордсет он в Африке рекордсет

    Public MySupplierCode As String                       'код поставщика
    Public MyWH As String                                 'код склада

    Public MyErrorForm As ErrorForm                       'реализация окна вывода сообщения об ошибке
End Module
