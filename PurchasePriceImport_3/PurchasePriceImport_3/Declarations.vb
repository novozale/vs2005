Module Declarations
    Public CompanyID As String                            'строка - номер компании в Scala
    Public Year As String                                 'cтрока - год в Scala
    Public UserCode As String                             'код пользователя Scala

    Public MyConnStr As String                            'строка соединения с БД
    Public MyNETConnStr As String                         '.NET строка соединения с БД
    Public MyConn As ADODB.Connection                     'соединение с БД
    Public MyRec As ADODB.Recordset                       'Рекордсет он в Африке рекордсет

    Public MyImportPriceShort As ImportPriceShort         'Реализация окна сокращенного импорта
    Public MyImportPriceFull As ImportPriceFull           'Реализация окна полного импорта
    Public MyErrorMessage As ErrorMessage                 'Реализация окна с сообщением об ошибках
End Module
