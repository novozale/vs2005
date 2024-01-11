Module declarations
    Public CompanyID As String                            'строка - номер компании в Scala
    Public Year As String                                 'cтрока - год в Scala
    Public UserCode As String                             'код пользователя Scala
    Public SalesmanCode As String                         'код продавца Scala
    Public SalesmanName As String                         'имя продавца Scala
    Public ScalaDate As Date                              'системная дата Scala

    Public MyConnStr As String                            'строка соединения с БД
    Public MyNETConnStr As String                         '.NET строка соединения с БД
    Public MyConn As ADODB.Connection                     'соединение с БД
    Public MyRec As ADODB.Recordset                       'Рекордсет он в Африке рекордсет

    Public MyDoc As Xml.XmlDocument                       'XML - документ
    Public MyHeaderNode As Xml.XmlNode                    'Узел - заголовок СФ
    Public MyFirstItemNode As Xml.XmlNode                 'Узел - заголовок запасов из СФ
    Public MyItemNodeList As Xml.XmlNodeList              'список узлов - запасов СФ

    Public MyErrorForm As ErrorForm                       'реализация окна вывода сообщения об ошибке

    Public appXLSRC As Object                             'Excel - документ
End Module
