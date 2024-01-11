Module Declarations
    Public CompanyID As String                            'строка - номер компании в Scala
    Public Year As String                                 'cтрока - год в Scala
    Public UserCode As String                             'код пользователя Scala
    Public PurchCode As String                            'код закупщика Scala
    Public PurchName As String                            'имя закупщика Scala
    Public IsManager As Integer                           'признак - является ли менеджером

    Public MyConnStr As String                            'строка соединения с БД
    Public MyNETConnStr As String                         '.NET строка соединения с БД
    Public MyConn As ADODB.Connection                     'соединение с БД
    Public MyRec As ADODB.Recordset                       'Рекордсет он в Африке рекордсет

    Public MyRequestNum As Integer                        'номер запроса на поиск поставщика
    Public MyItemSrchID As Integer                        'ID товара в предложенном решении
    Public ImportFileName As String                       'имя файла для импорта данных
    Public MySupplierID As Integer                        'код поставщика для импорта данных
    Public MySupplierCode As String                       'код поставщика в Scala для импорта данных
    Public MySupplierName As String                       'название поставщика для импорта данных
    Public ExcelVersion As String                         'версия файла для импорта 
    Public MyRowIndex As Integer                          'ID строки в таблице (для контекстного меню)

    Public MySupplierSelect As SupplierSelect             'реализация окна поиска поставщиков
    Public MySupplierSelectList As SupplierSelectList     'реализация окна выборки при поиске поставщиков
    Public MyAddItem As AddItem                           'реализация окна добавления товара в запрос
    Public MyItemSelect As ItemSelect                     'Реализация окна выбора товара Scala
    Public MyItemSelectList As ItemSelectList             'реализация окна выбора товара по критериям
    Public MyAttachmentsList As AttachmentsList           'реализация окна аттачментов
    Public MySearcherList As SearcherList                 'реализация окна списка поисковиков
    Public MyOrdersList As OrdersList                     'реализация окна проверки заказов 0 типа

    Public SortColumnNum As Integer                       'Номер колонки в пердложении для сортировки
    Public SortColOrder As System.ComponentModel.ListSortDirection 'направление сортировки
End Module
