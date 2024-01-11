Module Declarations

    Public CompanyID As String                            'строка - номер компании в Scala
    Public Year As String                                 'cтрока - год в Scala
    Public UserCode As String                             'код пользователя Scala

    Public MyConnStr As String                            'строка соединения с БД
    Public MyNETConnStr As String                         '.NET строка соединения с БД
    Public MyConn As ADODB.Connection                     'соединение с БД
    Public MyRec As ADODB.Recordset                       'Рекордсет он в Африке рекордсет

    Public MySuccess As Boolean                           'Успешность выполнения операции
    Public LoadFlag As Integer                            'Загрузка формы не завершена (0) или завершена (1)

    Public MyAddPriceValue As AddPriceValue               'реализация окна добавления значения прайс - листа
    Public MyEditPriceValue As EditPriceValue             'реализация окна редактирования значения прайс - листа
    Public MyShipmentCost As ShipmentCost                 'реализация окна ввода факта стоимости доставки (список)
    Public MyAddShipmentCost As AddShipmentCost           'реализация окна добавления стоимости доставки
    Public MySupplierSelect As SupplierSelect             'реализация окна поиска поставщиков
    Public MySupplierSelectList As SupplierSelectList     'реализация окна выборки при поиске поставщиков
    Public MyAddSInvoice As AddSInvoice                   'реализация окна добавления инвойса на продажу
    Public MyAddPInvoice As AddPInvoice                   'реализация окна добавления инвойса на продажу
    Public MyAddRelOrder As AddRelOrder                   'реализация окна добавления инвойса на продажу

    Public Destination As String                          'пункт назначения (или "Средняя по региону" за 100 километров  по региону при этом PriceType д.б. 1)
    Public PriceType As Integer                           'тип прайса 0 фиксированный, 1 - за 100 километров .
    Public PriceFrom As Double                            'от какого значения прайс (например, от 0 кг)по региону
    Public PriceTo As Double                              'до какого значения прайс (например, до 100 кг)
    Public PriceVal As Double                             'собственно значение прайса
    Public MinCost As Double                              'минимальное значение

    Public MyRecordID As String                           'GUID записи о стоимости доставки
End Module
