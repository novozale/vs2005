Module Declarations
    Public CompanyID As String                            'строка - номер компании в Scala
    Public Year As String                                 'cтрока - год в Scala
    Public UserCode As String                             'короткое имя пользователя Scala
    Public UserID As Integer                              'код пользователя Scala
    Public UserName As String                             'логин пользователя Scala
    Public FullName As String                             'ФИО пользователя Scala
    Public SalesmanCode As String                         'код продавца
    Public SalesmanName As String                         'имя продавца

    Public MyConnStr As String                            'строка соединения с БД
    Public MyNETConnStr As String                         '.NET строка соединения с БД
    Public MyConn As ADODB.Connection                     'соединение с БД
    Public MyRec As ADODB.Recordset                       'Рекордсет он в Африке рекордсет

    Public MyOrderNum As String                           'номер заказа
    Public CustomerID As String                           'код покупателя
    Public CreditAmount As Double                         'размер кредита в рублях
    Public CreditInDays As Integer                        'размер кредита в днях
    Public MinMargin As Double                            'минимальная маржа в заказе
    Public MyPermission As Boolean                        'можно или нет отгружать сверх
    Public MyMarginReason As String                       'причина отгрузки с маржой ниже установленной
    Public OrderID As String                              'номер заказа
    Public OrderSum As Double                             'сумма заказа
    Public CurrCode As String                             'валюта
    Public Avance1Type As Double                          'сумма авансов 1 типа по заказу
    Public Avance2Type As Double                          'сумма авансов 2 типа по заказу
    Public MyPayment As Double                            'сумма платежей по заказу
    Public InvoiceDebt As Double                          'денежный долг по счетам фактурам RUR
    Public OrderDebt As Double                            'денежный долг по заказам разрешенным к отгрузке (без счетов фактур) RUR
    Public OverduePaymentQTY As Integer                   'кол-во инвойсов с просроченной оплатой
    Public Overdue As Double                              'сумма инвойсов с просроченной оплатой
    Public CmdToShip As Boolean                           'дана команда на отгрузку сверх лимита или нет
    Public MyReason As String                             'причина отгрузки
    Public MyShipmentsID As String                        'Номер отгрузки
    Public MyOperationFlag As Integer                     'флаг - выполнена операция или нет
    Public MyCustomerCode As String                       'код покупателя
    Public MyWH As String                                 'код склада

    Public MyOrderDetails As OrderDetails                 'Реализация окна деталей по заказу
    Public MyLowMarginReason As LowMarginReason           'реализация окна ввода причины отгрузки с низкой маржой
    Public MyCreditDialog As CreditDialog                 'реализация окна вывода информации по кредитному клиенту
    Public MyNonCreditDialog As NonCreditDialog           'реализация окна вывода информации по некредитному клиенту
    Public MyCreditInfo As CreditInfo                     'реализация окна вывода детальной информации по кредитному клиенту
    Public MyNonCreditInfo As NonCreditInfo               'реализация окна вывода детальной информации по некредитному клиенту
    Public MySendInfo As SendInfo                         'реализация окна отправки информации клиенту по EMail
    Public MyShipment As Shipment                         'Реализация окна огрузки
    Public MyContactInfo As ContactInfo                   'реализация окна списка контактов из  CRM
    Public MyDelAddresses As DelAddresses                 'реализация окна списка адресов из Scala

    Public MyProjectIsApproved As Integer                 '---Утвержден ли заказ - проектный или непроектный
    Public MyProjectID As String                          '---ID проекта

    Public IsWEBOrder As Integer                          '---является ли заказ заказом с WEB сайта (1) или нет (0)
End Module
