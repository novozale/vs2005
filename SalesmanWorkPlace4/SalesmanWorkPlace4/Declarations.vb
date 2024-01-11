Module Declarations

    Public CompanyID As String                            'строка - номер компании в Scala
    Public Year As String                                 'cтрока - год в Scala
    Public UserCode As String                             'код пользователя Scala
    Public SalesmanCode As String                         'код продавца Scala
    Public SalesmanName As String                         'имя продавца Scala

    Public MyConnStr As String                            'строка соединения с БД
    Public MyNETConnStr As String                         '.NET строка соединения с БД
    Public MyConn As ADODB.Connection                     'соединение с БД
    Public MyRec As ADODB.Recordset                       'Рекордсет он в Африке рекордсет

    Public MyEditHeader As EditHeader                     'реализация окна редактирования заголовка
    Public MyCustomerSelect As CustomerSelect             'реализация окна поиска клиентов
    Public MyCustomerSelectList As CustomerSelectList     'реализация окна выборки при поиске клиентов
    Public MyOrderLines As OrderLines                     'реализация окна ввода строк предложения
    Public MySupplierSelect As SupplierSelect             'реализация окна поиска поставщиков
    Public MySupplierSelectList As SupplierSelectList     'реализация окна выборки при поиске поставщиков
    Public MyItemSelectList As ItemSelectList             'реализация окна выборки при поиске запасов
    Public MyShowWHDetails As ShowWHDetails               'реализация окна информации по складам
    Public MyShowBatchInfo As ShowBatchInfo               'реализация окна информации по партиям
    Public MyALTItems As ALTItems                         'реализация окна со списком альтернативных запасов
    Public MyAddToOrder As AddToOrder                     'реализация окна добавления запаса в заказ
    Public MyEditInOrder As EditInOrder                   'реализация окна редактирования запаса в заказе
    Public MyItemSelect As ItemSelect                     'реализация окна поиска запаса
    Public MySelectItemBySuppCode As SelectItemBySuppCode 'реализация окна выбора товара по коду товара поставщика
    Public MySendReturnProposal As SendReturnProposal     'реализация окна передачи / возврата коммерческого предложения
    Public MyShipmentsCost As ShipmentsCost               'реализация окна ввода стоимости доставки
    Public MyEDOInfo As EDOInfo                           'реализация окна ввода дополнительной информации по ЭДО
    Public MyEstimatedIncome As EstimatedIncome           'реализация окна с информацией об ожидаемом приходе
    Public MySearchSupplier As SearchSupplier             'реализация окна Списка заявок на поиск поставщика
    Public MyEditRequest As EditRequest                   'реализация окна создания / редактирования заявки на поиск поставщика
    Public MyAttachmentsList As AttachmentsList           'реализация окна присоединенных файлов
    Public MyAddItem As AddItem                           'реализация окна добавления товара в запрос
    Public MyCPList As CPList                             'Реализация окна выбора КП
    Public MyContactInfo As ContactInfo                   'реализация окна списка контактов из  CRM
    Public MySalesCommentsToProposal As SalesCommentsToProposal  'реализация окна ввода комментариев по строке предложения
    Public MyCommentAndCancelReason As CommentAndCancelReason 'реализация окна ввода комментариев и причин отказа в случае отказа
    Public MyRestoreSearch As RestoreSearch               'реализация окна восстановления поиска после паузы
    Public MyCorrectRequestDate As CorrectRequestDate     'реализация окна корректировки запрошенной даыт предоставления КП
    Public MyPrintProposal As PrintProposal               'реализация окна печати КП

    Public MyOrderNum As String                           'номер заказа
    Public CurrencyCode As Integer                        'код валюты (для сумм заказа)
    Public CurrencyName As String                         'название валюты
    Public CurrencyValue As Double                        'значение валюты в рублях на текущую дату
    Public CurrencyValueOrder As Double                   'значение валюты в рублях на дату создания заказа
    Public MyRequestNum As Integer                        'номер запроса на поиск поставщика
    Public MyItemSrchID As Integer                        'ID товара в запросе на поиск
    Public MyItemPropID As Integer                        'ID товара в предложении поиска
    Public MyCPID As String                               'ID коммерческого предложения для переноса в него предложения от поисковиков
    Public MyRowIndex As Integer                          'ID строки в таблице (для контекстного меню)
    Public MyRez1 As Integer                              'результат операции

    Public CustomerNumber                                 'код клиента
    Public WHNum                                          'номер склада
    Public MySuccess As Boolean                           'Успешность выполнения операции
    Public MyItemID As String                             'код запаса
    Public MyItemSuppID As String                         'код запаса поставщика
    Public MyItemName As String                           'имя запаса
    Public MyQty As Double                                'количество заказанного
    Public MyUOM As Integer                               'код единицы измерения
    Public MySum As Double                                'сумма заказанного
    Public MySS As Double                                 'себестоимость
    Public MyDiscount As String                           'скидка
    Public MyItemCode As String                           'код товара в Scala для импорта спецификации из Excel
    Public MySuppID As String                             'код поставщика
    Public MySuppName                                     'название поставщика

    Public ImportFileName As String                       'имя файла для импорта заказа
    Public ExcelVersion As String                         'версия файла для импорта 

    Public DeliveryDate As Date                           'предполагаемая дата отгрузки
    Public DeliveryDateFlag As Integer                    'флаг - применять дату ко всем строкам (1) или нет(0)
    Public WeekQTY As Double                              'срок поставки в неделях (должен быть >= 0; 0 - в наличии на складе)
    Public DelWeekQTY As Double                           'срок доставки до клиента в неделях (должен быть >= 0; 0 - если нет доставки (самовывоз))

    Public MinMarginLevelManager As Double                'минимальное значение маржи, утверждаемое менеджером
    Public MinMarginLevelDirector As Double               'минимальное значение маржи, утверждаемое директором

    Public MyCCPermission As Boolean                      'Входит  или нет в группу CRMManagers
    Public MyCPPermission As Boolean                      'Входит  или нет в группу ProposalManager
    Public MyPermission As Boolean                        'Входит  или нет в группу CRMDirector

    Public IsSelfDelivery As Integer                      'Самовывоз (1) или нет (0)
    Public MyOperationResult As Integer                   'результат выполнения операции (1) - выполнена (0) - нет

    Public SortColumnNum As Integer                       'Номер колонки в пердложении для сортировки
    Public SortColOrder As System.ComponentModel.ListSortDirection 'направление сортировки
End Module
