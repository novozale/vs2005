Module Declarations
    Public CompanyID As String                            'строка - номер компании в Scala
    Public Year As String                                 'cтрока - год в Scala
    Public UserCode As String                             'короткое имя пользователя Scala

    Public MyConnStr As String                            'строка соединения с БД
    Public MyNETConnStr As String                         '.NET строка соединения с БД
    Public MyConn As ADODB.Connection                     'соединение с БД
    Public MyRec As ADODB.Recordset                       'Рекордсет он в Африке рекордсет

    Public MyCity As City                                 'реализация окна ввода информации по городу
    Public MyCityID As Integer                            'код города
    Public MyManufacturer As Manufacturer                 'реализация окна ввода информации по производителю
    Public MyManufacturerID As Integer                    'код производителя
    Public MySalesman As Salesman                         'реализация окна ввода информации по продавцу
    Public MySalesmanID As String                         'код продавца
    Public MyProductGroup As ProductGroup                 'реализация окна ввода информации по группе продуктов
    Public MyProductGroupID As String                     'код группы продуктов
    Public MyProductSubGroup As ProductSubGroup           'реализация окна ввода информации по подгруппе продуктов
    Public MyProductSubGroupID As String                  'код подгруппы продуктов
    Public MyProduct As Product                           'реализация окна ввода информации по продукту
    Public MyProductID As String                          'код продукта в Scala
    Public MyCustomerSelectList As CustomerSelectList     'реализация окна выборки при поиске клиентов
    Public MyCustomer As Customer                         'реализация окна ввода информации по клиенту
    Public MyCustomerID As String                         'код клиента в Scala
    Public MyDiscountGroup As DiscountGroup               'реализация окна ввода / редактирования скидки на группу товаров
    Public MyDiscountSubgroup As DiscountSubgroup         'реализация окна ввода / редактирования скидки на подгруппу товаров
    Public MyDiscountItem As DiscountItem                 'реализация окна ввода / редактирования скидки на товар
    Public MyItemList As ItemList                         'реализация окна выбора товара
    Public MyItemSelectList As ItemSelectList             'реализация окна выбора товара с условиями
    Public MyAgreedRange As AgreedRange                   'реализация окна ввода согласованного ассортимента
    Public MyBasePrice As BasePrice                       'реализация окна выбора и печати базового прайс листа
    Public MyIndPrice As IndPrice                         'реализация окна выбора и печати индивидуального прайс листа
    Public MyUploadFilesToDB As UploadFilesToDB           'реализация окна загрузки картинок в БД
    Public MyUploadFilesToCatalog As UploadFilesToCatalog 'реализация окна выгрузки картинок в каталог
    Public MyUploadInfoToSaintGobain As UploadInfoToSaintGobain 'реализация окна выгрузки информации для Сен Гобена
    Public MyDeletePictures As DeletePictures             'Реализация окна удаления неверно корреспондирующихся картинок
    Public MyTransferNamesDescrToDB As TransferNamesDescrToDB 'реализация окна переноса названий и описаний, полученных из WEB, в БД
    Public MyMatchPictAndScalaCode As MatchPictAndScalaCode 'реализация окна связывания картинки с кодом Scala
    Public MyDeletePictureFromDB As DeletePictureFromDB     'реализация окна удаления карттинок из БД
    Public MyLoadOnePictToDB As LoadOnePictToDB             'реализация окна загрузки одной картинки в БД
    Public MyCheckUpdateNamesDescr As CheckUpdateNamesDescr 'реализация окна проверки корректности названий / описаний продуктов и их загрузки в БД
    Public MyDownloadInfoFromSE As DownloadInfoFromSE       'реализация окна получения данных с сервиса Шнейдер Электрик
    Public MyDownloadInfoFromABB As DownloadInfoFromABB     'реализация окна получения данных с сервиса ABB
    Public MyUploadDataToMagento As UploadDataToMagento     'реализация окна загрузки данных на сайт magento
    Public MyUploadPicturesToMagento As UploadPicturesToMagento 'реализация окна загрузки картинок на сайт magento
    Public MyUploadAvailabilityToMagento As UploadAvailabilityToMagento 'реализация окна загрузки информации о доступном количестве на сайт Magento
    Public MyErrWindow As ErrWindow                         'реализация окна вывода сообщения об ошибке
    Public MyCASH_FullUpload As CASH_FullUpload             'реализация окна выгрузки всей незаблокированной номенклатуры на кассу
    Public MyCASH_CustomUpload As CASH_CustomUpload         'реализация окна выгрузки обобщенной ручной номенклатуры на кассу
    Public MyErrorMessage As ErrorMessage                   'Реализация окна с сообщением об ошибках

    Public MyFilterColumn As Integer                      'Колонка, по которой выставляется фильтр
    Public MyAccessToken As String = "g7rvo6kkef82uwv5isvwokk3n2ohicyh" 'Токен для обмена данными с Magento
    Public MyGlobalStr As String                            'глобальная строка
End Module
