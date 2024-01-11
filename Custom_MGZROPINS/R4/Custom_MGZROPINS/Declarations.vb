Module Declarations

    Public CompanyID As String                            'строка - номер компании в Scala
    Public Year As String                                 'cтрока - год в Scala
    Public UserCode As String                             'код пользователя Scala

    Public MyConnStr As String                            'строка соединения с БД
    Public MyNETConnStr As String                         '.NET строка соединения с БД
    Public MyConn As ADODB.Connection                     'соединение с БД
    Public MyRec As ADODB.Recordset                       'Рекордсет он в Африке рекордсет

    Public MyItemSelectList As ItemSelectList             'реализация окна выборки при поиске запасов
    Public MyItemSelectList2 As ItemSelectList2           'реализация окна выборки при поиске запасов
    Public MyAddCustom As AddCustom                       'реализация окна добавления ручныз значений МЖЗ, ROP и страх. запаса
    Public MyEditCustom As EditCustom                     'реализация окна редактирования ручныз значений МЖЗ, ROP и страх. запаса

    Public MySuccess As Boolean                           'Успешность выполнения операции

    Public MyMGZ As Double                                'значение МЖЗ
    Public MyROP As Double                                'значение ROP
    Public MyInsuranceLVL As Double                       'значение страхового запаса
    Public MyDueDate As Date                              'До какой даты действуют ручные значения

    Public MyWorkLevel As Integer                         'Уровень работы (0) - DC, (1) - Локальный склад

End Module
