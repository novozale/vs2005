Public Module Declarations

    Public CompanyID As String                            'строка - номер компании в Scala
    Public Year As String                                 'cтрока - год в Scala
    Public UserCode As String                             'код пользователя Scala

    Public Conn As ADODB.Connection
    Public NETConnStr As String
    Public ConnStr As String

    Public FrmAddProduct As AddProduct           ' Форма для добавления запаса на конс. склад
    Public FrmEditProduct As EditProduct     ' Форма для обновления запаса на конс. склад    
    Public Rec As ADODB.Recordset

    Public SelectList As ItemSelectList
    Public SelectList2 As ItemSelectList2

    Public IsSuccess As Boolean

    Public MinQty As Double   'Минимальный уровень запасов
    Public MaxQty As Double   'Максимальный уровень запасов    


End Module
