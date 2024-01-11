Module Functions
    Public Sub InitMyConn(ByVal IsSystem As Boolean)
        '////////////////////////////////////////////////////////////////////////////////////////
        '// Инициализация соединения с БД, чтение глобальных переменных
        '//
        '////////////////////////////////////////////////////////////////////////////////////////
        Dim Scala As New SfwIII.Application

        On Error GoTo MyCatch
        If MyConn Is Nothing Then
            MyConn = New ADODB.Connection
            MyConn.CursorLocation = 3
            MyConn.CommandTimeout = 600
            MyConn.ConnectionTimeout = 300
            If Declarations.MyConnStr = "" Then
                Declarations.MyConnStr = Scala.ActiveProcess.UserContext.GetConnectionString(1)
                'Declarations.MyConnStr = "Provider=SQLOLEDB;User ID=sa;Password=sqladmin;DATABASE=ScaDataDB;SERVER=SQLCLS"
                Declarations.MyNETConnStr = Replace(Declarations.MyConnStr, "Provider=SQLOLEDB;", "")
                Declarations.MyNETConnStr = Declarations.MyNETConnStr & ";Timeout=0;"
            End If
            If IsSystem = True Then
                MyConn.Open(Replace(Declarations.MyConnStr, "ScaDataDB", "ScalaSystemDB"))
            Else
                MyConn.Open(Declarations.MyConnStr)
            End If
            If Declarations.CompanyID = "" Then
                Declarations.CompanyID = Scala.ActiveProcess.CommonVars.CompanyCode
            End If
            If Declarations.Year = "" Then
                Declarations.Year = Mid(Scala.ActiveProcess.CommonVars.FiscalYear, 3)
            End If
        End If
        Exit Sub
MyCatch:
        MsgBox(Err.Description, vbCritical, "Ошибка Functions 1")
    End Sub

    Public Sub trycloseMyRec()
        '////////////////////////////////////////////////////////////////////////////////////////
        '//Попытка закрытия рекордсета
        '//
        '////////////////////////////////////////////////////////////////////////////////////////
        On Error Resume Next
        MyRec.Close()
    End Sub

    Public Sub InitMyRec(ByVal IsSystem As Boolean, ByVal sql As String)
        '////////////////////////////////////////////////////////////////////////////////////////
        '//Открытие рекордсета
        '//
        '////////////////////////////////////////////////////////////////////////////////////////
        Dim MyErr

        On Error GoTo MyCatch
        InitMyConn(IsSystem)
        If MyRec Is Nothing Then
            MyRec = New ADODB.Recordset
        End If
        trycloseMyRec()
        MyRec.Open(sql, MyConn)
        If MyConn.Errors.Count > 0 Then
            For Each MyErr In MyConn.Errors
                Err.Raise(MyErr.Number, MyErr.Source, MyErr.Description)
            Next MyErr
        End If
        Exit Sub
MyCatch:
        MsgBox(Err.Description, vbCritical, "Ошибка Functions 2")
    End Sub

    Public Function CheckWEBOrNot(ByVal MyOrder As String) As Integer
        '////////////////////////////////////////////////////////////////////////////////////////
        '//
        '// проверка - является ли данный заказ заказом с WEB сайта
        '//
        '////////////////////////////////////////////////////////////////////////////////////////
        Dim MySQLStr As String

        MySQLStr = "SELECT     COUNT(ID) AS CC "
        MySQLStr = MySQLStr & "FROM tbl_WEB_OrderNum "
        MySQLStr = MySQLStr & "WHERE (ScaOrderNUm = N'" & MyOrder & "')"
        InitMyConn(False)
        InitMyRec(False, MySQLStr)
        If Declarations.MyRec.BOF = True And Declarations.MyRec.EOF = True Then
            CheckWEBOrNot = 0
        Else
            Declarations.MyRec.MoveFirst()
            If Declarations.MyRec.Fields("CC").Value = 0 Then
                CheckWEBOrNot = 0
            Else
                CheckWEBOrNot = 1
            End If
        End If
        trycloseMyRec()
    End Function

    Public Function GetCardPayment(ByVal OrderNumber As String) As Double
        '////////////////////////////////////////////////////////////////////////////////////////
        '//
        '// Получение суммы оплаченного по карте на WEB сайте
        '//
        '////////////////////////////////////////////////////////////////////////////////////////
        Dim MySQLStr As String

        MySQLStr = "SELECT tbl_WEB_MGPayments.OrderSumm "
        MySQLStr = MySQLStr & "FROM  tbl_WEB_MGPayments INNER JOIN "
        MySQLStr = MySQLStr & "tbl_WEB_OrderNum ON tbl_WEB_MGPayments.OrderNum = tbl_WEB_OrderNum.WebOrderNum "
        MySQLStr = MySQLStr & "WHERE (tbl_WEB_OrderNum.ScaOrderNUm = N'" & Trim(OrderNumber) & "')"
        InitMyConn(False)
        InitMyRec(False, MySQLStr)
        If Declarations.MyRec.BOF = True And Declarations.MyRec.EOF = True Then
            GetCardPayment = 0
        Else
            GetCardPayment = Declarations.MyRec.Fields("OrderSumm").Value
        End If
        trycloseMyRec()
    End Function
End Module
