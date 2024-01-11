Public Class EditConsolidatedOrder
    Public StartParam As String

    Private Sub Button13_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button13.Click
        '////////////////////////////////////////////////////////////////////////////////
        '//
        '// Выход из окна без сохранения
        '//
        '////////////////////////////////////////////////////////////////////////////////

        Me.Close()
    End Sub

    Private Sub EditConsolidatedOrder_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles Me.KeyDown
        '////////////////////////////////////////////////////////////////////////////////
        '//
        '// Запрет закрытия окна по ALT+F4
        '//
        '////////////////////////////////////////////////////////////////////////////////

        If e.KeyData = Keys.Alt + Keys.F4 Then
            e.Handled = True
        End If
    End Sub

    Private Sub EditConsolidatedOrder_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        '////////////////////////////////////////////////////////////////////////////////
        '//
        '// загрузка значений в окно
        '//
        '////////////////////////////////////////////////////////////////////////////////
        Dim MySQLStr As String

        If StartParam = "Create" Then '---создание нового заказа
            Declarations.MyOrderID = 0
            Label1.Text = ""
            Label4.Text = Declarations.WHFrom
            Label7.Text = Declarations.WHTo
            Label9.Text = Declarations.FullName
            DateTimePicker1.Value = Now()
            DateTimePicker2.Value = Now()
            TextBox1.Text = ""
            TextBox2.Text = ""
        Else
            MySQLStr = "SELECT tbl_DisplacementOrder_ShipmentInfo.ID, tbl_DisplacementOrder_ShipmentInfo.ShipmentDate, tbl_DisplacementOrder_ShipmentInfo.ReceivingDate, "
            MySQLStr = MySQLStr & "ISNULL(ScalaSystemDB.dbo.ScaUsers.FullName, '') AS FullName, ISNULL(tbl_DisplacementOrder_ShipmentInfo.TransportCompanyDocNum, '') "
            MySQLStr = MySQLStr & "AS TransportCompanyDocNum, ISNULL(tbl_DisplacementOrder_ShipmentInfo.Comments, '') AS Comments "
            MySQLStr = MySQLStr & "FROM ScalaSystemDB.dbo.ScaUsers RIGHT OUTER JOIN "
            MySQLStr = MySQLStr & "tbl_DisplacementOrder_ShipmentInfo ON ScalaSystemDB.dbo.ScaUsers.UserID = tbl_DisplacementOrder_ShipmentInfo.UserID "
            MySQLStr = MySQLStr & "WHERE (tbl_DisplacementOrder_ShipmentInfo.ID = " & Declarations.MyOrderID & ") "
            InitMyConn(False)
            InitMyRec(False, MySQLStr)
            If Declarations.MyRec.EOF = True And Declarations.MyRec.BOF = True Then
                trycloseMyRec()
                MsgBox("Данный заказ не существует, возможно, удален другим пользователем. Обновите данные в окне заказов.", MsgBoxStyle.Critical, "Внимание!")
                Me.Close()
            Else
                Declarations.MyRec.MoveFirst()
                Label1.Text = Declarations.MyOrderID
                Label4.Text = Declarations.WHFrom
                Label7.Text = Declarations.WHTo
                Label9.Text = Declarations.MyRec.Fields("FullName").Value
                DateTimePicker1.Value = Declarations.MyRec.Fields("ShipmentDate").Value
                DateTimePicker2.Value = Declarations.MyRec.Fields("ReceivingDate").Value
                TextBox1.Text = Declarations.MyRec.Fields("TransportCompanyDocNum").Value
                TextBox2.Text = Declarations.MyRec.Fields("Comments").Value
                trycloseMyRec()
                TextBox1.Select()
            End If
        End If
    End Sub

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        '////////////////////////////////////////////////////////////////////////////////
        '//
        '// Выход с сохранением значений
        '//
        '////////////////////////////////////////////////////////////////////////////////

        If CheckData() = True Then
            If StartParam = "Create" Then
                SaveNewData()
            Else
                UpdateData()
            End If
            Me.Close()
        End If
    End Sub

    Private Function Checkdata() As Boolean
        '////////////////////////////////////////////////////////////////////////////////
        '//
        '// Проверка введенных значений
        '//
        '////////////////////////////////////////////////////////////////////////////////

        If DateTimePicker1.Value > DateTimePicker2.Value Then
            MsgBox("Дата отправки заказа не должна быть больше даты приемки", MsgBoxStyle.Critical, "Внимание!")
            Checkdata = False
            DateTimePicker1.Focus()
        Else
            Checkdata = True
        End If
    End Function

    Private Sub SaveNewData()
        '////////////////////////////////////////////////////////////////////////////////
        '//
        '// Сохранение нового заказа
        '//
        '////////////////////////////////////////////////////////////////////////////////
        Dim MySQLStr As String

        MySQLStr = "INSERT INTO tbl_DisplacementOrder_ShipmentInfo "
        MySQLStr = MySQLStr & "(WHFrom, WHTo, ShipmentDate, ReceivingDate, UserID, TransportCompanyDocNum, Comments) "
        MySQLStr = MySQLStr & "VALUES (N'" & Declarations.WHFromCode & "', "
        MySQLStr = MySQLStr & "N'" & Declarations.WHToCode & "', "
        MySQLStr = MySQLStr & "CONVERT(DATETIME, '" & Microsoft.VisualBasic.Right("00" & CStr(DatePart(DateInterval.Day, DateTimePicker1.Value)), 2) & "/" & Microsoft.VisualBasic.Right("00" & CStr(DatePart(DateInterval.Month, DateTimePicker1.Value)), 2) & "/" & CStr(DatePart(DateInterval.Year, DateTimePicker1.Value)) & "', 103), "
        MySQLStr = MySQLStr & "CONVERT(DATETIME, '" & Microsoft.VisualBasic.Right("00" & CStr(DatePart(DateInterval.Day, DateTimePicker2.Value)), 2) & "/" & Microsoft.VisualBasic.Right("00" & CStr(DatePart(DateInterval.Month, DateTimePicker2.Value)), 2) & "/" & CStr(DatePart(DateInterval.Year, DateTimePicker2.Value)) & "', 103), "
        MySQLStr = MySQLStr & Declarations.UserID & ", "
        MySQLStr = MySQLStr & "N'" & TextBox1.Text & "', "
        MySQLStr = MySQLStr & "N'" & TextBox2.Text & "') "
        InitMyConn(False)
        Declarations.MyConn.Execute(MySQLStr)
    End Sub

    Private Sub UpdateData()
        '////////////////////////////////////////////////////////////////////////////////
        '//
        '// Обновление значений редактируемого заказа
        '//
        '////////////////////////////////////////////////////////////////////////////////
        Dim MySQLStr As String

        MySQLStr = "Update tbl_DisplacementOrder_ShipmentInfo "
        MySQLStr = MySQLStr & "SET ShipmentDate = CONVERT(DATETIME, '" & Microsoft.VisualBasic.Right("00" & CStr(DatePart(DateInterval.Day, DateTimePicker1.Value)), 2) & "/" & Microsoft.VisualBasic.Right("00" & CStr(DatePart(DateInterval.Month, DateTimePicker1.Value)), 2) & "/" & CStr(DatePart(DateInterval.Year, DateTimePicker1.Value)) & "', 103), "
        MySQLStr = MySQLStr & "ReceivingDate = CONVERT(DATETIME, '" & Microsoft.VisualBasic.Right("00" & CStr(DatePart(DateInterval.Day, DateTimePicker2.Value)), 2) & "/" & Microsoft.VisualBasic.Right("00" & CStr(DatePart(DateInterval.Month, DateTimePicker2.Value)), 2) & "/" & CStr(DatePart(DateInterval.Year, DateTimePicker2.Value)) & "', 103), "
        MySQLStr = MySQLStr & "UserID = " & Declarations.UserID & ", "
        MySQLStr = MySQLStr & "TransportCompanyDocNum = N'" & TextBox1.Text & "', "
        MySQLStr = MySQLStr & "Comments = N'" & TextBox2.Text & "' "
        MySQLStr = MySQLStr & "WHERE (ID = " & Declarations.MyOrderID & ") "
        InitMyConn(False)
        Declarations.MyConn.Execute(MySQLStr)
    End Sub
End Class