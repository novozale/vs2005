Public Class NonCreditDialog

    Private Sub Button3_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button3.Click
        '////////////////////////////////////////////////////////////////////////////////////////
        '// Выход из формы с разрешением на отгрузку
        '//
        '////////////////////////////////////////////////////////////////////////////////////////

        If TextBox5.Text = "" Then
            MsgBox("Заполните причину отгрузки товара при отсутствии оплаты.", vbCritical, "Внимание!")
            TextBox5.Select()
            Exit Sub
        Else
            Declarations.CmdToShip = True
            Declarations.MyReason = TextBox5.Text
            Me.Close()
        End If
    End Sub

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        '////////////////////////////////////////////////////////////////////////////////////////
        '// Просмотр информации о клиенте
        '//
        '////////////////////////////////////////////////////////////////////////////////////////

        MyNonCreditInfo = New NonCreditInfo
        MyNonCreditInfo.ShowDialog()
    End Sub

    Private Sub Button2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button2.Click
        '////////////////////////////////////////////////////////////////////////////////////////
        '// Выход из формы с разрешением на отгрузку
        '//
        '////////////////////////////////////////////////////////////////////////////////////////

        Me.Close()
    End Sub

    Private Sub NonCreditDialog_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        '////////////////////////////////////////////////////////////////////////////////////////
        '// Загрузка формы
        '//
        '////////////////////////////////////////////////////////////////////////////////////////
        Dim MySQLStr As String
        Dim CardPayedSum As Double          '---предоплата по карточке с WEB сайта

        On Error GoTo MyCatch
        TextBox1.Text = Declarations.OrderSum
        TextBox2.Text = Declarations.MyPayment
        TextBox3.Text = Declarations.Avance1Type + Declarations.Avance2Type
        TextBox6.Text = Declarations.InvoiceDebt
        If Declarations.MyPermission = True Then
            Button3.Visible = True
            Button3.Enabled = True
            TextBox5.Visible = True
            Label10.Visible = True
        Else
            Button3.Visible = False
            Button3.Enabled = False
            TextBox5.Visible = False
            Label10.Visible = False
        End If

        MySQLStr = "SELECT SYCD009 "
        MySQLStr = MySQLStr & "FROM SYCD0100 WITH (NOLOCK) "
        MySQLStr = MySQLStr & "WHERE (SYCD001 = N'" & Declarations.CurrCode & "') "
        InitMyConn(True)
        InitMyRec(True, MySQLStr)
        TextBox4.Text = Declarations.MyRec.Fields("SYCD009").Value
        trycloseMyRec()

        '---------информация по WEB
        Declarations.IsWEBOrder = CheckWEBOrNot(Declarations.OrderID)
        If Declarations.IsWEBOrder = 0 Then '---не является заказом с WEB сайта
            Label11.Visible = False
            Label12.Visible = False
        Else
            CardPayedSum = GetCardPayment(Declarations.OrderID)
            Label12.Text = "На WEB сайте оплачено " & CStr(CardPayedSum) & " руб"
            Label11.Visible = True
            Label12.Visible = True
        End If
        Exit Sub
MyCatch:
        MsgBox(Err.Description, vbCritical, "Ошибка Frm_NonCreditDialog 1")
    End Sub
End Class