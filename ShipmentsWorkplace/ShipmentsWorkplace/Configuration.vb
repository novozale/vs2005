Public Class Configuration

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        '////////////////////////////////////////////////////////////////////////////////
        '//
        '// Выход
        '//
        '////////////////////////////////////////////////////////////////////////////////

        Me.Close()
    End Sub

    Private Sub Configuration_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        '////////////////////////////////////////////////////////////////////////////////
        '//
        '// Загрузка данных
        '//
        '////////////////////////////////////////////////////////////////////////////////
        Dim MySQLStr As String

        '-----EMail для уведомления
        MySQLStr = "SELECT Value "
        MySQLStr = MySQLStr & "FROM tbl_Shipments_SalesmanWP_Config "
        MySQLStr = MySQLStr & "WHERE (UserId = " & Declarations.UserID & ") "
        MySQLStr = MySQLStr & "AND (Parameter = N'EMail для уведомления') "
        InitMyConn(False)
        InitMyRec(False, MySQLStr)
        If Declarations.MyRec.EOF = True And Declarations.MyRec.BOF = True Then
            ComboBoxEMail.Text = "Вручную"
            trycloseMyRec()
        Else
            If Declarations.MyRec.Fields("Value").Value = "0" Then
                ComboBoxEMail.Text = "Вручную"
            Else
                ComboBoxEMail.Text = "Выбранный из CRM контакт"
            End If
            trycloseMyRec()
        End If

        '-----Заполнение поля контакт
        MySQLStr = "SELECT Value "
        MySQLStr = MySQLStr & "FROM tbl_Shipments_SalesmanWP_Config "
        MySQLStr = MySQLStr & "WHERE (UserId = " & Declarations.UserID & ") "
        MySQLStr = MySQLStr & "AND (Parameter = N'Заполнение поля контакт') "
        InitMyConn(False)
        InitMyRec(False, MySQLStr)
        If Declarations.MyRec.EOF = True And Declarations.MyRec.BOF = True Then
            ComboBoxContact.Text = "Вручную"
            trycloseMyRec()
        Else
            If Declarations.MyRec.Fields("Value").Value = "0" Then
                ComboBoxContact.Text = "Вручную"
            Else
                ComboBoxContact.Text = "Из Scala"
            End If
            trycloseMyRec()
        End If
    End Sub

    Private Sub Button2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button2.Click
        '////////////////////////////////////////////////////////////////////////////////
        '//
        '// Выход с сохранением данных
        '//
        '////////////////////////////////////////////////////////////////////////////////

        SaveInfo()
        Me.Close()
    End Sub

    Private Sub SaveInfo()
        '////////////////////////////////////////////////////////////////////////////////
        '//
        '// Сохранение данных по конфигурации
        '//
        '////////////////////////////////////////////////////////////////////////////////
        Dim MySQLStr As String

        '-----EMail для уведомления
        MySQLStr = "DELETE FROM tbl_Shipments_SalesmanWP_Config "
        MySQLStr = MySQLStr & "WHERE (UserId = " & Declarations.UserID & ") "
        MySQLStr = MySQLStr & "AND (Parameter = N'EMail для уведомления') "
        InitMyConn(False)
        Declarations.MyConn.Execute(MySQLStr)

        MySQLStr = "INSERT INTO tbl_Shipments_SalesmanWP_Config "
        MySQLStr = MySQLStr & "(UserId, Parameter, Value) "
        MySQLStr = MySQLStr & "VALUES (" & Declarations.UserID & ", "
        MySQLStr = MySQLStr & "N'EMail для уведомления', "
        If ComboBoxEMail.Text = "Вручную" Then
            MySQLStr = MySQLStr & "N'0') "
            MyEmail = 0
        Else
            MySQLStr = MySQLStr & "N'1') "
            MyEmail = 1
        End If
        InitMyConn(False)
        Declarations.MyConn.Execute(MySQLStr)

        '-----Заполнение поля контакт
        MySQLStr = "DELETE FROM tbl_Shipments_SalesmanWP_Config "
        MySQLStr = MySQLStr & "WHERE (UserId = " & Declarations.UserID & ") "
        MySQLStr = MySQLStr & "AND (Parameter = N'Заполнение поля контакт') "
        InitMyConn(False)
        Declarations.MyConn.Execute(MySQLStr)

        MySQLStr = "INSERT INTO tbl_Shipments_SalesmanWP_Config "
        MySQLStr = MySQLStr & "(UserId, Parameter, Value) "
        MySQLStr = MySQLStr & "VALUES (" & Declarations.UserID & ", "
        MySQLStr = MySQLStr & "N'Заполнение поля контакт', "
        If ComboBoxContact.Text = "Вручную" Then
            MySQLStr = MySQLStr & "N'0') "
            MyContact = 0
        Else
            MySQLStr = MySQLStr & "N'1') "
            MyContact = 1
        End If
        InitMyConn(False)
        Declarations.MyConn.Execute(MySQLStr)
    End Sub
End Class