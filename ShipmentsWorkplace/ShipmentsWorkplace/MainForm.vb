Public Class MainForm
    Public LoadFlag As Integer

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        '////////////////////////////////////////////////////////////////////////////////
        '//
        '// Выход
        '//
        '////////////////////////////////////////////////////////////////////////////////

        Application.Exit()
    End Sub

    Private Sub MainForm_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles Me.KeyDown
        '////////////////////////////////////////////////////////////////////////////////
        '//
        '// Запрет закрытия окна по ALT+F4
        '//
        '////////////////////////////////////////////////////////////////////////////////

        If e.KeyData = Keys.Alt + Keys.F4 Then
            e.Handled = True
        End If
    End Sub

    Private Sub MainForm_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        '/////////////////////////////////////////////////////////////////////////////////////
        '//
        '// При запуске определяем параметры - Год, компания, пользователь и т.д.
        '// после чего выводим список поставщиков 
        '/////////////////////////////////////////////////////////////////////////////////////
        Dim MySQLStr As String                        'рабочая строка
        Dim MyAdapter As SqlClient.SqlDataAdapter     'для списка складов
        Dim MyDs As New DataSet                       '

        LoadFlag = 1
        '---параметры запуска
        Try
            Dim Scala As New SfwIII.Application

            Declarations.CompanyID = Scala.ActiveProcess.CommonVars.CompanyCode
            Declarations.Year = Mid(Scala.ActiveProcess.CommonVars.FiscalYear, 3)
            Declarations.UserCode = Scala.ActiveProcess.CommonVars.UserCode
            'Declarations.UserCode = "galkina"
        Catch
            MsgBox("Программа должна запускаться только из меню Scala", MsgBoxStyle.Critical, "Внимание!")
            Application.Exit()
        End Try

        '---ID пользователя
        MySQLStr = "SELECT UserID, FullName, UserName "
        MySQLStr = MySQLStr & "FROM  ScalaSystemDB.dbo.ScaUsers WITH(NOLOCK) "
        MySQLStr = MySQLStr & "WHERE (Upper(UserName) = N'" & UCase(Trim(Declarations.UserCode)) & "') "
        InitMyConn(False)
        InitMyRec(False, MySQLStr)
        If Declarations.MyRec.EOF = True And Declarations.MyRec.BOF = True Then
            MsgBox("Не найден ID сотрудника, соответствующий логину на вход в Scala. Обратитесь к администратору.", MsgBoxStyle.Critical, "Внимание!")
            trycloseMyRec()
            Application.Exit()
        Else
            Declarations.UserID = Declarations.MyRec.Fields("UserID").Value
            Declarations.FullName = Declarations.MyRec.Fields("FullName").Value
            Declarations.UserName = Declarations.MyRec.Fields("UserName").Value
            Declarations.SalesmanName = Declarations.FullName
            trycloseMyRec()
        End If

        '---Код продавца
        MySQLStr = "Select ST01001 "
        MySQLStr = MySQLStr & "FROM ST010300 "
        MySQLStr = MySQLStr & "WHERE (ST01002 = N'" & Trim(Declarations.FullName) & "')"
        InitMyConn(False)
        InitMyRec(False, MySQLStr)
        If Declarations.MyRec.EOF = True And Declarations.MyRec.BOF = True Then
            MsgBox("Не найден код продавца, соответствующий логину на вход в Scala. Обратитесь к администратору.", MsgBoxStyle.Critical, "Внимание!")
            trycloseMyRec()
            Application.Exit()
        Else
            Declarations.SalesmanCode = Declarations.MyRec.Fields("ST01001").Value
            trycloseMyRec()
        End If

        '---Список складов
        MySQLStr = "SELECT SC23001 AS WHCode, SC23001 + ' ' + SC23002 AS WHName "
        MySQLStr = MySQLStr & "FROM SC230300 WITH(NOLOCK) "
        MySQLStr = MySQLStr & "WHERE (SC23006 = N'1') AND (SC23001 IN('01','03')) "
        MySQLStr = MySQLStr & "ORDER BY WHCode "
        Try
            MyAdapter = New SqlClient.SqlDataAdapter(MySQLStr, Declarations.MyNETConnStr)
            MyAdapter.SelectCommand.CommandTimeout = 600
            MyAdapter.Fill(MyDs)
            ComboBox1.DisplayMember = "WHName" 'Это то что будет отображаться
            ComboBox1.ValueMember = "WHCode"   'это то что будет храниться
            ComboBox1.DataSource = MyDs.Tables(0).DefaultView
        Catch ex As Exception
            MsgBox(ex.ToString)
        End Try

        '---комбобокс активен - неактивен
        ComboBox2.Text = "Только активные покупатели"

        '---комбобокс работа в группе или индивидуально
        MySQLStr = "SELECT Value "
        MySQLStr = MySQLStr & "FROM tbl_Shipments_SalesmanWP_Config "
        MySQLStr = MySQLStr & "WHERE (UserId = " & Declarations.UserID & ") "
        MySQLStr = MySQLStr & "AND (Parameter = N'Работа в группе или индивидуально') "
        InitMyConn(False)
        InitMyRec(False, MySQLStr)
        If Declarations.MyRec.EOF = True And Declarations.MyRec.BOF = True Then
            ComboBox3.Text = "индивидуально"
            trycloseMyRec()
        Else
            If Declarations.MyRec.Fields("Value").Value = "0" Then
                ComboBox3.Text = "индивидуально"
            Else
                ComboBox3.Text = "в группе"
            End If
            trycloseMyRec()
        End If

        '--------------чтение конфигурации в глобальные переменные--------------------------
        '-----EMail для уведомления
        MySQLStr = "SELECT Value "
        MySQLStr = MySQLStr & "FROM tbl_Shipments_SalesmanWP_Config "
        MySQLStr = MySQLStr & "WHERE (UserId = " & Declarations.UserID & ") "
        MySQLStr = MySQLStr & "AND (Parameter = N'EMail для уведомления') "
        InitMyConn(False)
        InitMyRec(False, MySQLStr)
        If Declarations.MyRec.EOF = True And Declarations.MyRec.BOF = True Then
            MyEmail = 0
            trycloseMyRec()
        Else
            MyEmail = Declarations.MyRec.Fields("Value").Value
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
            MyContact = 0
            trycloseMyRec()
        Else
            MyContact = Declarations.MyRec.Fields("Value").Value
            trycloseMyRec()
        End If


        LoadFlag = 0
        Application.DoEvents()
        Windows.Forms.Cursor.Current = Cursors.WaitCursor
        '---Загрузка данных
        DataLoading()
        Application.DoEvents()
        Windows.Forms.Cursor.Current = Cursors.WaitCursor
        '---Проверка состояния кнопок
        CheckButtonsState()
        Application.DoEvents()
        Windows.Forms.Cursor.Current = Cursors.Default
    End Sub

    Public Function DataLoading()
        '/////////////////////////////////////////////////////////////////////////////////////
        '//
        '// Загрузка списка покупателей (в соответствии с параметрами)
        '// 
        '/////////////////////////////////////////////////////////////////////////////////////
        Dim MySQLStr As String
        Dim MyAdapter As SqlClient.SqlDataAdapter     'для списка поставщиков
        Dim MyDs As New DataSet                       '

        If LoadFlag = 0 Then
            If ComboBox2.Text = "Только активные покупатели" Then
                If ComboBox3.Text = "индивидуально" Then
                    MySQLStr = "Exec spp_Shipments_SalesmanWP_CommonInfo 1, N'" & Trim(ComboBox1.SelectedValue) & "', N'" & Declarations.SalesmanCode & "', 0, N'', N'' "
                Else
                    MySQLStr = "Exec spp_Shipments_SalesmanWP_CommonInfo 1, N'" & Trim(ComboBox1.SelectedValue) & "', N'" & Declarations.SalesmanCode & "', 1, N'', N'' "
                End If
            Else
                If ComboBox3.Text = "индивидуально" Then
                    MySQLStr = "Exec spp_Shipments_SalesmanWP_CommonInfo 0, N'" & Trim(ComboBox1.SelectedValue) & "', N'" & Declarations.SalesmanCode & "', 0, N'', N'' "
                Else
                    MySQLStr = "Exec spp_Shipments_SalesmanWP_CommonInfo 0, N'" & Trim(ComboBox1.SelectedValue) & "', N'" & Declarations.SalesmanCode & "', 1, N'', N'' "
                End If
            End If

            Try
                MyAdapter = New SqlClient.SqlDataAdapter(MySQLStr, Declarations.MyNETConnStr)
                MyAdapter.SelectCommand.CommandTimeout = 600
                MyAdapter.Fill(MyDs)
                DataGridView1.DataSource = MyDs.Tables(0)
            Catch ex As Exception
                MsgBox(ex.ToString)
            End Try

            '---заголовки
            DataGridView1.Columns(0).HeaderText = "Код поку пателя"
            DataGridView1.Columns(0).Width = 90
            DataGridView1.Columns(1).HeaderText = "Покупатель"
            DataGridView1.Columns(1).Width = 210
            DataGridView1.Columns(2).HeaderText = "Адрес покупателя"
            DataGridView1.Columns(2).Width = 361
            DataGridView1.Columns(3).HeaderText = "Заказов с отгрузкой в теч. 7 дней"
            DataGridView1.Columns(3).Width = 110
            DataGridView1.Columns(4).HeaderText = "заказов с просроченной отгрузкой"
            DataGridView1.Columns(4).Width = 110
            DataGridView1.Columns(5).HeaderText = "Заказов, у которых дата прихода больше даты отгрузки"
            DataGridView1.Columns(5).Width = 110
            DataGridView1.Columns(6).HeaderText = "Заказов, не вывезенных в течении 7 дней"
            DataGridView1.Columns(6).Width = 110
            DataGridView1.Columns(7).HeaderText = "Заказов, не отгруженных в течении 2 дней"
            DataGridView1.Columns(7).Width = 110
        End If
    End Function

    Public Function CheckButtonsState()
        '////////////////////////////////////////////////////////////////////////////////
        '//
        '// Проверка и выставление состояния кнопок
        '//
        '////////////////////////////////////////////////////////////////////////////////

        If DataGridView1.SelectedRows.Count = 0 Then
            Button2.Enabled = False
        Else
            Button2.Enabled = True
        End If
    End Function

    Private Sub Button3_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button3.Click
        '////////////////////////////////////////////////////////////////////////////////
        '//
        '// Обновление данных
        '//
        '////////////////////////////////////////////////////////////////////////////////

        Windows.Forms.Cursor.Current = Cursors.WaitCursor
        '---загрузка данных
        DataLoading()
        Application.DoEvents()
        '---проверка состояния кнопок
        CheckButtonsState()
        Application.DoEvents()
        Windows.Forms.Cursor.Current = Cursors.Default
    End Sub

    Private Sub ComboBox1_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles ComboBox1.SelectedIndexChanged
        '////////////////////////////////////////////////////////////////////////////////
        '//
        '// Загрузка данных в соответствии с выбранным складом
        '//
        '////////////////////////////////////////////////////////////////////////////////

        Windows.Forms.Cursor.Current = Cursors.WaitCursor
        '---загрузка данных
        DataLoading()
        Application.DoEvents()
        Windows.Forms.Cursor.Current = Cursors.WaitCursor
        '---проверка состояния кнопок
        CheckButtonsState()
        Application.DoEvents()
        Windows.Forms.Cursor.Current = Cursors.Default
    End Sub

    Private Sub ComboBox2_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles ComboBox2.SelectedIndexChanged
        '////////////////////////////////////////////////////////////////////////////////
        '//
        '// Загрузка данных в соответствии с выбранной опцией - отображать всех покупателей или только активных
        '//
        '////////////////////////////////////////////////////////////////////////////////

        Windows.Forms.Cursor.Current = Cursors.WaitCursor
        '---загрузка данных
        DataLoading()
        Application.DoEvents()
        Windows.Forms.Cursor.Current = Cursors.WaitCursor
        '---проверка состояния кнопок
        CheckButtonsState()
        Application.DoEvents()
        Windows.Forms.Cursor.Current = Cursors.Default
    End Sub

    Private Sub Button2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button2.Click
        '////////////////////////////////////////////////////////////////////////////////
        '//
        '// открытиие окна формирования отгрузки
        '//
        '////////////////////////////////////////////////////////////////////////////////

        OpenShipmentWindow()
    End Sub

    Private Sub DataGridView1_CellFormatting(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellFormattingEventArgs) Handles DataGridView1.CellFormatting
        '////////////////////////////////////////////////////////////////////////////////
        '//
        '// подсветка информации по покупателям
        '//
        '////////////////////////////////////////////////////////////////////////////////
        Dim row As DataGridViewRow

        row = DataGridView1.Rows(e.RowIndex)
        row.Cells(0).Style.Font = New Font(DataGridView1.DefaultCellStyle.Font, FontStyle.Bold)
        row.Cells(1).Style.Font = New Font(DataGridView1.DefaultCellStyle.Font, FontStyle.Bold)
        row.Cells(3).Style.Format = "n0"
        row.Cells(3).Style.Font = New Font(DataGridView1.DefaultCellStyle.Font, FontStyle.Bold)
        row.Cells(4).Style.Format = "n0"
        row.Cells(4).Style.Font = New Font(DataGridView1.DefaultCellStyle.Font, FontStyle.Bold)
        row.Cells(5).Style.Format = "n0"
        row.Cells(5).Style.Font = New Font(DataGridView1.DefaultCellStyle.Font, FontStyle.Bold)
        row.Cells(6).Style.Format = "n0"
        row.Cells(6).Style.Font = New Font(DataGridView1.DefaultCellStyle.Font, FontStyle.Bold)
        row.Cells(7).Style.Format = "n0"
        row.Cells(7).Style.Font = New Font(DataGridView1.DefaultCellStyle.Font, FontStyle.Bold)
        If Trim(row.Cells(3).Value.ToString) <> "" Then
            row.Cells(3).Style.BackColor = Color.LightGreen
        Else
            row.Cells(3).Style.BackColor = Color.Empty
        End If
        If Trim(row.Cells(4).Value.ToString) <> "" Then
            row.Cells(4).Style.BackColor = Color.Red
        Else
            row.Cells(4).Style.BackColor = Color.Empty
        End If
        If Trim(row.Cells(5).Value.ToString) <> "" Then
            row.Cells(5).Style.BackColor = Color.Orange
        Else
            row.Cells(5).Style.BackColor = Color.Empty
        End If
        If Trim(row.Cells(6).Value.ToString) <> "" Then
            row.Cells(6).Style.BackColor = Color.Yellow
        Else
            row.Cells(6).Style.BackColor = Color.Empty
        End If
        If Trim(row.Cells(7).Value.ToString) <> "" Then
            row.Cells(7).Style.BackColor = Color.Yellow
        Else
            row.Cells(7).Style.BackColor = Color.Empty
        End If
    End Sub

    Private Sub DataGridView1_ColumnHeaderMouseClick(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellMouseEventArgs) Handles DataGridView1.ColumnHeaderMouseClick
        '/////////////////////////////////////////////////////////////////////////////////////
        '//
        '// Щелчок по заголовку таблицы 
        '// 
        '/////////////////////////////////////////////////////////////////////////////////////

        Button6.Text = "Подсветить все"
    End Sub


    Private Sub DataGridView1_DoubleClick(ByVal sender As Object, ByVal e As System.EventArgs) Handles DataGridView1.DoubleClick
        '////////////////////////////////////////////////////////////////////////////////
        '//
        '// открытиие окна формирования консолидированных заказов
        '//
        '////////////////////////////////////////////////////////////////////////////////

        OpenShipmentWindow()
    End Sub

    Private Sub OpenShipmentWindow()
        '////////////////////////////////////////////////////////////////////////////////
        '//
        '// открытиие окна формирования отгрузок
        '//
        '////////////////////////////////////////////////////////////////////////////////

        Declarations.MyCustomerCode = Trim(Me.DataGridView1.SelectedRows.Item(0).Cells(0).Value.ToString())
        Declarations.MyWH = Trim(Me.ComboBox1.SelectedValue)
        If ComboBox3.Text = "индивидуально" Then
            Declarations.MyGroupOrIndividualFlag = 0
        Else
            Declarations.MyGroupOrIndividualFlag = 1
        End If
        MyShipmentsList = New ShipmentsList
        MyShipmentsList.ShowDialog()
        Application.DoEvents()
        Windows.Forms.Cursor.Current = Cursors.WaitCursor
        DataLoading()
        '---текущей строкой сделать редактированную
        For i As Integer = 0 To DataGridView1.Rows.Count - 1
            If Trim(DataGridView1.Item(0, i).Value.ToString) = Declarations.MyCustomerCode Then
                DataGridView1.CurrentCell = DataGridView1.Item(0, i)
            End If
        Next
        CheckButtonsState()
    End Sub

    Private Sub ComboBox3_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles ComboBox3.SelectedIndexChanged
        '////////////////////////////////////////////////////////////////////////////////
        '//
        '// Загрузка данных в соответствии с выбранной опцией - работа в группе или индивидуально
        '//
        '////////////////////////////////////////////////////////////////////////////////

        Windows.Forms.Cursor.Current = Cursors.WaitCursor
        '---Сохранение выбора
        SaveJobTypeChoice()
        Application.DoEvents()
        Windows.Forms.Cursor.Current = Cursors.WaitCursor
        '---загрузка данных
        DataLoading()
        Application.DoEvents()
        Windows.Forms.Cursor.Current = Cursors.WaitCursor
        '---проверка состояния кнопок
        CheckButtonsState()
        Application.DoEvents()
        Windows.Forms.Cursor.Current = Cursors.Default
    End Sub

    Public Sub SaveJobTypeChoice()
        '////////////////////////////////////////////////////////////////////////////////
        '//
        '// Сохранение выбранной опции - работа в группе или индивидуально
        '//
        '////////////////////////////////////////////////////////////////////////////////
        Dim MySQLStr As String

        MySQLStr = "DELETE FROM tbl_Shipments_SalesmanWP_Config "
        MySQLStr = MySQLStr & "WHERE (UserId = " & Declarations.UserID & ") "
        MySQLStr = MySQLStr & "AND (Parameter = N'Работа в группе или индивидуально') "
        InitMyConn(False)
        Declarations.MyConn.Execute(MySQLStr)

        MySQLStr = "INSERT INTO tbl_Shipments_SalesmanWP_Config "
        MySQLStr = MySQLStr & "(UserId, Parameter, Value) "
        MySQLStr = MySQLStr & "VALUES (" & Declarations.UserID & ", "
        MySQLStr = MySQLStr & "N'Работа в группе или индивидуально', "
        If ComboBox3.Text = "индивидуально" Then
            MySQLStr = MySQLStr & "N'0') "
        Else
            MySQLStr = MySQLStr & "N'1') "
        End If
        InitMyConn(False)
        Declarations.MyConn.Execute(MySQLStr)
    End Sub

    Private Sub Button8_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button8.Click
        '////////////////////////////////////////////////////////////////////////////////
        '//
        '// Поиск первого подходящего поставщика
        '//
        '////////////////////////////////////////////////////////////////////////////////

        If Trim(TextBox2.Text) = "" And Trim(TextBox3.Text) = "" Then
        Else
            Windows.Forms.Cursor.Current = Cursors.WaitCursor
            For i As Integer = 0 To DataGridView1.Rows.Count - 1
                If (InStr(UCase(Trim(DataGridView1.Item(0, i).Value.ToString)), UCase(Trim(TextBox2.Text))) <> 0 _
                    Or InStr(UCase(Trim(DataGridView1.Item(1, i).Value.ToString)), UCase(Trim(TextBox2.Text))) <> 0 _
                    Or InStr(UCase(Trim(DataGridView1.Item(2, i).Value.ToString)), UCase(Trim(TextBox2.Text))) <> 0) _
                    And (InStr(UCase(Trim(DataGridView1.Item(0, i).Value.ToString)), UCase(Trim(TextBox3.Text))) <> 0 _
                    Or InStr(UCase(Trim(DataGridView1.Item(1, i).Value.ToString)), UCase(Trim(TextBox3.Text))) <> 0 _
                    Or InStr(UCase(Trim(DataGridView1.Item(2, i).Value.ToString)), UCase(Trim(TextBox3.Text))) <> 0) Then
                    DataGridView1.CurrentCell = DataGridView1.Item(0, i)
                    Windows.Forms.Cursor.Current = Cursors.Default
                    Exit Sub
                End If
            Next
            Windows.Forms.Cursor.Current = Cursors.Default
        End If
    End Sub

    Private Sub Button7_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button7.Click
        '////////////////////////////////////////////////////////////////////////////////
        '//
        '// Поиск следующего подходящего события
        '//
        '////////////////////////////////////////////////////////////////////////////////
        Dim MyRez As Object

        If Trim(TextBox2.Text) = "" And Trim(TextBox3.Text) = "" Then
        Else
            Windows.Forms.Cursor.Current = Cursors.WaitCursor
            For i As Integer = DataGridView1.CurrentCellAddress.Y + 1 To DataGridView1.Rows.Count
                If i = DataGridView1.Rows.Count Then
                    MyRez = MsgBox("Поиск дошел до конца списка. Начать сначала?", MsgBoxStyle.YesNo, "Внимание!")
                    If MyRez = 6 Then
                        i = 0
                    Else
                        Windows.Forms.Cursor.Current = Cursors.Default
                        Exit Sub
                    End If
                End If
                If DataGridView1.Rows.Count = 0 Then
                Else
                    If (InStr(UCase(Trim(DataGridView1.Item(0, i).Value.ToString)), UCase(Trim(TextBox2.Text))) <> 0 _
                    Or InStr(UCase(Trim(DataGridView1.Item(1, i).Value.ToString)), UCase(Trim(TextBox2.Text))) <> 0 _
                    Or InStr(UCase(Trim(DataGridView1.Item(2, i).Value.ToString)), UCase(Trim(TextBox2.Text))) <> 0) _
                    And (InStr(UCase(Trim(DataGridView1.Item(0, i).Value.ToString)), UCase(Trim(TextBox3.Text))) <> 0 _
                    Or InStr(UCase(Trim(DataGridView1.Item(1, i).Value.ToString)), UCase(Trim(TextBox3.Text))) <> 0 _
                    Or InStr(UCase(Trim(DataGridView1.Item(2, i).Value.ToString)), UCase(Trim(TextBox3.Text))) <> 0) Then
                        DataGridView1.CurrentCell = DataGridView1.Item(0, i)
                        Windows.Forms.Cursor.Current = Cursors.Default
                        Exit Sub
                    End If
                End If
            Next i
            Windows.Forms.Cursor.Current = Cursors.Default
        End If
    End Sub

    Private Sub Button6_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button6.Click
        '////////////////////////////////////////////////////////////////////////////////
        '//
        '// Подсвечивание всех подходящих по критерию записей
        '//
        '////////////////////////////////////////////////////////////////////////////////

        If Button6.Text = "Подсветить все" Then
            Windows.Forms.Cursor.Current = Cursors.WaitCursor
            For i As Integer = 0 To DataGridView1.Rows.Count - 1
                If (InStr(UCase(Trim(DataGridView1.Item(0, i).Value.ToString)), UCase(Trim(TextBox2.Text))) <> 0 _
                    Or InStr(UCase(Trim(DataGridView1.Item(1, i).Value.ToString)), UCase(Trim(TextBox2.Text))) <> 0 _
                    Or InStr(UCase(Trim(DataGridView1.Item(2, i).Value.ToString)), UCase(Trim(TextBox2.Text))) <> 0) _
                    And (InStr(UCase(Trim(DataGridView1.Item(0, i).Value.ToString)), UCase(Trim(TextBox3.Text))) <> 0 _
                    Or InStr(UCase(Trim(DataGridView1.Item(1, i).Value.ToString)), UCase(Trim(TextBox3.Text))) <> 0 _
                    Or InStr(UCase(Trim(DataGridView1.Item(2, i).Value.ToString)), UCase(Trim(TextBox3.Text))) <> 0) Then
                    DataGridView1.Rows(i).Cells(0).Style.BackColor = Color.Yellow
                    DataGridView1.Rows(i).Cells(1).Style.BackColor = Color.Yellow
                    DataGridView1.Rows(i).Cells(2).Style.BackColor = Color.Yellow
                Else
                    DataGridView1.Rows(i).Cells(0).Style.BackColor = Color.Empty
                    DataGridView1.Rows(i).Cells(1).Style.BackColor = Color.Empty
                    DataGridView1.Rows(i).Cells(2).Style.BackColor = Color.Empty
                End If
            Next
            Windows.Forms.Cursor.Current = Cursors.Default
            Button6.Text = "Снять выдел."
        Else
            Windows.Forms.Cursor.Current = Cursors.WaitCursor
            For i As Integer = 0 To DataGridView1.Rows.Count - 1
                DataGridView1.Rows(i).Cells(0).Style.BackColor = Color.Empty
                DataGridView1.Rows(i).Cells(1).Style.BackColor = Color.Empty
                DataGridView1.Rows(i).Cells(2).Style.BackColor = Color.Empty
            Next
            Windows.Forms.Cursor.Current = Cursors.Default
            Button6.Text = "Подсветить все"
        End If
    End Sub

    Private Sub Button9_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button9.Click
        '////////////////////////////////////////////////////////////////////////////////
        '//
        '// Выбор всех подходящих по критерию поставщиков в отдельное окно
        '//
        '////////////////////////////////////////////////////////////////////////////////

        If Trim(TextBox2.Text) = "" And Trim(TextBox3.Text) = "" Then
            MsgBox("Необходимо ввести критерий поиска", MsgBoxStyle.OkOnly, "Внимание!")
            TextBox2.Select()
        Else
            MyCustomerSelectList = New CustomerSelectList
            MyCustomerSelectList.ShowDialog()
        End If
    End Sub

    Private Sub Button4_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button4.Click
        '////////////////////////////////////////////////////////////////////////////////
        '//
        '// Открытие окна настроек
        '//
        '////////////////////////////////////////////////////////////////////////////////

        MyConfiguration = New Configuration
        MyConfiguration.ShowDialog()
    End Sub
End Class