Public Class MainForm

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

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        '////////////////////////////////////////////////////////////////////////////////
        '//
        '// Выход
        '//
        '////////////////////////////////////////////////////////////////////////////////

        Application.Exit()
    End Sub

    Private Sub MainForm_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        '/////////////////////////////////////////////////////////////////////////////////////
        '//
        '// При запуске определяем параметры - Год, компания, пользователь и т.д.
        '// после чего выводим список производителей 
        '/////////////////////////////////////////////////////////////////////////////////////
        Dim MySQLStr As String                        'рабочая строка
        Dim MyAdapter As SqlClient.SqlDataAdapter     'для списка складов
        Dim MyDs As New DataSet                       '

        '---параметры запуска
        Try
            Dim Scala As New SfwIII.Application

            Declarations.CompanyID = Scala.ActiveProcess.CommonVars.CompanyCode
            Declarations.Year = Mid(Scala.ActiveProcess.CommonVars.FiscalYear, 3)
            Declarations.UserCode = Scala.ActiveProcess.CommonVars.UserCode

        Catch
            MsgBox("Программа должна запускаться только из меню Scala", MsgBoxStyle.Critical, "Внимание!")
            Application.Exit()
        End Try

        '---ID пользователя
        MySQLStr = "SELECT UserID, FullName "
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
            trycloseMyRec()
        End If

        '---Загрузка данных
        DataLoading()
        '---Проверка состояния кнопок
        CheckButtonsState()
    End Sub

    Public Function DataLoading()
        '/////////////////////////////////////////////////////////////////////////////////////
        '//
        '// Загрузка списка производителей
        '// 
        '/////////////////////////////////////////////////////////////////////////////////////
        Dim MySQLStr As String
        Dim MyAdapter As SqlClient.SqlDataAdapter     'для списка поставщиков
        Dim MyDs As New DataSet                       '

        MySQLStr = "SELECT * "
        MySQLStr = MySQLStr & "FROM tbl_Manufacturers "
        MySQLStr = MySQLStr & "ORDER BY ID "

        Try
            MyAdapter = New SqlClient.SqlDataAdapter(MySQLStr, Declarations.MyNETConnStr)
            MyAdapter.SelectCommand.CommandTimeout = 600
            MyAdapter.Fill(MyDs)
            DataGridView1.DataSource = MyDs.Tables(0)
        Catch ex As Exception
            MsgBox(ex.ToString)
        End Try

        '---заголовки
        DataGridView1.Columns(0).HeaderText = "Код произ водителя"
        DataGridView1.Columns(0).Width = 100
        DataGridView1.Columns(1).HeaderText = "Призводитель"
        DataGridView1.Columns(1).Width = 300
        DataGridView1.Columns(2).HeaderText = "Адрес производителя"
        DataGridView1.Columns(2).Width = 410
        DataGridView1.Columns(3).HeaderText = "Контактная информация"
        DataGridView1.Columns(3).Width = 410
        DataGridView1.Columns(4).HeaderText = "Стандартный"
        DataGridView1.Columns(4).Width = 100
    End Function

    Public Function CheckButtonsState()
        '////////////////////////////////////////////////////////////////////////////////
        '//
        '// Проверка и выставление состояния кнопок
        '//
        '////////////////////////////////////////////////////////////////////////////////

        If DataGridView1.SelectedRows.Count = 0 Then
            Button4.Enabled = False
            Button5.Enabled = False
        Else
            Button4.Enabled = True
            Button5.Enabled = True
        End If
    End Function

    Private Sub Button3_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button3.Click
        '////////////////////////////////////////////////////////////////////////////////
        '//
        '// Обновление данных
        '//
        '////////////////////////////////////////////////////////////////////////////////

        '---загрузка данных
        DataLoading()
        '---проверка состояния кнопок
        CheckButtonsState()
    End Sub

    Private Sub Button8_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button8.Click
        '////////////////////////////////////////////////////////////////////////////////
        '//
        '// Поиск первого подходящего производителя
        '//
        '////////////////////////////////////////////////////////////////////////////////

        If Trim(TextBox2.Text) = "" And Trim(TextBox3.Text) = "" Then
        Else
            Windows.Forms.Cursor.Current = Cursors.WaitCursor
            For i As Integer = 0 To DataGridView1.Rows.Count - 1
                If InStr(UCase(Trim(DataGridView1.Item(0, i).Value.ToString)), UCase(Trim(TextBox2.Text))) <> 0 And InStr(UCase(Trim(DataGridView1.Item(0, i).Value.ToString)), UCase(Trim(TextBox3.Text))) <> 0 _
                    Or InStr(UCase(Trim(DataGridView1.Item(1, i).Value.ToString)), UCase(Trim(TextBox2.Text))) <> 0 And InStr(UCase(Trim(DataGridView1.Item(1, i).Value.ToString)), UCase(Trim(TextBox3.Text))) <> 0 _
                    Or InStr(UCase(Trim(DataGridView1.Item(2, i).Value.ToString)), UCase(Trim(TextBox2.Text))) <> 0 And InStr(UCase(Trim(DataGridView1.Item(2, i).Value.ToString)), UCase(Trim(TextBox3.Text))) <> 0 _
                    Or InStr(UCase(Trim(DataGridView1.Item(3, i).Value.ToString)), UCase(Trim(TextBox2.Text))) <> 0 And InStr(UCase(Trim(DataGridView1.Item(3, i).Value.ToString)), UCase(Trim(TextBox3.Text))) <> 0 Then
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
        '// Поиск следующего подходящего производителя
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
                    If InStr(UCase(Trim(DataGridView1.Item(0, i).Value.ToString)), UCase(Trim(TextBox2.Text))) <> 0 And InStr(UCase(Trim(DataGridView1.Item(0, i).Value.ToString)), UCase(Trim(TextBox3.Text))) <> 0 _
                        Or InStr(UCase(Trim(DataGridView1.Item(1, i).Value.ToString)), UCase(Trim(TextBox2.Text))) <> 0 And InStr(UCase(Trim(DataGridView1.Item(1, i).Value.ToString)), UCase(Trim(TextBox3.Text))) <> 0 _
                        Or InStr(UCase(Trim(DataGridView1.Item(2, i).Value.ToString)), UCase(Trim(TextBox2.Text))) <> 0 And InStr(UCase(Trim(DataGridView1.Item(2, i).Value.ToString)), UCase(Trim(TextBox3.Text))) <> 0 _
                        Or InStr(UCase(Trim(DataGridView1.Item(3, i).Value.ToString)), UCase(Trim(TextBox2.Text))) <> 0 And InStr(UCase(Trim(DataGridView1.Item(3, i).Value.ToString)), UCase(Trim(TextBox3.Text))) <> 0 Then
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
        '// Подсвечивание всех подходящих по критерию производителей
        '//
        '////////////////////////////////////////////////////////////////////////////////

        If Button6.Text = "Подсветить все" Then
            Windows.Forms.Cursor.Current = Cursors.WaitCursor
            For i As Integer = 0 To DataGridView1.Rows.Count - 1
                If InStr(UCase(Trim(DataGridView1.Item(0, i).Value.ToString)), UCase(Trim(TextBox2.Text))) <> 0 And InStr(UCase(Trim(DataGridView1.Item(0, i).Value.ToString)), UCase(Trim(TextBox3.Text))) <> 0 _
                        Or InStr(UCase(Trim(DataGridView1.Item(1, i).Value.ToString)), UCase(Trim(TextBox2.Text))) <> 0 And InStr(UCase(Trim(DataGridView1.Item(1, i).Value.ToString)), UCase(Trim(TextBox3.Text))) <> 0 _
                        Or InStr(UCase(Trim(DataGridView1.Item(2, i).Value.ToString)), UCase(Trim(TextBox2.Text))) <> 0 And InStr(UCase(Trim(DataGridView1.Item(2, i).Value.ToString)), UCase(Trim(TextBox3.Text))) <> 0 _
                        Or InStr(UCase(Trim(DataGridView1.Item(3, i).Value.ToString)), UCase(Trim(TextBox2.Text))) <> 0 And InStr(UCase(Trim(DataGridView1.Item(3, i).Value.ToString)), UCase(Trim(TextBox3.Text))) <> 0 Then
                    DataGridView1.Rows(i).Cells(0).Style.BackColor = Color.Yellow
                    DataGridView1.Rows(i).Cells(1).Style.BackColor = Color.Yellow
                    DataGridView1.Rows(i).Cells(2).Style.BackColor = Color.Yellow
                    DataGridView1.Rows(i).Cells(3).Style.BackColor = Color.Yellow
                Else
                    DataGridView1.Rows(i).Cells(0).Style.BackColor = Color.Empty
                    DataGridView1.Rows(i).Cells(1).Style.BackColor = Color.Empty
                    DataGridView1.Rows(i).Cells(2).Style.BackColor = Color.Empty
                    DataGridView1.Rows(i).Cells(3).Style.BackColor = Color.Empty
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
                DataGridView1.Rows(i).Cells(3).Style.BackColor = Color.Empty
            Next
            Windows.Forms.Cursor.Current = Cursors.Default
            Button6.Text = "Подсветить все"
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
            MyManufacturersSelectList = New ManufacturersSelectList
            MyManufacturersSelectList.ShowDialog()
        End If
    End Sub

    Private Sub Button2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button2.Click
        '////////////////////////////////////////////////////////////////////////////////
        '//
        '// Создание нового производителя
        '//
        '////////////////////////////////////////////////////////////////////////////////

        MyEditManufacturer = New EditManufacturer
        MyEditManufacturer.NewManufacturer = 0
        MyEditManufacturer.ShowDialog()
        '---загрузка данных
        DataLoading()
        '---проверка состояния кнопок
        CheckButtonsState()
    End Sub

    Private Sub Button4_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button4.Click
        '////////////////////////////////////////////////////////////////////////////////
        '//
        '// редактирование производителя
        '//
        '////////////////////////////////////////////////////////////////////////////////

        MyEditManufacturer = New EditManufacturer
        MyEditManufacturer.NewManufacturer = 1
        MyEditManufacturer.ShowDialog()
        '---загрузка данных
        DataLoading()
        '---проверка состояния кнопок
        CheckButtonsState()
    End Sub

    Private Sub Button5_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button5.Click
        '////////////////////////////////////////////////////////////////////////////////
        '//
        '// удаление производителя
        '//
        '////////////////////////////////////////////////////////////////////////////////
        Dim MySQLStr As String
        Dim LinksPresent As Integer         'есть ссылки (1) или нет (0)

        Declarations.MyManufacturerCode = Trim(Me.DataGridView1.SelectedRows.Item(0).Cells(0).Value.ToString())

        '---смотрим, нет ли ссылок на производителя
        MySQLStr = "SELECT COUNT(tbl_ItemCardID) AS CC "
        MySQLStr = MySQLStr & "FROM tbl_ItemCard0300 "
        MySQLStr = MySQLStr & "WHERE (Manufacturer = " & Declarations.MyManufacturerCode & ") "
        InitMyConn(False)
        InitMyRec(False, MySQLStr)
        If Declarations.MyRec.BOF = True And Declarations.MyRec.EOF = True Then
            LinksPresent = 0
            trycloseMyRec()
        Else
            If Declarations.MyRec.Fields("CC").Value > 0 Then
                LinksPresent = 1
            Else
                LinksPresent = 0
            End If
            trycloseMyRec()
        End If

        '---если ссылок нет - удаляем.
        If LinksPresent = 0 Then
            MySQLStr = "DELETE FROM tbl_Manufacturers "
            MySQLStr = MySQLStr & "WHERE (ID = " & Declarations.MyManufacturerCode & ") "
            InitMyConn(False)
            Declarations.MyConn.Execute(MySQLStr)
        Else
            MsgBox("Данного производителя нельзя удалить, так как есть запасы, ссылающиеся на него.", MsgBoxStyle.Critical, "Внимание!")
        End If

        '---загрузка данных
        DataLoading()
        '---проверка состояния кнопок
        CheckButtonsState()
    End Sub
End Class
