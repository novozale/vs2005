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
        '// после чего выводим список складов получения 
        '/////////////////////////////////////////////////////////////////////////////////////
        Dim MySQLStr As String                        'рабочая строка
        Dim MyAdapter As SqlClient.SqlDataAdapter     'для списка складов
        Dim MyDs As New DataSet

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

        '---------администратор - изменение заказов после закрытия
        CheckRights(Declarations.UserCode, "Начальник логистики")


        '---Список складов
        MySQLStr = "SELECT SC23001 AS WHCode, SC23001 + ' ' + SC23002 AS WHName "
        MySQLStr = MySQLStr & "FROM SC230300 WITH(NOLOCK) "
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
        ComboBox2.Text = "Только активные склады приемки"

        '---Загрузка данных
        DataLoading()
        '---Проверка состояния кнопок
        CheckButtonsState()
    End Sub

    Public Function DataLoading()
        '/////////////////////////////////////////////////////////////////////////////////////
        '//
        '// Загрузка списка складов приемки с информацией
        '// 
        '/////////////////////////////////////////////////////////////////////////////////////
        Dim MySQLStr As String
        Dim MyAdapter As SqlClient.SqlDataAdapter     'для списка поставщиков
        Dim MyDs As New DataSet                       '

        If ComboBox2.Text = "Только активные склады приемки" Then
            MySQLStr = "Exec spp_DisplacementWorkplace_WHToListPrep 1, N'" & Trim(ComboBox1.SelectedValue) & "' "
        Else
            MySQLStr = "Exec spp_DisplacementWorkplace_WHToListPrep 0, N'" & Trim(ComboBox1.SelectedValue) & "' "
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
        DataGridView1.Columns(0).HeaderText = "Код склада"
        DataGridView1.Columns(0).Width = 100
        DataGridView1.Columns(1).HeaderText = "Склад"
        DataGridView1.Columns(1).Width = 224
        DataGridView1.Columns(2).HeaderText = "Несгруппи рованных заказов"
        DataGridView1.Columns(2).Width = 110
        DataGridView1.Columns(3).HeaderText = "Всего отправок в работе"
        DataGridView1.Columns(3).Width = 110
        DataGridView1.Columns(4).HeaderText = "Неотгруженных отправок"
        DataGridView1.Columns(4).Width = 110
        DataGridView1.Columns(5).HeaderText = "Непринятых отправок"
        DataGridView1.Columns(5).Width = 110
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

        '---загрузка данных
        DataLoading()
        '---проверка состояния кнопок
        CheckButtonsState()
    End Sub

    Private Sub ComboBox1_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles ComboBox1.SelectedIndexChanged
        '////////////////////////////////////////////////////////////////////////////////
        '//
        '// Загрузка данных в соответствии с выбранным складом
        '//
        '////////////////////////////////////////////////////////////////////////////////

        '---загрузка данных
        DataLoading()
        '---проверка состояния кнопок
        CheckButtonsState()
    End Sub

    Private Sub ComboBox2_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles ComboBox2.SelectedIndexChanged
        '////////////////////////////////////////////////////////////////////////////////
        '//
        '// Загрузка данных в соответствии с выбранной опцией - отображать все склады или только активные
        '//
        '////////////////////////////////////////////////////////////////////////////////

        '---загрузка данных
        DataLoading()
        '---проверка состояния кнопок
        CheckButtonsState()
    End Sub

    Private Sub DataGridView1_CellFormatting(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellFormattingEventArgs) Handles DataGridView1.CellFormatting
        '////////////////////////////////////////////////////////////////////////////////
        '//
        '// подсветка информации по складам
        '//
        '////////////////////////////////////////////////////////////////////////////////
        Dim row As DataGridViewRow

        row = DataGridView1.Rows(e.RowIndex)
        row.Cells(0).Style.Font = New Font(DataGridView1.DefaultCellStyle.Font, FontStyle.Bold)
        row.Cells(1).Style.Font = New Font(DataGridView1.DefaultCellStyle.Font, FontStyle.Bold)
        If Trim(row.Cells(2).Value.ToString) <> "0" Then
            row.Cells(2).Style.BackColor = Color.LightCoral
        Else
            row.Cells(2).Style.BackColor = Color.Empty
        End If
        If Trim(row.Cells(3).Value.ToString) <> "0" Then
            row.Cells(3).Style.BackColor = Color.LightGreen
        Else
            row.Cells(3).Style.BackColor = Color.Empty
        End If
        If Trim(row.Cells(4).Value.ToString) <> "0" Then
            row.Cells(4).Style.BackColor = Color.LightCoral
        Else
            row.Cells(4).Style.BackColor = Color.Empty
        End If
        If Trim(row.Cells(5).Value.ToString) <> "0" Then
            row.Cells(5).Style.BackColor = Color.LightCoral
        Else
            row.Cells(5).Style.BackColor = Color.Empty
        End If
    End Sub

    Private Sub Button2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button2.Click
        '////////////////////////////////////////////////////////////////////////////////
        '//
        '// открытиие окна формирования отгрузок
        '//
        '////////////////////////////////////////////////////////////////////////////////

        OpenConsolidationWindow()
    End Sub

    Private Sub DataGridView1_DoubleClick(ByVal sender As Object, ByVal e As System.EventArgs) Handles DataGridView1.DoubleClick
        '////////////////////////////////////////////////////////////////////////////////
        '//
        '// открытиие окна формирования отгрузок
        '//
        '////////////////////////////////////////////////////////////////////////////////

        OpenConsolidationWindow()
    End Sub

    Private Sub OpenConsolidationWindow()
        '////////////////////////////////////////////////////////////////////////////////
        '//
        '// открытиие окна формирования консолидированных заказов
        '//
        '////////////////////////////////////////////////////////////////////////////////

        Declarations.WHToCode = Trim(Me.DataGridView1.SelectedRows.Item(0).Cells(0).Value.ToString())
        Declarations.WHFromCode = Trim(Me.ComboBox1.SelectedValue)
        MyConsolidatedOrders = New ConsolidatedOrders
        MyConsolidatedOrders.ShowDialog()
        DataLoading()
        '---текущей строкой сделать редактированную
        For i As Integer = 0 To DataGridView1.Rows.Count - 1
            If Trim(DataGridView1.Item(0, i).Value.ToString) = Declarations.WHToCode Then
                DataGridView1.CurrentCell = DataGridView1.Item(0, i)
            End If
        Next
        CheckButtonsState()
    End Sub
End Class
