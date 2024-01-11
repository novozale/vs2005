Public Class ShipmentCost

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        '////////////////////////////////////////////////////////////////////////////////
        '//
        '// Закрытие окна ввода стоимостей доставки
        '//
        '////////////////////////////////////////////////////////////////////////////////

        Me.Close()
    End Sub

    Private Sub ShipmentCost_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles Me.KeyDown
        '/////////////////////////////////////////////////////////////////////////////////////
        '//
        '// Блокировка ALT + F4
        '// 
        '/////////////////////////////////////////////////////////////////////////////////////

        If e.KeyData = Keys.Alt + Keys.F4 Then
            e.Handled = True
        End If

    End Sub

    Private Sub ShipmentCost_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        '////////////////////////////////////////////////////////////////////////////////
        '//
        '// Загрузка формы, загрузка данных в форму
        '//
        '////////////////////////////////////////////////////////////////////////////////

        LoadFlag = 0

        '---Фильтр по датам
        DateTimePicker1.Value = DateAdd(DateInterval.Month, -1, CDate("01/" & DatePart(DateInterval.Month, Now()) & "/" & DatePart(DateInterval.Year, Now())))
        DateTimePicker2.Value = CDate(DatePart(DateInterval.Day, Now()) & "/" & DatePart(DateInterval.Month, Now()) & "/" & DatePart(DateInterval.Year, Now()))

        '---Список перевозчиков
        PrepareCarrierList()
        ComboBox1.SelectedValue = "----"

        LoadFlag = 1

        '---Загрузка данных
        DataLoading()

        '---Проверка состояния кнопок
        CheckButtonsState()
    End Sub

    Private Sub PrepareCarrierList()
        '////////////////////////////////////////////////////////////////////////////////
        '//
        '// Подготовка списка перевозчиков
        '//
        '////////////////////////////////////////////////////////////////////////////////
        Dim MySQLStr As String
        Dim MyAdapter As SqlClient.SqlDataAdapter     'для списка продавцов
        Dim MyDs As New DataSet                       '

        MySQLStr = "SELECT tbl_ShipmentsCost_Fact.PL01001, tbl_ShipmentsCost_Fact.PL01001 + ' ' + ISNULL(PL010300.PL01002, N'') AS PL01002 "
        MySQLStr = MySQLStr & "FROM tbl_ShipmentsCost_Fact WITH (NOLOCK) INNER JOIN "
        MySQLStr = MySQLStr & "PL010300 ON tbl_ShipmentsCost_Fact.PL01001 = PL010300.PL01001 "
        MySQLStr = MySQLStr & "UNION "
        MySQLStr = MySQLStr & "SELECT '----' AS PL01001, '    Все перевозчики' AS PL01002 "
        MySQLStr = MySQLStr & "ORDER BY tbl_ShipmentsCost_Fact.PL01001 "

        Try
            MyAdapter = New SqlClient.SqlDataAdapter(MySQLStr, Declarations.MyNETConnStr)
            MyAdapter.SelectCommand.CommandTimeout = 600
            MyAdapter.Fill(MyDs)
            ComboBox1.DisplayMember = "PL01002" 'Это то что будет отображаться
            ComboBox1.ValueMember = "PL01001"   'это то что будет храниться
            ComboBox1.DataSource = MyDs.Tables(0).DefaultView
        Catch ex As Exception
            MsgBox(ex.ToString)
        End Try

    End Sub

    Private Function DataLoading()
        '/////////////////////////////////////////////////////////////////////////////////////
        '//
        '// Загрузка списка действий (в соответствии с параметрами)
        '// 
        '/////////////////////////////////////////////////////////////////////////////////////
        Dim MySQLStr As String
        Dim MyAdapter As SqlClient.SqlDataAdapter     'для списка оплат
        Dim MyDs As New DataSet                       '

        If LoadFlag = 1 Then
            MySQLStr = "SELECT tbl_ShipmentsCost_Fact.ID, tbl_ShipmentsCost_Fact.PL01001 AS CarrID, ISNULL(PL010300.PL01002, N'') AS CarrName, "
            MySQLStr = MySQLStr & "LTRIM(RTRIM(LTRIM(RTRIM(ISNULL(PL010300.PL01003, N''))) + ' ' + LTRIM(RTRIM(ISNULL(PL010300.PL01004, N''))) "
            MySQLStr = MySQLStr & "+ ' ' + LTRIM(RTRIM(ISNULL(PL010300.PL01005, N''))))) AS CarrAddr, tbl_ShipmentsCost_Fact.PL03002, "
            MySQLStr = MySQLStr & "tbl_ShipmentsCost_Fact.DocDate, tbl_ShipmentsCost_Fact.DocSumm, ISNULL(View_1.InvoicesNum, 0) AS InvoicesNum, "
            MySQLStr = MySQLStr & "ISNULL(tbl_ShipmentsCost_ShipmentsType.Name, N'') AS DelWay "
            MySQLStr = MySQLStr & "FROM tbl_ShipmentsCost_Fact WITH (NOLOCK) LEFT OUTER JOIN "
            MySQLStr = MySQLStr & "(SELECT DocID, COUNT(DocID) AS InvoicesNum "
            MySQLStr = MySQLStr & "FROM tbl_ShipmentsCost_FactByInvoices WITH (NOLOCK) "
            MySQLStr = MySQLStr & "GROUP BY DocID) AS View_1 ON tbl_ShipmentsCost_Fact.ID = View_1.DocID LEFT OUTER JOIN "
            MySQLStr = MySQLStr & "tbl_ShipmentsCost_ShipmentsType ON tbl_ShipmentsCost_Fact.ShipmentsType = tbl_ShipmentsCost_ShipmentsType.ID LEFT OUTER JOIN "
            MySQLStr = MySQLStr & "PL010300 ON tbl_ShipmentsCost_Fact.PL01001 = PL010300.PL01001 "
            MySQLStr = MySQLStr & "WHERE (tbl_ShipmentsCost_Fact.DocDate >= CONVERT(DATETIME, '" & Format(DateTimePicker1.Value, "dd/MM/yyyy") & "', 103)) AND "
            MySQLStr = MySQLStr & "(tbl_ShipmentsCost_Fact.DocDate <= CONVERT(DATETIME, '" & Format(DateTimePicker2.Value, "dd/MM/yyyy") & "', 103)) "
            If ComboBox1.SelectedValue <> "----" Then
                MySQLStr = MySQLStr & " AND (tbl_ShipmentsCost_Fact.PL01001 = N'" & ComboBox1.SelectedValue & "') "
            End If
            MySQLStr = MySQLStr & "ORDER BY CarrID, DocDate "

            Try
                MyAdapter = New SqlClient.SqlDataAdapter(MySQLStr, Declarations.MyNETConnStr)
                MyAdapter.SelectCommand.CommandTimeout = 600
                MyAdapter.Fill(MyDs)
                DataGridView1.DataSource = MyDs.Tables(0)
            Catch ex As Exception
                MsgBox(ex.ToString)
            End Try

            '---заголовки
            DataGridView1.Columns(0).HeaderText = "ID"
            DataGridView1.Columns(0).Width = 0
            DataGridView1.Columns(0).Visible = False
            DataGridView1.Columns(1).HeaderText = "Код перево зчика"
            DataGridView1.Columns(1).Width = 70
            DataGridView1.Columns(2).HeaderText = "Перевозчик"
            DataGridView1.Columns(2).Width = 170
            DataGridView1.Columns(3).HeaderText = "Адрес перевозчика"
            DataGridView1.Columns(3).Width = 280
            DataGridView1.Columns(4).HeaderText = "N документа"
            DataGridView1.Columns(4).Width = 100
            DataGridView1.Columns(5).HeaderText = "Дата документа"
            DataGridView1.Columns(5).Width = 100
            DataGridView1.Columns(6).HeaderText = "Сумма документа (РУБ)"
            DataGridView1.Columns(6).Width = 100
            DataGridView1.Columns(7).HeaderText = "Кол-во доставл. документов"
            DataGridView1.Columns(7).Width = 100
            DataGridView1.Columns(8).HeaderText = "Способ доставки"
            DataGridView1.Columns(8).Width = 160

        End If
    End Function

    Private Sub CheckButtonsState()
        '/////////////////////////////////////////////////////////////////////////////////////
        '//
        '// Проверка и выставление состояния кнопок
        '// 
        '/////////////////////////////////////////////////////////////////////////////////////

        If LoadFlag = 1 Then
            If DataGridView1.SelectedRows.Count = 0 Then
                Button2.Enabled = False
                Button3.Enabled = False
                Button4.Enabled = False
            Else
                Button2.Enabled = True
                Button3.Enabled = True
                Button4.Enabled = True
            End If
        End If
    End Sub

    Private Sub Button5_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button5.Click
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

    Private Sub Button4_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button4.Click
        '////////////////////////////////////////////////////////////////////////////////
        '//
        '// просмотр записи о доставке
        '//
        '////////////////////////////////////////////////////////////////////////////////

        ViewAction()
    End Sub

    Private Sub ViewAction()
        '////////////////////////////////////////////////////////////////////////////////
        '//
        '// просмотр записи о доставке
        '//
        '////////////////////////////////////////////////////////////////////////////////

        MyAddShipmentCost = New AddShipmentCost
        MyAddShipmentCost.StartParam = "View"
        MyAddShipmentCost.ShowDialog()
    End Sub


    Private Sub DataGridView1_DoubleClick(ByVal sender As Object, ByVal e As System.EventArgs) Handles DataGridView1.DoubleClick
        '////////////////////////////////////////////////////////////////////////////////
        '//
        '// просмотр записи о доставке
        '//
        '////////////////////////////////////////////////////////////////////////////////

        ViewAction()
    End Sub

    Private Sub ComboBox1_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles ComboBox1.SelectedIndexChanged
        '////////////////////////////////////////////////////////////////////////////////
        '//
        '// Загрузка данных в соответствии с выбранным перевозчиком
        '//
        '////////////////////////////////////////////////////////////////////////////////

        '---загрузка данных
        DataLoading()
        '---проверка состояния кнопок
        CheckButtonsState()
    End Sub

    Private Sub DateTimePicker1_ValueChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles DateTimePicker1.ValueChanged
        '////////////////////////////////////////////////////////////////////////////////
        '//
        '// Загрузка данных в соответствии с выбранной начальной датой
        '//
        '////////////////////////////////////////////////////////////////////////////////

        '---загрузка данных
        DataLoading()
        '---проверка состояния кнопок
        CheckButtonsState()
    End Sub

    Private Sub DateTimePicker2_ValueChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles DateTimePicker2.ValueChanged
        '////////////////////////////////////////////////////////////////////////////////
        '//
        '// Загрузка данных в соответствии с выбранной конечной датой
        '//
        '////////////////////////////////////////////////////////////////////////////////

        '---загрузка данных
        DataLoading()
        '---проверка состояния кнопок
        CheckButtonsState()
    End Sub

    Private Sub Button6_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button6.Click
        '////////////////////////////////////////////////////////////////////////////////
        '//
        '// Нажатие кнопки добавление новой записи
        '//
        '////////////////////////////////////////////////////////////////////////////////

        CreateNewRecord()
    End Sub

    Private Sub CreateNewRecord()
        '////////////////////////////////////////////////////////////////////////////////
        '//
        '// Добавление новой записи о доставке
        '//
        '////////////////////////////////////////////////////////////////////////////////

        MyAddShipmentCost = New AddShipmentCost
        MyAddShipmentCost.StartParam = "Create"
        MyAddShipmentCost.ShowDialog()
        '---загрузка данных
        DataLoading()
        '---текущей строкой сделать созданную
        For i As Integer = 0 To DataGridView1.Rows.Count - 1
            If Trim(DataGridView1.Item(0, i).Value.ToString) = Declarations.MyRecordID Then
                DataGridView1.CurrentCell = DataGridView1.Item(1, i)
            End If
        Next
        '---проверка состояния кнопок
        CheckButtonsState()
    End Sub

    Private Sub Button3_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button3.Click
        '////////////////////////////////////////////////////////////////////////////////
        '//
        '// Нажатие кнопки редактирования новой записи
        '//
        '////////////////////////////////////////////////////////////////////////////////

        EditRecord()
    End Sub

    Private Sub EditRecord()
        '////////////////////////////////////////////////////////////////////////////////
        '//
        '// Редактирование новой записи о доставке
        '//
        '////////////////////////////////////////////////////////////////////////////////

        MyAddShipmentCost = New AddShipmentCost
        MyAddShipmentCost.StartParam = "Edit"
        MyAddShipmentCost.ShowDialog()
        '---загрузка данных
        DataLoading()
        '---текущей строкой сделать редактированную
        For i As Integer = 0 To DataGridView1.Rows.Count - 1
            If Trim(DataGridView1.Item(0, i).Value.ToString) = Declarations.MyRecordID Then
                DataGridView1.CurrentCell = DataGridView1.Item(1, i)
            End If
        Next
        '---проверка состояния кнопок
        CheckButtonsState()
    End Sub

    Private Sub Button2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button2.Click
        '////////////////////////////////////////////////////////////////////////////////
        '//
        '// Нажатие кнопки удаления записи о доставке
        '//
        '////////////////////////////////////////////////////////////////////////////////
        Dim MyRez As MsgBoxResult

        MyRez = MsgBox("Вы уверены, что хотите удалить выбранную запись?", MsgBoxStyle.YesNo, "Внимание!")
        If MyRez = MsgBoxResult.Yes Then
            DeleteRecord()
        End If
    End Sub

    Private Sub DeleteRecord()
        '////////////////////////////////////////////////////////////////////////////////
        '//
        '// Удаление записи о доставке
        '//
        '////////////////////////////////////////////////////////////////////////////////
        Dim MySQLStr As String

        Declarations.MyRecordID = Trim(MyShipmentCost.DataGridView1.SelectedRows.Item(0).Cells(0).Value.ToString())

        MySQLStr = "DELETE FROM tbl_ShipmentsCost_FactByInvoices "
        MySQLStr = MySQLStr & "WHERE (DocID = N'" & Declarations.MyRecordID & "')"
        InitMyConn(False)
        Declarations.MyConn.Execute(MySQLStr)

        MySQLStr = "DELETE FROM tbl_ShipmentsCost_Fact "
        MySQLStr = MySQLStr & "WHERE (ID = '" & Declarations.MyRecordID & "')"
        InitMyConn(False)
        Declarations.MyConn.Execute(MySQLStr)

        '---загрузка данных
        DataLoading()
        '---проверка состояния кнопок
        CheckButtonsState()
    End Sub
End Class