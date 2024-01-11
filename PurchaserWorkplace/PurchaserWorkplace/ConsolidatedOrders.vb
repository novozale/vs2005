Public Class ConsolidatedOrders

    Private Sub ConsolidatedOrders_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles Me.KeyDown
        '////////////////////////////////////////////////////////////////////////////////
        '//
        '// Запрет закрытия окна по ALT+F4
        '//
        '////////////////////////////////////////////////////////////////////////////////

        If e.KeyData = Keys.Alt + Keys.F4 Then
            e.Handled = True
        End If
    End Sub

    Private Sub ConsolidatedOrders_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        '////////////////////////////////////////////////////////////////////////////////
        '//
        '// открытиие окна формирования консолидированных заказов
        '//
        '////////////////////////////////////////////////////////////////////////////////
        Dim MySQLStr As String                        'рабочая строка
        Dim MyAdapter As SqlClient.SqlDataAdapter     'для списка складов
        Dim MyDs As New DataSet                       '

        '---Список складов
        MySQLStr = "SELECT SC23001 AS WHCode, SC23001 + ' ' + SC23002 AS WHName "
        MySQLStr = MySQLStr & "FROM SC230300 WITH(NOLOCK) "
        MySQLStr = MySQLStr & "WHERE (SC23006 = N'1') OR (LEFT(SC23006, 2) = N'TR') "
        MySQLStr = MySQLStr & "ORDER BY WHCode "

        Try
            MyAdapter = New SqlClient.SqlDataAdapter(MySQLStr, Declarations.MyNETConnStr)
            MyAdapter.SelectCommand.CommandTimeout = 600
            MyAdapter.Fill(MyDs)
            ComboBox2.DisplayMember = "WHName" 'Это то что будет отображаться
            ComboBox2.ValueMember = "WHCode"   'это то что будет храниться
            ComboBox2.DataSource = MyDs.Tables(0).DefaultView
        Catch ex As Exception
            MsgBox(ex.ToString)
        End Try

        Label9.Text = Declarations.MySupplierCode
        Label2.Text = Trim(MainForm.DataGridView1.SelectedRows.Item(0).Cells(1).Value.ToString())
        Dim rootWindow As MainForm = TryCast(Application.OpenForms(0), MainForm)
        ComboBox2.SelectedValue = rootWindow.ComboBox1.SelectedValue
        ComboBox1.SelectedText = "Только активные (непринятые)"

        LoadConsolidatedOrders()
        LoadIncludedOrders()
        LoadFreeOrders()
        CheckConsolidatedButtons()
        CheckRemoveButtons()
        CheckAddButtons()

    End Sub

    Private Sub LoadConsolidatedOrders()
        '////////////////////////////////////////////////////////////////////////////////
        '//
        '// Загрузка информации по обобщенным заказам
        '//
        '////////////////////////////////////////////////////////////////////////////////
        Dim MySQLStr As String
        Dim MyAdapter As SqlClient.SqlDataAdapter     'для списка заказов
        Dim MyDs As New DataSet

        MySQLStr = "spp_PurchaseWorkplace_TotalGroupOrdersPrep N'" & Declarations.MyWH & "', N'" & Declarations.MySupplierCode & "', "
        If ComboBox1.Text = "Только активные (непринятые)" Then
            MySQLStr = MySQLStr & "1, "
        Else
            MySQLStr = MySQLStr & "0, "
        End If
        MySQLStr = MySQLStr & "0 "

        Try
            MyAdapter = New SqlClient.SqlDataAdapter(MySQLStr, Declarations.MyNETConnStr)
            MyAdapter.SelectCommand.CommandTimeout = 600
            MyAdapter.Fill(MyDs)
            DataGridView1.DataSource = MyDs.Tables(0)
        Catch ex As Exception
            MsgBox(ex.ToString)
        End Try


        '---заголовки
        DataGridView1.Columns(0).HeaderText = "N заказа"
        DataGridView1.Columns(0).Width = 80
        DataGridView1.Columns(1).HeaderText = "Дата заказа"
        DataGridView1.Columns(1).Width = 80
        DataGridView1.Columns(2).HeaderText = "Сумма заказа РУБ"
        DataGridView1.Columns(2).Width = 130
        DataGridView1.Columns(3).HeaderText = "Закупщик"
        DataGridView1.Columns(3).Width = 140
        DataGridView1.Columns(4).HeaderText = "Дата размешения заказа"
        DataGridView1.Columns(4).Width = 80
        DataGridView1.Columns(5).HeaderText = "Дата подтверждения заказа"
        DataGridView1.Columns(5).Width = 80
        DataGridView1.Columns(6).HeaderText = "N заказа поставщика"
        DataGridView1.Columns(6).Width = 130
        DataGridView1.Columns(7).HeaderText = "Не принят"
        DataGridView1.Columns(7).Width = 100
        DataGridView1.Columns(7).Visible = False
        DataGridView1.Columns(8).HeaderText = "Подтвержденная дата поставки"
        DataGridView1.Columns(8).Width = 80
        DataGridView1.Columns(9).HeaderText = "Кол-во включенных заказов"
        DataGridView1.Columns(9).Width = 90
        DataGridView1.Columns(10).HeaderText = "Первона чальная дата поставки"
        DataGridView1.Columns(10).Width = 80
        DataGridView1.Columns(11).HeaderText = "Контрольная дата"
        DataGridView1.Columns(11).Width = 80
        DataGridView1.Columns(12).HeaderText = "Контактная информация"
        DataGridView1.Columns(12).Width = 320
        DataGridView1.Columns(13).HeaderText = "Комментарии"
        DataGridView1.Columns(13).Width = 320
        DataGridView1.Columns(14).HeaderText = "Файл со счетом"
        DataGridView1.Columns(14).Width = 150
    End Sub

    Private Sub CheckConsolidatedButtons()
        '////////////////////////////////////////////////////////////////////////////////
        '//
        '// Проверка состояния кнопок работы с обобщенными заказами
        '//
        '////////////////////////////////////////////////////////////////////////////////

        If DataGridView1.SelectedRows.Count = 0 Then
            Button2.Enabled = False
            Button4.Enabled = False
            Button7.Enabled = False
            Button6.Enabled = False
            Button14.Enabled = False
            Button16.Enabled = False
            Button17.Enabled = False
            Button19.Enabled = False
        Else
            Button17.Enabled = True
            If DataGridView1.SelectedRows.Count = 0 Then
                Button2.Enabled = False
                Button4.Enabled = False
            Else
                Button2.Enabled = True
                Button4.Enabled = True
            End If
            If Trim(DataGridView1.SelectedRows.Item(0).Cells(7).Value.ToString()) = "" Then
                Button7.Enabled = False
            Else
                Button7.Enabled = True
            End If
            If DataGridView2.SelectedRows.Count = 0 Then
                Button6.Enabled = True
            Else
                Button6.Enabled = False
            End If
            If Trim(DataGridView1.SelectedRows.Item(0).Cells(4).Value.ToString()) = "" And _
                Trim(DataGridView1.SelectedRows.Item(0).Cells(7).Value.ToString()) = "" Then
                Button14.Enabled = True
            Else
                Button14.Enabled = False
            End If
            Button16.Enabled = True
            If Trim(DataGridView1.SelectedRows.Item(0).Cells(14).Value.ToString()) <> "" Then
                Button19.Enabled = True
            Else
                Button19.Enabled = False
            End If
        End If
    End Sub

    Private Sub LoadIncludedOrders()
        '////////////////////////////////////////////////////////////////////////////////
        '//
        '// Загрузка информации по заказам, включенным в избранный обобщенный
        '//
        '////////////////////////////////////////////////////////////////////////////////
        Dim MySQLStr As String
        Dim MyAdapter As SqlClient.SqlDataAdapter     'для списка заказов
        Dim MyDs As New DataSet

        If DataGridView1.SelectedRows.Count <> 0 Then
            MySQLStr = "spp_PurchaseWorkplace_GroupOrdersPrep N'" & Trim(DataGridView1.SelectedRows.Item(0).Cells(0).Value.ToString()) & "'"
        Else
            MySQLStr = "spp_PurchaseWorkplace_GroupOrdersPrep N'0'"
        End If

        Try
            MyAdapter = New SqlClient.SqlDataAdapter(MySQLStr, Declarations.MyNETConnStr)
            MyAdapter.SelectCommand.CommandTimeout = 600
            MyAdapter.Fill(MyDs)
            DataGridView2.DataSource = MyDs.Tables(0)
        Catch ex As Exception
            MsgBox(ex.ToString)
        End Try


        '---заголовки
        DataGridView2.Columns(0).HeaderText = "N заказа"
        DataGridView2.Columns(0).Width = 75
        DataGridView2.Columns(1).HeaderText = "Дата заказа"
        DataGridView2.Columns(1).Width = 75
        DataGridView2.Columns(2).HeaderText = "Сумма РУБ"
        DataGridView2.Columns(2).Width = 100
        DataGridView2.Columns(3).HeaderText = "Закупщик"
        DataGridView2.Columns(3).Width = 150
        DataGridView2.Columns(4).HeaderText = "N заказа на продажу"
        DataGridView2.Columns(4).Width = 75
        DataGridView2.Columns(5).HeaderText = "Склад назна чения"
        DataGridView2.Columns(5).Width = 70
        DataGridView2.Columns(6).HeaderText = "Продавец"
        DataGridView2.Columns(6).Width = 200
        DataGridView2.Columns(7).HeaderText = "Покупатель"
        DataGridView2.Columns(7).Width = 200

    End Sub

    Private Sub LoadFreeOrders()
        '////////////////////////////////////////////////////////////////////////////////
        '//
        '// Загрузка информации по доступным к включению в обобщенный заказ заказам
        '//
        '////////////////////////////////////////////////////////////////////////////////
        Dim MySQLStr As String
        Dim MyAdapter As SqlClient.SqlDataAdapter     'для списка заказов
        Dim MyDs As New DataSet

        MySQLStr = "spp_PurchaseWorkplace_NonGroupOrdersPrep N'" & Declarations.MyWH & "', N'" & Declarations.MySupplierCode & "'"

        Try
            MyAdapter = New SqlClient.SqlDataAdapter(MySQLStr, Declarations.MyNETConnStr)
            MyAdapter.SelectCommand.CommandTimeout = 600
            MyAdapter.Fill(MyDs)
            DataGridView3.DataSource = MyDs.Tables(0)
        Catch ex As Exception
            MsgBox(ex.ToString)
        End Try

        '---заголовки
        DataGridView3.Columns(0).HeaderText = "N заказа"
        DataGridView3.Columns(0).Width = 75
        DataGridView3.Columns(1).HeaderText = "Дата заказа"
        DataGridView3.Columns(1).Width = 75
        DataGridView3.Columns(2).HeaderText = "Сумма РУБ"
        DataGridView3.Columns(2).Width = 100
        DataGridView3.Columns(3).HeaderText = "Закупщик"
        DataGridView3.Columns(3).Width = 150
        DataGridView3.Columns(4).HeaderText = "N заказа на продажу"
        DataGridView3.Columns(4).Width = 75
        DataGridView3.Columns(5).HeaderText = "Склад назна чения"
        DataGridView3.Columns(5).Width = 70
        DataGridView3.Columns(6).HeaderText = "Продавец"
        DataGridView3.Columns(6).Width = 200
        DataGridView3.Columns(7).HeaderText = "Покупатель"
        DataGridView3.Columns(7).Width = 200

    End Sub

    Private Sub CheckAddButtons()
        '////////////////////////////////////////////////////////////////////////////////
        '//
        '// Проверка состояния кнопок добавления заказа в обобщенный
        '//
        '////////////////////////////////////////////////////////////////////////////////

        If DataGridView1.SelectedRows.Count = 0 Or DataGridView3.SelectedRows.Count = 0 Then
            Button9.Enabled = False
            Button10.Enabled = False
        Else
            If DataGridView1.SelectedRows.Count = 0 Then
                Button9.Enabled = False
                Button10.Enabled = False
            Else
                If Trim(DataGridView1.SelectedRows.Item(0).Cells(4).Value.ToString()) <> "" Or _
                    Trim(DataGridView1.SelectedRows.Item(0).Cells(5).Value.ToString()) <> "" Or _
                    Trim(DataGridView1.SelectedRows.Item(0).Cells(6).Value.ToString()) <> "" Or _
                    Trim(DataGridView1.SelectedRows.Item(0).Cells(7).Value.ToString()) = "" Then
                    Button9.Enabled = False
                    Button10.Enabled = False
                Else

                    Button9.Enabled = True
                    Button10.Enabled = True
                End If
            End If
        End If
    End Sub

    Private Sub CheckRemoveButtons()
        '////////////////////////////////////////////////////////////////////////////////
        '//
        '// Проверка состояния кнопок удаления заказа из обобщенного
        '//
        '////////////////////////////////////////////////////////////////////////////////

        If DataGridView2.SelectedRows.Count = 0 Then
            Button11.Enabled = False
            Button12.Enabled = False
        Else
            If DataGridView1.SelectedRows.Count = 0 Then
                Button11.Enabled = False
                Button12.Enabled = False
            Else
                If Trim(DataGridView1.SelectedRows.Item(0).Cells(4).Value.ToString()) <> "" Or _
                    Trim(DataGridView1.SelectedRows.Item(0).Cells(5).Value.ToString()) <> "" Or _
                    Trim(DataGridView1.SelectedRows.Item(0).Cells(6).Value.ToString()) <> "" Or _
                    Trim(DataGridView1.SelectedRows.Item(0).Cells(7).Value.ToString()) = "" Then
                    Button11.Enabled = False
                    Button12.Enabled = False
                Else

                    Button11.Enabled = True
                    Button12.Enabled = True
                End If
            End If
        End If
    End Sub

    Private Sub Button13_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button13.Click
        '////////////////////////////////////////////////////////////////////////////////
        '//
        '// Выход из окна
        '//
        '////////////////////////////////////////////////////////////////////////////////

        Me.Close()
    End Sub

    Private Sub ComboBox1_SelectedValueChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles ComboBox1.SelectedValueChanged
        '////////////////////////////////////////////////////////////////////////////////
        '//
        '// Смена выбора - все заказы или только активные
        '//
        '////////////////////////////////////////////////////////////////////////////////

        LoadConsolidatedOrders()
        LoadIncludedOrders()
        LoadFreeOrders()
        CheckConsolidatedButtons()
        CheckRemoveButtons()
        CheckAddButtons()
    End Sub

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        '////////////////////////////////////////////////////////////////////////////////
        '//
        '// Обновление информации в окне
        '//
        '////////////////////////////////////////////////////////////////////////////////

        LoadConsolidatedOrders()
        LoadIncludedOrders()
        LoadFreeOrders()
        CheckConsolidatedButtons()
        CheckRemoveButtons()
        CheckAddButtons()
    End Sub

    Private Sub DataGridView1_CellFormatting(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellFormattingEventArgs) Handles DataGridView1.CellFormatting
        '////////////////////////////////////////////////////////////////////////////////
        '//
        '// подсветка информации по обобщенным заказам
        '//
        '////////////////////////////////////////////////////////////////////////////////
        Dim row As DataGridViewRow

        row = DataGridView1.Rows(e.RowIndex)
        If row.Cells(9).Value = 0 Then
            row.DefaultCellStyle.BackColor = Color.LightYellow
        Else
            If (Trim(row.Cells(4).Value.ToString) = "" Or Trim(row.Cells(5).Value.ToString) = "") And Trim(row.Cells(7).Value.ToString) = "" Then
                row.DefaultCellStyle.BackColor = Color.LightPink
            Else
                row.DefaultCellStyle.BackColor = Color.Empty
            End If
        End If
        If Trim(row.Cells(7).Value.ToString) <> "" Then
            If IsDBNull(row.Cells(8).Value) Then
                row.Cells(8).Style.BackColor = Color.Empty
            Else
                If row.Cells(8).Value < Now() Then
                    row.Cells(8).Style.BackColor = Color.LightPink
                ElseIf row.Cells(8).Value < DateAdd(DateInterval.Day, 3, Now()) Then
                    row.Cells(8).Style.BackColor = Color.Yellow
                Else
                    row.Cells(8).Style.BackColor = Color.Empty
                End If
            End If
            If IsDBNull(row.Cells(11).Value) Then
                row.Cells(11).Style.BackColor = Color.Empty
            Else
                If row.Cells(11).Value < DateAdd(DateInterval.Day, -2, Now()) Then
                    row.Cells(11).Style.BackColor = Color.Empty
                ElseIf row.Cells(11).Value <= DateAdd(DateInterval.Day, 3, Now()) Then
                    row.Cells(11).Style.BackColor = Color.Yellow
                Else
                    row.Cells(11).Style.BackColor = Color.Empty
                End If
            End If
        Else
            row.Cells(8).Style.BackColor = Color.Empty
            row.Cells(11).Style.BackColor = Color.Empty
        End If
    End Sub

    Private Sub DataGridView1_SelectionChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles DataGridView1.SelectionChanged
        '////////////////////////////////////////////////////////////////////////////////
        '//
        '// смена выбора обобщенного заказа
        '//
        '////////////////////////////////////////////////////////////////////////////////

        'LoadConsolidatedOrders()
        LoadIncludedOrders()
        LoadFreeOrders()
        CheckConsolidatedButtons()
        CheckRemoveButtons()
        CheckAddButtons()
    End Sub

    Private Sub Button8_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button8.Click
        '////////////////////////////////////////////////////////////////////////////////
        '//
        '// Создание нового обобщенного заказа
        '//
        '////////////////////////////////////////////////////////////////////////////////

        MyEditConsolidatedOrder = New EditConsolidatedOrder
        MyEditConsolidatedOrder.StartParam = "Create"
        MyEditConsolidatedOrder.ShowDialog()
        LoadConsolidatedOrders()
        '---текущей строкой сделать созданную
        For i As Integer = 0 To DataGridView1.Rows.Count - 1
            If Trim(DataGridView1.Item(0, i).Value.ToString) = Declarations.MyOrderID Then
                DataGridView1.CurrentCell = DataGridView1.Item(0, i)
            End If
        Next
        LoadIncludedOrders()
        LoadFreeOrders()
        CheckConsolidatedButtons()
        CheckRemoveButtons()
        CheckAddButtons()
    End Sub

    Private Sub Button7_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button7.Click
        '////////////////////////////////////////////////////////////////////////////////
        '//
        '// Редактирование обобщенного заказа
        '//
        '////////////////////////////////////////////////////////////////////////////////

        Declarations.MyOrderID = Trim(DataGridView1.SelectedRows.Item(0).Cells(0).Value.ToString())
        If CheckUOMInOrders(0) = True Then
            MyEditConsolidatedOrder = New EditConsolidatedOrder
            MyEditConsolidatedOrder.StartParam = "Edit"
            MyEditConsolidatedOrder.ShowDialog()
            LoadConsolidatedOrders()
            '---текущей строкой сделать редактированную
            For i As Integer = 0 To DataGridView1.Rows.Count - 1
                If Trim(DataGridView1.Item(0, i).Value.ToString) = Declarations.MyOrderID Then
                    DataGridView1.CurrentCell = DataGridView1.Item(0, i)
                End If
            Next
            LoadIncludedOrders()
            LoadFreeOrders()
            CheckConsolidatedButtons()
            CheckRemoveButtons()
            CheckAddButtons()
        End If
    End Sub

    Private Sub Button6_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button6.Click
        '////////////////////////////////////////////////////////////////////////////////
        '//
        '// Удаление обобщенного заказа
        '//
        '////////////////////////////////////////////////////////////////////////////////
        Dim MySQLStr As String
        Dim MyRez As MsgBoxResult

        MyRez = MsgBox("Вы действительно хотите удалить обобщенный заказ?", MsgBoxStyle.YesNo, "Внимание!")
        If MyRez = MsgBoxResult.Yes Then
            MySQLStr = "DELETE FROM tbl_PurchaseWorkplace_ConsolidatedOrders "
            MySQLStr = MySQLStr & "WHERE (ID = N'" & Trim(DataGridView1.SelectedRows.Item(0).Cells(0).Value.ToString()) & "') "
            InitMyConn(False)
            Declarations.MyConn.Execute(MySQLStr)
            LoadConsolidatedOrders()
            LoadIncludedOrders()
            LoadFreeOrders()
            CheckConsolidatedButtons()
            CheckRemoveButtons()
            CheckAddButtons()
        End If
    End Sub

    Private Sub Button9_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button9.Click
        '////////////////////////////////////////////////////////////////////////////////
        '//
        '// Включение заказа в обобщенный заказ
        '//
        '////////////////////////////////////////////////////////////////////////////////
        Dim MySQLStr As String

        Declarations.MyOrderID = Trim(DataGridView1.SelectedRows.Item(0).Cells(0).Value.ToString())
        If CheckUOMInOrders(1) = True Then
            MySQLStr = "UPDATE PC010300 "
            MySQLStr = MySQLStr & "SET PC01052 = N'" & Trim(DataGridView1.SelectedRows.Item(0).Cells(0).Value.ToString()) & "' "
            MySQLStr = MySQLStr & "WHERE (PC01001 = N'" & Trim(DataGridView3.SelectedRows.Item(0).Cells(0).Value.ToString()) & "') "
            InitMyConn(False)
            Declarations.MyConn.Execute(MySQLStr)
            LoadConsolidatedOrders()
            '---текущей строкой сделать редактированную
            For i As Integer = 0 To DataGridView1.Rows.Count - 1
                If Trim(DataGridView1.Item(0, i).Value.ToString) = Declarations.MyOrderID Then
                    DataGridView1.CurrentCell = DataGridView1.Item(0, i)
                End If
            Next
            LoadIncludedOrders()
            LoadFreeOrders()
            CheckConsolidatedButtons()
            CheckRemoveButtons()
            CheckAddButtons()
        End If
    End Sub

    Private Sub Button11_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button11.Click
        '////////////////////////////////////////////////////////////////////////////////
        '//
        '// Удаление заказа из обобщенного заказа
        '//
        '////////////////////////////////////////////////////////////////////////////////
        Dim MySQLStr As String

        Declarations.MyOrderID = Trim(DataGridView1.SelectedRows.Item(0).Cells(0).Value.ToString())
        MySQLStr = "UPDATE PC010300 "
        MySQLStr = MySQLStr & "SET PC01052 = N'' "
        MySQLStr = MySQLStr & "WHERE (PC01001 = N'" & Trim(DataGridView2.SelectedRows.Item(0).Cells(0).Value.ToString()) & "') "
        InitMyConn(False)
        Declarations.MyConn.Execute(MySQLStr)
        LoadConsolidatedOrders()
        '---текущей строкой сделать редактированную
        For i As Integer = 0 To DataGridView1.Rows.Count - 1
            If Trim(DataGridView1.Item(0, i).Value.ToString) = Declarations.MyOrderID Then
                DataGridView1.CurrentCell = DataGridView1.Item(0, i)
            End If
        Next
        LoadIncludedOrders()
        LoadFreeOrders()
        CheckConsolidatedButtons()
        CheckRemoveButtons()
        CheckAddButtons()
    End Sub

    Private Sub Button12_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button12.Click
        '////////////////////////////////////////////////////////////////////////////////
        '//
        '// Удаление всех заказов из обобщенного заказа
        '//
        '////////////////////////////////////////////////////////////////////////////////
        Dim MySQLStr As String

        Declarations.MyOrderID = Trim(DataGridView1.SelectedRows.Item(0).Cells(0).Value.ToString())
        MySQLStr = "UPDATE PC010300 "
        MySQLStr = MySQLStr & "SET PC01052 = N'' "
        MySQLStr = MySQLStr & "WHERE (PC01052 = N'" & Trim(DataGridView1.SelectedRows.Item(0).Cells(0).Value.ToString()) & "') "
        InitMyConn(False)
        Declarations.MyConn.Execute(MySQLStr)
        LoadConsolidatedOrders()
        '---текущей строкой сделать редактированную
        For i As Integer = 0 To DataGridView1.Rows.Count - 1
            If Trim(DataGridView1.Item(0, i).Value.ToString) = Declarations.MyOrderID Then
                DataGridView1.CurrentCell = DataGridView1.Item(0, i)
            End If
        Next
        LoadIncludedOrders()
        LoadFreeOrders()
        CheckConsolidatedButtons()
        CheckRemoveButtons()
        CheckAddButtons()
    End Sub

    Private Sub Button10_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button10.Click
        '////////////////////////////////////////////////////////////////////////////////
        '//
        '// Включение всех доступных заказов в обобщенный заказ
        '//
        '////////////////////////////////////////////////////////////////////////////////
        Dim MySQLStr As String

        Declarations.MyOrderID = Trim(DataGridView1.SelectedRows.Item(0).Cells(0).Value.ToString())
        If CheckUOMInOrders(2) = True Then
            MySQLStr = "UPDATE PC010300 "
            MySQLStr = MySQLStr & "SET PC01052 = N'" & Trim(DataGridView1.SelectedRows.Item(0).Cells(0).Value.ToString()) & "' "
            MySQLStr = MySQLStr & "WHERE (PC01001 IN "
            MySQLStr = MySQLStr & "(SELECT PC010300_1.PC01001 AS OrderN "
            MySQLStr = MySQLStr & "FROM PC010300 AS PC010300_1 INNER JOIN "
            MySQLStr = MySQLStr & "(SELECT PC03001 "
            MySQLStr = MySQLStr & "FROM PC030300 "
            MySQLStr = MySQLStr & "WHERE (PC03010 <> 0) OR "
            MySQLStr = MySQLStr & "(PC03011 <> 0) "
            MySQLStr = MySQLStr & "GROUP BY PC03001) AS View_2 ON PC010300_1.PC01001 = View_2.PC03001 "
            MySQLStr = MySQLStr & "WHERE (PC010300_1.PC01001 NOT IN "
            MySQLStr = MySQLStr & "(SELECT PC010300_2.PC01001 "
            MySQLStr = MySQLStr & "FROM tbl_PurchaseWorkplace_ConsolidatedOrders INNER JOIN "
            MySQLStr = MySQLStr & "PC010300 AS PC010300_2 ON tbl_PurchaseWorkplace_ConsolidatedOrders.ID = PC010300_2.PC01052 "
            MySQLStr = MySQLStr & "GROUP BY PC010300_2.PC01001)) AND (PC010300_1.PC01002 <> 0) AND "
            MySQLStr = MySQLStr & "(PC010300_1.PC01023 = N'" & Declarations.MyWH & "') AND "
            MySQLStr = MySQLStr & "(PC010300_1.PC01003 = N'" & Declarations.MySupplierCode & "'))) "
            InitMyConn(False)
            Declarations.MyConn.Execute(MySQLStr)
            LoadConsolidatedOrders()
            '---текущей строкой сделать редактированную
            For i As Integer = 0 To DataGridView1.Rows.Count - 1
                If Trim(DataGridView1.Item(0, i).Value.ToString) = Declarations.MyOrderID Then
                    DataGridView1.CurrentCell = DataGridView1.Item(0, i)
                End If
            Next
            LoadIncludedOrders()
            LoadFreeOrders()
            CheckConsolidatedButtons()
            CheckRemoveButtons()
            CheckAddButtons()
        End If
    End Sub

    Private Sub Button14_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button14.Click
        '////////////////////////////////////////////////////////////////////////////////
        '//
        '// Закрытие уже принятого заказа
        '//
        '////////////////////////////////////////////////////////////////////////////////
        Dim MySQLStr As String
        Dim MyRez As MsgBoxResult

        MyRez = MsgBox("Вы действительно хотите закрыть обобщенный заказ?", MsgBoxStyle.YesNo, "Внимание!")
        If MyRez = MsgBoxResult.Yes Then
            MySQLStr = "UPDATE tbl_PurchaseWorkplace_ConsolidatedOrders "
            MySQLStr = MySQLStr & "SET SupplierPlacedDate = CONVERT(DATETIME, '" & Microsoft.VisualBasic.Right("00" & CStr(DatePart(DateInterval.Day, Now())), 2) & "/" & Microsoft.VisualBasic.Right("00" & CStr(DatePart(DateInterval.Month, Now())), 2) & "/" & CStr(DatePart(DateInterval.Year, Now())) & "', 103), "
            MySQLStr = MySQLStr & "ConfirmedDate = CONVERT(DATETIME, '" & Microsoft.VisualBasic.Right("00" & CStr(DatePart(DateInterval.Day, Now())), 2) & "/" & Microsoft.VisualBasic.Right("00" & CStr(DatePart(DateInterval.Month, Now())), 2) & "/" & CStr(DatePart(DateInterval.Year, Now())) & "', 103) "
            MySQLStr = MySQLStr & "WHERE (ID = N'" & Trim(DataGridView1.SelectedRows.Item(0).Cells(0).Value.ToString()) & "') "
            InitMyConn(False)
            Declarations.MyConn.Execute(MySQLStr)
            LoadConsolidatedOrders()
            '---текущей строкой сделать редактированную
            For i As Integer = 0 To DataGridView1.Rows.Count - 1
                If Trim(DataGridView1.Item(0, i).Value.ToString) = Declarations.MyOrderID Then
                    DataGridView1.CurrentCell = DataGridView1.Item(0, i)
                End If
            Next
            LoadIncludedOrders()
            LoadFreeOrders()
            CheckConsolidatedButtons()
            CheckRemoveButtons()
            CheckAddButtons()
        End If
    End Sub

    Private Sub Button4_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button4.Click
        '////////////////////////////////////////////////////////////////////////////////
        '//
        '// Выгрузка шаблона для подтверждения поставки
        '//
        '////////////////////////////////////////////////////////////////////////////////

        If My.Settings.UseOffice = "LibreOffice" Then
            ExportToLOTemplate()
        Else
            ExportToExcelTemplate()
        End If

    End Sub

    Private Sub ExportToExcelTemplate()
        '////////////////////////////////////////////////////////////////////////////////
        '//
        '// Выгрузка в Excel шаблона для подтверждения поставки 
        '//
        '////////////////////////////////////////////////////////////////////////////////
        Dim MyObj As Object
        Dim MyWRKBook As Object
        Dim i As Integer
        Dim MySQLStr As String

        MyObj = CreateObject("Excel.Application")
        MyObj.SheetsInNewWorkbook = 1
        MyWRKBook = MyObj.Workbooks.Add

        '-----------------------Заголовок
        MyWRKBook.ActiveSheet.Range("A1") = "Подтверждение о поставке"
        MyWRKBook.ActiveSheet.Range("D1") = "Поставщик"
        MyWRKBook.ActiveSheet.Range("E1").NumberFormat = "@"
        MyWRKBook.ActiveSheet.Range("E1") = Declarations.MySupplierCode
        MyWRKBook.ActiveSheet.Range("A1").Select()
        MyWRKBook.ActiveSheet.Range("A1").Font.Bold = True
        MyWRKBook.ActiveSheet.Range("D1:E1").Select()
        MyWRKBook.ActiveSheet.Range("D1:E1").Font.Bold = True
        MyWRKBook.ActiveSheet.Range("E1:E1").Select()
        MyWRKBook.ActiveSheet.Range("E1:E1").Borders(5).LineStyle = -4142
        MyWRKBook.ActiveSheet.Range("E1:E1").Borders(6).LineStyle = -4142
        With MyWRKBook.ActiveSheet.Range("E1:E1").Borders(7)
            .LineStyle = 1
            .Weight = 4
            .ColorIndex = -4105
        End With
        With MyWRKBook.ActiveSheet.Range("E1:E1").Borders(8)
            .LineStyle = 1
            .Weight = 4
            .ColorIndex = -4105
        End With
        With MyWRKBook.ActiveSheet.Range("E1:E1").Borders(9)
            .LineStyle = 1
            .Weight = 4
            .ColorIndex = -4105
        End With
        With MyWRKBook.ActiveSheet.Range("E1:E1").Borders(10)
            .LineStyle = 1
            .Weight = 4
            .ColorIndex = -4105
        End With
        With MyWRKBook.ActiveSheet.Range("E1:E1").Borders(11)
            .LineStyle = 1
            .Weight = 4
            .ColorIndex = -4105
        End With
        With MyWRKBook.ActiveSheet.Range("E1:E1").Interior
            .ColorIndex = 36
            .Pattern = 1
            .PatternColorIndex = -4105
        End With

        MyWRKBook.ActiveSheet.Range("B3") = "N заказа на закупку"
        MyWRKBook.ActiveSheet.Range("C3") = "Подтверждение всего заказа"
        MyWRKBook.ActiveSheet.Range("D3") = "Код товара поставщика"
        MyWRKBook.ActiveSheet.Range("E3") = "Подтвержденная дата поставки"
        MyWRKBook.ActiveSheet.Range("F3") = "Подтвержденная вторая дата поставки"
        MyWRKBook.ActiveSheet.Range("B3:F3").Select()
        MyWRKBook.ActiveSheet.Range("B3:F3").Font.Bold = True
        MyWRKBook.ActiveSheet.Range("B3:F3").Select()
        MyWRKBook.ActiveSheet.Range("B3:F3").Borders(5).LineStyle = -4142
        MyWRKBook.ActiveSheet.Range("B3:F3").Borders(6).LineStyle = -4142
        With MyWRKBook.ActiveSheet.Range("B3:F3").Borders(7)
            .LineStyle = 1
            .Weight = 4
            .ColorIndex = -4105
        End With
        With MyWRKBook.ActiveSheet.Range("B3:F3").Borders(8)
            .LineStyle = 1
            .Weight = 4
            .ColorIndex = -4105
        End With
        With MyWRKBook.ActiveSheet.Range("B3:F3").Borders(9)
            .LineStyle = 1
            .Weight = 4
            .ColorIndex = -4105
        End With
        With MyWRKBook.ActiveSheet.Range("B3:F3").Borders(10)
            .LineStyle = 1
            .Weight = 4
            .ColorIndex = -4105
        End With
        With MyWRKBook.ActiveSheet.Range("B3:F3").Borders(11)
            .LineStyle = 1
            .Weight = 4
            .ColorIndex = -4105
        End With
        With MyWRKBook.ActiveSheet.Range("B3:F3").Interior
            .ColorIndex = 36
            .Pattern = 1
            .PatternColorIndex = -4105
        End With

        MyWRKBook.ActiveSheet.Columns("B:B").ColumnWidth = 30
        MyWRKBook.ActiveSheet.Columns("C:C").ColumnWidth = 30
        MyWRKBook.ActiveSheet.Columns("D:D").ColumnWidth = 30
        MyWRKBook.ActiveSheet.Columns("E:E").ColumnWidth = 30
        MyWRKBook.ActiveSheet.Columns("F:F").ColumnWidth = 36

        '---------------------------------таблица
        i = 4
        MySQLStr = "SELECT PC010300.PC01001, SC010300.SC01060 "
        MySQLStr = MySQLStr & "FROM PC030300 WITH(NOLOCK) INNER JOIN "
        MySQLStr = MySQLStr & "PC010300 ON PC030300.PC03001 = PC010300.PC01001 INNER JOIN "
        MySQLStr = MySQLStr & "SC010300 ON PC030300.PC03005 = SC010300.SC01001 "
        MySQLStr = MySQLStr & "WHERE (PC010300.PC01052 = N'" & Trim(DataGridView1.SelectedRows.Item(0).Cells(0).Value.ToString()) & "') "
        MySQLStr = MySQLStr & "GROUP BY PC010300.PC01001, SC010300.SC01060 "
        MySQLStr = MySQLStr & "ORDER BY SC010300.SC01060, PC010300.PC01001 "
        InitMyConn(False)
        InitMyRec(False, MySQLStr)
        If Declarations.MyRec.BOF = True And Declarations.MyRec.EOF = True Then
            trycloseMyRec()
        Else
            Declarations.MyRec.MoveFirst()
            While Declarations.MyRec.EOF = False
                MyWRKBook.ActiveSheet.Range("B" & CStr(i)).NumberFormat = "@"
                MyWRKBook.ActiveSheet.Range("B" & CStr(i)) = Declarations.MyRec.Fields("PC01001").Value
                MyWRKBook.ActiveSheet.Range("D" & CStr(i)).NumberFormat = "@"
                MyWRKBook.ActiveSheet.Range("D" & CStr(i)) = Declarations.MyRec.Fields("SC01060").Value
                Try
                    MyWRKBook.ActiveSheet.Range("E" & CStr(i)).NumberFormat = "ДД/ММ/ГГГГ"
                    With MyWRKBook.ActiveSheet.Range("E" & CStr(i)).Validation
                        .Delete()
                        .Add(Type:=4, AlertStyle:=1, Operator:=1, Formula1:="01/01/1900", Formula2:="31/12/9999")
                        .IgnoreBlank = True
                        .InCellDropdown = True
                        .InputTitle = ""
                        .ErrorTitle = ""
                        .InputMessage = ""
                        .ErrorMessage = ""
                        .ShowInput = True
                        .ShowError = True
                    End With
                    MyWRKBook.ActiveSheet.Range("F" & CStr(i)).NumberFormat = "ДД/ММ/ГГГГ"
                    With MyWRKBook.ActiveSheet.Range("F" & CStr(i)).Validation
                        .Delete()
                        .Add(Type:=4, AlertStyle:=1, Operator:=1, Formula1:="01/01/1900", Formula2:="31/12/9999")
                        .IgnoreBlank = True
                        .InCellDropdown = True
                        .InputTitle = ""
                        .ErrorTitle = ""
                        .InputMessage = ""
                        .ErrorMessage = ""
                        .ShowInput = True
                        .ShowError = True
                    End With
                Catch
                End Try
                Declarations.MyRec.MoveNext()
                i = i + 1
            End While
            trycloseMyRec()
        End If
        MyWRKBook.ActiveSheet.Range("A1").Select()
        MyObj.Application.Visible = True
        MyObj = Nothing

    End Sub

    Private Sub ExportToLOTemplate()
        '////////////////////////////////////////////////////////////////////////////////
        '//
        '// Выгрузка в Libre Office шаблона для подтверждения поставки 
        '//
        '////////////////////////////////////////////////////////////////////////////////
        Dim oServiceManager As Object
        Dim oDispatcher As Object
        Dim oDesktop As Object
        Dim oWorkBook As Object
        Dim oSheet As Object
        Dim Counter As Integer
        Dim oFrame As Object
        Dim MySQLStr As String
        Dim i As Integer

        LOSetNotation(0)
        oServiceManager = CreateObject("com.sun.star.ServiceManager")
        oDesktop = oServiceManager.createInstance("com.sun.star.frame.Desktop")
        oDispatcher = oServiceManager.createInstance("com.sun.star.frame.DispatchHelper")
        Dim arg(1)
        arg(0) = mAkePropertyValue("Hidden", True, oServiceManager)
        arg(1) = mAkePropertyValue("MacroExecutionMode", 4, oServiceManager)
        oWorkBook = oDesktop.loadComponentFromURL("private:factory/scalc", "_blank", 0, arg)
        oSheet = oWorkBook.getSheets().getByIndex(0)
        oFrame = oWorkBook.getCurrentController.getFrame

        '-----Ширина колонок
        oSheet.getColumns().getByName("B").Width = 6000
        oSheet.getColumns().getByName("C").Width = 6000
        oSheet.getColumns().getByName("D").Width = 6000
        oSheet.getColumns().getByName("E").Width = 6000
        oSheet.getColumns().getByName("F").Width = 7000

        '-----Заголовок
        oSheet.getCellRangeByName("A1").String = "Подтверждение о поставке"
        oSheet.getCellRangeByName("D1").String = "Поставщик"
        oSheet.getCellRangeByName("E1").String = Declarations.MySupplierCode
        LOFontSetFamilyName(oServiceManager, oDispatcher, oFrame, "A1:E1", "Calibri")
        LOFontSetBold(oServiceManager, oDispatcher, oFrame, "A1:E1")
        LOFontSetSize(oServiceManager, oDispatcher, oFrame, "A1:E1", 11)
        Dim LineFormat As Object
        LineFormat = oServiceManager.Bridge_GetStruct("com.sun.star.table.BorderLine2")
        LineFormat.LineStyle = 0
        LineFormat.LineWidth = 70
        oSheet.getCellRangeByName("E1").TopBorder = LineFormat
        oSheet.getCellRangeByName("E1").RightBorder = LineFormat
        oSheet.getCellRangeByName("E1").LeftBorder = LineFormat
        oSheet.getCellRangeByName("E1").BottomBorder = LineFormat
        oSheet.getCellRangeByName("E1").CellBackColor = 16775598

        '-----заголовок таблицы
        oSheet.getCellRangeByName("B3").String = "N заказа на закупку"
        oSheet.getCellRangeByName("C3").String = "Подтверждение всего заказа"
        oSheet.getCellRangeByName("D3").String = "Код товара поставщика"
        oSheet.getCellRangeByName("E3").String = "Подтвержденная дата поставки"
        oSheet.getCellRangeByName("F3").String = "Подтвержденная вторая дата поставки"
        LOFontSetFamilyName(oServiceManager, oDispatcher, oFrame, "B3:F3", "Calibri")
        LOFontSetBold(oServiceManager, oDispatcher, oFrame, "B3:F3")
        LOFontSetSize(oServiceManager, oDispatcher, oFrame, "B3:F3", 11)
        LineFormat.LineStyle = 0
        LineFormat.LineWidth = 70
        oSheet.getCellRangeByName("B3:F3").TopBorder = LineFormat
        oSheet.getCellRangeByName("B3:F3").RightBorder = LineFormat
        oSheet.getCellRangeByName("B3:F3").LeftBorder = LineFormat
        oSheet.getCellRangeByName("B3:F3").BottomBorder = LineFormat
        oSheet.getCellRangeByName("B3:F3").CellBackColor = 16775598
        oSheet.getCellRangeByName("B3:F3").VertJustify = 2
        oSheet.getCellRangeByName("B3:F3").HoriJustify = 2
        LOWrapText(oServiceManager, oDispatcher, oFrame, "B3:F3")

        '---------------------------------таблица
        i = 4
        'MySQLStr = "SELECT PC010300.PC01001, SC010300.SC01060 "
        'MySQLStr = MySQLStr & "FROM PC030300 WITH(NOLOCK) INNER JOIN "
        'MySQLStr = MySQLStr & "PC010300 ON PC030300.PC03001 = PC010300.PC01001 INNER JOIN "
        'MySQLStr = MySQLStr & "SC010300 ON PC030300.PC03005 = SC010300.SC01001 "
        'MySQLStr = MySQLStr & "WHERE (PC010300.PC01052 = N'" & Trim(DataGridView1.SelectedRows.Item(0).Cells(0).Value.ToString()) & "') "
        'MySQLStr = MySQLStr & "GROUP BY PC010300.PC01001, SC010300.SC01060 "
        'MySQLStr = MySQLStr & "ORDER BY SC010300.SC01060, PC010300.PC01001 "
        MySQLStr = "SELECT PC010300.PC01001, SC010300.SC01060, PC030300.PC03024, PC030300.PC03031 "
        MySQLStr = MySQLStr & "FROM  PC030300 WITH (NOLOCK) INNER JOIN "
        MySQLStr = MySQLStr & "PC010300 ON PC030300.PC03001 = PC010300.PC01001 INNER JOIN "
        MySQLStr = MySQLStr & "SC010300 ON PC030300.PC03005 = SC010300.SC01001 "
        MySQLStr = MySQLStr & "WHERE (PC010300.PC01052 = N'" & Trim(DataGridView1.SelectedRows.Item(0).Cells(0).Value.ToString()) & "') "
        MySQLStr = MySQLStr & "AND (PC030300.PC03010 - PC030300.PC03011 > 0) "
        MySQLStr = MySQLStr & "GROUP BY PC010300.PC01001, SC010300.SC01060, PC030300.PC03024, PC030300.PC03031 "
        MySQLStr = MySQLStr & "ORDER BY SC010300.SC01060, PC010300.PC01001 "
        InitMyConn(False)
        InitMyRec(False, MySQLStr)

        If Declarations.MyRec.BOF = True And Declarations.MyRec.EOF = True Then
            trycloseMyRec()
        Else
            Declarations.MyRec.MoveFirst()
            While Declarations.MyRec.EOF = False
                '-----N заказа на закупку
                oSheet.getCellRangeByName("B" & CStr(i)).String = Declarations.MyRec.Fields("PC01001").Value
                '-----Код товара поставщика
                oSheet.getCellRangeByName("D" & CStr(i)).String = Declarations.MyRec.Fields("SC01060").Value
                '-----Дата первой поставки
                oSheet.getCellRangeByName("E" & CStr(i)).Value = Declarations.MyRec.Fields("PC03024").Value
                '-----Дата последней поставки
                oSheet.getCellRangeByName("F" & CStr(i)).Value = Declarations.MyRec.Fields("PC03031").Value
                '-----форматы
                LOFormatCells(oServiceManager, oDispatcher, oFrame, "E" & i & ":F" & i, 36)


                Declarations.MyRec.MoveNext()
                i = i + 1
            End While
            trycloseMyRec()
        End If
        Dim args() As Object
        ReDim args(0)
        args(0) = oServiceManager.Bridge_GetStruct("com.sun.star.beans.PropertyValue")
        args(0).Name = "ToPoint"
        args(0).Value = "$A$1"
        oDispatcher.executeDispatch(oFrame, ".uno:GoToCell", "", 0, args)

        oWorkBook.CurrentController.Frame.ContainerWindow.Visible = True
        oWorkBook.CurrentController.Frame.ContainerWindow.toFront()
    End Sub

    Private Sub Button5_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button5.Click
        '////////////////////////////////////////////////////////////////////////////////
        '//
        '// Импорт подтверждения поставки из стандартного шаблона
        '//
        '////////////////////////////////////////////////////////////////////////////////

        Declarations.MyOrderID = Trim(DataGridView1.SelectedRows.Item(0).Cells(0).Value.ToString())
        MyImportCommonConfirmation = New ImportCommonConfirmation
        MyImportCommonConfirmation.ShowDialog()
        LoadConsolidatedOrders()
        '---текущей строкой сделать редактированную
        For i As Integer = 0 To DataGridView1.Rows.Count - 1
            If Trim(DataGridView1.Item(0, i).Value.ToString) = Declarations.MyOrderID Then
                DataGridView1.CurrentCell = DataGridView1.Item(0, i)
            End If
        Next
        LoadIncludedOrders()
        LoadFreeOrders()
        CheckConsolidatedButtons()
        CheckRemoveButtons()
        CheckAddButtons()
        '---Выгрузка полный Excel
        If My.Settings.UseOffice = "LibreOffice" Then
            ExportOrderToLOFull(Trim(DataGridView1.SelectedRows.Item(0).Cells(0).Value.ToString()))
        Else
            ExportOrderToExcelFull(Trim(DataGridView1.SelectedRows.Item(0).Cells(0).Value.ToString()))
        End If
    End Sub

    Private Sub Button2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button2.Click
        '////////////////////////////////////////////////////////////////////////////////
        '//
        '// Выгрузка обобщенного заказа в Excel
        '//
        '////////////////////////////////////////////////////////////////////////////////

        If CheckUOMInOrders(0) = True Then
            If My.Settings.UseOffice = "LibreOffice" Then
                ExportOrderToLO(Trim(DataGridView1.SelectedRows.Item(0).Cells(0).Value.ToString()))
            Else
                ExportOrderToExcel(Trim(DataGridView1.SelectedRows.Item(0).Cells(0).Value.ToString()))
            End If
        End If
    End Sub

    Private Function CheckUOMInOrders(ByVal MyType As Integer) As Boolean
        '////////////////////////////////////////////////////////////////////////////////
        '//
        '// Проверка - есть ли в заказах запасы с единицей измерения, отличающейся от проставленных в карточке
        '// 0 - только в обобщенном заказе
        '// 1 - в обобщенном заказе и одном добавляемом
        '// 2 - в обобщенном заказе и всех добавляемых
        '// разработчик Новожилов А.Н. 2012
        '//
        '////////////////////////////////////////////////////////////////////////////////
        Dim MySQLStr As String
        Dim MyERRStr As String                      'сообщения об ошибках 

        If MyType = 0 Then               '---только в обобщенном заказе
            MySQLStr = "SELECT PC030300.PC03001 AS OrderN, PC030300.PC03005 AS ItemN, View_1_1.txt AS OrderUOM, View_1.txt AS CardUOM "
            MySQLStr = MySQLStr & "FROM PC030300 WITH(NOLOCK) INNER JOIN "
            MySQLStr = MySQLStr & "PC010300 ON PC030300.PC03001 = PC010300.PC01001 INNER JOIN "
            MySQLStr = MySQLStr & "SC010300 ON PC030300.PC03005 = SC010300.SC01001 AND PC030300.PC03009 <> SC010300.SC01134 INNER JOIN "
            MySQLStr = MySQLStr & "(SELECT 0 AS num, SC09002 AS txt "
            MySQLStr = MySQLStr & "FROM SC090300 WITH(NOLOCK) "
            MySQLStr = MySQLStr & "WHERE   (SC09001 = 'RUS') "
            MySQLStr = MySQLStr & "UNION "
            MySQLStr = MySQLStr & "SELECT     1 AS Expr1, SC09003 "
            MySQLStr = MySQLStr & "FROM         SC090300 AS SC090300_40 WITH(NOLOCK) "
            MySQLStr = MySQLStr & "WHERE     (SC09001 = 'RUS') "
            MySQLStr = MySQLStr & "UNION "
            MySQLStr = MySQLStr & "SELECT     2 AS Expr1, SC09004 "
            MySQLStr = MySQLStr & "FROM         SC090300 AS SC090300_39 WITH(NOLOCK) "
            MySQLStr = MySQLStr & "WHERE     (SC09001 = 'RUS') "
            MySQLStr = MySQLStr & "UNION "
            MySQLStr = MySQLStr & "SELECT     3 AS Expr1, SC09005 "
            MySQLStr = MySQLStr & "FROM         SC090300 AS SC090300_38 WITH(NOLOCK) "
            MySQLStr = MySQLStr & "WHERE     (SC09001 = 'RUS') "
            MySQLStr = MySQLStr & "UNION "
            MySQLStr = MySQLStr & "SELECT     4 AS Expr1, SC09006 "
            MySQLStr = MySQLStr & "FROM         SC090300 AS SC090300_37 WITH(NOLOCK) "
            MySQLStr = MySQLStr & "WHERE     (SC09001 = 'RUS')"
            MySQLStr = MySQLStr & "UNION "
            MySQLStr = MySQLStr & "SELECT     5 AS Expr1, SC09007 "
            MySQLStr = MySQLStr & "FROM         SC090300 AS SC090300_36 WITH(NOLOCK) "
            MySQLStr = MySQLStr & "WHERE     (SC09001 = 'RUS') "
            MySQLStr = MySQLStr & "UNION "
            MySQLStr = MySQLStr & "SELECT     6 AS Expr1, SC09008 "
            MySQLStr = MySQLStr & "FROM         SC090300 AS SC090300_35 WITH(NOLOCK) "
            MySQLStr = MySQLStr & "WHERE     (SC09001 = 'RUS') "
            MySQLStr = MySQLStr & "UNION "
            MySQLStr = MySQLStr & "SELECT     7 AS Expr1, SC09009 "
            MySQLStr = MySQLStr & "FROM         SC090300 AS SC090300_34 WITH(NOLOCK) "
            MySQLStr = MySQLStr & "WHERE     (SC09001 = 'RUS') "
            MySQLStr = MySQLStr & "UNION "
            MySQLStr = MySQLStr & "SELECT     8 AS Expr1, SC09010 "
            MySQLStr = MySQLStr & "FROM         SC090300 AS SC090300_33 WITH(NOLOCK) "
            MySQLStr = MySQLStr & "WHERE     (SC09001 = 'RUS') "
            MySQLStr = MySQLStr & "UNION "
            MySQLStr = MySQLStr & "SELECT     9 AS Expr1, SC09011 "
            MySQLStr = MySQLStr & "FROM         SC090300 AS SC090300_32 WITH(NOLOCK) "
            MySQLStr = MySQLStr & "WHERE     (SC09001 = 'RUS') "
            MySQLStr = MySQLStr & "UNION "
            MySQLStr = MySQLStr & "SELECT     10 AS Expr1, SC09012 "
            MySQLStr = MySQLStr & "FROM         SC090300 AS SC090300_31 WITH(NOLOCK) "
            MySQLStr = MySQLStr & "WHERE     (SC09001 = 'RUS') "
            MySQLStr = MySQLStr & "UNION "
            MySQLStr = MySQLStr & "SELECT     11 AS Expr1, SC09013 "
            MySQLStr = MySQLStr & "FROM         SC090300 AS SC090300_30 WITH(NOLOCK) "
            MySQLStr = MySQLStr & "WHERE     (SC09001 = 'RUS') "
            MySQLStr = MySQLStr & "UNION "
            MySQLStr = MySQLStr & "SELECT     12 AS Expr1, SC09014 "
            MySQLStr = MySQLStr & "FROM         SC090300 AS SC090300_29 WITH(NOLOCK) "
            MySQLStr = MySQLStr & "WHERE     (SC09001 = 'RUS') "
            MySQLStr = MySQLStr & "UNION "
            MySQLStr = MySQLStr & "SELECT     13 AS Expr1, SC09015 "
            MySQLStr = MySQLStr & "FROM         SC090300 AS SC090300_28 WITH(NOLOCK) "
            MySQLStr = MySQLStr & "WHERE     (SC09001 = 'RUS') "
            MySQLStr = MySQLStr & "UNION "
            MySQLStr = MySQLStr & "SELECT     14 AS Expr1, SC09016 "
            MySQLStr = MySQLStr & "FROM         SC090300 AS SC090300_27 WITH(NOLOCK) "
            MySQLStr = MySQLStr & "WHERE     (SC09001 = 'RUS') "
            MySQLStr = MySQLStr & "UNION "
            MySQLStr = MySQLStr & "SELECT     15 AS Expr1, SC09017 "
            MySQLStr = MySQLStr & "FROM         SC090300 AS SC090300_26 WITH(NOLOCK) "
            MySQLStr = MySQLStr & "WHERE     (SC09001 = 'RUS') "
            MySQLStr = MySQLStr & "UNION "
            MySQLStr = MySQLStr & "SELECT     16 AS Expr1, SC09018 "
            MySQLStr = MySQLStr & "FROM         SC090300 AS SC090300_25 WITH(NOLOCK) "
            MySQLStr = MySQLStr & "WHERE     (SC09001 = 'RUS') "
            MySQLStr = MySQLStr & "UNION "
            MySQLStr = MySQLStr & "SELECT     17 AS Expr1, SC09019 "
            MySQLStr = MySQLStr & "FROM         SC090300 AS SC090300_24 WITH(NOLOCK) "
            MySQLStr = MySQLStr & "WHERE     (SC09001 = 'RUS') "
            MySQLStr = MySQLStr & "UNION "
            MySQLStr = MySQLStr & "SELECT     18 AS Expr1, SC09020 "
            MySQLStr = MySQLStr & "FROM         SC090300 AS SC090300_23 WITH(NOLOCK) "
            MySQLStr = MySQLStr & "WHERE     (SC09001 = 'RUS') "
            MySQLStr = MySQLStr & "UNION "
            MySQLStr = MySQLStr & "SELECT     19 AS Expr1, SC09021 "
            MySQLStr = MySQLStr & "FROM         SC090300 AS SC090300_22 WITH(NOLOCK) "
            MySQLStr = MySQLStr & "WHERE     (SC09001 = 'RUS') "
            MySQLStr = MySQLStr & "UNION "
            MySQLStr = MySQLStr & "SELECT     20 AS Expr1, SC09022 "
            MySQLStr = MySQLStr & "FROM         SC090300 AS SC090300_21 WITH(NOLOCK) "
            MySQLStr = MySQLStr & "WHERE     (SC09001 = 'RUS') "
            MySQLStr = MySQLStr & "UNION "
            MySQLStr = MySQLStr & "SELECT     21 AS Expr1, SC09023 "
            MySQLStr = MySQLStr & "FROM         SC090300 AS SC090300_20 WITH(NOLOCK) "
            MySQLStr = MySQLStr & "WHERE     (SC09001 = 'RUS') "
            MySQLStr = MySQLStr & "UNION "
            MySQLStr = MySQLStr & "SELECT     22 AS Expr1, SC09024 "
            MySQLStr = MySQLStr & "FROM         SC090300 AS SC090300_19 WITH(NOLOCK) "
            MySQLStr = MySQLStr & "WHERE     (SC09001 = 'RUS') "
            MySQLStr = MySQLStr & "UNION "
            MySQLStr = MySQLStr & "SELECT     23 AS Expr1, SC09025 "
            MySQLStr = MySQLStr & "FROM         SC090300 AS SC090300_18 WITH(NOLOCK) "
            MySQLStr = MySQLStr & "WHERE     (SC09001 = 'RUS') "
            MySQLStr = MySQLStr & "UNION "
            MySQLStr = MySQLStr & "SELECT     24 AS Expr1, SC09026 "
            MySQLStr = MySQLStr & "FROM         SC090300 AS SC090300_17 WITH(NOLOCK) "
            MySQLStr = MySQLStr & "WHERE     (SC09001 = 'RUS') "
            MySQLStr = MySQLStr & "UNION "
            MySQLStr = MySQLStr & "SELECT     25 AS Expr1, SC09027 "
            MySQLStr = MySQLStr & "FROM         SC090300 AS SC090300_16 WITH(NOLOCK) "
            MySQLStr = MySQLStr & "WHERE     (SC09001 = 'RUS') "
            MySQLStr = MySQLStr & "UNION "
            MySQLStr = MySQLStr & "SELECT     26 AS Expr1, SC09028 "
            MySQLStr = MySQLStr & "FROM         SC090300 AS SC090300_15 WITH(NOLOCK) "
            MySQLStr = MySQLStr & "WHERE     (SC09001 = 'RUS') "
            MySQLStr = MySQLStr & "UNION "
            MySQLStr = MySQLStr & "SELECT     27 AS Expr1, SC09029 "
            MySQLStr = MySQLStr & "FROM         SC090300 AS SC090300_14 WITH(NOLOCK) "
            MySQLStr = MySQLStr & "WHERE     (SC09001 = 'RUS') "
            MySQLStr = MySQLStr & "UNION "
            MySQLStr = MySQLStr & "SELECT     28 AS Expr1, SC09030 "
            MySQLStr = MySQLStr & "FROM         SC090300 AS SC090300_13 WITH(NOLOCK) "
            MySQLStr = MySQLStr & "WHERE     (SC09001 = 'RUS') "
            MySQLStr = MySQLStr & "UNION "
            MySQLStr = MySQLStr & "SELECT     29 AS Expr1, SC09031 "
            MySQLStr = MySQLStr & "FROM         SC090300 AS SC090300_12 WITH(NOLOCK) "
            MySQLStr = MySQLStr & "WHERE     (SC09001 = 'RUS') "
            MySQLStr = MySQLStr & "UNION "
            MySQLStr = MySQLStr & "SELECT     30 AS Expr1, SC09032 "
            MySQLStr = MySQLStr & "FROM         SC090300 AS SC090300_11 WITH(NOLOCK) "
            MySQLStr = MySQLStr & "WHERE     (SC09001 = 'RUS') "
            MySQLStr = MySQLStr & "UNION "
            MySQLStr = MySQLStr & "SELECT     31 AS Expr1, SC09033 "
            MySQLStr = MySQLStr & "FROM         SC090300 AS SC090300_10 WITH(NOLOCK) "
            MySQLStr = MySQLStr & "WHERE     (SC09001 = 'RUS') "
            MySQLStr = MySQLStr & "UNION "
            MySQLStr = MySQLStr & "SELECT     32 AS Expr1, SC09034 "
            MySQLStr = MySQLStr & "FROM         SC090300 AS SC090300_9 WITH(NOLOCK) "
            MySQLStr = MySQLStr & "WHERE     (SC09001 = 'RUS') "
            MySQLStr = MySQLStr & "UNION "
            MySQLStr = MySQLStr & "SELECT     33 AS Expr1, SC09035 "
            MySQLStr = MySQLStr & "FROM         SC090300 AS SC090300_8 WITH(NOLOCK) "
            MySQLStr = MySQLStr & "WHERE     (SC09001 = 'RUS') "
            MySQLStr = MySQLStr & "UNION "
            MySQLStr = MySQLStr & "SELECT     34 AS Expr1, SC09036 "
            MySQLStr = MySQLStr & "FROM         SC090300 AS SC090300_7 WITH(NOLOCK) "
            MySQLStr = MySQLStr & "WHERE     (SC09001 = 'RUS') "
            MySQLStr = MySQLStr & "UNION "
            MySQLStr = MySQLStr & "SELECT     35 AS Expr1, SC09037 "
            MySQLStr = MySQLStr & "FROM         SC090300 AS SC090300_6 WITH(NOLOCK) "
            MySQLStr = MySQLStr & "WHERE     (SC09001 = 'RUS') "
            MySQLStr = MySQLStr & "UNION "
            MySQLStr = MySQLStr & "SELECT     36 AS Expr1, SC09038 "
            MySQLStr = MySQLStr & "FROM         SC090300 AS SC090300_5 WITH(NOLOCK) "
            MySQLStr = MySQLStr & "WHERE     (SC09001 = 'RUS') "
            MySQLStr = MySQLStr & "UNION "
            MySQLStr = MySQLStr & "SELECT     37 AS Expr1, SC09039 "
            MySQLStr = MySQLStr & "FROM         SC090300 AS SC090300_4 WITH(NOLOCK) "
            MySQLStr = MySQLStr & "WHERE     (SC09001 = 'RUS') "
            MySQLStr = MySQLStr & "UNION "
            MySQLStr = MySQLStr & "SELECT     38 AS Expr1, SC09040 "
            MySQLStr = MySQLStr & "FROM         SC090300 AS SC090300_3 WITH(NOLOCK) "
            MySQLStr = MySQLStr & "WHERE     (SC09001 = 'RUS') "
            MySQLStr = MySQLStr & "UNION "
            MySQLStr = MySQLStr & "SELECT     39 AS Expr1, SC09041 "
            MySQLStr = MySQLStr & "FROM         SC090300 AS SC090300_2 WITH(NOLOCK) "
            MySQLStr = MySQLStr & "WHERE     (SC09001 = 'RUS') "
            MySQLStr = MySQLStr & "UNION "
            MySQLStr = MySQLStr & "SELECT     40 AS Expr1, SC09042 "
            MySQLStr = MySQLStr & "FROM         SC090300 AS SC090300_1 WITH(NOLOCK) "
            MySQLStr = MySQLStr & "WHERE     (SC09001 = 'RUS')) AS View_1 ON SC010300.SC01134 = View_1.num INNER JOIN "
            MySQLStr = MySQLStr & "(SELECT 0 AS num, SC09002 AS txt "
            MySQLStr = MySQLStr & "FROM SC090300 WITH(NOLOCK) "
            MySQLStr = MySQLStr & "WHERE   (SC09001 = 'RUS') "
            MySQLStr = MySQLStr & "UNION "
            MySQLStr = MySQLStr & "SELECT     1 AS Expr1, SC09003 "
            MySQLStr = MySQLStr & "FROM         SC090300 AS SC090300_40 WITH(NOLOCK) "
            MySQLStr = MySQLStr & "WHERE     (SC09001 = 'RUS') "
            MySQLStr = MySQLStr & "UNION "
            MySQLStr = MySQLStr & "SELECT     2 AS Expr1, SC09004 "
            MySQLStr = MySQLStr & "FROM         SC090300 AS SC090300_39 WITH(NOLOCK) "
            MySQLStr = MySQLStr & "WHERE     (SC09001 = 'RUS') "
            MySQLStr = MySQLStr & "UNION "
            MySQLStr = MySQLStr & "SELECT     3 AS Expr1, SC09005 "
            MySQLStr = MySQLStr & "FROM         SC090300 AS SC090300_38 WITH(NOLOCK) "
            MySQLStr = MySQLStr & "WHERE     (SC09001 = 'RUS') "
            MySQLStr = MySQLStr & "UNION "
            MySQLStr = MySQLStr & "SELECT     4 AS Expr1, SC09006 "
            MySQLStr = MySQLStr & "FROM         SC090300 AS SC090300_37 WITH(NOLOCK) "
            MySQLStr = MySQLStr & "WHERE     (SC09001 = 'RUS')"
            MySQLStr = MySQLStr & "UNION "
            MySQLStr = MySQLStr & "SELECT     5 AS Expr1, SC09007 "
            MySQLStr = MySQLStr & "FROM         SC090300 AS SC090300_36 WITH(NOLOCK) "
            MySQLStr = MySQLStr & "WHERE     (SC09001 = 'RUS') "
            MySQLStr = MySQLStr & "UNION "
            MySQLStr = MySQLStr & "SELECT     6 AS Expr1, SC09008 "
            MySQLStr = MySQLStr & "FROM         SC090300 AS SC090300_35 WITH(NOLOCK) "
            MySQLStr = MySQLStr & "WHERE     (SC09001 = 'RUS') "
            MySQLStr = MySQLStr & "UNION "
            MySQLStr = MySQLStr & "SELECT     7 AS Expr1, SC09009 "
            MySQLStr = MySQLStr & "FROM         SC090300 AS SC090300_34 WITH(NOLOCK) "
            MySQLStr = MySQLStr & "WHERE     (SC09001 = 'RUS') "
            MySQLStr = MySQLStr & "UNION "
            MySQLStr = MySQLStr & "SELECT     8 AS Expr1, SC09010 "
            MySQLStr = MySQLStr & "FROM         SC090300 AS SC090300_33 WITH(NOLOCK) "
            MySQLStr = MySQLStr & "WHERE     (SC09001 = 'RUS') "
            MySQLStr = MySQLStr & "UNION "
            MySQLStr = MySQLStr & "SELECT     9 AS Expr1, SC09011 "
            MySQLStr = MySQLStr & "FROM         SC090300 AS SC090300_32 WITH(NOLOCK) "
            MySQLStr = MySQLStr & "WHERE     (SC09001 = 'RUS') "
            MySQLStr = MySQLStr & "UNION "
            MySQLStr = MySQLStr & "SELECT     10 AS Expr1, SC09012 "
            MySQLStr = MySQLStr & "FROM         SC090300 AS SC090300_31 WITH(NOLOCK) "
            MySQLStr = MySQLStr & "WHERE     (SC09001 = 'RUS') "
            MySQLStr = MySQLStr & "UNION "
            MySQLStr = MySQLStr & "SELECT     11 AS Expr1, SC09013 "
            MySQLStr = MySQLStr & "FROM         SC090300 AS SC090300_30 WITH(NOLOCK) "
            MySQLStr = MySQLStr & "WHERE     (SC09001 = 'RUS') "
            MySQLStr = MySQLStr & "UNION "
            MySQLStr = MySQLStr & "SELECT     12 AS Expr1, SC09014 "
            MySQLStr = MySQLStr & "FROM         SC090300 AS SC090300_29 WITH(NOLOCK) "
            MySQLStr = MySQLStr & "WHERE     (SC09001 = 'RUS') "
            MySQLStr = MySQLStr & "UNION "
            MySQLStr = MySQLStr & "SELECT     13 AS Expr1, SC09015 "
            MySQLStr = MySQLStr & "FROM         SC090300 AS SC090300_28 WITH(NOLOCK) "
            MySQLStr = MySQLStr & "WHERE     (SC09001 = 'RUS') "
            MySQLStr = MySQLStr & "UNION "
            MySQLStr = MySQLStr & "SELECT     14 AS Expr1, SC09016 "
            MySQLStr = MySQLStr & "FROM         SC090300 AS SC090300_27 WITH(NOLOCK) "
            MySQLStr = MySQLStr & "WHERE     (SC09001 = 'RUS') "
            MySQLStr = MySQLStr & "UNION "
            MySQLStr = MySQLStr & "SELECT     15 AS Expr1, SC09017 "
            MySQLStr = MySQLStr & "FROM         SC090300 AS SC090300_26 WITH(NOLOCK) "
            MySQLStr = MySQLStr & "WHERE     (SC09001 = 'RUS') "
            MySQLStr = MySQLStr & "UNION "
            MySQLStr = MySQLStr & "SELECT     16 AS Expr1, SC09018 "
            MySQLStr = MySQLStr & "FROM         SC090300 AS SC090300_25 WITH(NOLOCK) "
            MySQLStr = MySQLStr & "WHERE     (SC09001 = 'RUS') "
            MySQLStr = MySQLStr & "UNION "
            MySQLStr = MySQLStr & "SELECT     17 AS Expr1, SC09019 "
            MySQLStr = MySQLStr & "FROM         SC090300 AS SC090300_24 WITH(NOLOCK) "
            MySQLStr = MySQLStr & "WHERE     (SC09001 = 'RUS') "
            MySQLStr = MySQLStr & "UNION "
            MySQLStr = MySQLStr & "SELECT     18 AS Expr1, SC09020 "
            MySQLStr = MySQLStr & "FROM         SC090300 AS SC090300_23 WITH(NOLOCK) "
            MySQLStr = MySQLStr & "WHERE     (SC09001 = 'RUS') "
            MySQLStr = MySQLStr & "UNION "
            MySQLStr = MySQLStr & "SELECT     19 AS Expr1, SC09021 "
            MySQLStr = MySQLStr & "FROM         SC090300 AS SC090300_22 WITH(NOLOCK) "
            MySQLStr = MySQLStr & "WHERE     (SC09001 = 'RUS') "
            MySQLStr = MySQLStr & "UNION "
            MySQLStr = MySQLStr & "SELECT     20 AS Expr1, SC09022 "
            MySQLStr = MySQLStr & "FROM         SC090300 AS SC090300_21 WITH(NOLOCK) "
            MySQLStr = MySQLStr & "WHERE     (SC09001 = 'RUS') "
            MySQLStr = MySQLStr & "UNION "
            MySQLStr = MySQLStr & "SELECT     21 AS Expr1, SC09023 "
            MySQLStr = MySQLStr & "FROM         SC090300 AS SC090300_20 WITH(NOLOCK) "
            MySQLStr = MySQLStr & "WHERE     (SC09001 = 'RUS') "
            MySQLStr = MySQLStr & "UNION "
            MySQLStr = MySQLStr & "SELECT     22 AS Expr1, SC09024 "
            MySQLStr = MySQLStr & "FROM         SC090300 AS SC090300_19 WITH(NOLOCK) "
            MySQLStr = MySQLStr & "WHERE     (SC09001 = 'RUS') "
            MySQLStr = MySQLStr & "UNION "
            MySQLStr = MySQLStr & "SELECT     23 AS Expr1, SC09025 "
            MySQLStr = MySQLStr & "FROM         SC090300 AS SC090300_18 WITH(NOLOCK) "
            MySQLStr = MySQLStr & "WHERE     (SC09001 = 'RUS') "
            MySQLStr = MySQLStr & "UNION "
            MySQLStr = MySQLStr & "SELECT     24 AS Expr1, SC09026 "
            MySQLStr = MySQLStr & "FROM         SC090300 AS SC090300_17 WITH(NOLOCK) "
            MySQLStr = MySQLStr & "WHERE     (SC09001 = 'RUS') "
            MySQLStr = MySQLStr & "UNION "
            MySQLStr = MySQLStr & "SELECT     25 AS Expr1, SC09027 "
            MySQLStr = MySQLStr & "FROM         SC090300 AS SC090300_16 WITH(NOLOCK) "
            MySQLStr = MySQLStr & "WHERE     (SC09001 = 'RUS') "
            MySQLStr = MySQLStr & "UNION "
            MySQLStr = MySQLStr & "SELECT     26 AS Expr1, SC09028 "
            MySQLStr = MySQLStr & "FROM         SC090300 AS SC090300_15 WITH(NOLOCK) "
            MySQLStr = MySQLStr & "WHERE     (SC09001 = 'RUS') "
            MySQLStr = MySQLStr & "UNION "
            MySQLStr = MySQLStr & "SELECT     27 AS Expr1, SC09029 "
            MySQLStr = MySQLStr & "FROM         SC090300 AS SC090300_14 WITH(NOLOCK) "
            MySQLStr = MySQLStr & "WHERE     (SC09001 = 'RUS') "
            MySQLStr = MySQLStr & "UNION "
            MySQLStr = MySQLStr & "SELECT     28 AS Expr1, SC09030 "
            MySQLStr = MySQLStr & "FROM         SC090300 AS SC090300_13 WITH(NOLOCK) "
            MySQLStr = MySQLStr & "WHERE     (SC09001 = 'RUS') "
            MySQLStr = MySQLStr & "UNION "
            MySQLStr = MySQLStr & "SELECT     29 AS Expr1, SC09031 "
            MySQLStr = MySQLStr & "FROM         SC090300 AS SC090300_12 WITH(NOLOCK) "
            MySQLStr = MySQLStr & "WHERE     (SC09001 = 'RUS') "
            MySQLStr = MySQLStr & "UNION "
            MySQLStr = MySQLStr & "SELECT     30 AS Expr1, SC09032 "
            MySQLStr = MySQLStr & "FROM         SC090300 AS SC090300_11 WITH(NOLOCK) "
            MySQLStr = MySQLStr & "WHERE     (SC09001 = 'RUS') "
            MySQLStr = MySQLStr & "UNION "
            MySQLStr = MySQLStr & "SELECT     31 AS Expr1, SC09033 "
            MySQLStr = MySQLStr & "FROM         SC090300 AS SC090300_10 WITH(NOLOCK) "
            MySQLStr = MySQLStr & "WHERE     (SC09001 = 'RUS') "
            MySQLStr = MySQLStr & "UNION "
            MySQLStr = MySQLStr & "SELECT     32 AS Expr1, SC09034 "
            MySQLStr = MySQLStr & "FROM         SC090300 AS SC090300_9 WITH(NOLOCK) "
            MySQLStr = MySQLStr & "WHERE     (SC09001 = 'RUS') "
            MySQLStr = MySQLStr & "UNION "
            MySQLStr = MySQLStr & "SELECT     33 AS Expr1, SC09035 "
            MySQLStr = MySQLStr & "FROM         SC090300 AS SC090300_8 WITH(NOLOCK) "
            MySQLStr = MySQLStr & "WHERE     (SC09001 = 'RUS') "
            MySQLStr = MySQLStr & "UNION "
            MySQLStr = MySQLStr & "SELECT     34 AS Expr1, SC09036 "
            MySQLStr = MySQLStr & "FROM         SC090300 AS SC090300_7 WITH(NOLOCK) "
            MySQLStr = MySQLStr & "WHERE     (SC09001 = 'RUS') "
            MySQLStr = MySQLStr & "UNION "
            MySQLStr = MySQLStr & "SELECT     35 AS Expr1, SC09037 "
            MySQLStr = MySQLStr & "FROM         SC090300 AS SC090300_6 WITH(NOLOCK) "
            MySQLStr = MySQLStr & "WHERE     (SC09001 = 'RUS') "
            MySQLStr = MySQLStr & "UNION "
            MySQLStr = MySQLStr & "SELECT     36 AS Expr1, SC09038 "
            MySQLStr = MySQLStr & "FROM         SC090300 AS SC090300_5 WITH(NOLOCK) "
            MySQLStr = MySQLStr & "WHERE     (SC09001 = 'RUS') "
            MySQLStr = MySQLStr & "UNION "
            MySQLStr = MySQLStr & "SELECT     37 AS Expr1, SC09039 "
            MySQLStr = MySQLStr & "FROM         SC090300 AS SC090300_4 WITH(NOLOCK) "
            MySQLStr = MySQLStr & "WHERE     (SC09001 = 'RUS') "
            MySQLStr = MySQLStr & "UNION "
            MySQLStr = MySQLStr & "SELECT     38 AS Expr1, SC09040 "
            MySQLStr = MySQLStr & "FROM         SC090300 AS SC090300_3 WITH(NOLOCK) "
            MySQLStr = MySQLStr & "WHERE     (SC09001 = 'RUS') "
            MySQLStr = MySQLStr & "UNION "
            MySQLStr = MySQLStr & "SELECT     39 AS Expr1, SC09041 "
            MySQLStr = MySQLStr & "FROM         SC090300 AS SC090300_2 WITH(NOLOCK) "
            MySQLStr = MySQLStr & "WHERE     (SC09001 = 'RUS') "
            MySQLStr = MySQLStr & "UNION "
            MySQLStr = MySQLStr & "SELECT     40 AS Expr1, SC09042 "
            MySQLStr = MySQLStr & "FROM         SC090300 AS SC090300_1 WITH(NOLOCK) "
            MySQLStr = MySQLStr & "WHERE     (SC09001 = 'RUS')) AS View_1_1 ON PC030300.PC03009 = View_1_1.num "
            MySQLStr = MySQLStr & "WHERE (PC010300.PC01052 = N'" & Trim(DataGridView1.SelectedRows.Item(0).Cells(0).Value.ToString()) & "') "
            MySQLStr = MySQLStr & "ORDER BY PC030300.PC03001, PC030300.PC03005 "
        ElseIf MyType = 1 Then           '---в обобщенном заказе и одном добавляемом
            MySQLStr = "SELECT View_2.PC03001 AS OrderN, View_2.PC03005 AS ItemN, View_1_1.txt AS OrderUOM, View_1.txt AS CardUOM "
            MySQLStr = MySQLStr & "FROM SC010300 WITH(NOLOCK) INNER JOIN "
            MySQLStr = MySQLStr & "(SELECT 0 AS num, SC09002 AS txt "
            MySQLStr = MySQLStr & "FROM SC090300 WITH(NOLOCK) "
            MySQLStr = MySQLStr & "WHERE   (SC09001 = 'RUS') "
            MySQLStr = MySQLStr & "UNION "
            MySQLStr = MySQLStr & "SELECT     1 AS Expr1, SC09003 "
            MySQLStr = MySQLStr & "FROM         SC090300 AS SC090300_40 WITH(NOLOCK) "
            MySQLStr = MySQLStr & "WHERE     (SC09001 = 'RUS') "
            MySQLStr = MySQLStr & "UNION "
            MySQLStr = MySQLStr & "SELECT     2 AS Expr1, SC09004 "
            MySQLStr = MySQLStr & "FROM         SC090300 AS SC090300_39 WITH(NOLOCK) "
            MySQLStr = MySQLStr & "WHERE     (SC09001 = 'RUS') "
            MySQLStr = MySQLStr & "UNION "
            MySQLStr = MySQLStr & "SELECT     3 AS Expr1, SC09005 "
            MySQLStr = MySQLStr & "FROM         SC090300 AS SC090300_38 WITH(NOLOCK) "
            MySQLStr = MySQLStr & "WHERE     (SC09001 = 'RUS') "
            MySQLStr = MySQLStr & "UNION "
            MySQLStr = MySQLStr & "SELECT     4 AS Expr1, SC09006 "
            MySQLStr = MySQLStr & "FROM         SC090300 AS SC090300_37 WITH(NOLOCK) "
            MySQLStr = MySQLStr & "WHERE     (SC09001 = 'RUS')"
            MySQLStr = MySQLStr & "UNION "
            MySQLStr = MySQLStr & "SELECT     5 AS Expr1, SC09007 "
            MySQLStr = MySQLStr & "FROM         SC090300 AS SC090300_36 WITH(NOLOCK) "
            MySQLStr = MySQLStr & "WHERE     (SC09001 = 'RUS') "
            MySQLStr = MySQLStr & "UNION "
            MySQLStr = MySQLStr & "SELECT     6 AS Expr1, SC09008 "
            MySQLStr = MySQLStr & "FROM         SC090300 AS SC090300_35 WITH(NOLOCK) "
            MySQLStr = MySQLStr & "WHERE     (SC09001 = 'RUS') "
            MySQLStr = MySQLStr & "UNION "
            MySQLStr = MySQLStr & "SELECT     7 AS Expr1, SC09009 "
            MySQLStr = MySQLStr & "FROM         SC090300 AS SC090300_34 WITH(NOLOCK) "
            MySQLStr = MySQLStr & "WHERE     (SC09001 = 'RUS') "
            MySQLStr = MySQLStr & "UNION "
            MySQLStr = MySQLStr & "SELECT     8 AS Expr1, SC09010 "
            MySQLStr = MySQLStr & "FROM         SC090300 AS SC090300_33 WITH(NOLOCK) "
            MySQLStr = MySQLStr & "WHERE     (SC09001 = 'RUS') "
            MySQLStr = MySQLStr & "UNION "
            MySQLStr = MySQLStr & "SELECT     9 AS Expr1, SC09011 "
            MySQLStr = MySQLStr & "FROM         SC090300 AS SC090300_32 WITH(NOLOCK) "
            MySQLStr = MySQLStr & "WHERE     (SC09001 = 'RUS') "
            MySQLStr = MySQLStr & "UNION "
            MySQLStr = MySQLStr & "SELECT     10 AS Expr1, SC09012 "
            MySQLStr = MySQLStr & "FROM         SC090300 AS SC090300_31 WITH(NOLOCK) "
            MySQLStr = MySQLStr & "WHERE     (SC09001 = 'RUS') "
            MySQLStr = MySQLStr & "UNION "
            MySQLStr = MySQLStr & "SELECT     11 AS Expr1, SC09013 "
            MySQLStr = MySQLStr & "FROM         SC090300 AS SC090300_30 WITH(NOLOCK) "
            MySQLStr = MySQLStr & "WHERE     (SC09001 = 'RUS') "
            MySQLStr = MySQLStr & "UNION "
            MySQLStr = MySQLStr & "SELECT     12 AS Expr1, SC09014 "
            MySQLStr = MySQLStr & "FROM         SC090300 AS SC090300_29 WITH(NOLOCK) "
            MySQLStr = MySQLStr & "WHERE     (SC09001 = 'RUS') "
            MySQLStr = MySQLStr & "UNION "
            MySQLStr = MySQLStr & "SELECT     13 AS Expr1, SC09015 "
            MySQLStr = MySQLStr & "FROM         SC090300 AS SC090300_28 WITH(NOLOCK) "
            MySQLStr = MySQLStr & "WHERE     (SC09001 = 'RUS') "
            MySQLStr = MySQLStr & "UNION "
            MySQLStr = MySQLStr & "SELECT     14 AS Expr1, SC09016 "
            MySQLStr = MySQLStr & "FROM         SC090300 AS SC090300_27 WITH(NOLOCK) "
            MySQLStr = MySQLStr & "WHERE     (SC09001 = 'RUS') "
            MySQLStr = MySQLStr & "UNION "
            MySQLStr = MySQLStr & "SELECT     15 AS Expr1, SC09017 "
            MySQLStr = MySQLStr & "FROM         SC090300 AS SC090300_26 WITH(NOLOCK) "
            MySQLStr = MySQLStr & "WHERE     (SC09001 = 'RUS') "
            MySQLStr = MySQLStr & "UNION "
            MySQLStr = MySQLStr & "SELECT     16 AS Expr1, SC09018 "
            MySQLStr = MySQLStr & "FROM         SC090300 AS SC090300_25 WITH(NOLOCK) "
            MySQLStr = MySQLStr & "WHERE     (SC09001 = 'RUS') "
            MySQLStr = MySQLStr & "UNION "
            MySQLStr = MySQLStr & "SELECT     17 AS Expr1, SC09019 "
            MySQLStr = MySQLStr & "FROM         SC090300 AS SC090300_24 WITH(NOLOCK) "
            MySQLStr = MySQLStr & "WHERE     (SC09001 = 'RUS') "
            MySQLStr = MySQLStr & "UNION "
            MySQLStr = MySQLStr & "SELECT     18 AS Expr1, SC09020 "
            MySQLStr = MySQLStr & "FROM         SC090300 AS SC090300_23 WITH(NOLOCK) "
            MySQLStr = MySQLStr & "WHERE     (SC09001 = 'RUS') "
            MySQLStr = MySQLStr & "UNION "
            MySQLStr = MySQLStr & "SELECT     19 AS Expr1, SC09021 "
            MySQLStr = MySQLStr & "FROM         SC090300 AS SC090300_22 WITH(NOLOCK) "
            MySQLStr = MySQLStr & "WHERE     (SC09001 = 'RUS') "
            MySQLStr = MySQLStr & "UNION "
            MySQLStr = MySQLStr & "SELECT     20 AS Expr1, SC09022 "
            MySQLStr = MySQLStr & "FROM         SC090300 AS SC090300_21 WITH(NOLOCK) "
            MySQLStr = MySQLStr & "WHERE     (SC09001 = 'RUS') "
            MySQLStr = MySQLStr & "UNION "
            MySQLStr = MySQLStr & "SELECT     21 AS Expr1, SC09023 "
            MySQLStr = MySQLStr & "FROM         SC090300 AS SC090300_20 WITH(NOLOCK) "
            MySQLStr = MySQLStr & "WHERE     (SC09001 = 'RUS') "
            MySQLStr = MySQLStr & "UNION "
            MySQLStr = MySQLStr & "SELECT     22 AS Expr1, SC09024 "
            MySQLStr = MySQLStr & "FROM         SC090300 AS SC090300_19 WITH(NOLOCK) "
            MySQLStr = MySQLStr & "WHERE     (SC09001 = 'RUS') "
            MySQLStr = MySQLStr & "UNION "
            MySQLStr = MySQLStr & "SELECT     23 AS Expr1, SC09025 "
            MySQLStr = MySQLStr & "FROM         SC090300 AS SC090300_18 WITH(NOLOCK) "
            MySQLStr = MySQLStr & "WHERE     (SC09001 = 'RUS') "
            MySQLStr = MySQLStr & "UNION "
            MySQLStr = MySQLStr & "SELECT     24 AS Expr1, SC09026 "
            MySQLStr = MySQLStr & "FROM         SC090300 AS SC090300_17 WITH(NOLOCK) "
            MySQLStr = MySQLStr & "WHERE     (SC09001 = 'RUS') "
            MySQLStr = MySQLStr & "UNION "
            MySQLStr = MySQLStr & "SELECT     25 AS Expr1, SC09027 "
            MySQLStr = MySQLStr & "FROM         SC090300 AS SC090300_16 WITH(NOLOCK) "
            MySQLStr = MySQLStr & "WHERE     (SC09001 = 'RUS') "
            MySQLStr = MySQLStr & "UNION "
            MySQLStr = MySQLStr & "SELECT     26 AS Expr1, SC09028 "
            MySQLStr = MySQLStr & "FROM         SC090300 AS SC090300_15 WITH(NOLOCK) "
            MySQLStr = MySQLStr & "WHERE     (SC09001 = 'RUS') "
            MySQLStr = MySQLStr & "UNION "
            MySQLStr = MySQLStr & "SELECT     27 AS Expr1, SC09029 "
            MySQLStr = MySQLStr & "FROM         SC090300 AS SC090300_14 WITH(NOLOCK) "
            MySQLStr = MySQLStr & "WHERE     (SC09001 = 'RUS') "
            MySQLStr = MySQLStr & "UNION "
            MySQLStr = MySQLStr & "SELECT     28 AS Expr1, SC09030 "
            MySQLStr = MySQLStr & "FROM         SC090300 AS SC090300_13 WITH(NOLOCK) "
            MySQLStr = MySQLStr & "WHERE     (SC09001 = 'RUS') "
            MySQLStr = MySQLStr & "UNION "
            MySQLStr = MySQLStr & "SELECT     29 AS Expr1, SC09031 "
            MySQLStr = MySQLStr & "FROM         SC090300 AS SC090300_12 WITH(NOLOCK) "
            MySQLStr = MySQLStr & "WHERE     (SC09001 = 'RUS') "
            MySQLStr = MySQLStr & "UNION "
            MySQLStr = MySQLStr & "SELECT     30 AS Expr1, SC09032 "
            MySQLStr = MySQLStr & "FROM         SC090300 AS SC090300_11 WITH(NOLOCK) "
            MySQLStr = MySQLStr & "WHERE     (SC09001 = 'RUS') "
            MySQLStr = MySQLStr & "UNION "
            MySQLStr = MySQLStr & "SELECT     31 AS Expr1, SC09033 "
            MySQLStr = MySQLStr & "FROM         SC090300 AS SC090300_10 WITH(NOLOCK) "
            MySQLStr = MySQLStr & "WHERE     (SC09001 = 'RUS') "
            MySQLStr = MySQLStr & "UNION "
            MySQLStr = MySQLStr & "SELECT     32 AS Expr1, SC09034 "
            MySQLStr = MySQLStr & "FROM         SC090300 AS SC090300_9 WITH(NOLOCK) "
            MySQLStr = MySQLStr & "WHERE     (SC09001 = 'RUS') "
            MySQLStr = MySQLStr & "UNION "
            MySQLStr = MySQLStr & "SELECT     33 AS Expr1, SC09035 "
            MySQLStr = MySQLStr & "FROM         SC090300 AS SC090300_8 WITH(NOLOCK) "
            MySQLStr = MySQLStr & "WHERE     (SC09001 = 'RUS') "
            MySQLStr = MySQLStr & "UNION "
            MySQLStr = MySQLStr & "SELECT     34 AS Expr1, SC09036 "
            MySQLStr = MySQLStr & "FROM         SC090300 AS SC090300_7 WITH(NOLOCK) "
            MySQLStr = MySQLStr & "WHERE     (SC09001 = 'RUS') "
            MySQLStr = MySQLStr & "UNION "
            MySQLStr = MySQLStr & "SELECT     35 AS Expr1, SC09037 "
            MySQLStr = MySQLStr & "FROM         SC090300 AS SC090300_6 WITH(NOLOCK) "
            MySQLStr = MySQLStr & "WHERE     (SC09001 = 'RUS') "
            MySQLStr = MySQLStr & "UNION "
            MySQLStr = MySQLStr & "SELECT     36 AS Expr1, SC09038 "
            MySQLStr = MySQLStr & "FROM         SC090300 AS SC090300_5 WITH(NOLOCK) "
            MySQLStr = MySQLStr & "WHERE     (SC09001 = 'RUS') "
            MySQLStr = MySQLStr & "UNION "
            MySQLStr = MySQLStr & "SELECT     37 AS Expr1, SC09039 "
            MySQLStr = MySQLStr & "FROM         SC090300 AS SC090300_4 WITH(NOLOCK) "
            MySQLStr = MySQLStr & "WHERE     (SC09001 = 'RUS') "
            MySQLStr = MySQLStr & "UNION "
            MySQLStr = MySQLStr & "SELECT     38 AS Expr1, SC09040 "
            MySQLStr = MySQLStr & "FROM         SC090300 AS SC090300_3 WITH(NOLOCK) "
            MySQLStr = MySQLStr & "WHERE     (SC09001 = 'RUS') "
            MySQLStr = MySQLStr & "UNION "
            MySQLStr = MySQLStr & "SELECT     39 AS Expr1, SC09041 "
            MySQLStr = MySQLStr & "FROM         SC090300 AS SC090300_2 WITH(NOLOCK) "
            MySQLStr = MySQLStr & "WHERE     (SC09001 = 'RUS') "
            MySQLStr = MySQLStr & "UNION "
            MySQLStr = MySQLStr & "SELECT     40 AS Expr1, SC09042 "
            MySQLStr = MySQLStr & "FROM         SC090300 AS SC090300_1 WITH(NOLOCK) "
            MySQLStr = MySQLStr & "WHERE     (SC09001 = 'RUS')) AS View_1 ON SC010300.SC01134 = View_1.num INNER JOIN "
            MySQLStr = MySQLStr & "(SELECT PC030300.PC03001, PC030300.PC03005, PC030300.PC03009 "
            MySQLStr = MySQLStr & "FROM PC030300 WITH(NOLOCK) INNER JOIN "
            MySQLStr = MySQLStr & "PC010300 ON PC030300.PC03001 = PC010300.PC01001 "
            MySQLStr = MySQLStr & "WHERE (PC010300.PC01052 = N'" & Trim(DataGridView1.SelectedRows.Item(0).Cells(0).Value.ToString()) & "') "
            MySQLStr = MySQLStr & "UNION ALL "
            MySQLStr = MySQLStr & "SELECT PC030300_1.PC03001, PC030300_1.PC03005, PC030300_1.PC03009 "
            MySQLStr = MySQLStr & "FROM PC030300 AS PC030300_1 WITH(NOLOCK) INNER JOIN "
            MySQLStr = MySQLStr & "PC010300 AS PC010300_1 ON PC030300_1.PC03001 = PC010300_1.PC01001 "
            MySQLStr = MySQLStr & "WHERE (PC010300_1.PC01001 = N'" & Trim(DataGridView3.SelectedRows.Item(0).Cells(0).Value.ToString()) & "')) AS View_2 ON SC010300.SC01001 = View_2.PC03005 AND "
            MySQLStr = MySQLStr & "SC010300.SC01134 <> View_2.PC03009 INNER JOIN "
            MySQLStr = MySQLStr & "(SELECT 0 AS num, SC09002 AS txt "
            MySQLStr = MySQLStr & "FROM SC090300 WITH(NOLOCK) "
            MySQLStr = MySQLStr & "WHERE   (SC09001 = 'RUS') "
            MySQLStr = MySQLStr & "UNION "
            MySQLStr = MySQLStr & "SELECT     1 AS Expr1, SC09003 "
            MySQLStr = MySQLStr & "FROM         SC090300 AS SC090300_40 WITH(NOLOCK) "
            MySQLStr = MySQLStr & "WHERE     (SC09001 = 'RUS') "
            MySQLStr = MySQLStr & "UNION "
            MySQLStr = MySQLStr & "SELECT     2 AS Expr1, SC09004 "
            MySQLStr = MySQLStr & "FROM         SC090300 AS SC090300_39 WITH(NOLOCK) "
            MySQLStr = MySQLStr & "WHERE     (SC09001 = 'RUS') "
            MySQLStr = MySQLStr & "UNION "
            MySQLStr = MySQLStr & "SELECT     3 AS Expr1, SC09005 "
            MySQLStr = MySQLStr & "FROM         SC090300 AS SC090300_38 WITH(NOLOCK) "
            MySQLStr = MySQLStr & "WHERE     (SC09001 = 'RUS') "
            MySQLStr = MySQLStr & "UNION "
            MySQLStr = MySQLStr & "SELECT     4 AS Expr1, SC09006 "
            MySQLStr = MySQLStr & "FROM         SC090300 AS SC090300_37 WITH(NOLOCK) "
            MySQLStr = MySQLStr & "WHERE     (SC09001 = 'RUS')"
            MySQLStr = MySQLStr & "UNION "
            MySQLStr = MySQLStr & "SELECT     5 AS Expr1, SC09007 "
            MySQLStr = MySQLStr & "FROM         SC090300 AS SC090300_36 WITH(NOLOCK) "
            MySQLStr = MySQLStr & "WHERE     (SC09001 = 'RUS') "
            MySQLStr = MySQLStr & "UNION "
            MySQLStr = MySQLStr & "SELECT     6 AS Expr1, SC09008 "
            MySQLStr = MySQLStr & "FROM         SC090300 AS SC090300_35 WITH(NOLOCK) "
            MySQLStr = MySQLStr & "WHERE     (SC09001 = 'RUS') "
            MySQLStr = MySQLStr & "UNION "
            MySQLStr = MySQLStr & "SELECT     7 AS Expr1, SC09009 "
            MySQLStr = MySQLStr & "FROM         SC090300 AS SC090300_34 WITH(NOLOCK) "
            MySQLStr = MySQLStr & "WHERE     (SC09001 = 'RUS') "
            MySQLStr = MySQLStr & "UNION "
            MySQLStr = MySQLStr & "SELECT     8 AS Expr1, SC09010 "
            MySQLStr = MySQLStr & "FROM         SC090300 AS SC090300_33 WITH(NOLOCK) "
            MySQLStr = MySQLStr & "WHERE     (SC09001 = 'RUS') "
            MySQLStr = MySQLStr & "UNION "
            MySQLStr = MySQLStr & "SELECT     9 AS Expr1, SC09011 "
            MySQLStr = MySQLStr & "FROM         SC090300 AS SC090300_32 WITH(NOLOCK) "
            MySQLStr = MySQLStr & "WHERE     (SC09001 = 'RUS') "
            MySQLStr = MySQLStr & "UNION "
            MySQLStr = MySQLStr & "SELECT     10 AS Expr1, SC09012 "
            MySQLStr = MySQLStr & "FROM         SC090300 AS SC090300_31 WITH(NOLOCK) "
            MySQLStr = MySQLStr & "WHERE     (SC09001 = 'RUS') "
            MySQLStr = MySQLStr & "UNION "
            MySQLStr = MySQLStr & "SELECT     11 AS Expr1, SC09013 "
            MySQLStr = MySQLStr & "FROM         SC090300 AS SC090300_30 WITH(NOLOCK) "
            MySQLStr = MySQLStr & "WHERE     (SC09001 = 'RUS') "
            MySQLStr = MySQLStr & "UNION "
            MySQLStr = MySQLStr & "SELECT     12 AS Expr1, SC09014 "
            MySQLStr = MySQLStr & "FROM         SC090300 AS SC090300_29 WITH(NOLOCK) "
            MySQLStr = MySQLStr & "WHERE     (SC09001 = 'RUS') "
            MySQLStr = MySQLStr & "UNION "
            MySQLStr = MySQLStr & "SELECT     13 AS Expr1, SC09015 "
            MySQLStr = MySQLStr & "FROM         SC090300 AS SC090300_28 WITH(NOLOCK) "
            MySQLStr = MySQLStr & "WHERE     (SC09001 = 'RUS') "
            MySQLStr = MySQLStr & "UNION "
            MySQLStr = MySQLStr & "SELECT     14 AS Expr1, SC09016 "
            MySQLStr = MySQLStr & "FROM         SC090300 AS SC090300_27 WITH(NOLOCK) "
            MySQLStr = MySQLStr & "WHERE     (SC09001 = 'RUS') "
            MySQLStr = MySQLStr & "UNION "
            MySQLStr = MySQLStr & "SELECT     15 AS Expr1, SC09017 "
            MySQLStr = MySQLStr & "FROM         SC090300 AS SC090300_26 WITH(NOLOCK) "
            MySQLStr = MySQLStr & "WHERE     (SC09001 = 'RUS') "
            MySQLStr = MySQLStr & "UNION "
            MySQLStr = MySQLStr & "SELECT     16 AS Expr1, SC09018 "
            MySQLStr = MySQLStr & "FROM         SC090300 AS SC090300_25 WITH(NOLOCK) "
            MySQLStr = MySQLStr & "WHERE     (SC09001 = 'RUS') "
            MySQLStr = MySQLStr & "UNION "
            MySQLStr = MySQLStr & "SELECT     17 AS Expr1, SC09019 "
            MySQLStr = MySQLStr & "FROM         SC090300 AS SC090300_24 WITH(NOLOCK) "
            MySQLStr = MySQLStr & "WHERE     (SC09001 = 'RUS') "
            MySQLStr = MySQLStr & "UNION "
            MySQLStr = MySQLStr & "SELECT     18 AS Expr1, SC09020 "
            MySQLStr = MySQLStr & "FROM         SC090300 AS SC090300_23 WITH(NOLOCK) "
            MySQLStr = MySQLStr & "WHERE     (SC09001 = 'RUS') "
            MySQLStr = MySQLStr & "UNION "
            MySQLStr = MySQLStr & "SELECT     19 AS Expr1, SC09021 "
            MySQLStr = MySQLStr & "FROM         SC090300 AS SC090300_22 WITH(NOLOCK) "
            MySQLStr = MySQLStr & "WHERE     (SC09001 = 'RUS') "
            MySQLStr = MySQLStr & "UNION "
            MySQLStr = MySQLStr & "SELECT     20 AS Expr1, SC09022 "
            MySQLStr = MySQLStr & "FROM         SC090300 AS SC090300_21 WITH(NOLOCK) "
            MySQLStr = MySQLStr & "WHERE     (SC09001 = 'RUS') "
            MySQLStr = MySQLStr & "UNION "
            MySQLStr = MySQLStr & "SELECT     21 AS Expr1, SC09023 "
            MySQLStr = MySQLStr & "FROM         SC090300 AS SC090300_20 WITH(NOLOCK) "
            MySQLStr = MySQLStr & "WHERE     (SC09001 = 'RUS') "
            MySQLStr = MySQLStr & "UNION "
            MySQLStr = MySQLStr & "SELECT     22 AS Expr1, SC09024 "
            MySQLStr = MySQLStr & "FROM         SC090300 AS SC090300_19 WITH(NOLOCK) "
            MySQLStr = MySQLStr & "WHERE     (SC09001 = 'RUS') "
            MySQLStr = MySQLStr & "UNION "
            MySQLStr = MySQLStr & "SELECT     23 AS Expr1, SC09025 "
            MySQLStr = MySQLStr & "FROM         SC090300 AS SC090300_18 WITH(NOLOCK) "
            MySQLStr = MySQLStr & "WHERE     (SC09001 = 'RUS') "
            MySQLStr = MySQLStr & "UNION "
            MySQLStr = MySQLStr & "SELECT     24 AS Expr1, SC09026 "
            MySQLStr = MySQLStr & "FROM         SC090300 AS SC090300_17 WITH(NOLOCK) "
            MySQLStr = MySQLStr & "WHERE     (SC09001 = 'RUS') "
            MySQLStr = MySQLStr & "UNION "
            MySQLStr = MySQLStr & "SELECT     25 AS Expr1, SC09027 "
            MySQLStr = MySQLStr & "FROM         SC090300 AS SC090300_16 WITH(NOLOCK) "
            MySQLStr = MySQLStr & "WHERE     (SC09001 = 'RUS') "
            MySQLStr = MySQLStr & "UNION "
            MySQLStr = MySQLStr & "SELECT     26 AS Expr1, SC09028 "
            MySQLStr = MySQLStr & "FROM         SC090300 AS SC090300_15 WITH(NOLOCK) "
            MySQLStr = MySQLStr & "WHERE     (SC09001 = 'RUS') "
            MySQLStr = MySQLStr & "UNION "
            MySQLStr = MySQLStr & "SELECT     27 AS Expr1, SC09029 "
            MySQLStr = MySQLStr & "FROM         SC090300 AS SC090300_14 WITH(NOLOCK) "
            MySQLStr = MySQLStr & "WHERE     (SC09001 = 'RUS') "
            MySQLStr = MySQLStr & "UNION "
            MySQLStr = MySQLStr & "SELECT     28 AS Expr1, SC09030 "
            MySQLStr = MySQLStr & "FROM         SC090300 AS SC090300_13 WITH(NOLOCK) "
            MySQLStr = MySQLStr & "WHERE     (SC09001 = 'RUS') "
            MySQLStr = MySQLStr & "UNION "
            MySQLStr = MySQLStr & "SELECT     29 AS Expr1, SC09031 "
            MySQLStr = MySQLStr & "FROM         SC090300 AS SC090300_12 WITH(NOLOCK) "
            MySQLStr = MySQLStr & "WHERE     (SC09001 = 'RUS') "
            MySQLStr = MySQLStr & "UNION "
            MySQLStr = MySQLStr & "SELECT     30 AS Expr1, SC09032 "
            MySQLStr = MySQLStr & "FROM         SC090300 AS SC090300_11 WITH(NOLOCK) "
            MySQLStr = MySQLStr & "WHERE     (SC09001 = 'RUS') "
            MySQLStr = MySQLStr & "UNION "
            MySQLStr = MySQLStr & "SELECT     31 AS Expr1, SC09033 "
            MySQLStr = MySQLStr & "FROM         SC090300 AS SC090300_10 WITH(NOLOCK) "
            MySQLStr = MySQLStr & "WHERE     (SC09001 = 'RUS') "
            MySQLStr = MySQLStr & "UNION "
            MySQLStr = MySQLStr & "SELECT     32 AS Expr1, SC09034 "
            MySQLStr = MySQLStr & "FROM         SC090300 AS SC090300_9 WITH(NOLOCK) "
            MySQLStr = MySQLStr & "WHERE     (SC09001 = 'RUS') "
            MySQLStr = MySQLStr & "UNION "
            MySQLStr = MySQLStr & "SELECT     33 AS Expr1, SC09035 "
            MySQLStr = MySQLStr & "FROM         SC090300 AS SC090300_8 WITH(NOLOCK) "
            MySQLStr = MySQLStr & "WHERE     (SC09001 = 'RUS') "
            MySQLStr = MySQLStr & "UNION "
            MySQLStr = MySQLStr & "SELECT     34 AS Expr1, SC09036 "
            MySQLStr = MySQLStr & "FROM         SC090300 AS SC090300_7 WITH(NOLOCK) "
            MySQLStr = MySQLStr & "WHERE     (SC09001 = 'RUS') "
            MySQLStr = MySQLStr & "UNION "
            MySQLStr = MySQLStr & "SELECT     35 AS Expr1, SC09037 "
            MySQLStr = MySQLStr & "FROM         SC090300 AS SC090300_6 WITH(NOLOCK) "
            MySQLStr = MySQLStr & "WHERE     (SC09001 = 'RUS') "
            MySQLStr = MySQLStr & "UNION "
            MySQLStr = MySQLStr & "SELECT     36 AS Expr1, SC09038 "
            MySQLStr = MySQLStr & "FROM         SC090300 AS SC090300_5 WITH(NOLOCK) "
            MySQLStr = MySQLStr & "WHERE     (SC09001 = 'RUS') "
            MySQLStr = MySQLStr & "UNION "
            MySQLStr = MySQLStr & "SELECT     37 AS Expr1, SC09039 "
            MySQLStr = MySQLStr & "FROM         SC090300 AS SC090300_4 WITH(NOLOCK) "
            MySQLStr = MySQLStr & "WHERE     (SC09001 = 'RUS') "
            MySQLStr = MySQLStr & "UNION "
            MySQLStr = MySQLStr & "SELECT     38 AS Expr1, SC09040 "
            MySQLStr = MySQLStr & "FROM         SC090300 AS SC090300_3 WITH(NOLOCK) "
            MySQLStr = MySQLStr & "WHERE     (SC09001 = 'RUS') "
            MySQLStr = MySQLStr & "UNION "
            MySQLStr = MySQLStr & "SELECT     39 AS Expr1, SC09041 "
            MySQLStr = MySQLStr & "FROM         SC090300 AS SC090300_2 WITH(NOLOCK) "
            MySQLStr = MySQLStr & "WHERE     (SC09001 = 'RUS') "
            MySQLStr = MySQLStr & "UNION "
            MySQLStr = MySQLStr & "SELECT     40 AS Expr1, SC09042 "
            MySQLStr = MySQLStr & "FROM         SC090300 AS SC090300_1 WITH(NOLOCK) "
            MySQLStr = MySQLStr & "WHERE     (SC09001 = 'RUS')) AS View_1_1 ON View_2.PC03009 = View_1_1.num "
        Else                             '---в обобщенном заказе и всех добавляемых
            MySQLStr = "SELECT View_2.PC03001 AS OrderN, View_2.PC03005 AS ItemN, View_1_1.txt AS OrderUOM, View_1.txt AS CardUOM "
            MySQLStr = MySQLStr & "FROM SC010300 WITH(NOLOCK) INNER JOIN "
            MySQLStr = MySQLStr & "(SELECT 0 AS num, SC09002 AS txt "
            MySQLStr = MySQLStr & "FROM SC090300 WITH(NOLOCK) "
            MySQLStr = MySQLStr & "WHERE   (SC09001 = 'RUS') "
            MySQLStr = MySQLStr & "UNION "
            MySQLStr = MySQLStr & "SELECT     1 AS Expr1, SC09003 "
            MySQLStr = MySQLStr & "FROM         SC090300 AS SC090300_40 WITH(NOLOCK) "
            MySQLStr = MySQLStr & "WHERE     (SC09001 = 'RUS') "
            MySQLStr = MySQLStr & "UNION "
            MySQLStr = MySQLStr & "SELECT     2 AS Expr1, SC09004 "
            MySQLStr = MySQLStr & "FROM         SC090300 AS SC090300_39 WITH(NOLOCK) "
            MySQLStr = MySQLStr & "WHERE     (SC09001 = 'RUS') "
            MySQLStr = MySQLStr & "UNION "
            MySQLStr = MySQLStr & "SELECT     3 AS Expr1, SC09005 "
            MySQLStr = MySQLStr & "FROM         SC090300 AS SC090300_38 WITH(NOLOCK) "
            MySQLStr = MySQLStr & "WHERE     (SC09001 = 'RUS') "
            MySQLStr = MySQLStr & "UNION "
            MySQLStr = MySQLStr & "SELECT     4 AS Expr1, SC09006 "
            MySQLStr = MySQLStr & "FROM         SC090300 AS SC090300_37 WITH(NOLOCK) "
            MySQLStr = MySQLStr & "WHERE     (SC09001 = 'RUS')"
            MySQLStr = MySQLStr & "UNION "
            MySQLStr = MySQLStr & "SELECT     5 AS Expr1, SC09007 "
            MySQLStr = MySQLStr & "FROM         SC090300 AS SC090300_36 WITH(NOLOCK) "
            MySQLStr = MySQLStr & "WHERE     (SC09001 = 'RUS') "
            MySQLStr = MySQLStr & "UNION "
            MySQLStr = MySQLStr & "SELECT     6 AS Expr1, SC09008 "
            MySQLStr = MySQLStr & "FROM         SC090300 AS SC090300_35 WITH(NOLOCK) "
            MySQLStr = MySQLStr & "WHERE     (SC09001 = 'RUS') "
            MySQLStr = MySQLStr & "UNION "
            MySQLStr = MySQLStr & "SELECT     7 AS Expr1, SC09009 "
            MySQLStr = MySQLStr & "FROM         SC090300 AS SC090300_34 WITH(NOLOCK) "
            MySQLStr = MySQLStr & "WHERE     (SC09001 = 'RUS') "
            MySQLStr = MySQLStr & "UNION "
            MySQLStr = MySQLStr & "SELECT     8 AS Expr1, SC09010 "
            MySQLStr = MySQLStr & "FROM         SC090300 AS SC090300_33 WITH(NOLOCK) "
            MySQLStr = MySQLStr & "WHERE     (SC09001 = 'RUS') "
            MySQLStr = MySQLStr & "UNION "
            MySQLStr = MySQLStr & "SELECT     9 AS Expr1, SC09011 "
            MySQLStr = MySQLStr & "FROM         SC090300 AS SC090300_32 WITH(NOLOCK) "
            MySQLStr = MySQLStr & "WHERE     (SC09001 = 'RUS') "
            MySQLStr = MySQLStr & "UNION "
            MySQLStr = MySQLStr & "SELECT     10 AS Expr1, SC09012 "
            MySQLStr = MySQLStr & "FROM         SC090300 AS SC090300_31 WITH(NOLOCK) "
            MySQLStr = MySQLStr & "WHERE     (SC09001 = 'RUS') "
            MySQLStr = MySQLStr & "UNION "
            MySQLStr = MySQLStr & "SELECT     11 AS Expr1, SC09013 "
            MySQLStr = MySQLStr & "FROM         SC090300 AS SC090300_30 WITH(NOLOCK) "
            MySQLStr = MySQLStr & "WHERE     (SC09001 = 'RUS') "
            MySQLStr = MySQLStr & "UNION "
            MySQLStr = MySQLStr & "SELECT     12 AS Expr1, SC09014 "
            MySQLStr = MySQLStr & "FROM         SC090300 AS SC090300_29 WITH(NOLOCK) "
            MySQLStr = MySQLStr & "WHERE     (SC09001 = 'RUS') "
            MySQLStr = MySQLStr & "UNION "
            MySQLStr = MySQLStr & "SELECT     13 AS Expr1, SC09015 "
            MySQLStr = MySQLStr & "FROM         SC090300 AS SC090300_28 WITH(NOLOCK) "
            MySQLStr = MySQLStr & "WHERE     (SC09001 = 'RUS') "
            MySQLStr = MySQLStr & "UNION "
            MySQLStr = MySQLStr & "SELECT     14 AS Expr1, SC09016 "
            MySQLStr = MySQLStr & "FROM         SC090300 AS SC090300_27 WITH(NOLOCK) "
            MySQLStr = MySQLStr & "WHERE     (SC09001 = 'RUS') "
            MySQLStr = MySQLStr & "UNION "
            MySQLStr = MySQLStr & "SELECT     15 AS Expr1, SC09017 "
            MySQLStr = MySQLStr & "FROM         SC090300 AS SC090300_26 WITH(NOLOCK) "
            MySQLStr = MySQLStr & "WHERE     (SC09001 = 'RUS') "
            MySQLStr = MySQLStr & "UNION "
            MySQLStr = MySQLStr & "SELECT     16 AS Expr1, SC09018 "
            MySQLStr = MySQLStr & "FROM         SC090300 AS SC090300_25 WITH(NOLOCK) "
            MySQLStr = MySQLStr & "WHERE     (SC09001 = 'RUS') "
            MySQLStr = MySQLStr & "UNION "
            MySQLStr = MySQLStr & "SELECT     17 AS Expr1, SC09019 "
            MySQLStr = MySQLStr & "FROM         SC090300 AS SC090300_24 WITH(NOLOCK) "
            MySQLStr = MySQLStr & "WHERE     (SC09001 = 'RUS') "
            MySQLStr = MySQLStr & "UNION "
            MySQLStr = MySQLStr & "SELECT     18 AS Expr1, SC09020 "
            MySQLStr = MySQLStr & "FROM         SC090300 AS SC090300_23 WITH(NOLOCK) "
            MySQLStr = MySQLStr & "WHERE     (SC09001 = 'RUS') "
            MySQLStr = MySQLStr & "UNION "
            MySQLStr = MySQLStr & "SELECT     19 AS Expr1, SC09021 "
            MySQLStr = MySQLStr & "FROM         SC090300 AS SC090300_22 WITH(NOLOCK) "
            MySQLStr = MySQLStr & "WHERE     (SC09001 = 'RUS') "
            MySQLStr = MySQLStr & "UNION "
            MySQLStr = MySQLStr & "SELECT     20 AS Expr1, SC09022 "
            MySQLStr = MySQLStr & "FROM         SC090300 AS SC090300_21 WITH(NOLOCK) "
            MySQLStr = MySQLStr & "WHERE     (SC09001 = 'RUS') "
            MySQLStr = MySQLStr & "UNION "
            MySQLStr = MySQLStr & "SELECT     21 AS Expr1, SC09023 "
            MySQLStr = MySQLStr & "FROM         SC090300 AS SC090300_20 WITH(NOLOCK) "
            MySQLStr = MySQLStr & "WHERE     (SC09001 = 'RUS') "
            MySQLStr = MySQLStr & "UNION "
            MySQLStr = MySQLStr & "SELECT     22 AS Expr1, SC09024 "
            MySQLStr = MySQLStr & "FROM         SC090300 AS SC090300_19 WITH(NOLOCK) "
            MySQLStr = MySQLStr & "WHERE     (SC09001 = 'RUS') "
            MySQLStr = MySQLStr & "UNION "
            MySQLStr = MySQLStr & "SELECT     23 AS Expr1, SC09025 "
            MySQLStr = MySQLStr & "FROM         SC090300 AS SC090300_18 WITH(NOLOCK) "
            MySQLStr = MySQLStr & "WHERE     (SC09001 = 'RUS') "
            MySQLStr = MySQLStr & "UNION "
            MySQLStr = MySQLStr & "SELECT     24 AS Expr1, SC09026 "
            MySQLStr = MySQLStr & "FROM         SC090300 AS SC090300_17 WITH(NOLOCK) "
            MySQLStr = MySQLStr & "WHERE     (SC09001 = 'RUS') "
            MySQLStr = MySQLStr & "UNION "
            MySQLStr = MySQLStr & "SELECT     25 AS Expr1, SC09027 "
            MySQLStr = MySQLStr & "FROM         SC090300 AS SC090300_16 WITH(NOLOCK) "
            MySQLStr = MySQLStr & "WHERE     (SC09001 = 'RUS') "
            MySQLStr = MySQLStr & "UNION "
            MySQLStr = MySQLStr & "SELECT     26 AS Expr1, SC09028 "
            MySQLStr = MySQLStr & "FROM         SC090300 AS SC090300_15 WITH(NOLOCK) "
            MySQLStr = MySQLStr & "WHERE     (SC09001 = 'RUS') "
            MySQLStr = MySQLStr & "UNION "
            MySQLStr = MySQLStr & "SELECT     27 AS Expr1, SC09029 "
            MySQLStr = MySQLStr & "FROM         SC090300 AS SC090300_14 WITH(NOLOCK) "
            MySQLStr = MySQLStr & "WHERE     (SC09001 = 'RUS') "
            MySQLStr = MySQLStr & "UNION "
            MySQLStr = MySQLStr & "SELECT     28 AS Expr1, SC09030 "
            MySQLStr = MySQLStr & "FROM         SC090300 AS SC090300_13 WITH(NOLOCK) "
            MySQLStr = MySQLStr & "WHERE     (SC09001 = 'RUS') "
            MySQLStr = MySQLStr & "UNION "
            MySQLStr = MySQLStr & "SELECT     29 AS Expr1, SC09031 "
            MySQLStr = MySQLStr & "FROM         SC090300 AS SC090300_12 WITH(NOLOCK) "
            MySQLStr = MySQLStr & "WHERE     (SC09001 = 'RUS') "
            MySQLStr = MySQLStr & "UNION "
            MySQLStr = MySQLStr & "SELECT     30 AS Expr1, SC09032 "
            MySQLStr = MySQLStr & "FROM         SC090300 AS SC090300_11 WITH(NOLOCK) "
            MySQLStr = MySQLStr & "WHERE     (SC09001 = 'RUS') "
            MySQLStr = MySQLStr & "UNION "
            MySQLStr = MySQLStr & "SELECT     31 AS Expr1, SC09033 "
            MySQLStr = MySQLStr & "FROM         SC090300 AS SC090300_10 WITH(NOLOCK) "
            MySQLStr = MySQLStr & "WHERE     (SC09001 = 'RUS') "
            MySQLStr = MySQLStr & "UNION "
            MySQLStr = MySQLStr & "SELECT     32 AS Expr1, SC09034 "
            MySQLStr = MySQLStr & "FROM         SC090300 AS SC090300_9 WITH(NOLOCK) "
            MySQLStr = MySQLStr & "WHERE     (SC09001 = 'RUS') "
            MySQLStr = MySQLStr & "UNION "
            MySQLStr = MySQLStr & "SELECT     33 AS Expr1, SC09035 "
            MySQLStr = MySQLStr & "FROM         SC090300 AS SC090300_8 WITH(NOLOCK) "
            MySQLStr = MySQLStr & "WHERE     (SC09001 = 'RUS') "
            MySQLStr = MySQLStr & "UNION "
            MySQLStr = MySQLStr & "SELECT     34 AS Expr1, SC09036 "
            MySQLStr = MySQLStr & "FROM         SC090300 AS SC090300_7 WITH(NOLOCK) "
            MySQLStr = MySQLStr & "WHERE     (SC09001 = 'RUS') "
            MySQLStr = MySQLStr & "UNION "
            MySQLStr = MySQLStr & "SELECT     35 AS Expr1, SC09037 "
            MySQLStr = MySQLStr & "FROM         SC090300 AS SC090300_6 WITH(NOLOCK) "
            MySQLStr = MySQLStr & "WHERE     (SC09001 = 'RUS') "
            MySQLStr = MySQLStr & "UNION "
            MySQLStr = MySQLStr & "SELECT     36 AS Expr1, SC09038 "
            MySQLStr = MySQLStr & "FROM         SC090300 AS SC090300_5 WITH(NOLOCK) "
            MySQLStr = MySQLStr & "WHERE     (SC09001 = 'RUS') "
            MySQLStr = MySQLStr & "UNION "
            MySQLStr = MySQLStr & "SELECT     37 AS Expr1, SC09039 "
            MySQLStr = MySQLStr & "FROM         SC090300 AS SC090300_4 WITH(NOLOCK) "
            MySQLStr = MySQLStr & "WHERE     (SC09001 = 'RUS') "
            MySQLStr = MySQLStr & "UNION "
            MySQLStr = MySQLStr & "SELECT     38 AS Expr1, SC09040 "
            MySQLStr = MySQLStr & "FROM         SC090300 AS SC090300_3 WITH(NOLOCK) "
            MySQLStr = MySQLStr & "WHERE     (SC09001 = 'RUS') "
            MySQLStr = MySQLStr & "UNION "
            MySQLStr = MySQLStr & "SELECT     39 AS Expr1, SC09041 "
            MySQLStr = MySQLStr & "FROM         SC090300 AS SC090300_2 WITH(NOLOCK) "
            MySQLStr = MySQLStr & "WHERE     (SC09001 = 'RUS') "
            MySQLStr = MySQLStr & "UNION "
            MySQLStr = MySQLStr & "SELECT     40 AS Expr1, SC09042 "
            MySQLStr = MySQLStr & "FROM         SC090300 AS SC090300_1 WITH(NOLOCK) "
            MySQLStr = MySQLStr & "WHERE     (SC09001 = 'RUS')) AS View_1 ON SC010300.SC01134 = View_1.num INNER JOIN "
            MySQLStr = MySQLStr & "(SELECT PC030300.PC03001, PC030300.PC03005, PC030300.PC03009 "
            MySQLStr = MySQLStr & "FROM PC030300 WITH(NOLOCK) INNER JOIN "
            MySQLStr = MySQLStr & "PC010300 ON PC030300.PC03001 = PC010300.PC01001 "
            MySQLStr = MySQLStr & "WHERE (PC010300.PC01052 = N'" & Trim(DataGridView1.SelectedRows.Item(0).Cells(0).Value.ToString()) & "') "
            MySQLStr = MySQLStr & "UNION ALL "
            MySQLStr = MySQLStr & "SELECT PC010300_1.PC01001, PC030300_1.PC03005, PC030300_1.PC03009 "
            MySQLStr = MySQLStr & "FROM PC010300 AS PC010300_1 WITH(NOLOCK) INNER JOIN "
            MySQLStr = MySQLStr & "(SELECT PC03001 "
            MySQLStr = MySQLStr & "FROM PC030300 AS PC030300_2 WITH(NOLOCK) "
            MySQLStr = MySQLStr & "WHERE (PC03010 <> 0) OR "
            MySQLStr = MySQLStr & "(PC03011 <> 0) "
            MySQLStr = MySQLStr & "GROUP BY PC03001) AS View_2_1 ON PC010300_1.PC01001 = View_2_1.PC03001 INNER JOIN "
            MySQLStr = MySQLStr & "PC030300 AS PC030300_1 ON PC010300_1.PC01001 = PC030300_1.PC03001 "
            MySQLStr = MySQLStr & "WHERE (PC010300_1.PC01001 NOT IN "
            MySQLStr = MySQLStr & "(SELECT PC010300_2.PC01001 "
            MySQLStr = MySQLStr & "FROM tbl_PurchaseWorkplace_ConsolidatedOrders WITH(NOLOCK) INNER JOIN "
            MySQLStr = MySQLStr & "PC010300 AS PC010300_2 ON tbl_PurchaseWorkplace_ConsolidatedOrders.ID = PC010300_2.PC01052 "
            MySQLStr = MySQLStr & "GROUP BY PC010300_2.PC01001)) AND (PC010300_1.PC01002 <> 0) AND (PC010300_1.PC01023 = N'" & Declarations.MyWH & "') AND "
            MySQLStr = MySQLStr & "(PC010300_1.PC01003 = N'" & Declarations.MySupplierCode & "')) AS View_2 ON SC010300.SC01001 = View_2.PC03005 AND "
            MySQLStr = MySQLStr & "SC010300.SC01134 <> View_2.PC03009 INNER JOIN "
            MySQLStr = MySQLStr & "(SELECT 0 AS num, SC09002 AS txt "
            MySQLStr = MySQLStr & "FROM SC090300 WITH(NOLOCK) "
            MySQLStr = MySQLStr & "WHERE   (SC09001 = 'RUS') "
            MySQLStr = MySQLStr & "UNION "
            MySQLStr = MySQLStr & "SELECT     1 AS Expr1, SC09003 "
            MySQLStr = MySQLStr & "FROM         SC090300 AS SC090300_40 WITH(NOLOCK) "
            MySQLStr = MySQLStr & "WHERE     (SC09001 = 'RUS') "
            MySQLStr = MySQLStr & "UNION "
            MySQLStr = MySQLStr & "SELECT     2 AS Expr1, SC09004 "
            MySQLStr = MySQLStr & "FROM         SC090300 AS SC090300_39 WITH(NOLOCK) "
            MySQLStr = MySQLStr & "WHERE     (SC09001 = 'RUS') "
            MySQLStr = MySQLStr & "UNION "
            MySQLStr = MySQLStr & "SELECT     3 AS Expr1, SC09005 "
            MySQLStr = MySQLStr & "FROM         SC090300 AS SC090300_38 WITH(NOLOCK) "
            MySQLStr = MySQLStr & "WHERE     (SC09001 = 'RUS') "
            MySQLStr = MySQLStr & "UNION "
            MySQLStr = MySQLStr & "SELECT     4 AS Expr1, SC09006 "
            MySQLStr = MySQLStr & "FROM         SC090300 AS SC090300_37 WITH(NOLOCK) "
            MySQLStr = MySQLStr & "WHERE     (SC09001 = 'RUS')"
            MySQLStr = MySQLStr & "UNION "
            MySQLStr = MySQLStr & "SELECT     5 AS Expr1, SC09007 "
            MySQLStr = MySQLStr & "FROM         SC090300 AS SC090300_36 WITH(NOLOCK) "
            MySQLStr = MySQLStr & "WHERE     (SC09001 = 'RUS') "
            MySQLStr = MySQLStr & "UNION "
            MySQLStr = MySQLStr & "SELECT     6 AS Expr1, SC09008 "
            MySQLStr = MySQLStr & "FROM         SC090300 AS SC090300_35 WITH(NOLOCK) "
            MySQLStr = MySQLStr & "WHERE     (SC09001 = 'RUS') "
            MySQLStr = MySQLStr & "UNION "
            MySQLStr = MySQLStr & "SELECT     7 AS Expr1, SC09009 "
            MySQLStr = MySQLStr & "FROM         SC090300 AS SC090300_34 WITH(NOLOCK) "
            MySQLStr = MySQLStr & "WHERE     (SC09001 = 'RUS') "
            MySQLStr = MySQLStr & "UNION "
            MySQLStr = MySQLStr & "SELECT     8 AS Expr1, SC09010 "
            MySQLStr = MySQLStr & "FROM         SC090300 AS SC090300_33 WITH(NOLOCK) "
            MySQLStr = MySQLStr & "WHERE     (SC09001 = 'RUS') "
            MySQLStr = MySQLStr & "UNION "
            MySQLStr = MySQLStr & "SELECT     9 AS Expr1, SC09011 "
            MySQLStr = MySQLStr & "FROM         SC090300 AS SC090300_32 WITH(NOLOCK) "
            MySQLStr = MySQLStr & "WHERE     (SC09001 = 'RUS') "
            MySQLStr = MySQLStr & "UNION "
            MySQLStr = MySQLStr & "SELECT     10 AS Expr1, SC09012 "
            MySQLStr = MySQLStr & "FROM         SC090300 AS SC090300_31 WITH(NOLOCK) "
            MySQLStr = MySQLStr & "WHERE     (SC09001 = 'RUS') "
            MySQLStr = MySQLStr & "UNION "
            MySQLStr = MySQLStr & "SELECT     11 AS Expr1, SC09013 "
            MySQLStr = MySQLStr & "FROM         SC090300 AS SC090300_30 WITH(NOLOCK) "
            MySQLStr = MySQLStr & "WHERE     (SC09001 = 'RUS') "
            MySQLStr = MySQLStr & "UNION "
            MySQLStr = MySQLStr & "SELECT     12 AS Expr1, SC09014 "
            MySQLStr = MySQLStr & "FROM         SC090300 AS SC090300_29 WITH(NOLOCK) "
            MySQLStr = MySQLStr & "WHERE     (SC09001 = 'RUS') "
            MySQLStr = MySQLStr & "UNION "
            MySQLStr = MySQLStr & "SELECT     13 AS Expr1, SC09015 "
            MySQLStr = MySQLStr & "FROM         SC090300 AS SC090300_28 WITH(NOLOCK) "
            MySQLStr = MySQLStr & "WHERE     (SC09001 = 'RUS') "
            MySQLStr = MySQLStr & "UNION "
            MySQLStr = MySQLStr & "SELECT     14 AS Expr1, SC09016 "
            MySQLStr = MySQLStr & "FROM         SC090300 AS SC090300_27 WITH(NOLOCK) "
            MySQLStr = MySQLStr & "WHERE     (SC09001 = 'RUS') "
            MySQLStr = MySQLStr & "UNION "
            MySQLStr = MySQLStr & "SELECT     15 AS Expr1, SC09017 "
            MySQLStr = MySQLStr & "FROM         SC090300 AS SC090300_26 WITH(NOLOCK) "
            MySQLStr = MySQLStr & "WHERE     (SC09001 = 'RUS') "
            MySQLStr = MySQLStr & "UNION "
            MySQLStr = MySQLStr & "SELECT     16 AS Expr1, SC09018 "
            MySQLStr = MySQLStr & "FROM         SC090300 AS SC090300_25 WITH(NOLOCK) "
            MySQLStr = MySQLStr & "WHERE     (SC09001 = 'RUS') "
            MySQLStr = MySQLStr & "UNION "
            MySQLStr = MySQLStr & "SELECT     17 AS Expr1, SC09019 "
            MySQLStr = MySQLStr & "FROM         SC090300 AS SC090300_24 WITH(NOLOCK) "
            MySQLStr = MySQLStr & "WHERE     (SC09001 = 'RUS') "
            MySQLStr = MySQLStr & "UNION "
            MySQLStr = MySQLStr & "SELECT     18 AS Expr1, SC09020 "
            MySQLStr = MySQLStr & "FROM         SC090300 AS SC090300_23 WITH(NOLOCK) "
            MySQLStr = MySQLStr & "WHERE     (SC09001 = 'RUS') "
            MySQLStr = MySQLStr & "UNION "
            MySQLStr = MySQLStr & "SELECT     19 AS Expr1, SC09021 "
            MySQLStr = MySQLStr & "FROM         SC090300 AS SC090300_22 WITH(NOLOCK) "
            MySQLStr = MySQLStr & "WHERE     (SC09001 = 'RUS') "
            MySQLStr = MySQLStr & "UNION "
            MySQLStr = MySQLStr & "SELECT     20 AS Expr1, SC09022 "
            MySQLStr = MySQLStr & "FROM         SC090300 AS SC090300_21 WITH(NOLOCK) "
            MySQLStr = MySQLStr & "WHERE     (SC09001 = 'RUS') "
            MySQLStr = MySQLStr & "UNION "
            MySQLStr = MySQLStr & "SELECT     21 AS Expr1, SC09023 "
            MySQLStr = MySQLStr & "FROM         SC090300 AS SC090300_20 WITH(NOLOCK) "
            MySQLStr = MySQLStr & "WHERE     (SC09001 = 'RUS') "
            MySQLStr = MySQLStr & "UNION "
            MySQLStr = MySQLStr & "SELECT     22 AS Expr1, SC09024 "
            MySQLStr = MySQLStr & "FROM         SC090300 AS SC090300_19 WITH(NOLOCK) "
            MySQLStr = MySQLStr & "WHERE     (SC09001 = 'RUS') "
            MySQLStr = MySQLStr & "UNION "
            MySQLStr = MySQLStr & "SELECT     23 AS Expr1, SC09025 "
            MySQLStr = MySQLStr & "FROM         SC090300 AS SC090300_18 WITH(NOLOCK) "
            MySQLStr = MySQLStr & "WHERE     (SC09001 = 'RUS') "
            MySQLStr = MySQLStr & "UNION "
            MySQLStr = MySQLStr & "SELECT     24 AS Expr1, SC09026 "
            MySQLStr = MySQLStr & "FROM         SC090300 AS SC090300_17 WITH(NOLOCK) "
            MySQLStr = MySQLStr & "WHERE     (SC09001 = 'RUS') "
            MySQLStr = MySQLStr & "UNION "
            MySQLStr = MySQLStr & "SELECT     25 AS Expr1, SC09027 "
            MySQLStr = MySQLStr & "FROM         SC090300 AS SC090300_16 WITH(NOLOCK) "
            MySQLStr = MySQLStr & "WHERE     (SC09001 = 'RUS') "
            MySQLStr = MySQLStr & "UNION "
            MySQLStr = MySQLStr & "SELECT     26 AS Expr1, SC09028 "
            MySQLStr = MySQLStr & "FROM         SC090300 AS SC090300_15 WITH(NOLOCK) "
            MySQLStr = MySQLStr & "WHERE     (SC09001 = 'RUS') "
            MySQLStr = MySQLStr & "UNION "
            MySQLStr = MySQLStr & "SELECT     27 AS Expr1, SC09029 "
            MySQLStr = MySQLStr & "FROM         SC090300 AS SC090300_14 WITH(NOLOCK) "
            MySQLStr = MySQLStr & "WHERE     (SC09001 = 'RUS') "
            MySQLStr = MySQLStr & "UNION "
            MySQLStr = MySQLStr & "SELECT     28 AS Expr1, SC09030 "
            MySQLStr = MySQLStr & "FROM         SC090300 AS SC090300_13 WITH(NOLOCK) "
            MySQLStr = MySQLStr & "WHERE     (SC09001 = 'RUS') "
            MySQLStr = MySQLStr & "UNION "
            MySQLStr = MySQLStr & "SELECT     29 AS Expr1, SC09031 "
            MySQLStr = MySQLStr & "FROM         SC090300 AS SC090300_12 WITH(NOLOCK) "
            MySQLStr = MySQLStr & "WHERE     (SC09001 = 'RUS') "
            MySQLStr = MySQLStr & "UNION "
            MySQLStr = MySQLStr & "SELECT     30 AS Expr1, SC09032 "
            MySQLStr = MySQLStr & "FROM         SC090300 AS SC090300_11 WITH(NOLOCK) "
            MySQLStr = MySQLStr & "WHERE     (SC09001 = 'RUS') "
            MySQLStr = MySQLStr & "UNION "
            MySQLStr = MySQLStr & "SELECT     31 AS Expr1, SC09033 "
            MySQLStr = MySQLStr & "FROM         SC090300 AS SC090300_10 WITH(NOLOCK) "
            MySQLStr = MySQLStr & "WHERE     (SC09001 = 'RUS') "
            MySQLStr = MySQLStr & "UNION "
            MySQLStr = MySQLStr & "SELECT     32 AS Expr1, SC09034 "
            MySQLStr = MySQLStr & "FROM         SC090300 AS SC090300_9 WITH(NOLOCK) "
            MySQLStr = MySQLStr & "WHERE     (SC09001 = 'RUS') "
            MySQLStr = MySQLStr & "UNION "
            MySQLStr = MySQLStr & "SELECT     33 AS Expr1, SC09035 "
            MySQLStr = MySQLStr & "FROM         SC090300 AS SC090300_8 WITH(NOLOCK) "
            MySQLStr = MySQLStr & "WHERE     (SC09001 = 'RUS') "
            MySQLStr = MySQLStr & "UNION "
            MySQLStr = MySQLStr & "SELECT     34 AS Expr1, SC09036 "
            MySQLStr = MySQLStr & "FROM         SC090300 AS SC090300_7 WITH(NOLOCK) "
            MySQLStr = MySQLStr & "WHERE     (SC09001 = 'RUS') "
            MySQLStr = MySQLStr & "UNION "
            MySQLStr = MySQLStr & "SELECT     35 AS Expr1, SC09037 "
            MySQLStr = MySQLStr & "FROM         SC090300 AS SC090300_6 WITH(NOLOCK) "
            MySQLStr = MySQLStr & "WHERE     (SC09001 = 'RUS') "
            MySQLStr = MySQLStr & "UNION "
            MySQLStr = MySQLStr & "SELECT     36 AS Expr1, SC09038 "
            MySQLStr = MySQLStr & "FROM         SC090300 AS SC090300_5 WITH(NOLOCK) "
            MySQLStr = MySQLStr & "WHERE     (SC09001 = 'RUS') "
            MySQLStr = MySQLStr & "UNION "
            MySQLStr = MySQLStr & "SELECT     37 AS Expr1, SC09039 "
            MySQLStr = MySQLStr & "FROM         SC090300 AS SC090300_4 WITH(NOLOCK) "
            MySQLStr = MySQLStr & "WHERE     (SC09001 = 'RUS') "
            MySQLStr = MySQLStr & "UNION "
            MySQLStr = MySQLStr & "SELECT     38 AS Expr1, SC09040 "
            MySQLStr = MySQLStr & "FROM         SC090300 AS SC090300_3 WITH(NOLOCK) "
            MySQLStr = MySQLStr & "WHERE     (SC09001 = 'RUS') "
            MySQLStr = MySQLStr & "UNION "
            MySQLStr = MySQLStr & "SELECT     39 AS Expr1, SC09041 "
            MySQLStr = MySQLStr & "FROM         SC090300 AS SC090300_2 WITH(NOLOCK) "
            MySQLStr = MySQLStr & "WHERE     (SC09001 = 'RUS') "
            MySQLStr = MySQLStr & "UNION "
            MySQLStr = MySQLStr & "SELECT     40 AS Expr1, SC09042 "
            MySQLStr = MySQLStr & "FROM         SC090300 AS SC090300_1 WITH(NOLOCK) "
            MySQLStr = MySQLStr & "WHERE     (SC09001 = 'RUS')) AS View_1_1 ON View_2.PC03009 = View_1_1.num "
        End If
        InitMyConn(False)
        InitMyRec(False, MySQLStr)
        If Declarations.MyRec.BOF = True And Declarations.MyRec.EOF = True Then
            trycloseMyRec()
            CheckUOMInOrders = True
            Exit Function
        Else
            Declarations.MyRec.MoveFirst()
            MyERRStr = ""
            While Declarations.MyRec.EOF = False
                MyERRStr = MyERRStr & "Заказ " & Declarations.MyRec.Fields("OrderN").Value & "   Код товара " & Declarations.MyRec.Fields("ItemN").Value & "   Единица измерения в заказе " & Declarations.MyRec.Fields("OrderUOM").Value & "   Единица измерения в карточке " & Declarations.MyRec.Fields("CardUOM").Value & " " & Chr(13)
                Declarations.MyRec.MoveNext()
            End While
            trycloseMyRec()
            CheckUOMInOrders = False
            MyErrorForm = New ErrorForm
            If MyType = 0 Then
                MyERRStr = "В обобщенном заказе есть запасы, единица измерения которых отличается от единицы измерения, указанной в карточке товара. " & Chr(13) & Chr(13) & MyERRStr
            Else
                MyERRStr = "В обобщенном заказе или в заказах, которые вы пытаетесь добавить в обобщенный, есть запасы, единица измерения которых отличается от единицы измерения, указанной в карточке товара. " & Chr(13) & Chr(13) & MyERRStr
            End If
            MyErrorForm.MyMsg = MyERRStr
            MyErrorForm.ShowDialog()
        End If
    End Function

    Private Sub ExportOrderToExcel(ByVal ComOrder As String)
        '////////////////////////////////////////////////////////////////////////////////
        '//
        '// Процедура выгрузки обобщенного заказа в Excel
        '//
        '////////////////////////////////////////////////////////////////////////////////
        Dim MyObj As Object
        Dim MyWRKBook As Object
        Dim i As Integer                              'счетчик строк

        MyObj = CreateObject("Excel.Application")
        MyObj.SheetsInNewWorkbook = 1
        MyWRKBook = MyObj.Workbooks.Add

        ExportOrderHeaderToExcel(MyWRKBook, ComOrder, i)
        ExportOrderBodyToExcel(MyWRKBook, ComOrder, i)
        ExportOrderFooterToExcel(MyWRKBook, ComOrder, i)

        MyWRKBook.ActiveSheet.Range("A1").Select()
        MyObj.Application.Visible = True
        MyObj = Nothing
    End Sub

    Private Sub ExportOrderToLO(ByVal ComOrder As String)
        '////////////////////////////////////////////////////////////////////////////////
        '//
        '// Процедура выгрузки обобщенного заказа в Libre Office
        '//
        '////////////////////////////////////////////////////////////////////////////////
        Dim oServiceManager As Object
        Dim oDispatcher As Object
        Dim oDesktop As Object
        Dim oWorkBook As Object
        Dim oSheet As Object
        Dim oFrame As Object
        Dim i As Integer                              'счетчик строк

        LOSetNotation(0)
        oServiceManager = CreateObject("com.sun.star.ServiceManager")
        oDesktop = oServiceManager.createInstance("com.sun.star.frame.Desktop")
        oDispatcher = oServiceManager.createInstance("com.sun.star.frame.DispatchHelper")
        Dim arg(1)
        arg(0) = mAkePropertyValue("Hidden", True, oServiceManager)
        arg(1) = mAkePropertyValue("MacroExecutionMode", 4, oServiceManager)
        oWorkBook = oDesktop.loadComponentFromURL("private:factory/scalc", "_blank", 0, arg)
        oSheet = oWorkBook.getSheets().getByIndex(0)
        oFrame = oWorkBook.getCurrentController.getFrame

        ExportOrderHeaderToLO(oSheet, oServiceManager, oWorkBook, oDispatcher, ComOrder, i)
        ExportOrderBodyToLO(oSheet, oServiceManager, oWorkBook, oDispatcher, ComOrder, i)
        ExportOrderFooterToLO(oSheet, oServiceManager, oWorkBook, oDispatcher, ComOrder, i)

        Dim args() As Object
        ReDim args(0)
        args(0) = oServiceManager.Bridge_GetStruct("com.sun.star.beans.PropertyValue")
        args(0).Name = "ToPoint"
        args(0).Value = "$A$1"
        oDispatcher.executeDispatch(oFrame, ".uno:GoToCell", "", 0, args)

        oWorkBook.CurrentController.Frame.ContainerWindow.Visible = True
        oWorkBook.CurrentController.Frame.ContainerWindow.toFront()
    End Sub

    Private Sub ExportOrderHeaderToExcel(ByRef MyWRKBook As Object, ByVal ComOrder As String, ByRef i As Integer)
        '////////////////////////////////////////////////////////////////////////////////
        '//
        '// Процедура выгрузки заголовка обобщенного заказа в Excel
        '//
        '////////////////////////////////////////////////////////////////////////////////
        Dim MySQLStr As String
        Dim SupplierName As String
        Dim SupplierAddress As String
        Dim DelName As String
        Dim DelAddr As String
        Dim CurrName As String
        Dim MyLanguage As String

        '------Язык документа---------------------
        MySQLStr = "SELECT PL01027 AS CC "
        MySQLStr = MySQLStr & "FROM PL010300 "
        MySQLStr = MySQLStr & "WHERE (PL01001 = N'" & Trim(Label9.Text) & "') "
        InitMyConn(False)
        InitMyRec(False, MySQLStr)
        If Declarations.MyRec.BOF = True And Declarations.MyRec.EOF = True Then
            MyLanguage = "RUS"
            trycloseMyRec()
        Else
            MyLanguage = Declarations.MyRec.Fields("CC").Value
            trycloseMyRec()
        End If

        '------наша компания и адрес--------------
        MyWRKBook.ActiveSheet.Range("B2:J2").MergeCells = True
        If MyLanguage = "RUS" Then
            MyWRKBook.ActiveSheet.Range("B2") = "ООО ""Скандика"""
        Else
            MyWRKBook.ActiveSheet.Range("B2") = "Skandika LLC"
        End If
        MyWRKBook.ActiveSheet.Range("B2:J2").Font.Name = "Arial"
        MyWRKBook.ActiveSheet.Range("B2:J2").Font.Size = 12
        MyWRKBook.ActiveSheet.Range("B2:J2").Font.Bold = True
        MyWRKBook.ActiveSheet.Range("B2:J2").WrapText = True

        MyWRKBook.ActiveSheet.Range("B3:D3").MergeCells = True
        MyWRKBook.ActiveSheet.Range("B3") = "Address/Адрес:"
        MyWRKBook.ActiveSheet.Range("B3:D3").Font.Size = 10
        MyWRKBook.ActiveSheet.Range("B3:D3").Font.Bold = True

        MyWRKBook.ActiveSheet.Range("E3:J3").MergeCells = True
        If MyLanguage = "RUS" Then
            MyWRKBook.ActiveSheet.Range("E3") = "195027, РФ, г. Санкт-Петербург, Шаумяна проспект, дом 4, корпус 1, литер А, помещение 25Н., Tel: +7 (812) 325 2040, Fax: +7 (812) 325 2039, www.skandikagroup.ru"
        Else
            MyWRKBook.ActiveSheet.Range("E3") = "195027, Russia, St. Petersburg, Shaumyana prospect, house 4, building 1, liter А, room 25Н., Tel: +7 (812) 325 2040, Fax: +7 (812) 325 2039, www.skandikagroup.ru"
        End If

        MyWRKBook.ActiveSheet.Range("E3:J3").Font.Size = 10
        MyWRKBook.ActiveSheet.Range("B3:J3").WrapText = True
        MyWRKBook.ActiveSheet.Rows("3:3").RowHeight = 30
        MyWRKBook.ActiveSheet.Range("B3:J3").VerticalAlignment = -4108

        '-------Номер и дата заказа на закупку--------------
        MySQLStr = "SELECT CONVERT(nvarchar(30),OrderDate,103) AS OrderDate "
        MySQLStr = MySQLStr & "FROM tbl_PurchaseWorkplace_ConsolidatedOrders "
        MySQLStr = MySQLStr & "WHERE (ID = N'" & ComOrder & "')"
        InitMyConn(False)
        InitMyRec(False, MySQLStr)
        MyWRKBook.ActiveSheet.Range("B4:J4").MergeCells = True
        If Declarations.MyRec.BOF = True And Declarations.MyRec.EOF = True Then
            MyWRKBook.ActiveSheet.Range("B4") = "Purchase Order / Заказ на закупку № " & ComOrder & " от  "
            trycloseMyRec()
        Else
            MyWRKBook.ActiveSheet.Range("B4") = "Purchase Order / Заказ на закупку № " & ComOrder & " от  " & Declarations.MyRec.Fields("OrderDate").Value
            trycloseMyRec()
        End If
        MyWRKBook.ActiveSheet.Range("B4:J4").Font.Name = "Arial"
        MyWRKBook.ActiveSheet.Range("B4:J4").Font.Size = 12
        MyWRKBook.ActiveSheet.Range("B4:J4").Font.Bold = True
        MyWRKBook.ActiveSheet.Range("B4:J4").WrapText = True
        MyWRKBook.ActiveSheet.Range("B4:J4").HorizontalAlignment = -4108

        MyWRKBook.ActiveSheet.Rows("5:5").RowHeight = 5

        '-----------поставщик-----------------------------
        MyWRKBook.ActiveSheet.Range("B6:D6").MergeCells = True
        MyWRKBook.ActiveSheet.Range("B6") = "Supplier / Поставщик:"
        MyWRKBook.ActiveSheet.Range("B6:D6").Font.Size = 10
        MyWRKBook.ActiveSheet.Range("B6:D6").Font.Bold = True

        MySQLStr = "SELECT PL01002 AS SuppName, LTRIM(RTRIM(LTRIM(RTRIM(PL01003)) + ' ' + LTRIM(RTRIM(PL01004)) + ' ' + LTRIM(RTRIM(PL01005)))) AS SuppAddress "
        MySQLStr = MySQLStr & "FROM PL010300 "
        MySQLStr = MySQLStr & "WHERE (PL01001 = N'" & Declarations.MySupplierCode & "') "
        InitMyConn(False)
        InitMyRec(False, MySQLStr)
        If Declarations.MyRec.BOF = True And Declarations.MyRec.EOF = True Then
            SupplierName = ""
            SupplierAddress = ""
            trycloseMyRec()
        Else
            SupplierName = Declarations.MyRec.Fields("SuppName").Value
            SupplierAddress = Declarations.MyRec.Fields("SuppAddress").Value
            trycloseMyRec()
        End If
        MyWRKBook.ActiveSheet.Range("E6:J6").MergeCells = True
        MyWRKBook.ActiveSheet.Range("E6") = SupplierName & Chr(10) & SupplierAddress
        MyWRKBook.ActiveSheet.Range("E6:J6").Font.Size = 10
        MyWRKBook.ActiveSheet.Range("B6:J6").WrapText = True
        MyWRKBook.ActiveSheet.Rows("6:6").RowHeight = 45
        MyWRKBook.ActiveSheet.Range("B6:J6").VerticalAlignment = -4108

        MyWRKBook.ActiveSheet.Rows("7:7").RowHeight = 5

        '---------Адрес поставки--------------------------------
        MyWRKBook.ActiveSheet.Range("B8:D8").MergeCells = True
        MyWRKBook.ActiveSheet.Range("B8") = "Delivery Address / Адрес поставки"
        MyWRKBook.ActiveSheet.Range("B8:D8").Font.Size = 10
        MyWRKBook.ActiveSheet.Range("B8:D8").Font.Bold = True

        MySQLStr = "SELECT LTRIM(RTRIM(ConsignorOfGoodsName)) AS DelName, LTRIM(RTRIM(ConsignorOfGoodsAddr)) AS DelAddr "
        MySQLStr = MySQLStr & "FROM tbl_WarehouseAccessoires0300 "
        MySQLStr = MySQLStr & "WHERE (SC23001 = N'" & Declarations.MyWH & "') "
        InitMyConn(False)
        InitMyRec(False, MySQLStr)
        If Declarations.MyRec.BOF = True And Declarations.MyRec.EOF = True Then
            DelName = ""
            DelAddr = ""
            trycloseMyRec()
        Else
            DelName = Declarations.MyRec.Fields("DelName").Value
            DelAddr = Declarations.MyRec.Fields("DelAddr").Value
            trycloseMyRec()
        End If
        If MyLanguage <> "RUS" And Trim(Declarations.MyWH) = "01" Then
            DelName = ""
            DelAddr = "Marshala Blyukhera prospekt, 78-D, 195253, Saint Petersburg, Russia Tel: +7 (812)325-20-40, Fax: +7 (812)325-03-22"
        End If
        MyWRKBook.ActiveSheet.Range("E8:J8").MergeCells = True
        MyWRKBook.ActiveSheet.Range("E8") = DelName & Chr(10) & DelAddr
        MyWRKBook.ActiveSheet.Range("E8:J8").Font.Size = 10
        MyWRKBook.ActiveSheet.Range("B8:J8").WrapText = True
        MyWRKBook.ActiveSheet.Rows("8:8").RowHeight = 45
        MyWRKBook.ActiveSheet.Range("B8:J8").VerticalAlignment = -4108

        MyWRKBook.ActiveSheet.Rows("9:9").RowHeight = 5

        '---------Условия поставки------------------------------
        MyWRKBook.ActiveSheet.Range("B10:D10").MergeCells = True
        MyWRKBook.ActiveSheet.Range("B10") = "Terms Of Delivery / Условия поставки"
        MyWRKBook.ActiveSheet.Range("B10:D10").Font.Size = 7

        MySQLStr = "SELECT View_1.PL23004 "
        MySQLStr = MySQLStr & "FROM PL010300 WITH(NOLOCK) INNER JOIN "
        MySQLStr = MySQLStr & "(SELECT PL23001, PL23002, PL23003, PL23004 "
        MySQLStr = MySQLStr & "FROM PL230300 WITH(NOLOCK) "
        MySQLStr = MySQLStr & "WHERE (PL23001 = N'1') AND (PL23002 = N'" & MyLanguage & "')) AS View_1 ON CONVERT(int, PL010300.PL01029) = CONVERT(int, View_1.PL23003) "
        MySQLStr = MySQLStr & "WHERE (PL010300.PL01001 = N'" & Declarations.MySupplierCode & "')"
        InitMyConn(False)
        InitMyRec(False, MySQLStr)
        MyWRKBook.ActiveSheet.Range("E10:J10").MergeCells = True
        If Declarations.MyRec.BOF = True And Declarations.MyRec.EOF = True Then
            trycloseMyRec()
        Else
            MyWRKBook.ActiveSheet.Range("E10") = Declarations.MyRec.Fields("PL23004").Value
            trycloseMyRec()
        End If
        MyWRKBook.ActiveSheet.Range("E10:J10").Font.Size = 8
        MyWRKBook.ActiveSheet.Range("B10:J10").WrapText = True
        MyWRKBook.ActiveSheet.Range("B10:J10").VerticalAlignment = -4108

        '------Условия оплаты------------------------------
        MyWRKBook.ActiveSheet.Range("B11:D11").MergeCells = True
        MyWRKBook.ActiveSheet.Range("B11") = "Terms Of Payment / Условия оплаты"
        MyWRKBook.ActiveSheet.Range("B11:D11").Font.Size = 7

        MySQLStr = "SELECT View_1.PL23004 "
        MySQLStr = MySQLStr & "FROM PL010300 WITH(NOLOCK) INNER JOIN "
        MySQLStr = MySQLStr & "(SELECT PL23001, PL23002, PL23003, PL23004 "
        MySQLStr = MySQLStr & "FROM PL230300 WITH(NOLOCK) "
        MySQLStr = MySQLStr & "WHERE (PL23001 = N'0') AND (PL23002 = N'" & MyLanguage & "')) AS View_1 ON PL010300.PL01028 = View_1.PL23003 "
        MySQLStr = MySQLStr & "WHERE (PL010300.PL01001 = N'" & Declarations.MySupplierCode & "')"
        InitMyConn(False)
        InitMyRec(False, MySQLStr)
        MyWRKBook.ActiveSheet.Range("E11:J11").MergeCells = True
        If Declarations.MyRec.BOF = True And Declarations.MyRec.EOF = True Then
            trycloseMyRec()
        Else
            MyWRKBook.ActiveSheet.Range("E11") = Declarations.MyRec.Fields("PL23004").Value
            trycloseMyRec()
        End If
        MyWRKBook.ActiveSheet.Range("E11:J11").Font.Size = 8
        MyWRKBook.ActiveSheet.Range("B11:J11").WrapText = True
        MyWRKBook.ActiveSheet.Range("B11:J11").VerticalAlignment = -4108

        MyWRKBook.ActiveSheet.Rows("12:12").RowHeight = 5

        '-------Закупщик------------------------------------
        MyWRKBook.ActiveSheet.Range("B13:D13").MergeCells = True
        MyWRKBook.ActiveSheet.Range("B13") = "Purchaser / Закупщик:"
        MyWRKBook.ActiveSheet.Range("B13:D13").Font.Size = 10

        MySQLStr = "SELECT LTRIM(RTRIM(LTRIM(RTRIM(ISNULL(View_1.SYPD001, ''))) + ' ' + LTRIM(RTRIM(ISNULL(View_1.SYPD003, ''))))) AS Purchaser "
        MySQLStr = MySQLStr & "FROM tbl_SupplierCard0300 WITH(NOLOCK) LEFT OUTER JOIN "
        MySQLStr = MySQLStr & "(SELECT SYPD001, SYPD002, SYPD003 "
        MySQLStr = MySQLStr & "FROM SYPD0300 WITH(NOLOCK) "
        If MyLanguage = "RUS" Then
            MySQLStr = MySQLStr & "WHERE (SYPD002 = N'ENG')) "
        Else
            MySQLStr = MySQLStr & "WHERE (SYPD002 = N'RUS')) "
        End If
        MySQLStr = MySQLStr & "AS View_1 ON UPPER(tbl_SupplierCard0300.Purchaser) = UPPER(View_1.SYPD001) "
        MySQLStr = MySQLStr & "WHERE (tbl_SupplierCard0300.PL01001 = N'" & Declarations.MySupplierCode & "')"
        InitMyConn(False)
        InitMyRec(False, MySQLStr)
        MyWRKBook.ActiveSheet.Range("E13:J13").MergeCells = True
        If Declarations.MyRec.BOF = True And Declarations.MyRec.EOF = True Then
            trycloseMyRec()
        Else
            MyWRKBook.ActiveSheet.Range("E13") = Declarations.MyRec.Fields("Purchaser").Value
            trycloseMyRec()
        End If
        MyWRKBook.ActiveSheet.Range("E13:J13").Font.Size = 10
        MyWRKBook.ActiveSheet.Range("B13:J13").WrapText = True
        MyWRKBook.ActiveSheet.Range("B13:J13").VerticalAlignment = -4108

        MyWRKBook.ActiveSheet.Rows("14:14").RowHeight = 5
        '-------Заголовок таблицы---------------------------
        MySQLStr = "SELECT SYCD0100.SYCD009 "
        MySQLStr = MySQLStr & "FROM PL010300 WITH(NOLOCK) INNER JOIN "
        MySQLStr = MySQLStr & "SYCD0100 ON PL010300.PL01026 = SYCD0100.SYCD001 "
        MySQLStr = MySQLStr & "WHERE (PL010300.PL01001 = N'" & Declarations.MySupplierCode & "') "
        InitMyConn(False)
        InitMyRec(False, MySQLStr)
        If Declarations.MyRec.BOF = True And Declarations.MyRec.EOF = True Then
            CurrName = "RUB"
            trycloseMyRec()
        Else
            CurrName = Declarations.MyRec.Fields("SYCD009").Value
            trycloseMyRec()
        End If
        MyWRKBook.ActiveSheet.Range("B15") = "N п/п"
        MyWRKBook.ActiveSheet.Range("C15") = "Supp Item Code / Код товара поставщ"
        MyWRKBook.ActiveSheet.Range("D15") = "Item Code / Код товара"
        MyWRKBook.ActiveSheet.Range("E15") = "Item Name / Наименование товара"
        MyWRKBook.ActiveSheet.Range("F15") = "UOM / Ед. измер-я"
        MyWRKBook.ActiveSheet.Range("G15") = "QTY / Количество"
        MyWRKBook.ActiveSheet.Range("H15") = "Price / Цена заказа, " & CurrName
        MyWRKBook.ActiveSheet.Range("I15") = "Price / Цена поставщика"
        MyWRKBook.ActiveSheet.Range("J15") = "Summa / Сумма, " & CurrName
        MyWRKBook.ActiveSheet.Range("B15:J15").Font.Size = 7
        MyWRKBook.ActiveSheet.Range("B15:J15").WrapText = True
        MyWRKBook.ActiveSheet.Range("B15:J15").VerticalAlignment = -4108
        MyWRKBook.ActiveSheet.Range("B15:J15").HorizontalAlignment = -4108
        MyWRKBook.ActiveSheet.Rows("15:15").RowHeight = 40
        MyWRKBook.ActiveSheet.Range("B15:J15").Select()
        MyWRKBook.ActiveSheet.Range("B15:J15").Borders(5).LineStyle = -4142
        MyWRKBook.ActiveSheet.Range("B15:J15").Borders(6).LineStyle = -4142
        With MyWRKBook.ActiveSheet.Range("B15:J15").Borders(7)
            .LineStyle = 1
            .Weight = 4
            .ColorIndex = -4105
        End With
        With MyWRKBook.ActiveSheet.Range("B15:J15").Borders(8)
            .LineStyle = 1
            .Weight = 4
            .ColorIndex = -4105
        End With
        With MyWRKBook.ActiveSheet.Range("B15:J15").Borders(9)
            .LineStyle = 1
            .Weight = 4
            .ColorIndex = -4105
        End With
        With MyWRKBook.ActiveSheet.Range("B15:J15").Borders(10)
            .LineStyle = 1
            .Weight = 4
            .ColorIndex = -4105
        End With
        With MyWRKBook.ActiveSheet.Range("B15:J15").Borders(11)
            .LineStyle = 1
            .Weight = 4
            .ColorIndex = -4105
        End With
        With MyWRKBook.ActiveSheet.Range("B15:J15").Interior
            .ColorIndex = 36
            .Pattern = 1
            .PatternColorIndex = -4105
        End With


        MyWRKBook.ActiveSheet.Columns("A:A").ColumnWidth = 1
        MyWRKBook.ActiveSheet.Columns("B:B").ColumnWidth = 3
        MyWRKBook.ActiveSheet.Columns("C:C").ColumnWidth = 8
        MyWRKBook.ActiveSheet.Columns("D:D").ColumnWidth = 8
        MyWRKBook.ActiveSheet.Columns("E:E").ColumnWidth = 27
        MyWRKBook.ActiveSheet.Columns("F:F").ColumnWidth = 4
        MyWRKBook.ActiveSheet.Columns("G:G").ColumnWidth = 4
        MyWRKBook.ActiveSheet.Columns("H:H").ColumnWidth = 7
        MyWRKBook.ActiveSheet.Columns("I:I").ColumnWidth = 7
        MyWRKBook.ActiveSheet.Columns("J:J").ColumnWidth = 12
        i = 16
    End Sub

    Private Sub ExportOrderHeaderToLO(ByRef oSheet As Object, ByRef oServiceManager As Object, ByRef oWorkBook As Object, _
        ByRef oDispatcher As Object, ByVal ComOrder As String, ByRef i As Integer)
        '////////////////////////////////////////////////////////////////////////////////
        '//
        '// Процедура выгрузки заголовка обобщенного заказа в Libre Office
        '//
        '////////////////////////////////////////////////////////////////////////////////
        Dim oFrame As Object
        Dim MySQLStr As String
        Dim SupplierName As String
        Dim SupplierAddress As String
        Dim DelName As String
        Dim DelAddr As String
        Dim CurrName As String
        Dim MyLanguage As String

        oFrame = oWorkBook.getCurrentController.getFrame

        '-----Ширина колонок
        oSheet.getColumns().getByName("A").Width = 400
        oSheet.getColumns().getByName("B").Width = 600
        oSheet.getColumns().getByName("C").Width = 1800
        oSheet.getColumns().getByName("D").Width = 1800
        oSheet.getColumns().getByName("E").Width = 6000
        oSheet.getColumns().getByName("F").Width = 1000
        oSheet.getColumns().getByName("G").Width = 1000
        oSheet.getColumns().getByName("H").Width = 1600
        oSheet.getColumns().getByName("I").Width = 1600
        oSheet.getColumns().getByName("J").Width = 2000

        '------Язык документа---------------------
        MySQLStr = "SELECT PL01027 AS CC "
        MySQLStr = MySQLStr & "FROM PL010300 "
        MySQLStr = MySQLStr & "WHERE (PL01001 = N'" & Trim(Label9.Text) & "') "
        InitMyConn(False)
        InitMyRec(False, MySQLStr)
        If Declarations.MyRec.BOF = True And Declarations.MyRec.EOF = True Then
            MyLanguage = "RUS"
            trycloseMyRec()
        Else
            MyLanguage = Declarations.MyRec.Fields("CC").Value
            trycloseMyRec()
        End If

        LOMergeCells(oServiceManager, oDispatcher, oFrame, "B2:I2")
        If MyLanguage = "RUS" Then
            oSheet.getCellRangeByName("B2").String = "ООО ""Скандика"""
        Else
            oSheet.getCellRangeByName("B2").String = "Skandika LLC"
        End If
        LOFontSetFamilyName(oServiceManager, oDispatcher, oFrame, "B2", "Arial")
        LOFontSetBold(oServiceManager, oDispatcher, oFrame, "B2")
        LOFontSetSize(oServiceManager, oDispatcher, oFrame, "B2", 12)
        oSheet.getCellRangeByName("B2").VertJustify = 2

        LOMergeCells(oServiceManager, oDispatcher, oFrame, "B3:D3")
        oSheet.getCellRangeByName("B3").String = "Address/Адрес:"
        LOFontSetFamilyName(oServiceManager, oDispatcher, oFrame, "B3:D3", "Calibri")
        LOFontSetBold(oServiceManager, oDispatcher, oFrame, "B3:D3")
        LOFontSetSize(oServiceManager, oDispatcher, oFrame, "B3:D3", 10)
        LOWrapText(oServiceManager, oDispatcher, oFrame, "B3:D3")
        oSheet.getCellRangeByName("B3:D3").VertJustify = 2

        LOMergeCells(oServiceManager, oDispatcher, oFrame, "E3:I3")
        If MyLanguage = "RUS" Then
            oSheet.getCellRangeByName("E3").String = "195027, РФ, г. Санкт-Петербург, Шаумяна проспект, дом 4, корпус 1, литер А, помещение 25Н., Tel: +7 (812) 325 2040, Fax: +7 (812) 325 2039, www.skandikagroup.ru"
        Else
            oSheet.getCellRangeByName("E3").String = "195027, Russia, St. Petersburg, Shaumyana prospect, house 4, building 1, liter А, room 25Н., Tel: +7 (812) 325 2040, Fax: +7 (812) 325 2039, www.skandikagroup.ru"
        End If
        LOFontSetFamilyName(oServiceManager, oDispatcher, oFrame, "E3:I3", "Calibri")
        LOFontSetSize(oServiceManager, oDispatcher, oFrame, "E3:I3", 10)
        LOWrapText(oServiceManager, oDispatcher, oFrame, "E3:I3")
        oSheet.getCellRangeByName("E3:I3").VertJustify = 2

        '-------Номер и дата заказа на закупку--------------
        MySQLStr = "SELECT CONVERT(nvarchar(30),OrderDate,103) AS OrderDate, SupplierCode, WH "
        MySQLStr = MySQLStr & "FROM tbl_PurchaseWorkplace_ConsolidatedOrders WITH(NOLOCK) "
        MySQLStr = MySQLStr & "WHERE (ID = N'" & ComOrder & "')"
        InitMyConn(False)
        InitMyRec(False, MySQLStr)
        LOMergeCells(oServiceManager, oDispatcher, oFrame, "B4:I4")
        If Declarations.MyRec.BOF = True And Declarations.MyRec.EOF = True Then
            Declarations.MySupplierCode = ""
            Declarations.MyWH = ""
            oSheet.getCellRangeByName("B4").String = "Purchase Order / Заказ на закупку № " & ComOrder
            trycloseMyRec()
        Else
            Declarations.MySupplierCode = Declarations.MyRec.Fields("SupplierCode").Value
            Declarations.MyWH = Declarations.MyRec.Fields("WH").Value
            oSheet.getCellRangeByName("B4").String = "Purchase Order / Заказ на закупку № " & ComOrder & " от  " & Declarations.MyRec.Fields("OrderDate").Value
        End If
        LOFontSetFamilyName(oServiceManager, oDispatcher, oFrame, "B4:I4", "Arial")
        LOFontSetBold(oServiceManager, oDispatcher, oFrame, "B4:I4")
        LOFontSetSize(oServiceManager, oDispatcher, oFrame, "B4:I4", 12)
        LOWrapText(oServiceManager, oDispatcher, oFrame, "B4:I4")
        oSheet.getCellRangeByName("B4:I4").VertJustify = 2
        oSheet.getCellRangeByName("A5").Rows.Height = 200

        '-----------поставщик-----------------------------
        LOMergeCells(oServiceManager, oDispatcher, oFrame, "B6:D6")
        oSheet.getCellRangeByName("B6").String = "Supplier / Поставщик:"
        LOFontSetFamilyName(oServiceManager, oDispatcher, oFrame, "B6:D6", "Calibri")
        LOFontSetBold(oServiceManager, oDispatcher, oFrame, "B6:D6")
        LOFontSetSize(oServiceManager, oDispatcher, oFrame, "B6:D6", 10)
        LOWrapText(oServiceManager, oDispatcher, oFrame, "B6:D6")
        oSheet.getCellRangeByName("B6:D6").VertJustify = 2

        MySQLStr = "SELECT PL01002 AS SuppName, LTRIM(RTRIM(LTRIM(RTRIM(PL01003)) + ' ' + LTRIM(RTRIM(PL01004)) + ' ' + LTRIM(RTRIM(PL01005)))) AS SuppAddress "
        MySQLStr = MySQLStr & "FROM PL010300 WITH(NOLOCK) "
        MySQLStr = MySQLStr & "WHERE (PL01001 = N'" & Declarations.MySupplierCode & "') "
        InitMyConn(False)
        InitMyRec(False, MySQLStr)
        If Declarations.MyRec.BOF = True And Declarations.MyRec.EOF = True Then
            SupplierName = ""
            SupplierAddress = ""
            trycloseMyRec()
        Else
            SupplierName = Declarations.MyRec.Fields("SuppName").Value
            SupplierAddress = Declarations.MyRec.Fields("SuppAddress").Value
            trycloseMyRec()
        End If
        LOMergeCells(oServiceManager, oDispatcher, oFrame, "E6:I6")
        oSheet.getCellRangeByName("E6").String = SupplierName & Chr(10) & SupplierAddress
        LOFontSetFamilyName(oServiceManager, oDispatcher, oFrame, "E6:I6", "Calibri")
        LOFontSetSize(oServiceManager, oDispatcher, oFrame, "E6:I6", 10)
        LOWrapText(oServiceManager, oDispatcher, oFrame, "E6:I6")
        oSheet.getCellRangeByName("E6:I6").VertJustify = 2
        oSheet.getCellRangeByName("A7").Rows.Height = 200

        '---------Адрес поставки--------------------------------
        LOMergeCells(oServiceManager, oDispatcher, oFrame, "B8:D8")
        oSheet.getCellRangeByName("B8").String = "Delivery Address / Адрес поставки"
        LOFontSetFamilyName(oServiceManager, oDispatcher, oFrame, "B8:D8", "Calibri")
        LOFontSetBold(oServiceManager, oDispatcher, oFrame, "B8:D8")
        LOFontSetSize(oServiceManager, oDispatcher, oFrame, "B8:D8", 10)
        LOWrapText(oServiceManager, oDispatcher, oFrame, "B8:D8")
        oSheet.getCellRangeByName("B8:D8").VertJustify = 2

        MySQLStr = "SELECT LTRIM(RTRIM(ConsignorOfGoodsName)) AS DelName, LTRIM(RTRIM(ConsignorOfGoodsAddr)) AS DelAddr "
        MySQLStr = MySQLStr & "FROM tbl_WarehouseAccessoires0300 WITH(NOLOCK) "
        MySQLStr = MySQLStr & "WHERE (SC23001 = N'" & Declarations.MyWH & "') "
        InitMyConn(False)
        InitMyRec(False, MySQLStr)
        If Declarations.MyRec.BOF = True And Declarations.MyRec.EOF = True Then
            DelName = ""
            DelAddr = ""
            trycloseMyRec()
        Else
            DelName = Declarations.MyRec.Fields("DelName").Value
            DelAddr = Declarations.MyRec.Fields("DelAddr").Value
            trycloseMyRec()
        End If
        If MyLanguage <> "RUS" And Trim(Declarations.MyWH) = "01" Then
            DelName = ""
            DelAddr = "Marshala Blyukhera prospekt, 78-D, 195253, Saint Petersburg, Russia Tel: +7 (812)325-20-40, Fax: +7 (812)325-03-22"
        End If
        LOMergeCells(oServiceManager, oDispatcher, oFrame, "E8:I8")
        oSheet.getCellRangeByName("E8").String = DelName & Chr(10) & DelAddr
        LOFontSetFamilyName(oServiceManager, oDispatcher, oFrame, "E8:I8", "Calibri")
        LOFontSetSize(oServiceManager, oDispatcher, oFrame, "E8:I8", 10)
        LOWrapText(oServiceManager, oDispatcher, oFrame, "E8:I8")
        oSheet.getCellRangeByName("E8:I8").VertJustify = 2
        oSheet.getCellRangeByName("A9").Rows.Height = 200

        '---------Условия поставки------------------------------
        LOMergeCells(oServiceManager, oDispatcher, oFrame, "B10:D10")
        oSheet.getCellRangeByName("B10").String = "Terms Of Delivery / Условия поставки"
        LOFontSetFamilyName(oServiceManager, oDispatcher, oFrame, "B10:D10", "Calibri")
        LOFontSetSize(oServiceManager, oDispatcher, oFrame, "B10:D10", 7)
        LOWrapText(oServiceManager, oDispatcher, oFrame, "B10:D10")
        oSheet.getCellRangeByName("B10:D10").VertJustify = 2

        MySQLStr = "SELECT View_1.PL23004 "
        MySQLStr = MySQLStr & "FROM PL010300 WITH(NOLOCK) INNER JOIN "
        MySQLStr = MySQLStr & "(SELECT PL23001, PL23002, PL23003, PL23004 "
        MySQLStr = MySQLStr & "FROM PL230300 WITH(NOLOCK) "
        MySQLStr = MySQLStr & "WHERE (PL23001 = N'1') AND (PL23002 = N'" & MyLanguage & "')) AS View_1 ON CONVERT(int, PL010300.PL01029) = CONVERT(int, View_1.PL23003) "
        MySQLStr = MySQLStr & "WHERE (PL010300.PL01001 = N'" & Declarations.MySupplierCode & "')"
        InitMyConn(False)
        InitMyRec(False, MySQLStr)
        LOMergeCells(oServiceManager, oDispatcher, oFrame, "E10:I10")
        If Declarations.MyRec.BOF = True And Declarations.MyRec.EOF = True Then
            trycloseMyRec()
        Else
            oSheet.getCellRangeByName("E10").String = Declarations.MyRec.Fields("PL23004").Value
            trycloseMyRec()
        End If
        LOFontSetFamilyName(oServiceManager, oDispatcher, oFrame, "E10:I10", "Calibri")
        LOFontSetSize(oServiceManager, oDispatcher, oFrame, "E10:I10", 7)
        LOWrapText(oServiceManager, oDispatcher, oFrame, "E10:I10")
        oSheet.getCellRangeByName("E10:I10").VertJustify = 2

        '------Условия оплаты------------------------------
        LOMergeCells(oServiceManager, oDispatcher, oFrame, "B11:D11")
        oSheet.getCellRangeByName("B11").String = "Terms Of Payment / Условия оплаты"
        LOFontSetFamilyName(oServiceManager, oDispatcher, oFrame, "B11:D11", "Calibri")
        LOFontSetSize(oServiceManager, oDispatcher, oFrame, "B11:D11", 7)
        LOWrapText(oServiceManager, oDispatcher, oFrame, "B11:D11")
        oSheet.getCellRangeByName("B11:D11").VertJustify = 2

        MySQLStr = "SELECT View_1.PL23004 "
        MySQLStr = MySQLStr & "FROM PL010300 WITH(NOLOCK) INNER JOIN "
        MySQLStr = MySQLStr & "(SELECT PL23001, PL23002, PL23003, PL23004 "
        MySQLStr = MySQLStr & "FROM PL230300 WITH(NOLOCK) "
        MySQLStr = MySQLStr & "WHERE (PL23001 = N'0') AND (PL23002 = N'" & MyLanguage & "')) AS View_1 ON PL010300.PL01028 = View_1.PL23003 "
        MySQLStr = MySQLStr & "WHERE (PL010300.PL01001 = N'" & Declarations.MySupplierCode & "')"
        InitMyConn(False)
        InitMyRec(False, MySQLStr)
        LOMergeCells(oServiceManager, oDispatcher, oFrame, "E11:I11")
        If Declarations.MyRec.BOF = True And Declarations.MyRec.EOF = True Then
            trycloseMyRec()
        Else
            oSheet.getCellRangeByName("E11").String = Declarations.MyRec.Fields("PL23004").Value
            trycloseMyRec()
        End If
        LOFontSetFamilyName(oServiceManager, oDispatcher, oFrame, "E11:I11", "Calibri")
        LOFontSetSize(oServiceManager, oDispatcher, oFrame, "E11:I11", 7)
        LOWrapText(oServiceManager, oDispatcher, oFrame, "E11:I11")
        oSheet.getCellRangeByName("E11:I11").VertJustify = 2
        oSheet.getCellRangeByName("A12").Rows.Height = 200

        '-------Закупщик------------------------------------
        LOMergeCells(oServiceManager, oDispatcher, oFrame, "B13:D13")
        oSheet.getCellRangeByName("B13").String = "Purchaser / Закупщик:"
        LOFontSetFamilyName(oServiceManager, oDispatcher, oFrame, "B13:D13", "Calibri")
        LOFontSetSize(oServiceManager, oDispatcher, oFrame, "B13:D13", 10)
        LOWrapText(oServiceManager, oDispatcher, oFrame, "B13:D13")
        oSheet.getCellRangeByName("B13:D13").VertJustify = 2

        MySQLStr = "SELECT LTRIM(RTRIM(LTRIM(RTRIM(ISNULL(View_1.SYPD001, ''))) + ' ' + LTRIM(RTRIM(ISNULL(View_1.SYPD003, ''))))) AS Purchaser "
        MySQLStr = MySQLStr & "FROM tbl_SupplierCard0300 WITH(NOLOCK) LEFT OUTER JOIN "
        MySQLStr = MySQLStr & "(SELECT SYPD001, SYPD002, SYPD003 "
        MySQLStr = MySQLStr & "FROM SYPD0300 WITH(NOLOCK) "
        MySQLStr = MySQLStr & "WHERE (SYPD002 = N'ENG')) AS View_1 ON UPPER(tbl_SupplierCard0300.Purchaser) = UPPER(View_1.SYPD001) "
        MySQLStr = MySQLStr & "WHERE (tbl_SupplierCard0300.PL01001 = N'" & Declarations.MySupplierCode & "')"
        InitMyConn(False)
        InitMyRec(False, MySQLStr)
        LOMergeCells(oServiceManager, oDispatcher, oFrame, "E13:I13")
        If Declarations.MyRec.BOF = True And Declarations.MyRec.EOF = True Then
            trycloseMyRec()
        Else
            oSheet.getCellRangeByName("E13").String = Declarations.MyRec.Fields("Purchaser").Value
            trycloseMyRec()
        End If
        LOFontSetFamilyName(oServiceManager, oDispatcher, oFrame, "E13:I13", "Calibri")
        LOFontSetSize(oServiceManager, oDispatcher, oFrame, "E13:I13", 10)
        LOWrapText(oServiceManager, oDispatcher, oFrame, "E13:I13")
        oSheet.getCellRangeByName("E13:I13").VertJustify = 2
        oSheet.getCellRangeByName("A14").Rows.Height = 200

        '-------Заголовок таблицы---------------------------
        MySQLStr = "SELECT SYCD0100.SYCD009 "
        MySQLStr = MySQLStr & "FROM PL010300 WITH(NOLOCK) INNER JOIN "
        MySQLStr = MySQLStr & "SYCD0100 ON PL010300.PL01026 = SYCD0100.SYCD001 "
        MySQLStr = MySQLStr & "WHERE (PL010300.PL01001 = N'" & Declarations.MySupplierCode & "') "
        InitMyConn(False)
        InitMyRec(False, MySQLStr)
        If Declarations.MyRec.BOF = True And Declarations.MyRec.EOF = True Then
            CurrName = "RUB"
            trycloseMyRec()
        Else
            CurrName = Declarations.MyRec.Fields("SYCD009").Value
            trycloseMyRec()
        End If

        oSheet.getCellRangeByName("B15").String = "N п/п"
        oSheet.getCellRangeByName("C15").String = "Supp Item Code / Код товара поставщ"
        oSheet.getCellRangeByName("D15").String = "Item Code / Код товара"
        oSheet.getCellRangeByName("E15").String = "Item Name / Наименование товара"
        oSheet.getCellRangeByName("F15").String = "UOM / Ед. измер-я"
        oSheet.getCellRangeByName("G15").String = "QTY / Количество"
        oSheet.getCellRangeByName("H15").String = "Price / Цена заказа, " & CurrName
        oSheet.getCellRangeByName("I15").String = "Price / Цена поставщика"
        oSheet.getCellRangeByName("J15").String = "Summa / Сумма, " & CurrName

        LOFontSetFamilyName(oServiceManager, oDispatcher, oFrame, "B15:J15", "Calibri")
        LOFontSetSize(oServiceManager, oDispatcher, oFrame, "B15:J15", 7)
        oSheet.getCellRangeByName("B15:J15").CellBackColor = 16775598
        Dim LineFormat As Object
        LineFormat = oServiceManager.Bridge_GetStruct("com.sun.star.table.BorderLine2")
        LineFormat.LineStyle = 0
        LineFormat.LineWidth = 70
        oSheet.getCellRangeByName("B15:J15").TopBorder = LineFormat
        oSheet.getCellRangeByName("B15:J15").RightBorder = LineFormat
        oSheet.getCellRangeByName("B15:J15").LeftBorder = LineFormat
        oSheet.getCellRangeByName("B15:J15").BottomBorder = LineFormat
        oSheet.getCellRangeByName("B15:J15").VertJustify = 2
        oSheet.getCellRangeByName("B15:J15").HoriJustify = 2
        LOWrapText(oServiceManager, oDispatcher, oFrame, "B15:J15")

        i = 16
    End Sub

    Private Sub ExportOrderBodyToExcel(ByRef MyWRKBook As Object, ByVal ComOrder As String, ByRef i As Integer)
        '////////////////////////////////////////////////////////////////////////////////
        '//
        '// Процедура выгрузки тела обобщенного заказа в Excel
        '//
        '////////////////////////////////////////////////////////////////////////////////
        Dim MySQLStr As String
        Dim j As Integer                        'счетчик строк

        j = i
        MySQLStr = "IF exists(select * from tempdb..sysobjects where "
        MySQLStr = MySQLStr & "id = object_id(N'tempdb..#_MyPCOrder')  "
        MySQLStr = MySQLStr & "and xtype = N'U') "
        MySQLStr = MySQLStr & "DROP TABLE #_MyPCOrder "
        InitMyConn(False)
        Declarations.MyConn.Execute(MySQLStr)

        MySQLStr = "CREATE TABLE #_MyPCOrder( "
        MySQLStr = MySQLStr & "[PC03005] [nvarchar](35) COLLATE Cyrillic_General_BIN, "
        MySQLStr = MySQLStr & "[SC01060] [nvarchar](35) COLLATE Cyrillic_General_BIN, "
        MySQLStr = MySQLStr & "[PC03006] [nvarchar] (52) COLLATE Cyrillic_General_BIN, "
        MySQLStr = MySQLStr & "[QTY] [numeric](20, 8), "
        MySQLStr = MySQLStr & "[PC03009] [int], "
        MySQLStr = MySQLStr & "[PC03009_Name][nvarchar](10), "
        MySQLStr = MySQLStr & "[Price] [numeric](28, 8), "
        MySQLStr = MySQLStr & "[EUPrice] [numeric](28, 8), "
        MySQLStr = MySQLStr & "[StrSum] [numeric](28, 8), "
        MySQLStr = MySQLStr & "[OrderSum] [numeric](28, 8), "
        MySQLStr = MySQLStr & "[MinQTY] [numeric](28, 8), "
        MySQLStr = MySQLStr & "[IsWH] [nvarchar](25) "
        MySQLStr = MySQLStr & ") "
        InitMyConn(False)
        Declarations.MyConn.Execute(MySQLStr)

        MySQLStr = "EXEC spp_PurchaseWorkplace_PurchaseGroupOrderPreparation N'" & ComOrder & "' "
        InitMyConn(False)
        Declarations.MyConn.Execute(MySQLStr)

        MySQLStr = "SELECT * "
        MySQLStr = MySQLStr & "FROM #_MyPCOrder "
        MySQLStr = MySQLStr & "Order BY SC01060 "
        InitMyConn(False)
        InitMyRec(False, MySQLStr)
        If Declarations.MyRec.BOF = True And Declarations.MyRec.EOF = True Then
            trycloseMyRec()
        Else
            Declarations.MyRec.MoveFirst()
            MyPurchOrderSum = Declarations.MyRec.Fields("OrderSum").Value
            While Declarations.MyRec.EOF = False
                '-----N п/п
                MyWRKBook.ActiveSheet.Range("B" & CStr(i)) = i - j + 1
                '-----код товара
                MyWRKBook.ActiveSheet.Range("C" & CStr(i)).NumberFormat = "@"
                MyWRKBook.ActiveSheet.Range("C" & CStr(i)) = Declarations.MyRec.Fields("SC01060").Value
                '-----код товара поставщика
                MyWRKBook.ActiveSheet.Range("D" & CStr(i)).NumberFormat = "@"
                MyWRKBook.ActiveSheet.Range("D" & CStr(i)) = Declarations.MyRec.Fields("PC03005").Value
                '-----название товара
                MyWRKBook.ActiveSheet.Range("E" & CStr(i)).NumberFormat = "@"
                MyWRKBook.ActiveSheet.Range("E" & CStr(i)) = Declarations.MyRec.Fields("PC03006").Value
                '-----единица измерения товара
                MyWRKBook.ActiveSheet.Range("F" & CStr(i)).NumberFormat = "@"
                MyWRKBook.ActiveSheet.Range("F" & CStr(i)) = Declarations.MyRec.Fields("PC03009_Name").Value
                '-----количество товара
                MyWRKBook.ActiveSheet.Range("G" & CStr(i)) = Declarations.MyRec.Fields("QTY").Value
                '-----цена
                MyWRKBook.ActiveSheet.Range("H" & CStr(i)) = Declarations.MyRec.Fields("Price").Value
                '-----цена в евро(если есть)
                If Declarations.MyRec.Fields("EUPrice").Value = 0 Then
                Else
                    MyWRKBook.ActiveSheet.Range("I" & CStr(i)) = Declarations.MyRec.Fields("EUPrice").Value
                End If
                '-----сумма строки
                MyWRKBook.ActiveSheet.Range("J" & CStr(i)) = Declarations.MyRec.Fields("StrSum").Value


                Declarations.MyRec.MoveNext()
                i = i + 1
            End While
            trycloseMyRec()

            MySQLStr = "IF exists(select * from tempdb..sysobjects where "
            MySQLStr = MySQLStr & "id = object_id(N'tempdb..#_MyPCOrder')  "
            MySQLStr = MySQLStr & "and xtype = N'U') "
            MySQLStr = MySQLStr & "DROP TABLE #_MyPCOrder "
            InitMyConn(False)
            Declarations.MyConn.Execute(MySQLStr)

            MyWRKBook.ActiveSheet.Range("B" & CStr(j) & ":J" & CStr(i - 1)).Select()
            MyWRKBook.ActiveSheet.Range("B" & CStr(j) & ":J" & CStr(i - 1)).Borders(5).LineStyle = -4142
            MyWRKBook.ActiveSheet.Range("B" & CStr(j) & ":J" & CStr(i - 1)).Borders(6).LineStyle = -4142
            With MyWRKBook.ActiveSheet.Range("B" & CStr(j) & ":J" & CStr(i - 1)).Borders(7)
                .LineStyle = 1
                .Weight = 2
                .ColorIndex = -4105
            End With
            With MyWRKBook.ActiveSheet.Range("B" & CStr(j) & ":J" & CStr(i - 1)).Borders(8)
                .LineStyle = 1
                .Weight = 4
                .ColorIndex = -4105
            End With
            With MyWRKBook.ActiveSheet.Range("B" & CStr(j) & ":J" & CStr(i - 1)).Borders(9)
                .LineStyle = 1
                .Weight = 2
                .ColorIndex = -4105
            End With
            With MyWRKBook.ActiveSheet.Range("B" & CStr(j) & ":J" & CStr(i - 1)).Borders(10)
                .LineStyle = 1
                .Weight = 2
                .ColorIndex = -4105
            End With
            With MyWRKBook.ActiveSheet.Range("B" & CStr(j) & ":J" & CStr(i - 1)).Borders(11)
                .LineStyle = 1
                .Weight = 2
                .ColorIndex = -4105
            End With
            With MyWRKBook.ActiveSheet.Range("B" & CStr(j) & ":J" & CStr(i - 1)).Borders(12)
                .LineStyle = 1
                .Weight = 2
                .ColorIndex = -4105
            End With

            MyWRKBook.ActiveSheet.Range("B" & CStr(j) & ":J" & CStr(i - 1)).Select()
            With MyWRKBook.ActiveSheet.Range("B" & CStr(j) & ":J" & CStr(i - 1)).Font
                .Name = "Arial"
                .Size = 7
            End With

            MyWRKBook.ActiveSheet.Range("C" & CStr(j) & ":D" & CStr(i - 1)).Font.Bold = True
            MyWRKBook.ActiveSheet.Range("B" & CStr(j) & ":J" & CStr(i - 1)).WrapText = True

        End If
    End Sub

    Private Sub ExportOrderBodyToLO(ByRef oSheet As Object, ByRef oServiceManager As Object, ByRef oWorkBook As Object, _
        ByRef oDispatcher As Object, ByVal ComOrder As String, ByRef i As Integer)
        '////////////////////////////////////////////////////////////////////////////////
        '//
        '// Процедура выгрузки тела обобщенного заказа в Libre Office
        '//
        '////////////////////////////////////////////////////////////////////////////////
        Dim oFrame As Object
        Dim MySQLStr As String
        Dim j As Integer                        'счетчик строк

        oFrame = oWorkBook.getCurrentController.getFrame
        j = i
        MySQLStr = "IF exists(select * from tempdb..sysobjects where "
        MySQLStr = MySQLStr & "id = object_id(N'tempdb..#_MyPCOrder')  "
        MySQLStr = MySQLStr & "and xtype = N'U') "
        MySQLStr = MySQLStr & "DROP TABLE #_MyPCOrder "
        InitMyConn(False)
        Declarations.MyConn.Execute(MySQLStr)

        MySQLStr = "CREATE TABLE #_MyPCOrder( "
        MySQLStr = MySQLStr & "[PC03005] [nvarchar](35) COLLATE Cyrillic_General_BIN, "
        MySQLStr = MySQLStr & "[SC01060] [nvarchar](35) COLLATE Cyrillic_General_BIN, "
        MySQLStr = MySQLStr & "[PC03006] [nvarchar] (52) COLLATE Cyrillic_General_BIN, "
        MySQLStr = MySQLStr & "[QTY] [numeric](20, 8), "
        MySQLStr = MySQLStr & "[PC03009] [int], "
        MySQLStr = MySQLStr & "[PC03009_Name][nvarchar](10), "
        MySQLStr = MySQLStr & "[Price] [numeric](28, 8), "
        MySQLStr = MySQLStr & "[EUPrice] [numeric](28, 8), "
        MySQLStr = MySQLStr & "[StrSum] [numeric](28, 8), "
        MySQLStr = MySQLStr & "[OrderSum] [numeric](28, 8), "
        MySQLStr = MySQLStr & "[MinQTY] [numeric](28, 8), "
        MySQLStr = MySQLStr & "[IsWH] [nvarchar](25) "
        MySQLStr = MySQLStr & ") "
        InitMyConn(False)
        Declarations.MyConn.Execute(MySQLStr)

        MySQLStr = "EXEC spp_PurchaseWorkplace_PurchaseGroupOrderPreparation N'" & ComOrder & "' "
        InitMyConn(False)
        Declarations.MyConn.Execute(MySQLStr)

        MySQLStr = "SELECT * "
        MySQLStr = MySQLStr & "FROM #_MyPCOrder "
        MySQLStr = MySQLStr & "Order BY SC01060 "
        InitMyConn(False)
        InitMyRec(False, MySQLStr)
        If Declarations.MyRec.BOF = True And Declarations.MyRec.EOF = True Then
            trycloseMyRec()
        Else
            Declarations.MyRec.MoveFirst()
            MyPurchOrderSum = Declarations.MyRec.Fields("OrderSum").Value
            While Declarations.MyRec.EOF = False
                '-----N п/п
                oSheet.getCellRangeByName("B" & CStr(i)).value = i - j + 1
                '-----код товара
                oSheet.getCellRangeByName("C" & CStr(i)).String = Declarations.MyRec.Fields("SC01060").Value
                '-----код товара поставщика
                oSheet.getCellRangeByName("D" & CStr(i)).String = Declarations.MyRec.Fields("PC03005").Value
                '-----название товара
                oSheet.getCellRangeByName("E" & CStr(i)).String = Declarations.MyRec.Fields("PC03006").Value
                '-----единица измерения товара
                oSheet.getCellRangeByName("F" & CStr(i)).String = Declarations.MyRec.Fields("PC03009_Name").Value
                '-----количество товара
                oSheet.getCellRangeByName("G" & CStr(i)).value = Declarations.MyRec.Fields("QTY").Value
                '-----цена
                oSheet.getCellRangeByName("H" & CStr(i)).value = Declarations.MyRec.Fields("Price").Value
                '-----цена в евро(если есть)
                If Not Declarations.MyRec.Fields("EUPrice").Value = 0 Then
                    oSheet.getCellRangeByName("I" & CStr(i)).value = Declarations.MyRec.Fields("EUPrice").Value
                End If
                '-----сумма строки
                oSheet.getCellRangeByName("J" & CStr(i)).value = Declarations.MyRec.Fields("StrSum").Value
                Declarations.MyRec.MoveNext()
                i = i + 1
            End While

            MySQLStr = "IF exists(select * from tempdb..sysobjects where "
            MySQLStr = MySQLStr & "id = object_id(N'tempdb..#_MyPCOrder')  "
            MySQLStr = MySQLStr & "and xtype = N'U') "
            MySQLStr = MySQLStr & "DROP TABLE #_MyPCOrder "
            InitMyConn(False)
            Declarations.MyConn.Execute(MySQLStr)

            LOFontSetFamilyName(oServiceManager, oDispatcher, oFrame, "B" & CStr(j) & ":J" & CStr(i - 1), "Arial")
            LOFontSetSize(oServiceManager, oDispatcher, oFrame, "B" & CStr(j) & ":J" & CStr(i - 1), 7)
            Dim LineFormat As Object
            LineFormat = oServiceManager.Bridge_GetStruct("com.sun.star.table.BorderLine2")
            LineFormat.LineStyle = 0
            LineFormat.LineWidth = 20
            oSheet.getCellRangeByName("B" & CStr(j) & ":J" & CStr(i - 1)).TopBorder = LineFormat
            oSheet.getCellRangeByName("B" & CStr(j) & ":J" & CStr(i - 1)).RightBorder = LineFormat
            oSheet.getCellRangeByName("B" & CStr(j) & ":J" & CStr(i - 1)).LeftBorder = LineFormat
            oSheet.getCellRangeByName("B" & CStr(j) & ":J" & CStr(i - 1)).BottomBorder = LineFormat
            LOWrapText(oServiceManager, oDispatcher, oFrame, "B" & CStr(j) & ":J" & CStr(i - 1))
            LOFontSetBold(oServiceManager, oDispatcher, oFrame, "C" & CStr(j) & ":D" & CStr(i - 1))
            LOFormatCells(oServiceManager, oDispatcher, oFrame, "I" & CStr(j) & ":J" & CStr(i - 1), 4)
        End If
    End Sub

    Private Sub ExportOrderFooterToExcel(ByRef MyWRKBook As Object, ByVal ComOrder As String, ByRef i As Integer)
        '////////////////////////////////////////////////////////////////////////////////
        '//
        '// Процедура выгрузки подвала обобщенного заказа в Excel
        '//
        '////////////////////////////////////////////////////////////////////////////////
        Dim MySQLStr As String
        Dim CurrName As String
        Dim PurchaserName As String
        Dim j As Integer
        Dim MyLanguage As String

        '------Язык документа---------------------
        MySQLStr = "SELECT PL01027 AS CC "
        MySQLStr = MySQLStr & "FROM PL010300 "
        MySQLStr = MySQLStr & "WHERE (PL01001 = N'" & Trim(Label9.Text) & "') "
        InitMyConn(False)
        InitMyRec(False, MySQLStr)
        If Declarations.MyRec.BOF = True And Declarations.MyRec.EOF = True Then
            MyLanguage = "RUS"
            trycloseMyRec()
        Else
            MyLanguage = Declarations.MyRec.Fields("CC").Value
            trycloseMyRec()
        End If

        j = i
        MySQLStr = "SELECT LTRIM(RTRIM(LTRIM(RTRIM(ISNULL(View_1.SYPD001, ''))) + ' ' + LTRIM(RTRIM(ISNULL(View_1.SYPD003, ''))))) AS Purchaser "
        MySQLStr = MySQLStr & "FROM tbl_SupplierCard0300 WITH(NOLOCK) LEFT OUTER JOIN "
        MySQLStr = MySQLStr & "(SELECT SYPD001, SYPD002, SYPD003 "
        MySQLStr = MySQLStr & "FROM SYPD0300 WITH(NOLOCK) "
        If MyLanguage = "RUS" Then
            MySQLStr = MySQLStr & "WHERE (SYPD002 = N'ENG')) "
        Else
            MySQLStr = MySQLStr & "WHERE (SYPD002 = N'RUS')) "
        End If
        MySQLStr = MySQLStr & "AS View_1 ON UPPER(tbl_SupplierCard0300.Purchaser) = UPPER(View_1.SYPD001) "
        MySQLStr = MySQLStr & "WHERE (tbl_SupplierCard0300.PL01001 = N'" & Declarations.MySupplierCode & "')"
        InitMyConn(False)
        InitMyRec(False, MySQLStr)
        If Declarations.MyRec.BOF = True And Declarations.MyRec.EOF = True Then
            PurchaserName = ""
            trycloseMyRec()
        Else
            PurchaserName = Declarations.MyRec.Fields("Purchaser").Value
            trycloseMyRec()
        End If

        MySQLStr = "SELECT SYCD0100.SYCD009 "
        MySQLStr = MySQLStr & "FROM PL010300 WITH(NOLOCK) INNER JOIN "
        MySQLStr = MySQLStr & "SYCD0100 ON PL010300.PL01026 = SYCD0100.SYCD001 "
        MySQLStr = MySQLStr & "WHERE (PL010300.PL01001 = N'" & Declarations.MySupplierCode & "') "
        InitMyConn(False)
        InitMyRec(False, MySQLStr)
        If Declarations.MyRec.BOF = True And Declarations.MyRec.EOF = True Then
            CurrName = "RUB"
            trycloseMyRec()
        Else
            CurrName = Declarations.MyRec.Fields("SYCD009").Value
            trycloseMyRec()
        End If

        MyWRKBook.ActiveSheet.Range("B" & CStr(i) & ":H" & CStr(i)).MergeCells = True
        MyWRKBook.ActiveSheet.Range("B" & CStr(i)) = "Total/Итого, " & CurrName & ":"
        MyWRKBook.ActiveSheet.Range("B" & CStr(i) & ":H" & CStr(i)).Font.Size = 10
        MyWRKBook.ActiveSheet.Range("B" & CStr(i) & ":H" & CStr(i)).Font.Bold = True
        MyWRKBook.ActiveSheet.Range("B" & CStr(i) & ":H" & CStr(i)).HorizontalAlignment = -4152
        MyWRKBook.ActiveSheet.Range("I" & CStr(i)) = Declarations.MyPurchOrderSum
        MyWRKBook.ActiveSheet.Range("I" & CStr(i)).Font.Size = 10
        MyWRKBook.ActiveSheet.Range("I" & CStr(i)).Font.Bold = True

        MyWRKBook.ActiveSheet.Rows(CStr(i + 1) & ":" & CStr(i + 1)).RowHeight = 5

        MyWRKBook.ActiveSheet.Range("B" & CStr(i + 2) & ":I" & CStr(i + 2)).MergeCells = True
        MyWRKBook.ActiveSheet.Range("B" & CStr(i + 2)) = "Total lines / Всего наименований " & CStr(i - 16) & " In amount of / На сумму " & CStr(Declarations.MyPurchOrderSum) & " " & CurrName
        MyWRKBook.ActiveSheet.Range("B" & CStr(i + 4) & ":I" & CStr(i + 4)).MergeCells = True
        MyWRKBook.ActiveSheet.Range("B" & CStr(i + 4)) = "Purchaser / Закупщик:______________________" & PurchaserName
    End Sub

    Private Sub ExportOrderFooterToLO(ByRef oSheet As Object, ByRef oServiceManager As Object, ByRef oWorkBook As Object, _
        ByRef oDispatcher As Object, ByVal ComOrder As String, ByRef i As Integer)
        '////////////////////////////////////////////////////////////////////////////////
        '//
        '// Процедура выгрузки подвала обобщенного заказа в Libre Office
        '//
        '////////////////////////////////////////////////////////////////////////////////
        Dim MyLanguage As String
        Dim MySQLStr As String
        Dim CurrName As String
        Dim PurchaserName As String
        Dim j As Integer
        Dim oFrame As Object

        oFrame = oWorkBook.getCurrentController.getFrame

        '------Язык документа---------------------
        MySQLStr = "SELECT PL01027 AS CC "
        MySQLStr = MySQLStr & "FROM PL010300 "
        MySQLStr = MySQLStr & "WHERE (PL01001 = N'" & Trim(Label9.Text) & "') "
        InitMyConn(False)
        InitMyRec(False, MySQLStr)
        If Declarations.MyRec.BOF = True And Declarations.MyRec.EOF = True Then
            MyLanguage = "RUS"
            trycloseMyRec()
        Else
            MyLanguage = Declarations.MyRec.Fields("CC").Value
            trycloseMyRec()
        End If

        j = i
        MySQLStr = "SELECT LTRIM(RTRIM(LTRIM(RTRIM(ISNULL(View_1.SYPD001, ''))) + ' ' + LTRIM(RTRIM(ISNULL(View_1.SYPD003, ''))))) AS Purchaser "
        MySQLStr = MySQLStr & "FROM tbl_SupplierCard0300 WITH(NOLOCK) LEFT OUTER JOIN "
        MySQLStr = MySQLStr & "(SELECT SYPD001, SYPD002, SYPD003 "
        MySQLStr = MySQLStr & "FROM SYPD0300 WITH(NOLOCK) "
        If MyLanguage = "RUS" Then
            MySQLStr = MySQLStr & "WHERE (SYPD002 = N'ENG')) "
        Else
            MySQLStr = MySQLStr & "WHERE (SYPD002 = N'RUS')) "
        End If
        MySQLStr = MySQLStr & "AS View_1 ON UPPER(tbl_SupplierCard0300.Purchaser) = UPPER(View_1.SYPD001) "
        MySQLStr = MySQLStr & "WHERE (tbl_SupplierCard0300.PL01001 = N'" & Declarations.MySupplierCode & "')"
        InitMyConn(False)
        InitMyRec(False, MySQLStr)
        If Declarations.MyRec.BOF = True And Declarations.MyRec.EOF = True Then
            PurchaserName = ""
            trycloseMyRec()
        Else
            PurchaserName = Declarations.MyRec.Fields("Purchaser").Value
            trycloseMyRec()
        End If

        MySQLStr = "SELECT SYCD0100.SYCD009 "
        MySQLStr = MySQLStr & "FROM PL010300 WITH(NOLOCK) INNER JOIN "
        MySQLStr = MySQLStr & "SYCD0100 ON PL010300.PL01026 = SYCD0100.SYCD001 "
        MySQLStr = MySQLStr & "WHERE (PL010300.PL01001 = N'" & Declarations.MySupplierCode & "') "
        InitMyConn(False)
        InitMyRec(False, MySQLStr)
        If Declarations.MyRec.BOF = True And Declarations.MyRec.EOF = True Then
            CurrName = "RUB"
            trycloseMyRec()
        Else
            CurrName = Declarations.MyRec.Fields("SYCD009").Value
            trycloseMyRec()
        End If

        LOMergeCells(oServiceManager, oDispatcher, oFrame, "B" & CStr(i) & ":H" & CStr(i))
        oSheet.getCellRangeByName("B" & CStr(i)).String = "Total/Итого, " & CurrName & ":"
        oSheet.getCellRangeByName("I" & CStr(i)).String = Declarations.MyPurchOrderSum
        LOFontSetFamilyName(oServiceManager, oDispatcher, oFrame, "B" & CStr(i) & ":I" & CStr(i), "Arial")
        LOFontSetBold(oServiceManager, oDispatcher, oFrame, "B" & CStr(i) & ":I" & CStr(i))
        LOFontSetSize(oServiceManager, oDispatcher, oFrame, "B" & CStr(i) & ":I" & CStr(i), 10)
        oSheet.getCellRangeByName("B" & CStr(i) & ":I" & CStr(i)).VertJustify = 2

        i = i + 1
        oSheet.getCellRangeByName("A" & CStr(i)).Rows.Height = 200

        i = i + 1
        LOMergeCells(oServiceManager, oDispatcher, oFrame, "B" & CStr(i) & ":I" & CStr(i))
        oSheet.getCellRangeByName("B" & CStr(i)).String = "Total lines / Всего наименований " & CStr(i - 18) & " In amount of / На сумму " & CStr(Declarations.MyPurchOrderSum) & " " & CurrName
        LOFontSetSize(oServiceManager, oDispatcher, oFrame, "B" & CStr(i), 10)
        oSheet.getCellRangeByName("B" & CStr(i)).VertJustify = 2

        i = i + 2
        LOMergeCells(oServiceManager, oDispatcher, oFrame, "B" & CStr(i) & ":I" & CStr(i))
        oSheet.getCellRangeByName("B" & CStr(i)).String = "Purchaser / Закупщик:______________________" & PurchaserName
        LOFontSetSize(oServiceManager, oDispatcher, oFrame, "B" & CStr(i), 10)
        oSheet.getCellRangeByName("B" & CStr(i)).VertJustify = 2
    End Sub

    Private Sub Button15_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button15.Click
        '////////////////////////////////////////////////////////////////////////////////
        '//
        '// Выгрузка обобщенного заказа в Excel - полная
        '//
        '////////////////////////////////////////////////////////////////////////////////

        If CheckUOMInOrders(0) = True Then
            If My.Settings.UseOffice = "LibreOffice" Then
                ExportOrderToLOFull(Trim(DataGridView1.SelectedRows.Item(0).Cells(0).Value.ToString()))
            Else
                ExportOrderToExcelFull(Trim(DataGridView1.SelectedRows.Item(0).Cells(0).Value.ToString()))
            End If

        End If
    End Sub

    Private Sub ExportOrderToExcelFull(ByVal ComOrder As String)
        '////////////////////////////////////////////////////////////////////////////////
        '//
        '// Процедура выгрузки обобщенного заказа в Excel - полная
        '//
        '////////////////////////////////////////////////////////////////////////////////
        Dim MyObj As Object
        Dim MyWRKBook As Object
        Dim i As Integer                              'счетчик строк
        Dim j As Integer                              'счетчик строк

        MyObj = CreateObject("Excel.Application")
        MyObj.SheetsInNewWorkbook = 1
        MyWRKBook = MyObj.Workbooks.Add

        ExportOrderHeaderToExcelFull(MyWRKBook, ComOrder, i)
        ExportOrderBodyToExcelFull(MyWRKBook, ComOrder, i, j)
        ExportOrderFooterToExcelFull(MyWRKBook, ComOrder, i, j)

        MyWRKBook.ActiveSheet.Range("A1").Select()
        MyObj.Application.Visible = True
        MyObj = Nothing
    End Sub

    Private Sub ExportOrderToLOFull(ByVal ComOrder As String)
        '////////////////////////////////////////////////////////////////////////////////
        '//
        '// Процедура выгрузки обобщенного заказа в Libre Office - полная
        '//
        '////////////////////////////////////////////////////////////////////////////////
        Dim oServiceManager As Object
        Dim oDispatcher As Object
        Dim oDesktop As Object
        Dim oWorkBook As Object
        Dim oSheet As Object
        Dim oFrame As Object
        Dim i As Integer                              'счетчик строк
        Dim j As Integer                              'счетчик строк

        LOSetNotation(0)
        oServiceManager = CreateObject("com.sun.star.ServiceManager")
        oDesktop = oServiceManager.createInstance("com.sun.star.frame.Desktop")
        oDispatcher = oServiceManager.createInstance("com.sun.star.frame.DispatchHelper")
        Dim arg(1)
        arg(0) = mAkePropertyValue("Hidden", True, oServiceManager)
        arg(1) = mAkePropertyValue("MacroExecutionMode", 4, oServiceManager)
        oWorkBook = oDesktop.loadComponentFromURL("private:factory/scalc", "_blank", 0, arg)
        oSheet = oWorkBook.getSheets().getByIndex(0)
        oFrame = oWorkBook.getCurrentController.getFrame

        ExportOrderHeaderToLOFull(oSheet, oServiceManager, oWorkBook, oDispatcher, ComOrder, i)
        ExportOrderBodyToLOFull(oSheet, oServiceManager, oWorkBook, oDispatcher, ComOrder, i, j)
        ExportOrderFooterToLOFull(oSheet, oServiceManager, oWorkBook, oDispatcher, ComOrder, i, j)

        Dim args() As Object
        ReDim args(0)
        args(0) = oServiceManager.Bridge_GetStruct("com.sun.star.beans.PropertyValue")
        args(0).Name = "ToPoint"
        args(0).Value = "$A$1"
        oDispatcher.executeDispatch(oFrame, ".uno:GoToCell", "", 0, args)

        oWorkBook.CurrentController.Frame.ContainerWindow.Visible = True
        oWorkBook.CurrentController.Frame.ContainerWindow.toFront()
    End Sub

    Private Sub ExportOrderHeaderToExcelFull(ByRef MyWRKBook As Object, ByVal ComOrder As String, ByRef i As Integer)
        '////////////////////////////////////////////////////////////////////////////////
        '//
        '// Процедура выгрузки заголовка обобщенного заказа в Excel полная
        '//
        '////////////////////////////////////////////////////////////////////////////////
        Dim MySQLStr As String
        Dim SupplierName As String
        Dim SupplierAddress As String
        Dim DelName As String
        Dim DelAddr As String
        Dim CurrName As String
        Dim MyLanguage As String

        '------Язык документа---------------------
        MySQLStr = "SELECT PL01027 AS CC "
        MySQLStr = MySQLStr & "FROM PL010300 "
        MySQLStr = MySQLStr & "WHERE (PL01001 = N'" & Trim(Label9.Text) & "') "
        InitMyConn(False)
        InitMyRec(False, MySQLStr)
        If Declarations.MyRec.BOF = True And Declarations.MyRec.EOF = True Then
            MyLanguage = "RUS"
            trycloseMyRec()
        Else
            MyLanguage = Declarations.MyRec.Fields("CC").Value
            trycloseMyRec()
        End If

        '------наша компания и адрес--------------
        MyWRKBook.ActiveSheet.Range("B2:J2").MergeCells = True
        If MyLanguage = "RUS" Then
            MyWRKBook.ActiveSheet.Range("B2") = "ООО ""Скандика"""
        Else
            MyWRKBook.ActiveSheet.Range("B2") = "Skandika LLC"
        End If
        MyWRKBook.ActiveSheet.Range("B2:J2").Font.Name = "Arial"
        MyWRKBook.ActiveSheet.Range("B2:J2").Font.Size = 12
        MyWRKBook.ActiveSheet.Range("B2:J2").Font.Bold = True
        MyWRKBook.ActiveSheet.Range("B2:J2").WrapText = True

        MyWRKBook.ActiveSheet.Range("B3:D3").MergeCells = True
        MyWRKBook.ActiveSheet.Range("B3") = "Address/Адрес:"
        MyWRKBook.ActiveSheet.Range("B3:D3").Font.Size = 10
        MyWRKBook.ActiveSheet.Range("B3:D3").Font.Bold = True

        MyWRKBook.ActiveSheet.Range("E3:J3").MergeCells = True
        If MyLanguage = "RUS" Then
            MyWRKBook.ActiveSheet.Range("E3") = "195027, РФ, г. Санкт-Петербург, Шаумяна проспект, дом 4, корпус 1, литер А, помещение 25Н., Tel: +7 (812) 325 2040, Fax: +7 (812) 325 2039, www.skandikagroup.ru"
        Else
            MyWRKBook.ActiveSheet.Range("E3") = "195027, Russia, St. Petersburg, Shaumyana prospect, house 4, building 1, liter А, room 25Н., Tel: +7 (812) 325 2040, Fax: +7 (812) 325 2039, www.skandikagroup.ru"
        End If

        MyWRKBook.ActiveSheet.Range("E3:J3").Font.Size = 10
        MyWRKBook.ActiveSheet.Range("B3:J3").WrapText = True
        MyWRKBook.ActiveSheet.Rows("3:3").RowHeight = 30
        MyWRKBook.ActiveSheet.Range("B3:J3").VerticalAlignment = -4108

        '-------Номер и дата заказа на закупку--------------
        MySQLStr = "SELECT CONVERT(nvarchar(30),OrderDate,103) AS OrderDate "
        MySQLStr = MySQLStr & "FROM tbl_PurchaseWorkplace_ConsolidatedOrders "
        MySQLStr = MySQLStr & "WHERE (ID = N'" & ComOrder & "')"
        InitMyConn(False)
        InitMyRec(False, MySQLStr)
        MyWRKBook.ActiveSheet.Range("B4:J4").MergeCells = True
        If Declarations.MyRec.BOF = True And Declarations.MyRec.EOF = True Then
            MyWRKBook.ActiveSheet.Range("B4") = "Purchase Order / Заказ на закупку № " & ComOrder & " от  "
            trycloseMyRec()
        Else
            MyWRKBook.ActiveSheet.Range("B4") = "Purchase Order / Заказ на закупку № " & ComOrder & " от  " & Declarations.MyRec.Fields("OrderDate").Value
            trycloseMyRec()
        End If
        MyWRKBook.ActiveSheet.Range("B4:J4").Font.Name = "Arial"
        MyWRKBook.ActiveSheet.Range("B4:J4").Font.Size = 12
        MyWRKBook.ActiveSheet.Range("B4:J4").Font.Bold = True
        MyWRKBook.ActiveSheet.Range("B4:J4").WrapText = True
        MyWRKBook.ActiveSheet.Range("B4:J4").HorizontalAlignment = -4108

        MyWRKBook.ActiveSheet.Rows("5:5").RowHeight = 5

        '-----------поставщик-----------------------------
        MyWRKBook.ActiveSheet.Range("B6:D6").MergeCells = True
        MyWRKBook.ActiveSheet.Range("B6") = "Supplier / Поставщик:"
        MyWRKBook.ActiveSheet.Range("B6:D6").Font.Size = 10
        MyWRKBook.ActiveSheet.Range("B6:D6").Font.Bold = True

        MySQLStr = "SELECT PL01002 AS SuppName, LTRIM(RTRIM(LTRIM(RTRIM(PL01003)) + ' ' + LTRIM(RTRIM(PL01004)) + ' ' + LTRIM(RTRIM(PL01005)))) AS SuppAddress "
        MySQLStr = MySQLStr & "FROM PL010300 "
        MySQLStr = MySQLStr & "WHERE (PL01001 = N'" & Declarations.MySupplierCode & "') "
        InitMyConn(False)
        InitMyRec(False, MySQLStr)
        If Declarations.MyRec.BOF = True And Declarations.MyRec.EOF = True Then
            SupplierName = ""
            SupplierAddress = ""
            trycloseMyRec()
        Else
            SupplierName = Declarations.MyRec.Fields("SuppName").Value
            SupplierAddress = Declarations.MyRec.Fields("SuppAddress").Value
            trycloseMyRec()
        End If
        MyWRKBook.ActiveSheet.Range("E6:J6").MergeCells = True
        MyWRKBook.ActiveSheet.Range("E6") = SupplierName & Chr(10) & SupplierAddress
        MyWRKBook.ActiveSheet.Range("E6:J6").Font.Size = 10
        MyWRKBook.ActiveSheet.Range("B6:J6").WrapText = True
        MyWRKBook.ActiveSheet.Rows("6:6").RowHeight = 45
        MyWRKBook.ActiveSheet.Range("B6:J6").VerticalAlignment = -4108

        MyWRKBook.ActiveSheet.Rows("7:7").RowHeight = 5

        '---------Адрес поставки--------------------------------
        MyWRKBook.ActiveSheet.Range("B8:D8").MergeCells = True
        MyWRKBook.ActiveSheet.Range("B8") = "Delivery Address / Адрес поставки"
        MyWRKBook.ActiveSheet.Range("B8:D8").Font.Size = 10
        MyWRKBook.ActiveSheet.Range("B8:D8").Font.Bold = True

        MySQLStr = "SELECT LTRIM(RTRIM(ConsignorOfGoodsName)) AS DelName, LTRIM(RTRIM(ConsignorOfGoodsAddr)) AS DelAddr "
        MySQLStr = MySQLStr & "FROM tbl_WarehouseAccessoires0300 "
        MySQLStr = MySQLStr & "WHERE (SC23001 = N'" & Declarations.MyWH & "') "
        InitMyConn(False)
        InitMyRec(False, MySQLStr)
        If Declarations.MyRec.BOF = True And Declarations.MyRec.EOF = True Then
            DelName = ""
            DelAddr = ""
            trycloseMyRec()
        Else
            DelName = Declarations.MyRec.Fields("DelName").Value
            DelAddr = Declarations.MyRec.Fields("DelAddr").Value
            trycloseMyRec()
        End If
        If MyLanguage <> "RUS" And Trim(Declarations.MyWH) = "01" Then
            DelName = ""
            DelAddr = "Marshala Blyukhera prospekt, 78-D, 195253, Saint Petersburg, Russia Tel: +7 (812)325-20-40, Fax: +7 (812)325-03-22"
        End If
        MyWRKBook.ActiveSheet.Range("E8:J8").MergeCells = True
        MyWRKBook.ActiveSheet.Range("E8") = DelName & Chr(10) & DelAddr
        MyWRKBook.ActiveSheet.Range("E8:J8").Font.Size = 10
        MyWRKBook.ActiveSheet.Range("B8:J8").WrapText = True
        MyWRKBook.ActiveSheet.Rows("8:8").RowHeight = 45
        MyWRKBook.ActiveSheet.Range("B8:J8").VerticalAlignment = -4108

        MyWRKBook.ActiveSheet.Rows("9:9").RowHeight = 5

        '---------Условия поставки------------------------------
        MyWRKBook.ActiveSheet.Range("B10:D10").MergeCells = True
        MyWRKBook.ActiveSheet.Range("B10") = "Terms Of Delivery / Условия поставки"
        MyWRKBook.ActiveSheet.Range("B10:D10").Font.Size = 7

        MySQLStr = "SELECT View_1.PL23004 "
        MySQLStr = MySQLStr & "FROM PL010300 WITH(NOLOCK) INNER JOIN "
        MySQLStr = MySQLStr & "(SELECT PL23001, PL23002, PL23003, PL23004 "
        MySQLStr = MySQLStr & "FROM PL230300 WITH(NOLOCK) "
        MySQLStr = MySQLStr & "WHERE (PL23001 = N'1') AND (PL23002 = N'" & MyLanguage & "')) AS View_1 ON CONVERT(int, PL010300.PL01029) = CONVERT(int, View_1.PL23003) "
        MySQLStr = MySQLStr & "WHERE (PL010300.PL01001 = N'" & Declarations.MySupplierCode & "')"
        InitMyConn(False)
        InitMyRec(False, MySQLStr)
        MyWRKBook.ActiveSheet.Range("E10:J10").MergeCells = True
        If Declarations.MyRec.BOF = True And Declarations.MyRec.EOF = True Then
            trycloseMyRec()
        Else
            MyWRKBook.ActiveSheet.Range("E10") = Declarations.MyRec.Fields("PL23004").Value
            trycloseMyRec()
        End If
        MyWRKBook.ActiveSheet.Range("E10:J10").Font.Size = 8
        MyWRKBook.ActiveSheet.Range("B10:J10").WrapText = True
        MyWRKBook.ActiveSheet.Range("B10:J10").VerticalAlignment = -4108

        '------Условия оплаты------------------------------
        MyWRKBook.ActiveSheet.Range("B11:D11").MergeCells = True
        MyWRKBook.ActiveSheet.Range("B11") = "Terms Of Payment / Условия оплаты"
        MyWRKBook.ActiveSheet.Range("B11:D11").Font.Size = 7

        MySQLStr = "SELECT View_1.PL23004 "
        MySQLStr = MySQLStr & "FROM PL010300 WITH(NOLOCK) INNER JOIN "
        MySQLStr = MySQLStr & "(SELECT PL23001, PL23002, PL23003, PL23004 "
        MySQLStr = MySQLStr & "FROM PL230300 WITH(NOLOCK) "
        MySQLStr = MySQLStr & "WHERE (PL23001 = N'0') AND (PL23002 = N'" & MyLanguage & "')) AS View_1 ON PL010300.PL01028 = View_1.PL23003 "
        MySQLStr = MySQLStr & "WHERE (PL010300.PL01001 = N'" & Declarations.MySupplierCode & "')"
        InitMyConn(False)
        InitMyRec(False, MySQLStr)
        MyWRKBook.ActiveSheet.Range("E11:J11").MergeCells = True
        If Declarations.MyRec.BOF = True And Declarations.MyRec.EOF = True Then
            trycloseMyRec()
        Else
            MyWRKBook.ActiveSheet.Range("E11") = Declarations.MyRec.Fields("PL23004").Value
            trycloseMyRec()
        End If
        MyWRKBook.ActiveSheet.Range("E11:J11").Font.Size = 8
        MyWRKBook.ActiveSheet.Range("B11:J11").WrapText = True
        MyWRKBook.ActiveSheet.Range("B11:J11").VerticalAlignment = -4108

        MyWRKBook.ActiveSheet.Rows("12:12").RowHeight = 5

        '-------Закупщик------------------------------------
        MyWRKBook.ActiveSheet.Range("B13:D13").MergeCells = True
        MyWRKBook.ActiveSheet.Range("B13") = "Purchaser / Закупщик:"
        MyWRKBook.ActiveSheet.Range("B13:D13").Font.Size = 10

        MySQLStr = "SELECT LTRIM(RTRIM(LTRIM(RTRIM(ISNULL(View_1.SYPD001, ''))) + ' ' + LTRIM(RTRIM(ISNULL(View_1.SYPD003, ''))))) AS Purchaser "
        MySQLStr = MySQLStr & "FROM tbl_SupplierCard0300 WITH(NOLOCK) LEFT OUTER JOIN "
        MySQLStr = MySQLStr & "(SELECT SYPD001, SYPD002, SYPD003 "
        MySQLStr = MySQLStr & "FROM SYPD0300 WITH(NOLOCK) "
        If MyLanguage = "RUS" Then
            MySQLStr = MySQLStr & "WHERE (SYPD002 = N'ENG')) "
        Else
            MySQLStr = MySQLStr & "WHERE (SYPD002 = N'RUS')) "
        End If
        MySQLStr = MySQLStr & "AS View_1 ON UPPER(tbl_SupplierCard0300.Purchaser) = UPPER(View_1.SYPD001) "
        MySQLStr = MySQLStr & "WHERE (tbl_SupplierCard0300.PL01001 = N'" & Declarations.MySupplierCode & "')"
        InitMyConn(False)
        InitMyRec(False, MySQLStr)
        MyWRKBook.ActiveSheet.Range("E13:J13").MergeCells = True
        If Declarations.MyRec.BOF = True And Declarations.MyRec.EOF = True Then
            trycloseMyRec()
        Else
            MyWRKBook.ActiveSheet.Range("E13") = Declarations.MyRec.Fields("Purchaser").Value
            trycloseMyRec()
        End If
        MyWRKBook.ActiveSheet.Range("E13:J13").Font.Size = 10
        MyWRKBook.ActiveSheet.Range("B13:J13").WrapText = True
        MyWRKBook.ActiveSheet.Range("B13:J13").VerticalAlignment = -4108

        MyWRKBook.ActiveSheet.Rows("14:14").RowHeight = 5
        '-------Заголовок таблицы---------------------------
        MySQLStr = "SELECT SYCD0100.SYCD009 "
        MySQLStr = MySQLStr & "FROM PL010300 WITH(NOLOCK) INNER JOIN "
        MySQLStr = MySQLStr & "SYCD0100 ON PL010300.PL01026 = SYCD0100.SYCD001 "
        MySQLStr = MySQLStr & "WHERE (PL010300.PL01001 = N'" & Declarations.MySupplierCode & "') "
        InitMyConn(False)
        InitMyRec(False, MySQLStr)
        If Declarations.MyRec.BOF = True And Declarations.MyRec.EOF = True Then
            CurrName = "RUB"
            trycloseMyRec()
        Else
            CurrName = Declarations.MyRec.Fields("SYCD009").Value
            trycloseMyRec()
        End If
        MyWRKBook.ActiveSheet.Range("B15") = "N п/п"
        MyWRKBook.ActiveSheet.Range("C15") = "Supp Item Code / Код товара поставщ"
        MyWRKBook.ActiveSheet.Range("D15") = "Item Code / Код товара"
        MyWRKBook.ActiveSheet.Range("E15") = "Item Name / Наименование товара"
        MyWRKBook.ActiveSheet.Range("F15") = "UOM / Ед. измер-я"
        MyWRKBook.ActiveSheet.Range("G15") = "QTY / Количество"
        MyWRKBook.ActiveSheet.Range("H15") = "Price / Цена заказа, " & CurrName
        MyWRKBook.ActiveSheet.Range("I15") = "Price / Цена поставщика"
        MyWRKBook.ActiveSheet.Range("J15") = "Summa / Сумма, " & CurrName
        MyWRKBook.ActiveSheet.Range("K15") = "Мин кол-во в заказе "
        MyWRKBook.ActiveSheet.Range("L15") = "складской "
        MyWRKBook.ActiveSheet.Range("B15:L15").Font.Size = 7
        MyWRKBook.ActiveSheet.Range("B15:L15").WrapText = True
        MyWRKBook.ActiveSheet.Range("B15:L15").VerticalAlignment = -4108
        MyWRKBook.ActiveSheet.Range("B15:L15").HorizontalAlignment = -4108
        MyWRKBook.ActiveSheet.Rows("15:15").RowHeight = 40
        MyWRKBook.ActiveSheet.Range("B15:L15").Select()
        MyWRKBook.ActiveSheet.Range("B15:L15").Borders(5).LineStyle = -4142
        MyWRKBook.ActiveSheet.Range("B15:L15").Borders(6).LineStyle = -4142
        With MyWRKBook.ActiveSheet.Range("B15:L15").Borders(7)
            .LineStyle = 1
            .Weight = 4
            .ColorIndex = -4105
        End With
        With MyWRKBook.ActiveSheet.Range("B15:L15").Borders(8)
            .LineStyle = 1
            .Weight = 4
            .ColorIndex = -4105
        End With
        With MyWRKBook.ActiveSheet.Range("B15:L15").Borders(9)
            .LineStyle = 1
            .Weight = 4
            .ColorIndex = -4105
        End With
        With MyWRKBook.ActiveSheet.Range("B15:L15").Borders(10)
            .LineStyle = 1
            .Weight = 4
            .ColorIndex = -4105
        End With
        With MyWRKBook.ActiveSheet.Range("B15:L15").Borders(11)
            .LineStyle = 1
            .Weight = 4
            .ColorIndex = -4105
        End With
        With MyWRKBook.ActiveSheet.Range("B15:L15").Interior
            .ColorIndex = 36
            .Pattern = 1
            .PatternColorIndex = -4105
        End With


        MyWRKBook.ActiveSheet.Columns("A:A").ColumnWidth = 1
        MyWRKBook.ActiveSheet.Columns("B:B").ColumnWidth = 3
        MyWRKBook.ActiveSheet.Columns("C:C").ColumnWidth = 8
        MyWRKBook.ActiveSheet.Columns("D:D").ColumnWidth = 8
        MyWRKBook.ActiveSheet.Columns("E:E").ColumnWidth = 27
        MyWRKBook.ActiveSheet.Columns("F:F").ColumnWidth = 4
        MyWRKBook.ActiveSheet.Columns("G:G").ColumnWidth = 4
        MyWRKBook.ActiveSheet.Columns("H:H").ColumnWidth = 7
        MyWRKBook.ActiveSheet.Columns("I:I").ColumnWidth = 7
        MyWRKBook.ActiveSheet.Columns("J:J").ColumnWidth = 12
        MyWRKBook.ActiveSheet.Columns("K:K").ColumnWidth = 7
        MyWRKBook.ActiveSheet.Columns("L:L").ColumnWidth = 5

        MyWRKBook.ActiveSheet.Range("C16") = "N заказа на продажу"
        MyWRKBook.ActiveSheet.Range("D16") = "N заказа на закупку"
        MyWRKBook.ActiveSheet.Range("E16") = "Продавец"
        MyWRKBook.ActiveSheet.Range("F16") = "UOM / Ед. измер-я"
        MyWRKBook.ActiveSheet.Range("G16") = "QTY / Количество"
        MyWRKBook.ActiveSheet.Range("H16") = "Дата отгрузки в заказе на продажу"
        MyWRKBook.ActiveSheet.Range("I16") = "Подтвержденная дата поставки"
        MyWRKBook.ActiveSheet.Range("J16") = "Подтвержденная задолженная дата"
        MyWRKBook.ActiveSheet.Range("C16:J16").Font.Size = 7
        MyWRKBook.ActiveSheet.Range("C16:J16").WrapText = True
        MyWRKBook.ActiveSheet.Range("C16:J16").VerticalAlignment = -4108
        MyWRKBook.ActiveSheet.Range("C16:J16").HorizontalAlignment = -4108
        MyWRKBook.ActiveSheet.Rows("16:16").RowHeight = 40
        MyWRKBook.ActiveSheet.Range("C16:J16").Select()
        MyWRKBook.ActiveSheet.Range("C16:J16").Borders(5).LineStyle = -4142
        MyWRKBook.ActiveSheet.Range("C16:J16").Borders(6).LineStyle = -4142
        With MyWRKBook.ActiveSheet.Range("C16:J16").Borders(7)
            .LineStyle = 1
            .Weight = 4
            .ColorIndex = -4105
        End With
        With MyWRKBook.ActiveSheet.Range("C16:J16").Borders(8)
            .LineStyle = 1
            .Weight = 4
            .ColorIndex = -4105
        End With
        With MyWRKBook.ActiveSheet.Range("C16:J16").Borders(9)
            .LineStyle = 1
            .Weight = 4
            .ColorIndex = -4105
        End With
        With MyWRKBook.ActiveSheet.Range("C16:J16").Borders(10)
            .LineStyle = 1
            .Weight = 4
            .ColorIndex = -4105
        End With
        With MyWRKBook.ActiveSheet.Range("C16:J16").Borders(11)
            .LineStyle = 1
            .Weight = 4
            .ColorIndex = -4105
        End With
        With MyWRKBook.ActiveSheet.Range("C16:J16").Interior
            .Color = 5296274
            .Pattern = 1
            .PatternColorIndex = -4105
        End With

        i = 17
    End Sub

    Private Sub ExportOrderHeaderToLOFull(ByRef oSheet As Object, ByRef oServiceManager As Object, ByRef oWorkBook As Object, _
        ByRef oDispatcher As Object, ByVal ComOrder As String, ByRef i As Integer)
        '////////////////////////////////////////////////////////////////////////////////
        '//
        '// Процедура выгрузки заголовка обобщенного заказа в Libre Office полная
        '//
        '////////////////////////////////////////////////////////////////////////////////
        Dim oFrame As Object
        Dim MySQLStr As String
        Dim SupplierName As String
        Dim SupplierAddress As String
        Dim DelName As String
        Dim DelAddr As String
        Dim CurrName As String
        Dim MyLanguage As String

        oFrame = oWorkBook.getCurrentController.getFrame

        '-----Ширина колонок
        oSheet.getColumns().getByName("A").Width = 400
        oSheet.getColumns().getByName("B").Width = 600
        oSheet.getColumns().getByName("C").Width = 1800
        oSheet.getColumns().getByName("D").Width = 1800
        oSheet.getColumns().getByName("E").Width = 6000
        oSheet.getColumns().getByName("F").Width = 1000
        oSheet.getColumns().getByName("G").Width = 1000
        oSheet.getColumns().getByName("H").Width = 1600
        oSheet.getColumns().getByName("I").Width = 1600
        oSheet.getColumns().getByName("J").Width = 2000
        oSheet.getColumns().getByName("K").Width = 1600
        oSheet.getColumns().getByName("L").Width = 1600

        '------Язык документа---------------------
        MySQLStr = "SELECT PL01027 AS CC "
        MySQLStr = MySQLStr & "FROM PL010300 "
        MySQLStr = MySQLStr & "WHERE (PL01001 = N'" & Trim(Label9.Text) & "') "
        InitMyConn(False)
        InitMyRec(False, MySQLStr)
        If Declarations.MyRec.BOF = True And Declarations.MyRec.EOF = True Then
            MyLanguage = "RUS"
            trycloseMyRec()
        Else
            MyLanguage = Declarations.MyRec.Fields("CC").Value
            trycloseMyRec()
        End If

        LOMergeCells(oServiceManager, oDispatcher, oFrame, "B2:I2")
        If MyLanguage = "RUS" Then
            oSheet.getCellRangeByName("B2").String = "ООО ""Скандика"""
        Else
            oSheet.getCellRangeByName("B2").String = "Skandika LLC"
        End If
        LOFontSetFamilyName(oServiceManager, oDispatcher, oFrame, "B2", "Arial")
        LOFontSetBold(oServiceManager, oDispatcher, oFrame, "B2")
        LOFontSetSize(oServiceManager, oDispatcher, oFrame, "B2", 12)
        oSheet.getCellRangeByName("B2").VertJustify = 2

        LOMergeCells(oServiceManager, oDispatcher, oFrame, "B3:D3")
        oSheet.getCellRangeByName("B3").String = "Address/Адрес:"
        LOFontSetFamilyName(oServiceManager, oDispatcher, oFrame, "B3:D3", "Calibri")
        LOFontSetBold(oServiceManager, oDispatcher, oFrame, "B3:D3")
        LOFontSetSize(oServiceManager, oDispatcher, oFrame, "B3:D3", 10)
        LOWrapText(oServiceManager, oDispatcher, oFrame, "B3:D3")
        oSheet.getCellRangeByName("B3:D3").VertJustify = 2

        LOMergeCells(oServiceManager, oDispatcher, oFrame, "E3:I3")
        If MyLanguage = "RUS" Then
            oSheet.getCellRangeByName("E3").String = "195027, РФ, г. Санкт-Петербург, Шаумяна проспект, дом 4, корпус 1, литер А, помещение 25Н., Tel: +7 (812) 325 2040, Fax: +7 (812) 325 2039, www.skandikagroup.ru"
        Else
            oSheet.getCellRangeByName("E3").String = "195027, Russia, St. Petersburg, Shaumyana prospect, house 4, building 1, liter А, room 25Н., Tel: +7 (812) 325 2040, Fax: +7 (812) 325 2039, www.skandikagroup.ru"
        End If
        LOFontSetFamilyName(oServiceManager, oDispatcher, oFrame, "E3:I3", "Calibri")
        LOFontSetSize(oServiceManager, oDispatcher, oFrame, "E3:I3", 10)
        LOWrapText(oServiceManager, oDispatcher, oFrame, "E3:I3")
        oSheet.getCellRangeByName("E3:I3").VertJustify = 2

        '-------Номер и дата заказа на закупку--------------
        MySQLStr = "SELECT CONVERT(nvarchar(30),OrderDate,103) AS OrderDate, SupplierCode, WH "
        MySQLStr = MySQLStr & "FROM tbl_PurchaseWorkplace_ConsolidatedOrders WITH(NOLOCK) "
        MySQLStr = MySQLStr & "WHERE (ID = N'" & ComOrder & "')"
        InitMyConn(False)
        InitMyRec(False, MySQLStr)
        LOMergeCells(oServiceManager, oDispatcher, oFrame, "B4:I4")
        If Declarations.MyRec.BOF = True And Declarations.MyRec.EOF = True Then
            Declarations.MySupplierCode = ""
            Declarations.MyWH = ""
            oSheet.getCellRangeByName("B4").String = "Purchase Order / Заказ на закупку № " & ComOrder
            trycloseMyRec()
        Else
            Declarations.MySupplierCode = Declarations.MyRec.Fields("SupplierCode").Value
            Declarations.MyWH = Declarations.MyRec.Fields("WH").Value
            oSheet.getCellRangeByName("B4").String = "Purchase Order / Заказ на закупку № " & ComOrder & " от  " & Declarations.MyRec.Fields("OrderDate").Value
        End If
        LOFontSetFamilyName(oServiceManager, oDispatcher, oFrame, "B4:I4", "Arial")
        LOFontSetBold(oServiceManager, oDispatcher, oFrame, "B4:I4")
        LOFontSetSize(oServiceManager, oDispatcher, oFrame, "B4:I4", 12)
        LOWrapText(oServiceManager, oDispatcher, oFrame, "B4:I4")
        oSheet.getCellRangeByName("B4:I4").VertJustify = 2
        oSheet.getCellRangeByName("A5").Rows.Height = 200

        '-----------поставщик-----------------------------
        LOMergeCells(oServiceManager, oDispatcher, oFrame, "B6:D6")
        oSheet.getCellRangeByName("B6").String = "Supplier / Поставщик:"
        LOFontSetFamilyName(oServiceManager, oDispatcher, oFrame, "B6:D6", "Calibri")
        LOFontSetBold(oServiceManager, oDispatcher, oFrame, "B6:D6")
        LOFontSetSize(oServiceManager, oDispatcher, oFrame, "B6:D6", 10)
        LOWrapText(oServiceManager, oDispatcher, oFrame, "B6:D6")
        oSheet.getCellRangeByName("B6:D6").VertJustify = 2

        MySQLStr = "SELECT PL01002 AS SuppName, LTRIM(RTRIM(LTRIM(RTRIM(PL01003)) + ' ' + LTRIM(RTRIM(PL01004)) + ' ' + LTRIM(RTRIM(PL01005)))) AS SuppAddress "
        MySQLStr = MySQLStr & "FROM PL010300 WITH(NOLOCK) "
        MySQLStr = MySQLStr & "WHERE (PL01001 = N'" & Declarations.MySupplierCode & "') "
        InitMyConn(False)
        InitMyRec(False, MySQLStr)
        If Declarations.MyRec.BOF = True And Declarations.MyRec.EOF = True Then
            SupplierName = ""
            SupplierAddress = ""
            trycloseMyRec()
        Else
            SupplierName = Declarations.MyRec.Fields("SuppName").Value
            SupplierAddress = Declarations.MyRec.Fields("SuppAddress").Value
            trycloseMyRec()
        End If
        LOMergeCells(oServiceManager, oDispatcher, oFrame, "E6:I6")
        oSheet.getCellRangeByName("E6").String = SupplierName & Chr(10) & SupplierAddress
        LOFontSetFamilyName(oServiceManager, oDispatcher, oFrame, "E6:I6", "Calibri")
        LOFontSetSize(oServiceManager, oDispatcher, oFrame, "E6:I6", 10)
        LOWrapText(oServiceManager, oDispatcher, oFrame, "E6:I6")
        oSheet.getCellRangeByName("E6:I6").VertJustify = 2
        oSheet.getCellRangeByName("A7").Rows.Height = 200

        '---------Адрес поставки--------------------------------
        LOMergeCells(oServiceManager, oDispatcher, oFrame, "B8:D8")
        oSheet.getCellRangeByName("B8").String = "Delivery Address / Адрес поставки"
        LOFontSetFamilyName(oServiceManager, oDispatcher, oFrame, "B8:D8", "Calibri")
        LOFontSetBold(oServiceManager, oDispatcher, oFrame, "B8:D8")
        LOFontSetSize(oServiceManager, oDispatcher, oFrame, "B8:D8", 10)
        LOWrapText(oServiceManager, oDispatcher, oFrame, "B8:D8")
        oSheet.getCellRangeByName("B8:D8").VertJustify = 2

        MySQLStr = "SELECT LTRIM(RTRIM(ConsignorOfGoodsName)) AS DelName, LTRIM(RTRIM(ConsignorOfGoodsAddr)) AS DelAddr "
        MySQLStr = MySQLStr & "FROM tbl_WarehouseAccessoires0300 WITH(NOLOCK) "
        MySQLStr = MySQLStr & "WHERE (SC23001 = N'" & Declarations.MyWH & "') "
        InitMyConn(False)
        InitMyRec(False, MySQLStr)
        If Declarations.MyRec.BOF = True And Declarations.MyRec.EOF = True Then
            DelName = ""
            DelAddr = ""
            trycloseMyRec()
        Else
            DelName = Declarations.MyRec.Fields("DelName").Value
            DelAddr = Declarations.MyRec.Fields("DelAddr").Value
            trycloseMyRec()
        End If
        If MyLanguage <> "RUS" And Trim(Declarations.MyWH) = "01" Then
            DelName = ""
            DelAddr = "Marshala Blyukhera prospekt, 78-D, 195253, Saint Petersburg, Russia Tel: +7 (812)325-20-40, Fax: +7 (812)325-03-22"
        End If
        LOMergeCells(oServiceManager, oDispatcher, oFrame, "E8:I8")
        oSheet.getCellRangeByName("E8").String = DelName & Chr(10) & DelAddr
        LOFontSetFamilyName(oServiceManager, oDispatcher, oFrame, "E8:I8", "Calibri")
        LOFontSetSize(oServiceManager, oDispatcher, oFrame, "E8:I8", 10)
        LOWrapText(oServiceManager, oDispatcher, oFrame, "E8:I8")
        oSheet.getCellRangeByName("E8:I8").VertJustify = 2
        oSheet.getCellRangeByName("A9").Rows.Height = 200

        '---------Условия поставки------------------------------
        LOMergeCells(oServiceManager, oDispatcher, oFrame, "B10:D10")
        oSheet.getCellRangeByName("B10").String = "Terms Of Delivery / Условия поставки"
        LOFontSetFamilyName(oServiceManager, oDispatcher, oFrame, "B10:D10", "Calibri")
        LOFontSetSize(oServiceManager, oDispatcher, oFrame, "B10:D10", 7)
        LOWrapText(oServiceManager, oDispatcher, oFrame, "B10:D10")
        oSheet.getCellRangeByName("B10:D10").VertJustify = 2

        MySQLStr = "SELECT View_1.PL23004 "
        MySQLStr = MySQLStr & "FROM PL010300 WITH(NOLOCK) INNER JOIN "
        MySQLStr = MySQLStr & "(SELECT PL23001, PL23002, PL23003, PL23004 "
        MySQLStr = MySQLStr & "FROM PL230300 WITH(NOLOCK) "
        MySQLStr = MySQLStr & "WHERE (PL23001 = N'1') AND (PL23002 = N'" & MyLanguage & "')) AS View_1 ON CONVERT(int, PL010300.PL01029) = CONVERT(int, View_1.PL23003) "
        MySQLStr = MySQLStr & "WHERE (PL010300.PL01001 = N'" & Declarations.MySupplierCode & "')"
        InitMyConn(False)
        InitMyRec(False, MySQLStr)
        LOMergeCells(oServiceManager, oDispatcher, oFrame, "E10:I10")
        If Declarations.MyRec.BOF = True And Declarations.MyRec.EOF = True Then
            trycloseMyRec()
        Else
            oSheet.getCellRangeByName("E10").String = Declarations.MyRec.Fields("PL23004").Value
            trycloseMyRec()
        End If
        LOFontSetFamilyName(oServiceManager, oDispatcher, oFrame, "E10:I10", "Calibri")
        LOFontSetSize(oServiceManager, oDispatcher, oFrame, "E10:I10", 7)
        LOWrapText(oServiceManager, oDispatcher, oFrame, "E10:I10")
        oSheet.getCellRangeByName("E10:I10").VertJustify = 2

        '------Условия оплаты------------------------------
        LOMergeCells(oServiceManager, oDispatcher, oFrame, "B11:D11")
        oSheet.getCellRangeByName("B11").String = "Terms Of Payment / Условия оплаты"
        LOFontSetFamilyName(oServiceManager, oDispatcher, oFrame, "B11:D11", "Calibri")
        LOFontSetSize(oServiceManager, oDispatcher, oFrame, "B11:D11", 7)
        LOWrapText(oServiceManager, oDispatcher, oFrame, "B11:D11")
        oSheet.getCellRangeByName("B11:D11").VertJustify = 2

        MySQLStr = "SELECT View_1.PL23004 "
        MySQLStr = MySQLStr & "FROM PL010300 WITH(NOLOCK) INNER JOIN "
        MySQLStr = MySQLStr & "(SELECT PL23001, PL23002, PL23003, PL23004 "
        MySQLStr = MySQLStr & "FROM PL230300 WITH(NOLOCK) "
        MySQLStr = MySQLStr & "WHERE (PL23001 = N'0') AND (PL23002 = N'" & MyLanguage & "')) AS View_1 ON PL010300.PL01028 = View_1.PL23003 "
        MySQLStr = MySQLStr & "WHERE (PL010300.PL01001 = N'" & Declarations.MySupplierCode & "')"
        InitMyConn(False)
        InitMyRec(False, MySQLStr)
        LOMergeCells(oServiceManager, oDispatcher, oFrame, "E11:I11")
        If Declarations.MyRec.BOF = True And Declarations.MyRec.EOF = True Then
            trycloseMyRec()
        Else
            oSheet.getCellRangeByName("E11").String = Declarations.MyRec.Fields("PL23004").Value
            trycloseMyRec()
        End If
        LOFontSetFamilyName(oServiceManager, oDispatcher, oFrame, "E11:I11", "Calibri")
        LOFontSetSize(oServiceManager, oDispatcher, oFrame, "E11:I11", 7)
        LOWrapText(oServiceManager, oDispatcher, oFrame, "E11:I11")
        oSheet.getCellRangeByName("E11:I11").VertJustify = 2
        oSheet.getCellRangeByName("A12").Rows.Height = 200

        '-------Закупщик------------------------------------
        LOMergeCells(oServiceManager, oDispatcher, oFrame, "B13:D13")
        oSheet.getCellRangeByName("B13").String = "Purchaser / Закупщик:"
        LOFontSetFamilyName(oServiceManager, oDispatcher, oFrame, "B13:D13", "Calibri")
        LOFontSetSize(oServiceManager, oDispatcher, oFrame, "B13:D13", 10)
        LOWrapText(oServiceManager, oDispatcher, oFrame, "B13:D13")
        oSheet.getCellRangeByName("B13:D13").VertJustify = 2

        MySQLStr = "SELECT LTRIM(RTRIM(LTRIM(RTRIM(ISNULL(View_1.SYPD001, ''))) + ' ' + LTRIM(RTRIM(ISNULL(View_1.SYPD003, ''))))) AS Purchaser "
        MySQLStr = MySQLStr & "FROM tbl_SupplierCard0300 WITH(NOLOCK) LEFT OUTER JOIN "
        MySQLStr = MySQLStr & "(SELECT SYPD001, SYPD002, SYPD003 "
        MySQLStr = MySQLStr & "FROM SYPD0300 WITH(NOLOCK) "
        MySQLStr = MySQLStr & "WHERE (SYPD002 = N'ENG')) AS View_1 ON UPPER(tbl_SupplierCard0300.Purchaser) = UPPER(View_1.SYPD001) "
        MySQLStr = MySQLStr & "WHERE (tbl_SupplierCard0300.PL01001 = N'" & Declarations.MySupplierCode & "')"
        InitMyConn(False)
        InitMyRec(False, MySQLStr)
        LOMergeCells(oServiceManager, oDispatcher, oFrame, "E13:I13")
        If Declarations.MyRec.BOF = True And Declarations.MyRec.EOF = True Then
            trycloseMyRec()
        Else
            oSheet.getCellRangeByName("E13").String = Declarations.MyRec.Fields("Purchaser").Value
            trycloseMyRec()
        End If
        LOFontSetFamilyName(oServiceManager, oDispatcher, oFrame, "E13:I13", "Calibri")
        LOFontSetSize(oServiceManager, oDispatcher, oFrame, "E13:I13", 10)
        LOWrapText(oServiceManager, oDispatcher, oFrame, "E13:I13")
        oSheet.getCellRangeByName("E13:I13").VertJustify = 2
        oSheet.getCellRangeByName("A14").Rows.Height = 200

        '-------Заголовок таблицы---------------------------
        MySQLStr = "SELECT SYCD0100.SYCD009 "
        MySQLStr = MySQLStr & "FROM PL010300 WITH(NOLOCK) INNER JOIN "
        MySQLStr = MySQLStr & "SYCD0100 ON PL010300.PL01026 = SYCD0100.SYCD001 "
        MySQLStr = MySQLStr & "WHERE (PL010300.PL01001 = N'" & Declarations.MySupplierCode & "') "
        InitMyConn(False)
        InitMyRec(False, MySQLStr)
        If Declarations.MyRec.BOF = True And Declarations.MyRec.EOF = True Then
            CurrName = "RUB"
            trycloseMyRec()
        Else
            CurrName = Declarations.MyRec.Fields("SYCD009").Value
            trycloseMyRec()
        End If

        oSheet.getCellRangeByName("B15").String = "N п/п"
        oSheet.getCellRangeByName("C15").String = "Supp Item Code / Код товара поставщ"
        oSheet.getCellRangeByName("D15").String = "Item Code / Код товара"
        oSheet.getCellRangeByName("E15").String = "Item Name / Наименование товара"
        oSheet.getCellRangeByName("F15").String = "UOM / Ед. измер-я"
        oSheet.getCellRangeByName("G15").String = "QTY / Количество"
        oSheet.getCellRangeByName("H15").String = "Price / Цена заказа, " & CurrName
        oSheet.getCellRangeByName("I15").String = "Price / Цена поставщика"
        oSheet.getCellRangeByName("J15").String = "Summa / Сумма, " & CurrName
        oSheet.getCellRangeByName("K15").String = "Мин кол-во в заказе"
        oSheet.getCellRangeByName("L15").String = "складской"

        LOFontSetFamilyName(oServiceManager, oDispatcher, oFrame, "B15:L15", "Calibri")
        LOFontSetSize(oServiceManager, oDispatcher, oFrame, "B15:L15", 7)
        oSheet.getCellRangeByName("B15:L15").CellBackColor = 16775598
        Dim LineFormat As Object
        LineFormat = oServiceManager.Bridge_GetStruct("com.sun.star.table.BorderLine2")
        LineFormat.LineStyle = 0
        LineFormat.LineWidth = 70
        oSheet.getCellRangeByName("B15:L15").TopBorder = LineFormat
        oSheet.getCellRangeByName("B15:L15").RightBorder = LineFormat
        oSheet.getCellRangeByName("B15:L15").LeftBorder = LineFormat
        oSheet.getCellRangeByName("B15:L15").BottomBorder = LineFormat
        oSheet.getCellRangeByName("B15:L15").VertJustify = 2
        oSheet.getCellRangeByName("B15:L15").HoriJustify = 2
        LOWrapText(oServiceManager, oDispatcher, oFrame, "B15:L15")

        oSheet.getCellRangeByName("C16").String = "N заказа на продажу"
        oSheet.getCellRangeByName("D16").String = "N заказа на закупку"
        oSheet.getCellRangeByName("E16").String = "Продавец"
        oSheet.getCellRangeByName("F16").String = "UOM / Ед. измер-я"
        oSheet.getCellRangeByName("G16").String = "QTY / Количество"
        oSheet.getCellRangeByName("H16").String = "Дата отгрузки в заказе на продажу"
        oSheet.getCellRangeByName("I16").String = "Подтвержденная дата поставки"
        oSheet.getCellRangeByName("J16").String = "Подтвержденная задолженная дата"

        LOFontSetFamilyName(oServiceManager, oDispatcher, oFrame, "C16:J16", "Calibri")
        LOFontSetSize(oServiceManager, oDispatcher, oFrame, "C16:J16", 7)
        oSheet.getCellRangeByName("C16:J16").CellBackColor = 14741460
        LineFormat = oServiceManager.Bridge_GetStruct("com.sun.star.table.BorderLine2")
        LineFormat.LineStyle = 0
        LineFormat.LineWidth = 70
        oSheet.getCellRangeByName("C16:J16").TopBorder = LineFormat
        oSheet.getCellRangeByName("C16:J16").RightBorder = LineFormat
        oSheet.getCellRangeByName("C16:J16").LeftBorder = LineFormat
        oSheet.getCellRangeByName("C16:J16").BottomBorder = LineFormat
        oSheet.getCellRangeByName("C16:J16").VertJustify = 2
        oSheet.getCellRangeByName("C16:J16").HoriJustify = 2
        LOWrapText(oServiceManager, oDispatcher, oFrame, "C16:J16")

        i = 17
    End Sub

    Private Sub ExportOrderBodyToExcelFull(ByRef MyWRKBook As Object, ByVal ComOrder As String, ByRef i As Integer, ByRef j As Integer)
        '////////////////////////////////////////////////////////////////////////////////
        '//
        '// Процедура выгрузки тела обобщенного заказа в Excel (полная)
        '//
        '////////////////////////////////////////////////////////////////////////////////
        Dim MySQLStr As String

        j = 0
        MySQLStr = "IF exists(select * from tempdb..sysobjects where "
        MySQLStr = MySQLStr & "id = object_id(N'tempdb..#_MyPCOrder')  "
        MySQLStr = MySQLStr & "and xtype = N'U') "
        MySQLStr = MySQLStr & "DROP TABLE #_MyPCOrder "
        InitMyConn(False)
        Declarations.MyConn.Execute(MySQLStr)

        MySQLStr = "CREATE TABLE #_MyPCOrder( "
        MySQLStr = MySQLStr & "[PC03005] [nvarchar](35) COLLATE Cyrillic_General_BIN, "
        MySQLStr = MySQLStr & "[SC01060] [nvarchar](35) COLLATE Cyrillic_General_BIN, "
        MySQLStr = MySQLStr & "[PC03006] [nvarchar] (52) COLLATE Cyrillic_General_BIN, "
        MySQLStr = MySQLStr & "[QTY] [numeric](20, 8), "
        MySQLStr = MySQLStr & "[PC03009] [int], "
        MySQLStr = MySQLStr & "[PC03009_Name][nvarchar](10), "
        MySQLStr = MySQLStr & "[Price] [numeric](28, 8), "
        MySQLStr = MySQLStr & "[EUPrice] [numeric](28, 8), "
        MySQLStr = MySQLStr & "[StrSum] [numeric](28, 8), "
        MySQLStr = MySQLStr & "[OrderSum] [numeric](28, 8), "
        MySQLStr = MySQLStr & "[MinQTY] [numeric](28, 8), "
        MySQLStr = MySQLStr & "[IsWH] [nvarchar](25) "
        MySQLStr = MySQLStr & ") "
        InitMyConn(False)
        Declarations.MyConn.Execute(MySQLStr)

        MySQLStr = "EXEC spp_PurchaseWorkplace_PurchaseGroupOrderPreparation N'" & ComOrder & "' "
        InitMyConn(False)
        Declarations.MyConn.Execute(MySQLStr)

        MySQLStr = "SELECT * "
        MySQLStr = MySQLStr & "FROM #_MyPCOrder "
        MySQLStr = MySQLStr & "Order BY SC01060 "
        InitMyConn(False)
        InitMyRec(False, MySQLStr)
        If Declarations.MyRec.BOF = True And Declarations.MyRec.EOF = True Then
            trycloseMyRec()
        Else
            Declarations.MyRec.MoveFirst()
            MyPurchOrderSum = Declarations.MyRec.Fields("OrderSum").Value
            While Declarations.MyRec.EOF = False
                '--------------------Вывод общей строки
                '-----N п/п
                MyWRKBook.ActiveSheet.Range("B" & CStr(i)) = j + 1
                '-----код товара поставщика
                MyWRKBook.ActiveSheet.Range("C" & CStr(i)).NumberFormat = "@"
                MyWRKBook.ActiveSheet.Range("C" & CStr(i)) = Declarations.MyRec.Fields("SC01060").Value
                '-----код товара 
                MyWRKBook.ActiveSheet.Range("D" & CStr(i)).NumberFormat = "@"
                MyWRKBook.ActiveSheet.Range("D" & CStr(i)) = Declarations.MyRec.Fields("PC03005").Value
                '-----название товара
                MyWRKBook.ActiveSheet.Range("E" & CStr(i)).NumberFormat = "@"
                MyWRKBook.ActiveSheet.Range("E" & CStr(i)) = Declarations.MyRec.Fields("PC03006").Value
                '-----единица измерения товара
                MyWRKBook.ActiveSheet.Range("F" & CStr(i)).NumberFormat = "@"
                MyWRKBook.ActiveSheet.Range("F" & CStr(i)) = Declarations.MyRec.Fields("PC03009_Name").Value
                '-----количество товара
                MyWRKBook.ActiveSheet.Range("G" & CStr(i)) = Declarations.MyRec.Fields("QTY").Value
                '-----цена
                MyWRKBook.ActiveSheet.Range("H" & CStr(i)) = Declarations.MyRec.Fields("Price").Value
                '-----цена в евро(если есть)
                If Declarations.MyRec.Fields("EUPrice").Value = 0 Then
                Else
                    MyWRKBook.ActiveSheet.Range("I" & CStr(i)) = Declarations.MyRec.Fields("EUPrice").Value
                End If
                '-----сумма строки
                MyWRKBook.ActiveSheet.Range("J" & CStr(i)) = Declarations.MyRec.Fields("StrSum").Value
                '-----минимальное количество в заказе
                MyWRKBook.ActiveSheet.Range("K" & CStr(i)) = Declarations.MyRec.Fields("MinQTY").Value
                '-----Признак - складской
                MyWRKBook.ActiveSheet.Range("L" & CStr(i)).NumberFormat = "@"
                MyWRKBook.ActiveSheet.Range("L" & CStr(i)) = Declarations.MyRec.Fields("IsWH").Value

                MyWRKBook.ActiveSheet.Range("B" & CStr(i) & ":L" & CStr(i)).Select()
                MyWRKBook.ActiveSheet.Range("B" & CStr(i) & ":L" & CStr(i)).Borders(5).LineStyle = -4142
                MyWRKBook.ActiveSheet.Range("B" & CStr(i) & ":L" & CStr(i)).Borders(6).LineStyle = -4142
                With MyWRKBook.ActiveSheet.Range("B" & CStr(i) & ":L" & CStr(i)).Borders(7)
                    .LineStyle = 1
                    .Weight = 2
                    .ColorIndex = -4105
                End With
                With MyWRKBook.ActiveSheet.Range("B" & CStr(i) & ":L" & CStr(i)).Borders(8)
                    .LineStyle = 1
                    .Weight = 2
                    .ColorIndex = -4105
                End With
                With MyWRKBook.ActiveSheet.Range("B" & CStr(i) & ":L" & CStr(i)).Borders(9)
                    .LineStyle = 1
                    .Weight = 2
                    .ColorIndex = -4105
                End With
                With MyWRKBook.ActiveSheet.Range("B" & CStr(i) & ":L" & CStr(i)).Borders(10)
                    .LineStyle = 1
                    .Weight = 2
                    .ColorIndex = -4105
                End With
                With MyWRKBook.ActiveSheet.Range("B" & CStr(i) & ":L" & CStr(i)).Borders(11)
                    .LineStyle = 1
                    .Weight = 2
                    .ColorIndex = -4105
                End With
                With MyWRKBook.ActiveSheet.Range("B" & CStr(i) & ":L" & CStr(i)).Borders(12)
                    .LineStyle = 1
                    .Weight = 2
                    .ColorIndex = -4105
                End With

                MyWRKBook.ActiveSheet.Range("B" & CStr(i) & ":L" & CStr(i)).Select()
                With MyWRKBook.ActiveSheet.Range("B" & CStr(i) & ":L" & CStr(i)).Font
                    .Name = "Arial"
                    .Size = 7
                End With

                MyWRKBook.ActiveSheet.Range("C" & CStr(i) & ":D" & CStr(i)).Font.Bold = True
                MyWRKBook.ActiveSheet.Range("B" & CStr(i) & ":L" & CStr(i)).WrapText = True

                With MyWRKBook.ActiveSheet.Range("B" & CStr(i) & ":L" & CStr(i)).Interior
                    .Color = 14277081
                    .Pattern = 1
                    .PatternColorIndex = -4105
                End With


                '--------------------Вывод подробностей (детальные строки
                MySQLStr = "SELECT PC010300.PC01001, PC010300.PC01060, View_1.OR03005, View_1.OR03011, View_1.OR03037, View_1.Salesman, "
                MySQLStr = MySQLStr & "CASE WHEN PC030300.PC03029 = 1 THEN PC030300.PC03024 ELSE NULL END AS ConfirmedDate, "
                MySQLStr = MySQLStr & "CASE WHEN PC030300.PC03029 = 1 THEN PC030300.PC03031 ELSE NULL END AS BackOrderedDate "
                MySQLStr = MySQLStr & "FROM PC010300 INNER JOIN "
                MySQLStr = MySQLStr & "PC030300 ON PC010300.PC01001 = PC030300.PC03001 INNER JOIN "
                MySQLStr = MySQLStr & "(SELECT OR030300.OR03001, OR030300.OR03005, OR030300.OR03011, OR030300.OR03037, LTRIM(RTRIM(OR010300.OR01020)) "
                MySQLStr = MySQLStr & "+ ' ' + LTRIM(RTRIM(ST010300.ST01002)) AS Salesman "
                MySQLStr = MySQLStr & "FROM OR030300 INNER JOIN "
                MySQLStr = MySQLStr & "OR010300 ON OR030300.OR03001 = OR010300.OR01001 INNER JOIN "
                MySQLStr = MySQLStr & "ST010300 ON OR010300.OR01019 = ST010300.ST01001) AS View_1 ON PC010300.PC01060 = View_1.OR03001 AND "
                MySQLStr = MySQLStr & "PC030300.PC03005 = View_1.OR03005 "
                MySQLStr = MySQLStr & "WHERE (PC010300.PC01052 = N'" & ComOrder & "') AND "
                MySQLStr = MySQLStr & "(PC010300.PC01060 <> N'') AND "
                MySQLStr = MySQLStr & "(View_1.OR03005 = N'" & Declarations.MyRec.Fields("PC03005").Value & "') "
                InitMyConn(False)
                Dim MyRec1 = New ADODB.Recordset
                MyRec1.LockType = LockTypeEnum.adLockOptimistic
                MyRec1.Open(MySQLStr, Declarations.MyConn)
                If MyRec1.BOF = True And MyRec1.EOF = True Then
                    MyRec1.Close()
                    MyRec1 = Nothing
                Else
                    MyRec1.MoveFirst()
                    While MyRec1.EOF = False
                        '---Номер заказа на продажу
                        MyWRKBook.ActiveSheet.Range("C" & CStr(i + 1)).NumberFormat = "@"
                        MyWRKBook.ActiveSheet.Range("C" & CStr(i + 1)) = MyRec1.Fields("PC01060").Value
                        '---Номер заказа на закупку
                        MyWRKBook.ActiveSheet.Range("D" & CStr(i + 1)).NumberFormat = "@"
                        MyWRKBook.ActiveSheet.Range("D" & CStr(i + 1)) = MyRec1.Fields("PC01001").Value
                        '---Продавец
                        MyWRKBook.ActiveSheet.Range("E" & CStr(i + 1)).NumberFormat = "@"
                        MyWRKBook.ActiveSheet.Range("E" & CStr(i + 1)) = MyRec1.Fields("Salesman").Value
                        '-----единица измерения товара
                        MyWRKBook.ActiveSheet.Range("F" & CStr(i + 1)).NumberFormat = "@"
                        MyWRKBook.ActiveSheet.Range("F" & CStr(i + 1)) = Declarations.MyRec.Fields("PC03009_Name").Value
                        '---Количество
                        MyWRKBook.ActiveSheet.Range("G" & CStr(i + 1)) = MyRec1.Fields("OR03011").Value
                        '---Дата отгрузки
                        MyWRKBook.ActiveSheet.Range("H" & CStr(i + 1)) = MyRec1.Fields("OR03037").Value
                        '---Подтвержденная дата поставки
                        MyWRKBook.ActiveSheet.Range("I" & CStr(i + 1)) = MyRec1.Fields("ConfirmedDate").Value
                        '---Подтвержденная задолженная дата поставки
                        MyWRKBook.ActiveSheet.Range("J" & CStr(i + 1)) = MyRec1.Fields("BackOrderedDate").Value

                        MyWRKBook.ActiveSheet.Range("C" & CStr(i + 1) & ":J" & CStr(i + 1)).Select()
                        MyWRKBook.ActiveSheet.Range("C" & CStr(i + 1) & ":J" & CStr(i + 1)).Borders(5).LineStyle = -4142
                        MyWRKBook.ActiveSheet.Range("C" & CStr(i + 1) & ":J" & CStr(i + 1)).Borders(6).LineStyle = -4142
                        With MyWRKBook.ActiveSheet.Range("C" & CStr(i + 1) & ":J" & CStr(i + 1)).Borders(7)
                            .LineStyle = 1
                            .Weight = 2
                            .ColorIndex = -4105
                        End With
                        With MyWRKBook.ActiveSheet.Range("C" & CStr(i + 1) & ":J" & CStr(i + 1)).Borders(8)
                            .LineStyle = 1
                            .Weight = 2
                            .ColorIndex = -4105
                        End With
                        With MyWRKBook.ActiveSheet.Range("C" & CStr(i + 1) & ":J" & CStr(i + 1)).Borders(9)
                            .LineStyle = 1
                            .Weight = 2
                            .ColorIndex = -4105
                        End With
                        With MyWRKBook.ActiveSheet.Range("C" & CStr(i + 1) & ":J" & CStr(i + 1)).Borders(10)
                            .LineStyle = 1
                            .Weight = 2
                            .ColorIndex = -4105
                        End With
                        With MyWRKBook.ActiveSheet.Range("C" & CStr(i + 1) & ":J" & CStr(i + 1)).Borders(11)
                            .LineStyle = 1
                            .Weight = 2
                            .ColorIndex = -4105
                        End With
                        With MyWRKBook.ActiveSheet.Range("C" & CStr(i + 1) & ":J" & CStr(i + 1)).Borders(12)
                            .LineStyle = 1
                            .Weight = 2
                            .ColorIndex = -4105
                        End With

                        MyWRKBook.ActiveSheet.Range("C" & CStr(i + 1) & ":J" & CStr(i + 1)).Select()
                        With MyWRKBook.ActiveSheet.Range("C" & CStr(i + 1) & ":J" & CStr(i + 1)).Font
                            .Name = "Arial"
                            .Size = 7
                        End With

                        '---подсветка дат
                        Try
                            MyWRKBook.ActiveSheet.Range("H" & CStr(i + 1) & ":H" & CStr(i + 1)).FormatConditions.Add(Type:=1, Operator:=6, Formula1:="=$J$" & CStr(i + 1))
                            MyWRKBook.ActiveSheet.Range("H" & CStr(i + 1) & ":H" & CStr(i + 1)).FormatConditions(MyWRKBook.ActiveSheet.Range("H" & CStr(i + 1) & ":H" & CStr(i + 1)).FormatConditions.Count).SetFirstPriority()
                            MyWRKBook.ActiveSheet.Range("H" & CStr(i + 1) & ":H" & CStr(i + 1)).FormatConditions(1).Font.Bold = True
                            MyWRKBook.ActiveSheet.Range("H" & CStr(i + 1) & ":H" & CStr(i + 1)).FormatConditions(1).Font.Italic = False
                            MyWRKBook.ActiveSheet.Range("H" & CStr(i + 1) & ":H" & CStr(i + 1)).FormatConditions(1).Font.ThemeColor = 1
                            MyWRKBook.ActiveSheet.Range("H" & CStr(i + 1) & ":H" & CStr(i + 1)).FormatConditions(1).Font.TintAndShade = 0
                            MyWRKBook.ActiveSheet.Range("H" & CStr(i + 1) & ":H" & CStr(i + 1)).FormatConditions(1).Interior.PatternColorIndex = -4105
                            MyWRKBook.ActiveSheet.Range("H" & CStr(i + 1) & ":H" & CStr(i + 1)).FormatConditions(1).Interior.Color = 255
                            MyWRKBook.ActiveSheet.Range("H" & CStr(i + 1) & ":H" & CStr(i + 1)).FormatConditions(1).Interior.TintAndShade = 0
                            MyWRKBook.ActiveSheet.Range("H" & CStr(i + 1) & ":H" & CStr(i + 1)).FormatConditions(1).StopIfTrue = False
                        Catch ex As Exception
                        End Try

                        MyRec1.MoveNext()
                        i = i + 1
                    End While
                    MyRec1.Close()
                    MyRec1 = Nothing
                End If

                Declarations.MyRec.MoveNext()
                i = i + 1
                j = j + 1
            End While
            trycloseMyRec()

            MySQLStr = "IF exists(select * from tempdb..sysobjects where "
            MySQLStr = MySQLStr & "id = object_id(N'tempdb..#_MyPCOrder')  "
            MySQLStr = MySQLStr & "and xtype = N'U') "
            MySQLStr = MySQLStr & "DROP TABLE #_MyPCOrder "
            InitMyConn(False)
            Declarations.MyConn.Execute(MySQLStr)
        End If
    End Sub

    Private Sub ExportOrderBodyToLOFull(ByRef oSheet As Object, ByRef oServiceManager As Object, ByRef oWorkBook As Object, _
        ByRef oDispatcher As Object, ByVal ComOrder As String, ByRef i As Integer, ByRef j As Integer)
        '////////////////////////////////////////////////////////////////////////////////
        '//
        '// Процедура выгрузки тела обобщенного заказа в Libre Office (полная)
        '//
        '////////////////////////////////////////////////////////////////////////////////
        Dim MySQLStr As String
        Dim LineFormat As Object
        Dim oFrame As Object

        LineFormat = oServiceManager.Bridge_GetStruct("com.sun.star.table.BorderLine2")
        oFrame = oWorkBook.getCurrentController.getFrame
        j = 0
        MySQLStr = "IF exists(select * from tempdb..sysobjects where "
        MySQLStr = MySQLStr & "id = object_id(N'tempdb..#_MyPCOrder')  "
        MySQLStr = MySQLStr & "and xtype = N'U') "
        MySQLStr = MySQLStr & "DROP TABLE #_MyPCOrder "
        InitMyConn(False)
        Declarations.MyConn.Execute(MySQLStr)

        MySQLStr = "CREATE TABLE #_MyPCOrder( "
        MySQLStr = MySQLStr & "[PC03005] [nvarchar](35) COLLATE Cyrillic_General_BIN, "
        MySQLStr = MySQLStr & "[SC01060] [nvarchar](35) COLLATE Cyrillic_General_BIN, "
        MySQLStr = MySQLStr & "[PC03006] [nvarchar] (52) COLLATE Cyrillic_General_BIN, "
        MySQLStr = MySQLStr & "[QTY] [numeric](20, 8), "
        MySQLStr = MySQLStr & "[PC03009] [int], "
        MySQLStr = MySQLStr & "[PC03009_Name][nvarchar](10), "
        MySQLStr = MySQLStr & "[Price] [numeric](28, 8), "
        MySQLStr = MySQLStr & "[EUPrice] [numeric](28, 8), "
        MySQLStr = MySQLStr & "[StrSum] [numeric](28, 8), "
        MySQLStr = MySQLStr & "[OrderSum] [numeric](28, 8), "
        MySQLStr = MySQLStr & "[MinQTY] [numeric](28, 8), "
        MySQLStr = MySQLStr & "[IsWH] [nvarchar](25) "
        MySQLStr = MySQLStr & ") "
        InitMyConn(False)
        Declarations.MyConn.Execute(MySQLStr)

        MySQLStr = "EXEC spp_PurchaseWorkplace_PurchaseGroupOrderPreparation N'" & ComOrder & "' "
        InitMyConn(False)
        Declarations.MyConn.Execute(MySQLStr)

        MySQLStr = "SELECT * "
        MySQLStr = MySQLStr & "FROM #_MyPCOrder "
        MySQLStr = MySQLStr & "Order BY SC01060 "
        InitMyConn(False)
        InitMyRec(False, MySQLStr)
        If Declarations.MyRec.BOF = True And Declarations.MyRec.EOF = True Then
            trycloseMyRec()
        Else
            Declarations.MyRec.MoveFirst()
            MyPurchOrderSum = Declarations.MyRec.Fields("OrderSum").Value
            While Declarations.MyRec.EOF = False
                '--------------------Вывод общей строки
                '-----N п/п
                oSheet.getCellRangeByName("B" & CStr(i)).String = j + 1
                '-----код товара поставщика
                oSheet.getCellRangeByName("C" & CStr(i)).String = Declarations.MyRec.Fields("SC01060").Value
                '-----код товара
                oSheet.getCellRangeByName("D" & CStr(i)).String = Declarations.MyRec.Fields("PC03005").Value
                '-----название товара
                oSheet.getCellRangeByName("E" & CStr(i)).String = Declarations.MyRec.Fields("PC03006").Value
                '-----единица измерения товара
                oSheet.getCellRangeByName("F" & CStr(i)).String = Declarations.MyRec.Fields("PC03009_Name").Value
                '-----количество товара
                oSheet.getCellRangeByName("G" & CStr(i)).Value = Declarations.MyRec.Fields("QTY").Value
                '-----цена
                oSheet.getCellRangeByName("H" & CStr(i)).Value = Declarations.MyRec.Fields("Price").Value
                '-----цена в евро(если есть)
                If Declarations.MyRec.Fields("EUPrice").Value = 0 Then
                Else
                    oSheet.getCellRangeByName("I" & CStr(i)).Value = Declarations.MyRec.Fields("EUPrice").Value
                End If
                '-----сумма строки
                oSheet.getCellRangeByName("J" & CStr(i)).Value = Declarations.MyRec.Fields("StrSum").Value
                '-----минимальное количество в заказе
                If Not IsDBNull(Declarations.MyRec.Fields("MinQTY").Value) Then
                    oSheet.getCellRangeByName("K" & CStr(i)).Value = Declarations.MyRec.Fields("MinQTY").Value
                End If
                '-----Признак - складской
                oSheet.getCellRangeByName("L" & CStr(i)).String = Declarations.MyRec.Fields("IsWH").Value
                '-----форматы
                LOFontSetFamilyName(oServiceManager, oDispatcher, oFrame, "B" & i & ":L" & i, "Arial")
                LineFormat.LineStyle = 0
                LineFormat.LineWidth = 20
                oSheet.getCellRangeByName("B" & i & ":L" & i).TopBorder = LineFormat
                oSheet.getCellRangeByName("B" & i & ":L" & i).RightBorder = LineFormat
                oSheet.getCellRangeByName("B" & i & ":L" & i).LeftBorder = LineFormat
                oSheet.getCellRangeByName("B" & i & ":L" & i).BottomBorder = LineFormat
                LOFontSetSize(oServiceManager, oDispatcher, oFrame, "B" & i & ":L" & i, 7)
                LOWrapText(oServiceManager, oDispatcher, oFrame, "B" & i & ":L" & i)
                oSheet.getCellRangeByName("B" & i & ":L" & i).CellBackColor = 14540253

                '--------------------Вывод подробностей (детальные строки
                MySQLStr = "SELECT distinct PC010300.PC01001, PC010300.PC01060, View_1.OR03005, View_1.OR03011, View_1.OR03037, View_1.Salesman, "
                MySQLStr = MySQLStr & "CASE WHEN PC030300.PC03029 = 1 THEN PC030300.PC03024 ELSE NULL END AS ConfirmedDate, "
                MySQLStr = MySQLStr & "CASE WHEN PC030300.PC03029 = 1 THEN PC030300.PC03031 ELSE NULL END AS BackOrderedDate "
                MySQLStr = MySQLStr & "FROM PC010300 INNER JOIN "
                MySQLStr = MySQLStr & "PC030300 ON PC010300.PC01001 = PC030300.PC03001 INNER JOIN "
                MySQLStr = MySQLStr & "(SELECT OR030300.OR03001, OR030300.OR03005, OR030300.OR03011, OR030300.OR03037, LTRIM(RTRIM(OR010300.OR01020)) "
                MySQLStr = MySQLStr & "+ ' ' + LTRIM(RTRIM(ST010300.ST01002)) AS Salesman "
                MySQLStr = MySQLStr & "FROM OR030300 INNER JOIN "
                MySQLStr = MySQLStr & "OR010300 ON OR030300.OR03001 = OR010300.OR01001 INNER JOIN "
                MySQLStr = MySQLStr & "ST010300 ON OR010300.OR01019 = ST010300.ST01001) AS View_1 ON PC010300.PC01060 = View_1.OR03001 AND "
                MySQLStr = MySQLStr & "PC030300.PC03005 = View_1.OR03005 "
                MySQLStr = MySQLStr & "WHERE (PC010300.PC01052 = N'" & ComOrder & "') AND "
                MySQLStr = MySQLStr & "(PC010300.PC01060 <> N'') AND "
                MySQLStr = MySQLStr & "(View_1.OR03005 = N'" & Declarations.MyRec.Fields("PC03005").Value & "') "
                InitMyConn(False)
                Dim MyRec1 = New ADODB.Recordset
                MyRec1.LockType = LockTypeEnum.adLockOptimistic
                MyRec1.Open(MySQLStr, Declarations.MyConn)
                If MyRec1.BOF = True And MyRec1.EOF = True Then
                    MyRec1.Close()
                    MyRec1 = Nothing
                Else
                    MyRec1.MoveFirst()
                    While MyRec1.EOF = False
                        '---Номер заказа на продажу
                        oSheet.getCellRangeByName("C" & CStr(i + 1)).String = MyRec1.Fields("PC01060").Value
                        '---Номер заказа на закупку
                        oSheet.getCellRangeByName("D" & CStr(i + 1)).String = MyRec1.Fields("PC01001").Value
                        '---Продавец
                        oSheet.getCellRangeByName("E" & CStr(i + 1)).String = MyRec1.Fields("Salesman").Value
                        '-----единица измерения товара
                        oSheet.getCellRangeByName("F" & CStr(i + 1)).String = Declarations.MyRec.Fields("PC03009_Name").Value
                        '---Количество
                        oSheet.getCellRangeByName("G" & CStr(i + 1)).Value = MyRec1.Fields("OR03011").Value
                        '---Дата отгрузки
                        oSheet.getCellRangeByName("H" & CStr(i + 1)).Value = MyRec1.Fields("OR03037").Value
                        '---Подтвержденная дата поставки
                        If Not IsDBNull(MyRec1.Fields("ConfirmedDate").Value) Then
                            oSheet.getCellRangeByName("I" & CStr(i + 1)).Value = MyRec1.Fields("ConfirmedDate").Value
                        End If
                        '---Подтвержденная задолженная дата поставки
                        If Not IsDBNull(MyRec1.Fields("BackOrderedDate").Value) Then
                            oSheet.getCellRangeByName("J" & CStr(i + 1)).Value = MyRec1.Fields("BackOrderedDate").Value
                        End If
                        '-----форматы
                        LOFontSetFamilyName(oServiceManager, oDispatcher, oFrame, "C" & (i + 1) & ":J" & (i + 1), "Arial")
                        LineFormat.LineStyle = 0
                        LineFormat.LineWidth = 20
                        oSheet.getCellRangeByName("C" & (i + 1) & ":J" & (i + 1)).TopBorder = LineFormat
                        oSheet.getCellRangeByName("C" & (i + 1) & ":J" & (i + 1)).RightBorder = LineFormat
                        oSheet.getCellRangeByName("C" & (i + 1) & ":J" & (i + 1)).LeftBorder = LineFormat
                        oSheet.getCellRangeByName("C" & (i + 1) & ":J" & (i + 1)).BottomBorder = LineFormat
                        LOFontSetSize(oServiceManager, oDispatcher, oFrame, "C" & (i + 1) & ":J" & (i + 1), 7)
                        LOWrapText(oServiceManager, oDispatcher, oFrame, "C" & (i + 1) & ":J" & (i + 1))
                        LOFormatCells(oServiceManager, oDispatcher, oFrame, "H" & (i + 1) & ":J" & (i + 1), 36)

                        '---подсветка дат
                        If Not IsDBNull(MyRec1.Fields("BackOrderedDate").Value) Then
                            If MyRec1.Fields("OR03037").Value < MyRec1.Fields("BackOrderedDate").Value Then
                                oSheet.getCellRangeByName("H" & (i + 1)).CellBackColor = 16296599
                            End If
                        End If

                        MyRec1.MoveNext()
                        i = i + 1
                    End While
                    MyRec1.Close()
                    MyRec1 = Nothing
                End If
                Declarations.MyRec.MoveNext()
                i = i + 1
                j = j + 1
            End While
            trycloseMyRec()

            MySQLStr = "IF exists(select * from tempdb..sysobjects where "
            MySQLStr = MySQLStr & "id = object_id(N'tempdb..#_MyPCOrder')  "
            MySQLStr = MySQLStr & "and xtype = N'U') "
            MySQLStr = MySQLStr & "DROP TABLE #_MyPCOrder "
            InitMyConn(False)
            Declarations.MyConn.Execute(MySQLStr)
        End If
    End Sub

    Private Sub ExportOrderFooterToExcelFull(ByRef MyWRKBook As Object, ByVal ComOrder As String, ByRef i As Integer, ByRef j As Integer)
        '////////////////////////////////////////////////////////////////////////////////
        '//
        '// Процедура выгрузки подвала обобщенного заказа в Excel (полная)
        '//
        '////////////////////////////////////////////////////////////////////////////////
        Dim MySQLStr As String
        Dim CurrName As String
        Dim PurchaserName As String
        Dim MyLanguage As String

        '------Язык документа---------------------
        MySQLStr = "SELECT PL01027 AS CC "
        MySQLStr = MySQLStr & "FROM PL010300 "
        MySQLStr = MySQLStr & "WHERE (PL01001 = N'" & Trim(Label9.Text) & "') "
        InitMyConn(False)
        InitMyRec(False, MySQLStr)
        If Declarations.MyRec.BOF = True And Declarations.MyRec.EOF = True Then
            MyLanguage = "RUS"
            trycloseMyRec()
        Else
            MyLanguage = Declarations.MyRec.Fields("CC").Value
            trycloseMyRec()
        End If

        'j = i
        MySQLStr = "SELECT LTRIM(RTRIM(LTRIM(RTRIM(ISNULL(View_1.SYPD001, ''))) + ' ' + LTRIM(RTRIM(ISNULL(View_1.SYPD003, ''))))) AS Purchaser "
        MySQLStr = MySQLStr & "FROM tbl_SupplierCard0300 WITH(NOLOCK) LEFT OUTER JOIN "
        MySQLStr = MySQLStr & "(SELECT SYPD001, SYPD002, SYPD003 "
        MySQLStr = MySQLStr & "FROM SYPD0300 WITH(NOLOCK) "
        If MyLanguage = "RUS" Then
            MySQLStr = MySQLStr & "WHERE (SYPD002 = N'ENG')) "
        Else
            MySQLStr = MySQLStr & "WHERE (SYPD002 = N'RUS')) "
        End If
        MySQLStr = MySQLStr & "AS View_1 ON UPPER(tbl_SupplierCard0300.Purchaser) = UPPER(View_1.SYPD001) "
        MySQLStr = MySQLStr & "WHERE (tbl_SupplierCard0300.PL01001 = N'" & Declarations.MySupplierCode & "')"
        InitMyConn(False)
        InitMyRec(False, MySQLStr)
        If Declarations.MyRec.BOF = True And Declarations.MyRec.EOF = True Then
            PurchaserName = ""
            trycloseMyRec()
        Else
            PurchaserName = Declarations.MyRec.Fields("Purchaser").Value
            trycloseMyRec()
        End If

        MySQLStr = "SELECT SYCD0100.SYCD009 "
        MySQLStr = MySQLStr & "FROM PL010300 WITH(NOLOCK) INNER JOIN "
        MySQLStr = MySQLStr & "SYCD0100 ON PL010300.PL01026 = SYCD0100.SYCD001 "
        MySQLStr = MySQLStr & "WHERE (PL010300.PL01001 = N'" & Declarations.MySupplierCode & "') "
        InitMyConn(False)
        InitMyRec(False, MySQLStr)
        If Declarations.MyRec.BOF = True And Declarations.MyRec.EOF = True Then
            CurrName = "RUB"
            trycloseMyRec()
        Else
            CurrName = Declarations.MyRec.Fields("SYCD009").Value
            trycloseMyRec()
        End If

        MyWRKBook.ActiveSheet.Range("B" & CStr(i) & ":H" & CStr(i)).MergeCells = True
        MyWRKBook.ActiveSheet.Range("B" & CStr(i)) = "Total/Итого, " & CurrName & ":"
        MyWRKBook.ActiveSheet.Range("B" & CStr(i) & ":H" & CStr(i)).Font.Size = 10
        MyWRKBook.ActiveSheet.Range("B" & CStr(i) & ":H" & CStr(i)).Font.Bold = True
        MyWRKBook.ActiveSheet.Range("B" & CStr(i) & ":H" & CStr(i)).HorizontalAlignment = -4152
        MyWRKBook.ActiveSheet.Range("I" & CStr(i)) = Declarations.MyPurchOrderSum
        MyWRKBook.ActiveSheet.Range("I" & CStr(i)).Font.Size = 10
        MyWRKBook.ActiveSheet.Range("I" & CStr(i)).Font.Bold = True

        MyWRKBook.ActiveSheet.Rows(CStr(i + 1) & ":" & CStr(i + 1)).RowHeight = 5

        MyWRKBook.ActiveSheet.Range("B" & CStr(i + 2) & ":I" & CStr(i + 2)).MergeCells = True
        MyWRKBook.ActiveSheet.Range("B" & CStr(i + 2)) = "Total lines / Всего наименований " & CStr(j) & " In amount of / На сумму " & CStr(Declarations.MyPurchOrderSum) & " " & CurrName
        MyWRKBook.ActiveSheet.Range("B" & CStr(i + 4) & ":I" & CStr(i + 4)).MergeCells = True
        MyWRKBook.ActiveSheet.Range("B" & CStr(i + 4)) = "Purchaser / Закупщик:______________________" & PurchaserName


        MyWRKBook.ActiveSheet.Range("A1:L" & CStr(i + 4)).Select()
        Try
            MyWRKBook.Application.PrintCommunication = True
        Catch
        End Try
        Try
            MyWRKBook.ActiveSheet.PageSetup.PrintArea = "$A$1:$L$" & CStr(i + 4)
        Catch ex As Exception
        End Try
        Try
            MyWRKBook.Application.PrintCommunication = False
        Catch
        End Try
        Try
            MyWRKBook.ActiveSheet.PageSetup.FitToPagesWide = 1
        Catch ex As Exception
        End Try
        Try
            MyWRKBook.ActiveSheet.PageSetup.FitToPagesTall = 0
        Catch
        End Try
        Try
            MyWRKBook.Application.PrintCommunication = True
        Catch
        End Try
    End Sub

    Private Sub ExportOrderFooterToLOFull(ByRef oSheet As Object, ByRef oServiceManager As Object, ByRef oWorkBook As Object, _
        ByRef oDispatcher As Object, ByVal ComOrder As String, ByRef i As Integer, ByRef j As Integer)
        '////////////////////////////////////////////////////////////////////////////////
        '//
        '// Процедура выгрузки подвала обобщенного заказа в Libre Office (полная)
        '//
        '////////////////////////////////////////////////////////////////////////////////
        Dim MyLanguage As String
        Dim MySQLStr As String
        Dim CurrName As String
        Dim PurchaserName As String
        Dim oFrame As Object

        oFrame = oWorkBook.getCurrentController.getFrame

        '------Язык документа---------------------
        MySQLStr = "SELECT PL01027 AS CC "
        MySQLStr = MySQLStr & "FROM PL010300 "
        MySQLStr = MySQLStr & "WHERE (PL01001 = N'" & Trim(Label9.Text) & "') "
        InitMyConn(False)
        InitMyRec(False, MySQLStr)
        If Declarations.MyRec.BOF = True And Declarations.MyRec.EOF = True Then
            MyLanguage = "RUS"
            trycloseMyRec()
        Else
            MyLanguage = Declarations.MyRec.Fields("CC").Value
            trycloseMyRec()
        End If

        'j = i
        MySQLStr = "SELECT LTRIM(RTRIM(LTRIM(RTRIM(ISNULL(View_1.SYPD001, ''))) + ' ' + LTRIM(RTRIM(ISNULL(View_1.SYPD003, ''))))) AS Purchaser "
        MySQLStr = MySQLStr & "FROM tbl_SupplierCard0300 WITH(NOLOCK) LEFT OUTER JOIN "
        MySQLStr = MySQLStr & "(SELECT SYPD001, SYPD002, SYPD003 "
        MySQLStr = MySQLStr & "FROM SYPD0300 WITH(NOLOCK) "
        If MyLanguage = "RUS" Then
            MySQLStr = MySQLStr & "WHERE (SYPD002 = N'ENG')) "
        Else
            MySQLStr = MySQLStr & "WHERE (SYPD002 = N'RUS')) "
        End If
        MySQLStr = MySQLStr & "AS View_1 ON UPPER(tbl_SupplierCard0300.Purchaser) = UPPER(View_1.SYPD001) "
        MySQLStr = MySQLStr & "WHERE (tbl_SupplierCard0300.PL01001 = N'" & Declarations.MySupplierCode & "')"
        InitMyConn(False)
        InitMyRec(False, MySQLStr)
        If Declarations.MyRec.BOF = True And Declarations.MyRec.EOF = True Then
            PurchaserName = ""
            trycloseMyRec()
        Else
            PurchaserName = Declarations.MyRec.Fields("Purchaser").Value
            trycloseMyRec()
        End If

        MySQLStr = "SELECT SYCD0100.SYCD009 "
        MySQLStr = MySQLStr & "FROM PL010300 WITH(NOLOCK) INNER JOIN "
        MySQLStr = MySQLStr & "SYCD0100 ON PL010300.PL01026 = SYCD0100.SYCD001 "
        MySQLStr = MySQLStr & "WHERE (PL010300.PL01001 = N'" & Declarations.MySupplierCode & "') "
        InitMyConn(False)
        InitMyRec(False, MySQLStr)
        If Declarations.MyRec.BOF = True And Declarations.MyRec.EOF = True Then
            CurrName = "RUB"
            trycloseMyRec()
        Else
            CurrName = Declarations.MyRec.Fields("SYCD009").Value
            trycloseMyRec()
        End If

        LOMergeCells(oServiceManager, oDispatcher, oFrame, "B" & CStr(i) & ":H" & CStr(i))
        oSheet.getCellRangeByName("B" & CStr(i)).String = "Total/Итого, " & CurrName & ":"
        oSheet.getCellRangeByName("I" & CStr(i)).String = Declarations.MyPurchOrderSum
        LOFontSetFamilyName(oServiceManager, oDispatcher, oFrame, "B" & CStr(i) & ":I" & CStr(i), "Arial")
        LOFontSetBold(oServiceManager, oDispatcher, oFrame, "B" & CStr(i) & ":I" & CStr(i))
        LOFontSetSize(oServiceManager, oDispatcher, oFrame, "B" & CStr(i) & ":I" & CStr(i), 10)
        oSheet.getCellRangeByName("B" & CStr(i) & ":I" & CStr(i)).VertJustify = 2

        i = i + 1
        oSheet.getCellRangeByName("A" & CStr(i)).Rows.Height = 200

        i = i + 1
        LOMergeCells(oServiceManager, oDispatcher, oFrame, "B" & CStr(i) & ":I" & CStr(i))
        oSheet.getCellRangeByName("B" & CStr(i)).String = "Total lines / Всего наименований " & CStr(j) & " In amount of / На сумму " & CStr(Declarations.MyPurchOrderSum) & " " & CurrName
        LOFontSetSize(oServiceManager, oDispatcher, oFrame, "B" & CStr(i), 10)
        oSheet.getCellRangeByName("B" & CStr(i)).VertJustify = 2

        i = i + 2
        LOMergeCells(oServiceManager, oDispatcher, oFrame, "B" & CStr(i) & ":I" & CStr(i))
        oSheet.getCellRangeByName("B" & CStr(i)).String = "Purchaser / Закупщик:______________________" & PurchaserName
        LOFontSetSize(oServiceManager, oDispatcher, oFrame, "B" & CStr(i), 10)
        oSheet.getCellRangeByName("B" & CStr(i)).VertJustify = 2
    End Sub

    Private Sub Button16_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button16.Click
        '////////////////////////////////////////////////////////////////////////////////
        '//
        '// Процедура рассылки уведомления об изменении дат поставки
        '//
        '////////////////////////////////////////////////////////////////////////////////
        Dim MyRez As Object

        MyRez = MsgBox("Вы уже загрузили новые даты поставки?", MsgBoxStyle.YesNo, "Внимание!")
        If MyRez = vbYes Then
            MyCreateSPTasks = New CreateSPTasks
            MyCreateSPTasks.CommonPurchOrderNum = Trim(DataGridView1.SelectedRows.Item(0).Cells(0).Value.ToString())
            MyCreateSPTasks.ShowDialog()
        End If
    End Sub

    Private Sub ComboBox2_SelectedValueChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles ComboBox2.SelectedValueChanged
        '////////////////////////////////////////////////////////////////////////////////
        '//
        '// Смена выбора - все заказы или только активные
        '//
        '////////////////////////////////////////////////////////////////////////////////

        Declarations.MyWH = ComboBox2.SelectedValue
        LoadConsolidatedOrders()
        LoadIncludedOrders()
        LoadFreeOrders()
        CheckConsolidatedButtons()
        CheckRemoveButtons()
        CheckAddButtons()
    End Sub

    Private Sub Button17_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button17.Click
        '////////////////////////////////////////////////////////////////////////////////
        '//
        '// Создание нового обобщенного заказа копированием существующего
        '//
        '////////////////////////////////////////////////////////////////////////////////

        MyEditConsolidatedOrder = New EditConsolidatedOrder
        MyEditConsolidatedOrder.StartParam = "Copy"
        MyEditConsolidatedOrder.ShowDialog()
        LoadConsolidatedOrders()
        '---текущей строкой сделать созданную
        For i As Integer = 0 To DataGridView1.Rows.Count - 1
            If Trim(DataGridView1.Item(0, i).Value.ToString) = Declarations.MyOrderID Then
                DataGridView1.CurrentCell = DataGridView1.Item(0, i)
            End If
        Next
        LoadIncludedOrders()
        LoadFreeOrders()
        CheckConsolidatedButtons()
        CheckRemoveButtons()
        CheckAddButtons()
    End Sub

    Private Sub Button19_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button19.Click
        '////////////////////////////////////////////////////////////////////////////////
        '//
        '// Открытие соответствующего файла.
        '//
        '////////////////////////////////////////////////////////////////////////////////

        If System.IO.Directory.Exists(My.Settings.BillPath + Trim(DataGridView1.SelectedRows.Item(0).Cells(14).Value.ToString())) = False _
            And System.IO.File.Exists(My.Settings.BillPath + Trim(DataGridView1.SelectedRows.Item(0).Cells(14).Value.ToString())) = False Then
            MsgBox("Файл или каталог " + My.Settings.BillPath + Trim(DataGridView1.SelectedRows.Item(0).Cells(14).Value.ToString()) + " не найден.", MsgBoxStyle.Critical, "Внимание!")
        Else
            If System.IO.Directory.Exists(My.Settings.BillPath + Trim(DataGridView1.SelectedRows.Item(0).Cells(14).Value.ToString())) Then
                Try
                    Process.Start("explorer.exe", My.Settings.BillPath + Trim(DataGridView1.SelectedRows.Item(0).Cells(14).Value.ToString()))
                Catch ex As Exception
                    MsgBox(ex.Message, MsgBoxStyle.Critical, "Внимание!")
                End Try
            Else
                Try
                    Dim startInfo As New ProcessStartInfo("CMD.EXE")
                    startInfo.WindowStyle = ProcessWindowStyle.Hidden
                    startInfo.CreateNoWindow = True
                    startInfo.UseShellExecute = False
                    startInfo.Arguments = "/c " + """" + My.Settings.BillPath + Trim(DataGridView1.SelectedRows.Item(0).Cells(14).Value.ToString()) + """"
                    Process.Start(startInfo)
                Catch ex As Exception
                    MsgBox(ex.Message, MsgBoxStyle.Critical, "Внимание!")
                End Try
            End If
        End If
    End Sub

    Private Sub Button18_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button18.Click
        '////////////////////////////////////////////////////////////////////////////////
        '//
        '// Открытие окна с заказами поставщика по всем складам
        '//
        '////////////////////////////////////////////////////////////////////////////////

        MySupplierInfo = New SupplierInfo
        MySupplierInfo.ShowDialog()
    End Sub
End Class