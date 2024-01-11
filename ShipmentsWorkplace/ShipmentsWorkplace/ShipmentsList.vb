Imports System.Net
Imports System.Xml

Public Class ShipmentsList
    Public LoadFlag As Integer

    Private Sub Button13_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button13.Click
        '////////////////////////////////////////////////////////////////////////////////
        '//
        '// Выход
        '//
        '////////////////////////////////////////////////////////////////////////////////

        Me.Close()
    End Sub

    Private Sub ShipmentsList_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles Me.KeyDown
        '////////////////////////////////////////////////////////////////////////////////
        '//
        '// Запрет закрытия окна по ALT+F4
        '//
        '////////////////////////////////////////////////////////////////////////////////

        If e.KeyData = Keys.Alt + Keys.F4 Then
            e.Handled = True
        End If
    End Sub

    Private Sub ShipmentsList_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        '////////////////////////////////////////////////////////////////////////////////
        '//
        '// открытиие окна формирования доставок / самовывозов
        '//
        '////////////////////////////////////////////////////////////////////////////////
        Dim MySQLStr As String                        'рабочая строка
        Dim MyAdapter As SqlClient.SqlDataAdapter     'для списка складов
        Dim MyDs As New DataSet                       '

        LoadFlag = 1
        '------------------склады---------------------
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

        Label9.Text = Declarations.SalesmanCode & " " & Declarations.UserName
        Label2.Text = Trim(MainForm.DataGridView1.SelectedRows.Item(0).Cells(0).Value.ToString()) & " " & Trim(MainForm.DataGridView1.SelectedRows.Item(0).Cells(1).Value.ToString())
        'ComboBoxAN.SelectedText = "Только активные (неотгруженные)"
        ComboBoxAN.SelectedIndex = 0 '--???!!!
        ComboBox1.SelectedValue = Declarations.MyWH
        LoadFlag = 0

        Application.DoEvents()
        Windows.Forms.Cursor.Current = Cursors.WaitCursor
        LoadShipments()
        LoadIncludedOrders()
        LoadFreeOrders()
        CheckShipmentsButtons()
        CheckRemoveButtons()
        CheckAddButtons()
        Windows.Forms.Cursor.Current = Cursors.Default
    End Sub

    Private Sub LoadShipments()
        '////////////////////////////////////////////////////////////////////////////////
        '//
        '// Загрузка данных по отгрузкам / доставкам
        '//
        '////////////////////////////////////////////////////////////////////////////////
        Dim MySQLStr As String
        Dim MyAdapter As SqlClient.SqlDataAdapter     'для списка поставщиков
        Dim MyDs As New DataSet                       '

        If LoadFlag = 0 Then
            If ComboBoxAN.Text = "Все доставки / самовывозы" Then
                MySQLStr = "Exec spp_Shipments_SalesmanWP_ShipmentInfo N'" & Trim(Declarations.MyCustomerCode) & "', N'" & Trim(Declarations.MyWH) & "', " & Declarations.SalesmanCode & ", 0, " & Declarations.MyGroupOrIndividualFlag
            Else
                MySQLStr = "Exec spp_Shipments_SalesmanWP_ShipmentInfo N'" & Trim(Declarations.MyCustomerCode) & "', N'" & Trim(Declarations.MyWH) & "', " & Declarations.SalesmanCode & ", 1, " & Declarations.MyGroupOrIndividualFlag
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
            DataGridView1.Columns(0).HeaderText = "N отгрузки"
            DataGridView1.Columns(0).Width = 70
            DataGridView1.Columns(1).HeaderText = "Код поку пателя"
            DataGridView1.Columns(1).Width = 70
            DataGridView1.Columns(2).HeaderText = "Покупатель"
            DataGridView1.Columns(2).Width = 175
            DataGridView1.Columns(3).HeaderText = "Доставка"
            DataGridView1.Columns(3).Width = 100
            DataGridView1.Columns(4).HeaderText = "Сумма на доставку"
            DataGridView1.Columns(4).Width = 80
            DataGridView1.Columns(4).DefaultCellStyle.Format = "n2"
            DataGridView1.Columns(5).HeaderText = "Сумма отгружаемого"
            DataGridView1.Columns(5).Width = 100
            DataGridView1.Columns(5).DefaultCellStyle.Format = "n2"
            DataGridView1.Columns(6).HeaderText = "Контактная информация"
            DataGridView1.Columns(6).Width = 205
            DataGridView1.Columns(7).HeaderText = "Адрес доставки"
            DataGridView1.Columns(7).Width = 265
            DataGridView1.Columns(8).HeaderText = "Комментарий складу"
            DataGridView1.Columns(8).Width = 195
            DataGridView1.Columns(9).HeaderText = "Печать счета"
            DataGridView1.Columns(9).Width = 70
            DataGridView1.Columns(10).HeaderText = "Печать справки - счета"
            DataGridView1.Columns(10).Width = 70
            DataGridView1.Columns(11).HeaderText = "Печать полного счета (восст.)"
            DataGridView1.Columns(11).Width = 70
            DataGridView1.Columns(12).HeaderText = "запро шенная дата поставки"
            DataGridView1.Columns(12).Width = 80
            DataGridView1.Columns(13).HeaderText = "Запрос на портал"
            DataGridView1.Columns(13).Width = 70
            DataGridView1.Columns(14).HeaderText = "Уведом ление клиенту"
            DataGridView1.Columns(14).Width = 70
            DataGridView1.Columns(15).HeaderText = "отгрузка произ ведена"
            DataGridView1.Columns(15).Width = 70
            DataGridView1.Columns(16).HeaderText = "Файл"
            DataGridView1.Columns(16).Width = 60
            DataGridView1.Columns(17).HeaderText = "Путь к файлу"
            DataGridView1.Columns(17).Visible = False
            DataGridView1.Columns(18).HeaderText = "Комментарий по транспорту"
            DataGridView1.Columns(18).Width = 195
            DataGridView1.Columns(19).HeaderText = "Комментарий по документам"
            DataGridView1.Columns(19).Width = 195

            FormatDataGridView1()
        End If
    End Sub

    Private Sub FormatDataGridView1()
        '////////////////////////////////////////////////////////////////////////////////
        '//
        '// подсветка информации по доставкам
        '//
        '////////////////////////////////////////////////////////////////////////////////
        Dim i As Integer

        For i = 0 To DataGridView1.Rows.Count - 1
            If DataGridView1.Rows(i).Cells(15).Value = 0 Then
                DataGridView1.Rows(i).DefaultCellStyle.BackColor = Color.White
                If DataGridView1.Rows(i).Cells(3).Value = "Доставка" Then
                    If DateDiff(DateInterval.Day, DataGridView1.Rows(i).Cells(12).Value, Now()) > 2 Then
                        DataGridView1.Rows(i).Cells(15).Style.BackColor = Color.LightPink
                    Else
                        DataGridView1.Rows(i).Cells(15).Style.BackColor = Color.White
                    End If

                Else            '---самовывоз или доставка с оплатой клиентом
                    If DateDiff(DateInterval.Day, DataGridView1.Rows(i).Cells(12).Value, Now()) > 7 Then
                        DataGridView1.Rows(i).Cells(15).Style.BackColor = Color.LightPink
                    Else
                        DataGridView1.Rows(i).Cells(15).Style.BackColor = Color.White
                    End If
                End If
            Else
                DataGridView1.Rows(i).DefaultCellStyle.BackColor = Color.LightGray
            End If
        Next
    End Sub

    Private Sub LoadIncludedOrders()
        '////////////////////////////////////////////////////////////////////////////////
        '//
        '// загрузка данных по заказам, включенным в отгрузку
        '//
        '////////////////////////////////////////////////////////////////////////////////
        Dim MySQLStr As String
        Dim MyAdapter As SqlClient.SqlDataAdapter     'для списка поставщиков
        Dim MyDs As New DataSet                       '

        If LoadFlag = 0 Then
            If DataGridView1.SelectedRows.Count <> 0 Then
                MySQLStr = "Exec spp_Shipments_SalesmanWP_ShipmentDetails N'" & Trim(DataGridView1.SelectedRows.Item(0).Cells(0).Value.ToString()) & "' "
            Else
                MySQLStr = "Exec spp_Shipments_SalesmanWP_ShipmentDetails N'0' "
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
            DataGridView2.Columns(0).HeaderText = "N з-за"
            DataGridView2.Columns(0).Width = 70
            DataGridView2.Columns(1).HeaderText = "Тип з-за"
            DataGridView2.Columns(1).Width = 30
            DataGridView2.Columns(2).HeaderText = "Продавец"
            DataGridView2.Columns(2).Width = 150
            DataGridView2.Columns(3).HeaderText = "Дата отгрузки"
            DataGridView2.Columns(3).Width = 80
            DataGridView2.Columns(4).HeaderText = "Сумма доставки (остаток)"
            DataGridView2.Columns(4).Width = 100
            DataGridView2.Columns(4).DefaultCellStyle.Format = "n2"
            DataGridView2.Columns(5).HeaderText = "Сумма распреде ленного"
            DataGridView2.Columns(5).Width = 100
            DataGridView2.Columns(5).DefaultCellStyle.Format = "n2"

            ReCalculateDelSumm()
        End If
    End Sub

    Private Sub ReCalculateDelSumm()
        '////////////////////////////////////////////////////////////////////////////////
        '//
        '// Вычисление суммы доставляемого и запланированной суммы на оплату перевозки
        '//
        '////////////////////////////////////////////////////////////////////////////////
        Dim DelSumm As Double
        Dim DelCostSumm As Double
        Dim i As Integer

        DelSumm = 0
        DelCostSumm = 0

        For i = 0 To DataGridView2.Rows.Count - 1
            DelSumm = DelSumm + DataGridView2.Rows.Item(i).Cells(5).Value
            DelCostSumm = DelCostSumm + DataGridView2.Rows.Item(i).Cells(4).Value
        Next

        TextBox1.Text = Math.Round(DelSumm, 2)
        TextBox2.Text = Math.Round(DelCostSumm, 2)
    End Sub

    Private Sub LoadFreeOrders()
        '////////////////////////////////////////////////////////////////////////////////
        '//
        '// загрузка данных по заказам, доступным к включению в отгрузки
        '//
        '////////////////////////////////////////////////////////////////////////////////
        Dim MySQLStr As String
        Dim MyAdapter As SqlClient.SqlDataAdapter     'для списка поставщиков
        Dim MyDs As New DataSet                       '

        If LoadFlag = 0 Then
            MySQLStr = "Exec spp_Shipments_SalesmanWP_AvlOrders N'" & Trim(Declarations.MyCustomerCode) & "', N'" & Trim(Declarations.MyWH) & "', " & Declarations.MyGroupOrIndividualFlag & ", N'" & Declarations.SalesmanCode & "' "

            Try
                MyAdapter = New SqlClient.SqlDataAdapter(MySQLStr, Declarations.MyNETConnStr)
                MyAdapter.SelectCommand.CommandTimeout = 600
                MyAdapter.Fill(MyDs)
                DataGridView3.DataSource = MyDs.Tables(0)
            Catch ex As Exception
                MsgBox(ex.ToString)
            End Try

            '---заголовки
            DataGridView3.Columns(0).HeaderText = "N з-за"
            DataGridView3.Columns(0).Width = 70
            DataGridView3.Columns(1).HeaderText = "Тип з-за"
            DataGridView3.Columns(1).Width = 30
            DataGridView3.Columns(2).HeaderText = "Продавец"
            DataGridView3.Columns(2).Width = 150
            DataGridView3.Columns(3).HeaderText = "Дата отгрузки"
            DataGridView3.Columns(3).DefaultCellStyle.Format = "dd/MM/yyyy"
            DataGridView3.Columns(3).Width = 80
            DataGridView3.Columns(4).HeaderText = "Макс дата прихода товара на склад"
            DataGridView3.Columns(4).Width = 80
            DataGridView3.Columns(4).DefaultCellStyle.Format = "dd/MM/yyyy"
            DataGridView3.Columns(5).HeaderText = "Макс задолж. дата прихода товара на склад"
            DataGridView3.Columns(5).Width = 80
            DataGridView3.Columns(5).DefaultCellStyle.Format = "dd/MM/yyyy"
            DataGridView3.Columns(6).HeaderText = "Сумма доставки (остаток)"
            DataGridView3.Columns(6).Width = 100
            DataGridView3.Columns(6).DefaultCellStyle.Format = "n2"
            DataGridView3.Columns(7).HeaderText = "Сумма заказа"
            DataGridView3.Columns(7).Width = 100
            DataGridView3.Columns(7).DefaultCellStyle.Format = "n2"
            DataGridView3.Columns(8).HeaderText = "Сумма распреде ленного"
            DataGridView3.Columns(8).Width = 100
            DataGridView3.Columns(8).DefaultCellStyle.Format = "n2"
            DataGridView3.Columns(9).HeaderText = "Разре шение на от грузку"
            DataGridView3.Columns(9).Width = 50
            DataGridView3.Columns(10).HeaderText = "Комментарий"
            DataGridView3.Columns(10).Width = 150

            FormatDataGridView3()
        End If
    End Sub

    Private Sub FormatDataGridView3()
        '////////////////////////////////////////////////////////////////////////////////
        '//
        '// подсветка информации по доступным заказам
        '//
        '////////////////////////////////////////////////////////////////////////////////

        For Each row As DataGridViewRow In DataGridView3.Rows
            If Trim(row.Cells(9).Value.ToString) <> "" Then
                row.Cells(9).Style.BackColor = Color.LightGreen
            Else
                row.Cells(9).Style.BackColor = Color.LightPink
            End If
            If row.Cells(3).Value < Now Then
                row.Cells(3).Style.BackColor = Color.LightGreen
            Else
                row.Cells(3).Style.BackColor = Color.Empty
            End If
            If IsDBNull(row.Cells(4).Value) = False Then
                If row.Cells(3).Value < row.Cells(4).Value Then
                    row.Cells(4).Style.BackColor = Color.LightYellow
                Else
                    row.Cells(4).Style.BackColor = Color.Empty
                End If
            Else
                row.Cells(4).Style.BackColor = Color.Empty
            End If
            If IsDBNull(row.Cells(5).Value) = False Then
                If row.Cells(3).Value < row.Cells(5).Value Then
                    row.Cells(5).Style.BackColor = Color.LightYellow
                Else
                    row.Cells(5).Style.BackColor = Color.Empty
                End If
            Else
                row.Cells(5).Style.BackColor = Color.Empty
            End If
            If row.Cells(7).Value = 0 Then
                row.Cells(7).Style.BackColor = Color.LightPink
            Else
                row.Cells(7).Style.BackColor = Color.Empty
            End If
            If row.Cells(8).Value = 0 Then
                row.Cells(8).Style.BackColor = Color.LightPink
            Else
                If row.Cells(8).Value < row.Cells(7).Value Then
                    row.Cells(8).Style.BackColor = Color.LightYellow
                Else
                    row.Cells(8).Style.BackColor = Color.Empty
                End If
            End If
        Next
    End Sub

    Private Sub CheckShipmentsButtons()
        '////////////////////////////////////////////////////////////////////////////////
        '//
        '// Проверка и выставление состояния кнопок отгрузок / доставок
        '//
        '////////////////////////////////////////////////////////////////////////////////

        If DataGridView1.SelectedRows.Count = 0 Then
            Button10.Enabled = False
            Button7.Enabled = False
            Button4.Enabled = False
            Button6.Enabled = False
            Button12.Enabled = False
            Button14.Enabled = False
        Else
            Button10.Enabled = True
            If DataGridView1.SelectedRows.Item(0).Cells(13).Value = 0 Then '---запрос на портал не сделан
                If DataGridView1.SelectedRows.Item(0).Cells(15).Value = 0 Then  '---отгрузка не произведена
                    Button7.Enabled = True
                    Button6.Enabled = True
                    If DataGridView1.SelectedRows.Item(0).Cells(5).Value = 0 Then
                        Button12.Enabled = False
                    Else
                        Button12.Enabled = True
                    End If
                Else
                    Button7.Enabled = False
                    Button6.Enabled = False
                    Button12.Enabled = False
                End If
            Else
                Button7.Enabled = False
                Button6.Enabled = False
                Button12.Enabled = False
            End If
            If DataGridView1.SelectedRows.Item(0).Cells(15).Value = 0 Then  '---отгрузка не произведена
                If DataGridView1.SelectedRows.Item(0).Cells(13).Value = 0 Then  '---запрос на портал не сделан
                    Button4.Enabled = False
                Else
                    Button4.Enabled = True
                End If
            Else
                Button4.Enabled = False
            End If
            If DataGridView1.SelectedRows.Item(0).Cells(14).Value = 0 Then  '---уведомление клиенту не отправлено
                If DataGridView1.SelectedRows.Item(0).Cells(15).Value = 0 Then
                    If DataGridView1.SelectedRows.Item(0).Cells(5).Value = 0 Then
                        Button14.Enabled = False
                    Else
                        Button14.Enabled = True
                    End If
                Else
                    Button14.Enabled = False
                End If
            Else
                Button14.Enabled = False
            End If
        End If
    End Sub

    Private Sub CheckRemoveButtons()
        '////////////////////////////////////////////////////////////////////////////////
        '//
        '// Проверка и выставление состояния кнопок удаления заказов из отгрузки
        '//
        '////////////////////////////////////////////////////////////////////////////////

        If DataGridView2.SelectedRows.Count = 0 Then
            Button11.Enabled = False
            Button5.Enabled = False
        Else
            If DataGridView1.SelectedRows.Item(0).Cells(15).Value = 0 _
                And DataGridView1.SelectedRows.Item(0).Cells(13).Value = 0 Then
                Button11.Enabled = True
            Else
                Button11.Enabled = False
            End If
            Button5.Enabled = True
        End If
    End Sub

    Private Sub CheckAddButtons()
        '////////////////////////////////////////////////////////////////////////////////
        '//
        '// Проверка и выставление состояния кнопок добавления заказов в отгрузку и доп действий
        '//
        '////////////////////////////////////////////////////////////////////////////////

        If DataGridView3.SelectedRows.Count = 0 Then
            Button9.Enabled = False
            Button2.Enabled = False
            Button3.Enabled = False
            Button15.Enabled = False
        Else
            If DataGridView1.SelectedRows.Count = 0 Then
                Button9.Enabled = False
            Else
                If DataGridView1.SelectedRows.Item(0).Cells(15).Value = 0 _
                And DataGridView1.SelectedRows.Item(0).Cells(13).Value = 0 Then
                    Button9.Enabled = True
                Else
                    Button9.Enabled = False
                End If
            End If
            Button2.Enabled = True
            Button3.Enabled = True
            Button15.Enabled = True
        End If
    End Sub

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        '////////////////////////////////////////////////////////////////////////////////
        '//
        '// Обновление информации в окне
        '//
        '////////////////////////////////////////////////////////////////////////////////

        Windows.Forms.Cursor.Current = Cursors.WaitCursor
        LoadShipments()
        LoadIncludedOrders()
        LoadFreeOrders()
        CheckShipmentsButtons()
        CheckRemoveButtons()
        CheckAddButtons()
        Windows.Forms.Cursor.Current = Cursors.Default
    End Sub

    Private Sub ComboBox1_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles ComboBox1.SelectedIndexChanged
        '////////////////////////////////////////////////////////////////////////////////
        '//
        '// Загрузка данных в соответствии с выбранным складом
        '//
        '////////////////////////////////////////////////////////////////////////////////

        If LoadFlag = 0 Then
            Windows.Forms.Cursor.Current = Cursors.WaitCursor
            Declarations.MyWH = Trim(ComboBox1.SelectedValue)
            LoadShipments()
            LoadIncludedOrders()
            LoadFreeOrders()
            CheckShipmentsButtons()
            CheckRemoveButtons()
            CheckAddButtons()
            Windows.Forms.Cursor.Current = Cursors.Default
        End If
    End Sub

    Private Sub Button2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button2.Click
        '////////////////////////////////////////////////////////////////////////////////
        '//
        '// Открытие окна с детальной информацией по выбранному заказу (первому из подсвеченных)
        '//
        '////////////////////////////////////////////////////////////////////////////////

        Try
            ShowOrderDetails(Trim(DataGridView3.SelectedRows.Item(0).Cells(0).Value.ToString()))
        Catch ex As Exception
        End Try
    End Sub

    Private Sub Button5_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button5.Click
        '////////////////////////////////////////////////////////////////////////////////
        '//
        '// Открытие окна с детальной информацией по выбранному заказу (первому из подсвеченных)
        '//
        '////////////////////////////////////////////////////////////////////////////////

        Try
            ShowOrderDetails(Trim(DataGridView2.SelectedRows.Item(0).Cells(0).Value.ToString()))
        Catch ex As Exception
        End Try
    End Sub

    Private Sub ShowOrderDetails(ByVal MyOrderNum As String)
        '////////////////////////////////////////////////////////////////////////////////
        '//
        '// Открытие окна с детальной информацией по выбранному заказу 
        '//
        '////////////////////////////////////////////////////////////////////////////////

        Declarations.MyOrderNum = MyOrderNum
        Dim MyOrderDetails = New OrderDetails
        MyOrderDetails.ShowDialog()
    End Sub

    Private Sub Button15_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button15.Click
        '////////////////////////////////////////////////////////////////////////////////
        '//
        '// Выдача разрешения на отгрузку выбранных заказов
        '//
        '////////////////////////////////////////////////////////////////////////////////
        Dim i As Integer

        For i = 0 To DataGridView3.SelectedRows.Count - 1
            If DataGridView3.SelectedRows.Item(i).Cells(9).Value.ToString = "" Then
                Windows.Forms.Cursor.Current = Cursors.WaitCursor
                ExecShippingAllovance(Trim(DataGridView3.SelectedRows.Item(i).Cells(0).Value.ToString))
                Application.DoEvents()
            End If
        Next
        Windows.Forms.Cursor.Current = Cursors.WaitCursor
        LoadFreeOrders()
        CheckAddButtons()
        Windows.Forms.Cursor.Current = Cursors.Default
    End Sub

    Private Sub Button3_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button3.Click
        '////////////////////////////////////////////////////////////////////////////////
        '//
        '// Выставление даты отгрузки в соответствии с подтвержденными датами из заказа на закупку
        '//
        '////////////////////////////////////////////////////////////////////////////////
        Dim i As Integer
        Dim MySQLStr As String

        For i = 0 To DataGridView3.SelectedRows.Count - 1
            If IsDBNull(DataGridView3.SelectedRows.Item(i).Cells(4).Value) = False _
                And IsDBNull(DataGridView3.SelectedRows.Item(i).Cells(5).Value) = False Then
                If DataGridView3.SelectedRows.Item(i).Cells(4).Value > DataGridView3.SelectedRows.Item(i).Cells(3).Value _
                    Or DataGridView3.SelectedRows.Item(i).Cells(5).Value > DataGridView3.SelectedRows.Item(i).Cells(3).Value Then
                    Windows.Forms.Cursor.Current = Cursors.WaitCursor
                    MySQLStr = "Exec spp_Shipments_SalesmanWP_ChangeDate N'" & Trim(DataGridView3.SelectedRows.Item(i).Cells(0).Value.ToString()) & "' "
                    InitMyConn(False)
                    Declarations.MyConn.Execute(MySQLStr)
                    Application.DoEvents()
                End If
            End If
        Next
        Windows.Forms.Cursor.Current = Cursors.WaitCursor
        LoadFreeOrders()
        CheckAddButtons()
        Windows.Forms.Cursor.Current = Cursors.Default
    End Sub

    Private Sub Button8_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button8.Click
        '////////////////////////////////////////////////////////////////////////////////
        '//
        '// Открытие окна создания отгрузки
        '//
        '////////////////////////////////////////////////////////////////////////////////
        Dim StrNum As Integer

        Declarations.MyShipmentsID = 0
        MyOperationFlag = 0
        MyShipment = New Shipment
        MyShipment.MyAction = "Create"
        MyShipment.ShowDialog()
        If MyOperationFlag <> 0 Then
            Windows.Forms.Cursor.Current = Cursors.WaitCursor
            LoadShipments()
            LoadIncludedOrders()
            LoadFreeOrders()
            CheckShipmentsButtons()
            CheckRemoveButtons()
            CheckAddButtons()
            Windows.Forms.Cursor.Current = Cursors.Default
            '---выставляем текущей созданную запись
            StrNum = DataGridView1.Rows.Count - 1
            DataGridView1.CurrentCell = DataGridView1.Item(0, StrNum)
        End If
    End Sub

    Private Sub ComboBoxAN_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles ComboBoxAN.SelectedIndexChanged
        '////////////////////////////////////////////////////////////////////////////////
        '//
        '// смена выбора - все доставки или только активные
        '//
        '////////////////////////////////////////////////////////////////////////////////

        If LoadFlag = 0 Then
            Windows.Forms.Cursor.Current = Cursors.WaitCursor
            LoadShipments()
            LoadIncludedOrders()
            LoadFreeOrders()
            CheckShipmentsButtons()
            CheckRemoveButtons()
            CheckAddButtons()
            Windows.Forms.Cursor.Current = Cursors.Default
        End If
    End Sub

    Private Sub Button6_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button6.Click
        '////////////////////////////////////////////////////////////////////////////////
        '//
        '// Удаление записи о доставке / отгрузке
        '//
        '////////////////////////////////////////////////////////////////////////////////
        Dim MySQLStr As String

        Declarations.MyShipmentsID = DataGridView1.SelectedRows.Item(0).Cells(0).Value
        '------детальные записи
        MySQLStr = "DELETE FROM tbl_Shipments_SalesmanWP_Details "
        MySQLStr = MySQLStr & "WHERE (ShipmentsID = " & Declarations.MyShipmentsID & ") "
        InitMyConn(False)
        Declarations.MyConn.Execute(MySQLStr)

        '------основная запись
        MySQLStr = "DELETE FROM tbl_Shipments_SalesmanWP_Info "
        MySQLStr = MySQLStr & "WHERE (ID = " & Declarations.MyShipmentsID & ") "
        InitMyConn(False)
        Declarations.MyConn.Execute(MySQLStr)

        Windows.Forms.Cursor.Current = Cursors.WaitCursor
        LoadShipments()
        LoadIncludedOrders()
        LoadFreeOrders()
        CheckShipmentsButtons()
        CheckRemoveButtons()
        CheckAddButtons()
        Windows.Forms.Cursor.Current = Cursors.Default
    End Sub

    Private Sub DataGridView1_ColumnHeaderMouseClick(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellMouseEventArgs) Handles DataGridView1.ColumnHeaderMouseClick
        '/////////////////////////////////////////////////////////////////////////////////////
        '//
        '// Щелчок по заголовку таблицы 
        '// 
        '/////////////////////////////////////////////////////////////////////////////////////

        FormatDataGridView1()
    End Sub

    Private Sub DataGridView1_SelectionChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles DataGridView1.SelectionChanged
        '////////////////////////////////////////////////////////////////////////////////
        '//
        '// смена выбора доставки / отгрузки
        '//
        '////////////////////////////////////////////////////////////////////////////////

        If LoadFlag = 0 Then
            Windows.Forms.Cursor.Current = Cursors.WaitCursor
            LoadIncludedOrders()
            CheckShipmentsButtons()
            CheckRemoveButtons()
            CheckAddButtons()
            Windows.Forms.Cursor.Current = Cursors.Default
        End If
    End Sub

    Private Sub Button10_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button10.Click
        '////////////////////////////////////////////////////////////////////////////////
        '//
        '// копирование выбранной отгрузки в новую
        '//
        '////////////////////////////////////////////////////////////////////////////////

        Dim StrNum As Integer

        Declarations.MyShipmentsID = DataGridView1.SelectedRows.Item(0).Cells(0).Value
        MyOperationFlag = 0
        MyShipment = New Shipment
        MyShipment.MyAction = "Copy"
        MyShipment.ShowDialog()
        If MyOperationFlag <> 0 Then
            Windows.Forms.Cursor.Current = Cursors.WaitCursor
            LoadShipments()
            LoadIncludedOrders()
            LoadFreeOrders()
            CheckShipmentsButtons()
            CheckRemoveButtons()
            CheckAddButtons()
            Windows.Forms.Cursor.Current = Cursors.Default
            '---выставляем текущей созданную запись
            StrNum = DataGridView1.Rows.Count - 1
            DataGridView1.CurrentCell = DataGridView1.Item(0, StrNum)
        End If
    End Sub

    Private Sub Button7_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button7.Click
        '////////////////////////////////////////////////////////////////////////////////
        '//
        '// Редактирование выбранной отгрузки
        '//
        '////////////////////////////////////////////////////////////////////////////////

        Declarations.MyShipmentsID = DataGridView1.SelectedRows.Item(0).Cells(0).Value
        MyOperationFlag = 0
        MyShipment = New Shipment
        MyShipment.MyAction = "Edit"
        MyShipment.ShowDialog()
        If MyOperationFlag <> 0 Then
            Windows.Forms.Cursor.Current = Cursors.WaitCursor
            LoadShipments()
            LoadIncludedOrders()
            LoadFreeOrders()
            CheckShipmentsButtons()
            CheckRemoveButtons()
            CheckAddButtons()
            Windows.Forms.Cursor.Current = Cursors.Default
            '---выставляем текущей редактируемую запись
            For i As Integer = 0 To DataGridView1.Rows.Count - 1
                If DataGridView1.Item(0, i).Value = Declarations.MyShipmentsID Then
                    DataGridView1.CurrentCell = DataGridView1.Item(0, i)
                    Exit For
                End If
            Next
        End If
    End Sub

    Private Sub Button9_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button9.Click
        '////////////////////////////////////////////////////////////////////////////////
        '//
        '// Добавление заказов в отгрузку
        '//
        '////////////////////////////////////////////////////////////////////////////////
        Dim i As Integer
        Dim MyFlag As Integer
        Dim MyMsg As String

        MyFlag = 0
        For i = 0 To DataGridView3.SelectedRows.Count - 1
            If (DataGridView1.SelectedRows.Item(0).Cells(3).Value = "Доставка" _
                And DataGridView3.SelectedRows.Item(i).Cells(9).Value.ToString = "+" _
                And DataGridView3.SelectedRows.Item(i).Cells(8).Value <> 0 _
                And DataGridView3.SelectedRows.Item(i).Cells(6).Value <> 0) _
                Or ((DataGridView1.SelectedRows.Item(0).Cells(3).Value = "Самовывоз" Or DataGridView1.SelectedRows.Item(0).Cells(3).Value = "Доставка с оплатой клиентом") _
                And DataGridView3.SelectedRows.Item(i).Cells(9).Value.ToString = "+" _
                And DataGridView3.SelectedRows.Item(i).Cells(8).Value <> 0) Then
                AddOrderToDelivery(DataGridView1.SelectedRows.Item(0).Cells(0).Value, DataGridView3.SelectedRows.Item(i).Cells(0).Value)
            Else
                MyFlag = MyFlag + 1
            End If
        Next
        Windows.Forms.Cursor.Current = Cursors.WaitCursor
        LoadIncludedOrders()
        LoadFreeOrders()
        ChangeSummsInDelivery()
        CheckShipmentsButtons()
        CheckRemoveButtons()
        CheckAddButtons()
        Windows.Forms.Cursor.Current = Cursors.Default
        If MyFlag > 0 Then
            MyMsg = "Один или более заказов не был добавлен в отгрузку. Возможные причины:" & Chr(13)
            MyMsg = MyMsg & "- Нет разрешения на отгрузку" & Chr(13)
            MyMsg = MyMsg & "- В заказе нет распределенных продуктов" & Chr(13)
            MyMsg = MyMsg & "- Стоимость доставки заказа 0 (самовывоз), а вы пытаетесь добавить счет в доставку, а не самовывоз" & Chr(13)
            MsgBox(MyMsg, MsgBoxStyle.Critical, "Внимание!")
        End If
    End Sub

    Private Sub AddOrderToDelivery(ByVal ShipmentsID As Integer, ByVal OrderNum As String)
        '////////////////////////////////////////////////////////////////////////////////
        '//
        '// Добавление одного заказа в отгрузку
        '//
        '////////////////////////////////////////////////////////////////////////////////
        Dim MySQLStr As String

        MySQLStr = "INSERT INTO tbl_Shipments_SalesmanWP_Details "
        MySQLStr = MySQLStr & "(ShipmentsID, OrderNum, InvoiceNum, IsClosed) "
        MySQLStr = MySQLStr & "VALUES (" & ShipmentsID.ToString & ", "
        MySQLStr = MySQLStr & "N'" & OrderNum & "', "
        MySQLStr = MySQLStr & "'', 0)"
        InitMyConn(False)
        Declarations.MyConn.Execute(MySQLStr)
    End Sub

    Private Sub ChangeSummsInDelivery()
        '////////////////////////////////////////////////////////////////////////////////
        '//
        '// изменение итоговых сумм в отгрузке
        '//
        '////////////////////////////////////////////////////////////////////////////////
        Dim MySQLStr As String

        MySQLStr = "UPDATE tbl_Shipments_SalesmanWP_Info "
        MySQLStr = MySQLStr & "SET DeliverySumm = " & Trim(Replace(TextBox2.Text, ",", ".")) & ", "
        MySQLStr = MySQLStr & "DeliveredSumm = " & Trim(Replace(TextBox1.Text, ",", ".")) & " "
        MySQLStr = MySQLStr & "WHERE (ID = " & DataGridView1.SelectedRows.Item(0).Cells(0).Value.ToString & ")"
        InitMyConn(False)
        Declarations.MyConn.Execute(MySQLStr)

        DataGridView1.SelectedRows.Item(0).Cells(4).Value = TextBox2.Text
        DataGridView1.SelectedRows.Item(0).Cells(5).Value = TextBox1.Text
    End Sub

    Private Sub Button11_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button11.Click
        '////////////////////////////////////////////////////////////////////////////////
        '//
        '// Удаление заказов из отгрузки
        '//
        '////////////////////////////////////////////////////////////////////////////////
        Dim i As Integer

        For i = 0 To DataGridView2.SelectedRows.Count - 1
            RemoveOrderFromDelivery(DataGridView1.SelectedRows.Item(0).Cells(0).Value, DataGridView2.SelectedRows.Item(i).Cells(0).Value)
        Next
        Windows.Forms.Cursor.Current = Cursors.WaitCursor
        LoadIncludedOrders()
        LoadFreeOrders()
        ChangeSummsInDelivery()
        CheckShipmentsButtons()
        CheckRemoveButtons()
        CheckAddButtons()
        Windows.Forms.Cursor.Current = Cursors.Default
    End Sub

    Private Sub RemoveOrderFromDelivery(ByVal ShipmentsID As Integer, ByVal OrderNum As String)
        '////////////////////////////////////////////////////////////////////////////////
        '//
        '// Удаление одного заказа из отгрузки
        '//
        '////////////////////////////////////////////////////////////////////////////////
        Dim MySQLStr As String

        MySQLStr = "DELETE FROM tbl_Shipments_SalesmanWP_Details "
        MySQLStr = MySQLStr & "WHERE (ShipmentsID = " & ShipmentsID.ToString & ") "
        MySQLStr = MySQLStr & "AND (OrderNum = N'" & OrderNum & "')"

        InitMyConn(False)
        Declarations.MyConn.Execute(MySQLStr)
    End Sub

    Private Sub Button4_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button4.Click
        '////////////////////////////////////////////////////////////////////////////////
        '//
        '// Закрытие отгрузки, если что то не будет отгружаться
        '//
        '////////////////////////////////////////////////////////////////////////////////
        Dim myrez As VariantType
        Dim MySQLStr As String

        myrez = MsgBox("Вы уверены, что хотите закрыть отгрузку?" + Chr(13) + Chr(10) + "Отгрузку надо закрывать только в том случае, если склад не может физически произвести требуемую отгрузку / сборку.", MsgBoxStyle.YesNo, "Внимание!")
        If myrez = vbYes Then
            MySQLStr = "UPDATE tbl_Shipments_SalesmanWP_Details "
            MySQLStr = MySQLStr & "SET IsClosed = 1 "
            MySQLStr = MySQLStr & "WHERE (ShipmentsID = " & DataGridView1.SelectedRows.Item(0).Cells(0).Value.ToString & ")"
            InitMyConn(False)
            Declarations.MyConn.Execute(MySQLStr)

            Windows.Forms.Cursor.Current = Cursors.WaitCursor
            LoadShipments()
            LoadIncludedOrders()
            LoadFreeOrders()
            CheckShipmentsButtons()
            CheckRemoveButtons()
            CheckAddButtons()
            Windows.Forms.Cursor.Current = Cursors.Default
        End If
    End Sub

    Private Sub Button14_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button14.Click
        '////////////////////////////////////////////////////////////////////////////////
        '//
        '// открытие окна отправки уведомления клиентам об отправке / готовности к самовывозу
        '//
        '////////////////////////////////////////////////////////////////////////////////

        MySendInfo = New SendInfo
        MySendInfo.ShowDialog()
        CheckShipmentsButtons()
        CheckRemoveButtons()
        CheckAddButtons()

    End Sub

    Private Sub Button12_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button12.Click
        '////////////////////////////////////////////////////////////////////////////////
        '//
        '// Вызов процедуры создания заявки на портале на отгрузку / самовывоз
        '//
        '////////////////////////////////////////////////////////////////////////////////
        Dim MySQLStr As String
        Dim OrderList As String
        Dim MyINN As String
        Dim MyLegalAddress As String
        Dim MyWH As String
        Dim MyDeliveryOrNot As String
        Dim PrintBillFlag As Boolean
        Dim PrintBillFlag1 As Boolean
        Dim PrintFullBillFlag As Boolean
        Dim MyRez As VariantType
        Dim MyFile As String
        Dim MyComment As String

        MyRez = Microsoft.VisualBasic.vbYes
        MyFile = DataGridView1.SelectedRows.Item(0).Cells(17).Value
        If CheckFilePresent() = False Then
            MyFile = ""
            MyRez = MsgBox("Файл, который вы хотите присоединить к заявке, отсутствует в указанном месте: " & DataGridView1.SelectedRows.Item(0).Cells(17).Value.ToString & " . Будете размещать заявку без файла?", MsgBoxStyle.YesNo, "Внимание!")
        End If

        If MyRez = Microsoft.VisualBasic.vbYes Then
            MyINN = ""
            MyLegalAddress = ""
            MyWH = ""
            MyDeliveryOrNot = ""
            PrintBillFlag = False

            '-----список заказов
            OrderList = ""
            MySQLStr = "SELECT Replace((SELECT OrderNum as 'data()' "
            MySQLStr = MySQLStr & "FROM tbl_Shipments_SalesmanWP_Details "
            MySQLStr = MySQLStr & "WHERE(ShipmentsID = " & DataGridView1.SelectedRows.Item(0).Cells(0).Value.ToString & ") "
            MySQLStr = MySQLStr & "FOR XML path('')), ' ', char(13) + char(10)) AS OrderList "
            InitMyConn(False)
            InitMyRec(False, MySQLStr)
            If Declarations.MyRec.EOF = True And Declarations.MyRec.BOF = True Then
            Else
                OrderList = Declarations.MyRec.Fields("OrderList").Value
            End If
            trycloseMyRec()

            '------склад
            MyWH = ComboBox1.Text

            '------ИНН и юр. адрес
            MySQLStr = "SELECT CustomerINN, CustomerLegalAddress "
            MySQLStr = MySQLStr & "FROM tbl_Shipments_SalesmanWP_Info "
            MySQLStr = MySQLStr & "WHERE (ID = " & DataGridView1.SelectedRows.Item(0).Cells(0).Value.ToString & ")"
            InitMyConn(False)
            InitMyRec(False, MySQLStr)
            If Declarations.MyRec.EOF = True And Declarations.MyRec.BOF = True Then
            Else
                MyINN = Declarations.MyRec.Fields("CustomerINN").Value
                MyLegalAddress = Declarations.MyRec.Fields("CustomerLegalAddress").Value
            End If
            trycloseMyRec()

            '------доставка
            If DataGridView1.SelectedRows.Item(0).Cells(3).Value = "Доставка" Then
                MyDeliveryOrNot = "Доставка"
            ElseIf DataGridView1.SelectedRows.Item(0).Cells(3).Value = "Доставка с оплатой клиентом" Then
                MyDeliveryOrNot = "Доставка с оплатой клиентом"
            Else
                MyDeliveryOrNot = "Самовывоз"
            End If

            '------печать счета
            If DataGridView1.SelectedRows.Item(0).Cells(9).Value = "+" Then
                PrintBillFlag = True
            Else
                PrintBillFlag = False
            End If

            '------печать справки - счета
            If DataGridView1.SelectedRows.Item(0).Cells(10).Value = "+" Then
                PrintBillFlag1 = True
            Else
                PrintBillFlag1 = False
            End If

            '------печать полного счета (восстановленного)
            If DataGridView1.SelectedRows.Item(0).Cells(11).Value = "+" Then
                PrintFullBillFlag = True
            Else
                PrintFullBillFlag = False
            End If

            '------Формирование комментария
            MyComment = ""
            If Not Trim(DataGridView1.SelectedRows.Item(0).Cells(8).Value).Equals("") Then
                MyComment = MyComment & "Складу:  " & Trim(DataGridView1.SelectedRows.Item(0).Cells(8).Value)
            End If
            If Not Trim(DataGridView1.SelectedRows.Item(0).Cells(18).Value).Equals("") Then
                If Not MyComment.Equals("") Then
                    MyComment = MyComment & Chr(13) & Chr(10)
                End If
                MyComment = MyComment & "По транспорту:  " & Trim(DataGridView1.SelectedRows.Item(0).Cells(18).Value)
            End If
            If Not Trim(DataGridView1.SelectedRows.Item(0).Cells(19).Value).Equals("") Then
                If Not MyComment.Equals("") Then
                    MyComment = MyComment & Chr(13) & Chr(10)
                End If
                MyComment = MyComment & "По документам:  " & Trim(DataGridView1.SelectedRows.Item(0).Cells(19).Value)
            End If


            If OrderList <> "" Then
                CreateRequest(DataGridView1.SelectedRows.Item(0).Cells(1).Value.ToString, _
                    DataGridView1.SelectedRows.Item(0).Cells(2).Value.ToString, _
                    MyINN, _
                    MyLegalAddress, _
                    "ESKRU\" + Declarations.UserName, _
                    MyWH, _
                    MyDeliveryOrNot, _
                    DataGridView1.SelectedRows.Item(0).Cells(4).Value, _
                    DataGridView1.SelectedRows.Item(0).Cells(6).Value, _
                    DataGridView1.SelectedRows.Item(0).Cells(7).Value, _
                    MyComment, _
                    PrintBillFlag, _
                    PrintBillFlag1, _
                    PrintFullBillFlag, _
                    DataGridView1.SelectedRows.Item(0).Cells(12).Value, _
                    OrderList, _
                    MyFile)

                '-------отрисовка + проверка
                MySQLStr = "UPDATE tbl_Shipments_SalesmanWP_Info "
                MySQLStr = MySQLStr & "SET IsRequestSend = 1 "
                MySQLStr = MySQLStr & "WHERE (ID = " & MyShipmentsList.DataGridView1.SelectedRows.Item(0).Cells(0).Value.ToString & ")"
                InitMyConn(False)
                Declarations.MyConn.Execute(MySQLStr)
                MyShipmentsList.DataGridView1.SelectedRows.Item(0).Cells(13).Value = 1
                CheckShipmentsButtons()
                CheckRemoveButtons()
                CheckAddButtons()
                MsgBox("Заявка на портале создана.", MsgBoxStyle.Information, "Внимание!")
            End If
        End If
    End Sub

    Private Function CheckFilePresent() As Boolean
        '////////////////////////////////////////////////////////////////////////////////
        '//
        '// Проверка наличия файла, который надо присоединить к заявке
        '//
        '////////////////////////////////////////////////////////////////////////////////

        If Trim(DataGridView1.SelectedRows.Item(0).Cells(17).Value.ToString) = "" Then
            CheckFilePresent = True
        Else
            If System.IO.File.Exists(Trim(DataGridView1.SelectedRows.Item(0).Cells(17).Value.ToString)) Then
                CheckFilePresent = True
            Else
                CheckFilePresent = False
            End If
        End If
    End Function

    Private Sub CreateRequest(ByVal MyCustomerCode As String, ByVal MyCustomerName As String, ByVal MyINN As String, ByVal MyLegalAddress As String, _
        ByVal MySalesman As String, ByVal MyWH As String, ByVal DeliveryOrNot As String, ByVal DeliverySumm As Double, ByVal MyContactInfo As String, _
        ByVal MyDeliveryAddress As String, ByVal MyComment As String, ByVal PrintBillOrNot As Boolean, ByVal PrintBillOrNot1 As Boolean, _
        ByVal PrintFullBillOrNot As Boolean, ByVal MyRequestedDate As DateTime, ByVal OrderList As String, ByVal MyFile As String)
        '////////////////////////////////////////////////////////////////////////////////
        '//
        '// Создание заявки на портале на отгрузку / самовывоз
        '//
        '////////////////////////////////////////////////////////////////////////////////
        Dim listWebService As spbprd4.Lists = New spbprd4.Lists()
        'Dim listWebService As spbprd41.Lists = New spbprd41.Lists()
        listWebService.Credentials = New System.Net.NetworkCredential("developer", "!Devpass", "ESKRU")
        'listWebService.Url = "http://spbprd4/logistics/_vti_bin/Lists.asmx"
        Dim listName = "{3e66b7ae-a55c-4e6f-92be-9f602e7d0417}"
        Dim listView = ""
        Dim listItemId As String = ""
        Dim FileName As String = ""
        Dim MyAttachment As Byte()
        Dim OrderListWithHref As String

        Dim strBatch As String = "<Method ID='1' Cmd='New'>"
        strBatch = strBatch + "<Field Name='ID'>New</Field>"
        strBatch = strBatch + "<Field Name='_x041a__x043e__x0434__x0020__x04'>" & MyCustomerCode & "</Field>"           '---Код клиента
        strBatch = strBatch + "<Field Name='Title'>" & MyCustomerName & "</Field>"                                      '---название клиента
        strBatch = strBatch + "<Field Name='_x0418__x041d__x041d__x0020__x04'>" & MyINN & "</Field>"                    '---ИНН клиента
        strBatch = strBatch + "<Field Name='_x042e__x0440__x0020__x0430__x04'>" & MyLegalAddress & "</Field>"           '---Юридический адрес клиента
        OrderListWithHref = ""
        Dim Orders As String() = OrderList.Split(New String() {Environment.NewLine}, StringSplitOptions.None)
        For Each Order As String In Orders
            OrderListWithHref = OrderListWithHref + "&lt;p style=""margin-bottom:0;margin-top:0""&gt;&lt;a href=""http://spbprd5/ReportServer/Pages/ReportViewer.aspx?/%d0%a1%d0%ba%d0%bb%d0%b0%d0%b4/%d0%a1%d0%bf%d0%b8%d1%81%d0%be%d0%ba+%d0%bf%d0%be%d0%b4%d0%b1%d0%be%d1%80%d0%ba%d0%b8/%d0%a1%d0%bf%d0%b8%d1%81%d0%be%d0%ba+%d0%bf%d0%be%d0%b4%d0%b1%d0%be%d1%80%d0%ba%d0%b8&amp;OrderNumber="
            OrderListWithHref = OrderListWithHref + Order
            OrderListWithHref = OrderListWithHref + """&gt;" + Order + "&lt;/a&gt;&lt;/p&gt;"
        Next
        'strBatch = strBatch + "<Field Name='_x041d__x043e__x043c__x0435__x04'>" & OrderList & "</Field>"                '---список заказов
        strBatch = strBatch + "<Field Name='_x041d__x043e__x043c__x0435__x04'>" & OrderListWithHref & "</Field>"        '---список заказов с гиперссылками
        strBatch = strBatch + "<Field Name='_x041f__x0440__x043e__x0434__x04'>" & MySalesman & "</Field>"               '---продавец
        strBatch = strBatch + "<Field Name='_x0421__x043a__x043b__x0430__x04'>" & MyWH & "</Field>"                     '---склад
        strBatch = strBatch + "<Field Name='_x0414__x043e__x0441__x0442__x040'>" & DeliveryOrNot & "</Field>"           '---доставка или самовывоз
        strBatch = strBatch + "<Field Name='_x0421__x0443__x043c__x043c__x04'>" & DeliverySumm.ToString & "</Field>"    '---сумма на доставку
        strBatch = strBatch + "<Field Name='_x041a__x043e__x043d__x0442__x04'>" & MyContactInfo & "</Field>"            '---контактная информация
        If DeliveryOrNot = "Самовывоз" Then
            strBatch = strBatch + "<Field Name='_x0410__x0434__x0440__x0435__x04'></Field>"                             '---адрес доставки
        Else
            strBatch = strBatch + "<Field Name='_x0410__x0434__x0440__x0435__x04'>" & MyDeliveryAddress & "</Field>"    '---адрес доставки
        End If
        strBatch = strBatch + "<Field Name='_x041a__x043e__x043c__x043c__x04'>" & MyComment & "</Field>"                '---комментарий
        strBatch = strBatch + "<Field Name='_x041f__x0435__x0447__x0430__x04'>" & PrintBillOrNot & "</Field>"           '---печатать счет или нет
        strBatch = strBatch + "<Field Name='_x041f__x0435__x0447__x0430__x040'>" & PrintBillOrNot1 & "</Field>"         '---печатать справку - счет или нет
        strBatch = strBatch + "<Field Name='_x041f__x0435__x0447__x0430__x041'>" & PrintFullBillOrNot & "</Field>"      '---печатать полный счет (восстановленный) или нет
        strBatch = strBatch + "<Field Name='_x0416__x0435__x043b__x0430__x04'>" & Format(MyRequestedDate, "yyyy-MM-dd HH:mm:ss") & "</Field>"    '---запрошенная дата отгрузки
        strBatch = strBatch + "</Method>"

        Dim xmlDoc As XmlDocument = New System.Xml.XmlDocument()
        Dim elBatch As System.Xml.XmlElement = xmlDoc.CreateElement("Batch")
        elBatch.SetAttribute("OnError", "Continue")
        elBatch.SetAttribute("ListVersion", "1")
        elBatch.SetAttribute("ViewName", listView)
        elBatch.InnerXml = strBatch

        Try
            Dim ndReturn As XmlNode = listWebService.UpdateListItems(listName, elBatch)

            '---Аттачмент
            If Trim(MyFile) <> "" Then
                Dim NewDoc As XmlDocument = New XmlDocument
                NewDoc.LoadXml(ndReturn.OuterXml)
                Dim NewNdList As XmlNodeList = NewDoc.GetElementsByTagName("z:row")
                listItemId = NewNdList(0).Attributes("ows_ID").Value.ToString
                '---имя файла
                FileName = System.IO.Path.GetFileName(MyFile)
                '---аттачмент
                MyAttachment = System.IO.File.ReadAllBytes(MyFile)
                listWebService.AddAttachment(listName, listItemId, FileName, MyAttachment)
            End If
        Catch ex As Exception
            MsgBox("Ошибка создания заявки на портале " + ex.Message, MsgBoxStyle.Critical, "Внимание!")
        End Try
    End Sub

    Private Sub DataGridView3_ColumnHeaderMouseClick(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellMouseEventArgs) Handles DataGridView3.ColumnHeaderMouseClick
        '/////////////////////////////////////////////////////////////////////////////////////
        '//
        '// Щелчок по заголовку таблицы 
        '// 
        '/////////////////////////////////////////////////////////////////////////////////////

        FormatDataGridView3()
    End Sub
End Class