Public Class MainForm
    Public LoadFlag As Integer

    Private Sub Button13_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button13.Click
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

        ComboBoxAN.SelectedIndex = 0

        LoadFlag = 0
        Application.DoEvents()
        Windows.Forms.Cursor.Current = Cursors.WaitCursor
        '---Загрузка данных
        LoadShipments()
        LoadFreeOrders()
        '---Проверка состояния кнопок
        CheckSHButtonsState()
        CheckOrderButtonsState()
        Windows.Forms.Cursor.Current = Cursors.Default
    End Sub

    Public Sub LoadShipments()
        '/////////////////////////////////////////////////////////////////////////////////////
        '//
        '// Загрузка списка Доставок / отгрузок
        '// 
        '/////////////////////////////////////////////////////////////////////////////////////
        Dim MySQLStr As String
        Dim MyAdapter As SqlClient.SqlDataAdapter     'для списка поставщиков
        Dim MyDs As New DataSet                       '

        If LoadFlag = 0 Then
            If ComboBoxAN.Text = "Все доставки / самовывозы" Then
                MySQLStr = "EXEC spp_WEBShipments_ShipmentInfo N'" & Trim(ComboBox1.SelectedValue) & "', 0"
            Else
                MySQLStr = "EXEC spp_WEBShipments_ShipmentInfo N'" & Trim(ComboBox1.SelectedValue) & "', 1"
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
            DataGridView1.Columns(0).HeaderText = "N отгр узки"
            DataGridView1.Columns(0).Width = 50
            DataGridView1.Columns(1).HeaderText = "Продавец по отгрузке"
            DataGridView1.Columns(1).Width = 80
            DataGridView1.Columns(2).HeaderText = "N заказа"
            DataGridView1.Columns(2).Width = 70
            DataGridView1.Columns(3).HeaderText = "N заказа с WEB"
            DataGridView1.Columns(3).Width = 70
            DataGridView1.Columns(4).HeaderText = "Код поку пателя"
            DataGridView1.Columns(4).Width = 70
            DataGridView1.Columns(5).HeaderText = "Покупатель"
            DataGridView1.Columns(5).Width = 140
            DataGridView1.Columns(6).HeaderText = "Доставка"
            DataGridView1.Columns(6).Width = 100
            DataGridView1.Columns(7).HeaderText = "Сумма на доставку"
            DataGridView1.Columns(7).Width = 80
            DataGridView1.Columns(7).DefaultCellStyle.Format = "n2"
            DataGridView1.Columns(8).HeaderText = "Сумма отгружаемого"
            DataGridView1.Columns(8).Width = 100
            DataGridView1.Columns(8).DefaultCellStyle.Format = "n2"
            DataGridView1.Columns(9).HeaderText = "Контактная информация"
            DataGridView1.Columns(9).Width = 200
            DataGridView1.Columns(10).HeaderText = "Адрес доставки"
            DataGridView1.Columns(10).Width = 235
            DataGridView1.Columns(11).HeaderText = "Комментарий"
            DataGridView1.Columns(11).Width = 155
            DataGridView1.Columns(12).HeaderText = "Печать счета"
            DataGridView1.Columns(12).Width = 60
            DataGridView1.Columns(13).HeaderText = "Печать справки - счета"
            DataGridView1.Columns(13).Width = 60
            DataGridView1.Columns(14).HeaderText = "Печать полного счета (восст.)"
            DataGridView1.Columns(14).Width = 60
            DataGridView1.Columns(15).HeaderText = "запро шенная дата поставки"
            DataGridView1.Columns(15).Width = 70
            DataGridView1.Columns(16).HeaderText = "Запрос на портал"
            DataGridView1.Columns(16).Width = 60
            DataGridView1.Columns(17).HeaderText = "Уведом ление клиенту"
            DataGridView1.Columns(17).Width = 60
            DataGridView1.Columns(18).HeaderText = "отгрузка произ ведена"
            DataGridView1.Columns(18).Width = 60
            DataGridView1.Columns(19).HeaderText = "Файл"
            DataGridView1.Columns(19).Width = 60
            DataGridView1.Columns(20).HeaderText = "Путь к файлу"
            DataGridView1.Columns(20).Visible = False

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
            If DataGridView1.Rows(i).Cells(18).Value = 0 Then
                DataGridView1.Rows(i).DefaultCellStyle.BackColor = Color.White
                If DataGridView1.Rows(i).Cells(6).Value = "Доставка" Then
                    If DateDiff(DateInterval.Day, DataGridView1.Rows(i).Cells(15).Value, Now()) > 2 Then
                        DataGridView1.Rows(i).Cells(18).Style.BackColor = Color.LightPink
                    Else
                        DataGridView1.Rows(i).Cells(18).Style.BackColor = Color.White
                    End If

                Else            '---самовывоз или доставка с оплатой клиентом
                    If DateDiff(DateInterval.Day, DataGridView1.Rows(i).Cells(15).Value, Now()) > 7 Then
                        DataGridView1.Rows(i).Cells(18).Style.BackColor = Color.LightPink
                    Else
                        DataGridView1.Rows(i).Cells(18).Style.BackColor = Color.White
                    End If
                End If
            Else
                DataGridView1.Rows(i).DefaultCellStyle.BackColor = Color.LightGray
            End If
        Next
    End Sub

    Public Sub LoadFreeOrders()
        '/////////////////////////////////////////////////////////////////////////////////////
        '//
        '// Загрузка списка заказов на отгрузку / доставку
        '// 
        '/////////////////////////////////////////////////////////////////////////////////////
        Dim MySQLStr As String
        Dim MyAdapter As SqlClient.SqlDataAdapter     'для списка поставщиков
        Dim MyDs As New DataSet                       '

        If LoadFlag = 0 Then
            MySQLStr = "spp_WEBShipments_AvlOrders N'" & Trim(ComboBox1.SelectedValue) & "'"

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
            DataGridView2.Columns(0).Width = 90
            DataGridView2.Columns(1).HeaderText = "Тип з-за"
            DataGridView2.Columns(1).Width = 60
            DataGridView2.Columns(2).HeaderText = "N з-за WEB"
            DataGridView2.Columns(2).Width = 90
            DataGridView2.Columns(3).HeaderText = "Клиент"
            DataGridView2.Columns(3).Width = 350
            DataGridView2.Columns(4).HeaderText = "Продавец"
            DataGridView2.Columns(4).Width = 200
            DataGridView2.Columns(5).HeaderText = "Дата отгрузки"
            DataGridView2.Columns(5).DefaultCellStyle.Format = "dd/MM/yyyy"
            DataGridView2.Columns(5).Width = 100
            DataGridView2.Columns(6).HeaderText = "Макс дата прихода товара на склад"
            DataGridView2.Columns(6).Width = 100
            DataGridView2.Columns(6).DefaultCellStyle.Format = "dd/MM/yyyy"
            DataGridView2.Columns(7).HeaderText = "Макс задолж. дата прихода товара на склад"
            DataGridView2.Columns(7).Width = 100
            DataGridView2.Columns(7).DefaultCellStyle.Format = "dd/MM/yyyy"
            DataGridView2.Columns(8).HeaderText = "Сумма доставки (остаток)"
            DataGridView2.Columns(8).Width = 120
            DataGridView2.Columns(8).DefaultCellStyle.Format = "n2"
            DataGridView2.Columns(9).HeaderText = "Сумма заказа"
            DataGridView2.Columns(9).Width = 120
            DataGridView2.Columns(9).DefaultCellStyle.Format = "n2"
            DataGridView2.Columns(10).HeaderText = "Сумма распреде ленного"
            DataGridView2.Columns(10).Width = 120
            DataGridView2.Columns(10).DefaultCellStyle.Format = "n2"
            DataGridView2.Columns(11).HeaderText = "Разрешение на отгрузку"
            DataGridView2.Columns(11).Width = 90

            FormatDataGridView2()
        End If
    End Sub

    Private Sub FormatDataGridView2()
        '////////////////////////////////////////////////////////////////////////////////
        '//
        '// подсветка информации по доступным заказам
        '//
        '////////////////////////////////////////////////////////////////////////////////

        For Each row As DataGridViewRow In DataGridView2.Rows
            If Trim(row.Cells(11).Value.ToString) <> "" Then
                row.Cells(11).Style.BackColor = Color.LightGreen
            Else
                row.Cells(11).Style.BackColor = Color.LightPink
            End If
            If row.Cells(5).Value < Now Then
                row.Cells(5).Style.BackColor = Color.LightGreen
            Else
                row.Cells(5).Style.BackColor = Color.Empty
            End If
            If IsDBNull(row.Cells(6).Value) = False Then
                If row.Cells(5).Value < row.Cells(6).Value Then
                    row.Cells(6).Style.BackColor = Color.LightYellow
                Else
                    row.Cells(6).Style.BackColor = Color.Empty
                End If
            Else
                row.Cells(6).Style.BackColor = Color.Empty
            End If
            If IsDBNull(row.Cells(7).Value) = False Then
                If row.Cells(5).Value < row.Cells(7).Value Then
                    row.Cells(7).Style.BackColor = Color.LightYellow
                Else
                    row.Cells(7).Style.BackColor = Color.Empty
                End If
            Else
                row.Cells(7).Style.BackColor = Color.Empty
            End If
            If row.Cells(9).Value = 0 Then
                row.Cells(9).Style.BackColor = Color.LightPink
            Else
                row.Cells(9).Style.BackColor = Color.Empty
            End If
            If row.Cells(10).Value = 0 Then
                row.Cells(10).Style.BackColor = Color.LightPink
            Else
                If row.Cells(10).Value < row.Cells(9).Value Then
                    row.Cells(10).Style.BackColor = Color.LightYellow
                Else
                    row.Cells(10).Style.BackColor = Color.Empty
                End If
            End If
        Next
    End Sub

    Public Sub CheckSHButtonsState()
        '////////////////////////////////////////////////////////////////////////////////
        '//
        '// Проверка и выставление состояния кнопок по отгрузкам
        '//
        '////////////////////////////////////////////////////////////////////////////////

        If DataGridView1.SelectedRows.Count = 0 Then
            Button5.Enabled = False
            Button6.Enabled = False
        Else
            '--------------уведомление клиенту
            If DataGridView1.SelectedRows.Item(0).Cells(17).Value = 0 Then  '---уведомление клиенту не отправлено
                If DataGridView1.SelectedRows.Item(0).Cells(18).Value = 0 Then '---отгрузка не закрыта (не произведена)
                    Button5.Enabled = True
                Else
                    Button5.Enabled = False
                End If
            Else
                Button5.Enabled = False
            End If
            '--------------принудительное закрытие отгрузки
            If DataGridView1.SelectedRows.Item(0).Cells(18).Value = 0 Then  '---отгрузка не произведена
                Button6.Enabled = True
            Else
                Button6.Enabled = False
            End If
        End If
    End Sub

    Public Sub CheckOrderButtonsState()
        '////////////////////////////////////////////////////////////////////////////////
        '//
        '// Проверка и выставление состояния кнопок по заказам
        '//
        '////////////////////////////////////////////////////////////////////////////////

        If DataGridView2.SelectedRows.Count = 0 Then
            Button2.Enabled = False
            Button3.Enabled = False
            Button4.Enabled = False
        Else
            Button2.Enabled = True
            If DataGridView2.SelectedRows.Item(0).Cells(11).Value.ToString = "+" Then
                Button3.Enabled = False
            Else
                Button3.Enabled = True
            End If
            If (DataGridView2.SelectedRows.Item(0).Cells(11).Value.ToString = "+" _
                And DataGridView2.SelectedRows.Item(0).Cells(10).Value <> 0) Then
                Button4.Enabled = True
            Else
                Button4.Enabled = False
            End If
        End If

    End Sub

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        '////////////////////////////////////////////////////////////////////////////////
        '//
        '// Обновление данных
        '//
        '////////////////////////////////////////////////////////////////////////////////

        Windows.Forms.Cursor.Current = Cursors.WaitCursor
        '---Загрузка данных
        LoadShipments()
        LoadFreeOrders()
        '---Проверка состояния кнопок
        CheckSHButtonsState()
        CheckOrderButtonsState()
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
            LoadShipments()
            LoadFreeOrders()
            '---Проверка состояния кнопок
            CheckSHButtonsState()
            CheckOrderButtonsState()
            Windows.Forms.Cursor.Current = Cursors.Default
        End If
    End Sub

    Private Sub ComboBoxAN_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles ComboBoxAN.SelectedIndexChanged
        '////////////////////////////////////////////////////////////////////////////////
        '//
        '// Загрузка данных в соответствии с выбранным параметром
        '//
        '////////////////////////////////////////////////////////////////////////////////

        If LoadFlag = 0 Then
            Windows.Forms.Cursor.Current = Cursors.WaitCursor
            LoadShipments()
            LoadFreeOrders()
            '---Проверка состояния кнопок
            CheckSHButtonsState()
            CheckOrderButtonsState()
            Windows.Forms.Cursor.Current = Cursors.Default
        End If
    End Sub

    Private Sub Button2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button2.Click
        '////////////////////////////////////////////////////////////////////////////////
        '//
        '// Открытие окна с детальной информацией по выбранному заказу 
        '//
        '////////////////////////////////////////////////////////////////////////////////

        Declarations.MyOrderNum = Trim(DataGridView2.SelectedRows.Item(0).Cells(0).Value.ToString())
        Dim MyOrderDetails = New OrderDetails
        MyOrderDetails.ShowDialog()
    End Sub

    Private Sub Button3_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button3.Click
        '////////////////////////////////////////////////////////////////////////////////
        '//
        '// Выдача разрешения на отгрузку выбранных заказов
        '//
        '////////////////////////////////////////////////////////////////////////////////

        If DataGridView2.SelectedRows.Item(0).Cells(11).Value.ToString = "" Then
            Windows.Forms.Cursor.Current = Cursors.WaitCursor
            ExecShippingAllovance(Trim(DataGridView2.SelectedRows.Item(0).Cells(0).Value.ToString))
            LoadFreeOrders()
            CheckOrderButtonsState()
            Windows.Forms.Cursor.Current = Cursors.Default
        End If

    End Sub

    Private Sub Button6_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button6.Click
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
            LoadFreeOrders()
            '---Проверка состояния кнопок
            CheckSHButtonsState()
            CheckOrderButtonsState()
            Windows.Forms.Cursor.Current = Cursors.Default
        End If
    End Sub

    Private Sub Button5_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button5.Click
        '////////////////////////////////////////////////////////////////////////////////
        '//
        '// открытие окна отправки уведомления клиентам об отправке / готовности к самовывозу
        '//
        '////////////////////////////////////////////////////////////////////////////////

        MySendInfo = New SendInfo
        MySendInfo.ShowDialog()
        CheckSHButtonsState()
        CheckOrderButtonsState()
    End Sub

    Private Sub Button4_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button4.Click
        '////////////////////////////////////////////////////////////////////////////////
        '//
        '// Открытие окна создания отгрузки
        '//
        '////////////////////////////////////////////////////////////////////////////////
        Dim MyMsg As String

        If (DataGridView2.SelectedRows.Item(0).Cells(11).Value.ToString = "+" _
                And DataGridView2.SelectedRows.Item(0).Cells(10).Value <> 0) Then
            Declarations.MyCustomerCode = GetFirstPartFromStr(DataGridView2.SelectedRows.Item(0).Cells(3).Value)
            Declarations.MyWH = Trim(ComboBox1.SelectedValue)
            Declarations.MyOrderNum = DataGridView2.SelectedRows.Item(0).Cells(0).Value
            Declarations.MyShipmentsID = 0
            MyOperationFlag = 0
            MyShipment = New Shipment
            MyShipment.ShowDialog()
            If MyOperationFlag <> 0 Then
                Application.DoEvents()
                Windows.Forms.Cursor.Current = Cursors.WaitCursor
                LoadShipments()
                LoadFreeOrders()
                CheckSHButtonsState()
                CheckOrderButtonsState()
                Windows.Forms.Cursor.Current = Cursors.Default
                '---выставляем текущей редактируемую запись
                For i As Integer = 0 To DataGridView1.Rows.Count - 1
                    If DataGridView1.Item(0, i).Value = Declarations.MyShipmentsID Then
                        DataGridView1.CurrentCell = DataGridView1.Item(0, i)
                        Exit For
                    End If
                Next
            End If
        Else
            If DataGridView2.SelectedRows.Item(0).Cells(11).Value.ToString <> "+" Then
                MyMsg = "Заказ не был добавлен в отгрузку. Причина:" & Chr(13)
                MyMsg = MyMsg & "- Нет разрешения на отгрузку" & Chr(13)
                MsgBox(MyMsg, MsgBoxStyle.Critical, "Внимание!")
            Else
                If DataGridView2.SelectedRows.Item(0).Cells(10).Value = 0 Then
                    MyMsg = "Заказ не был добавлен в отгрузку. Причина:" & Chr(13)
                    MyMsg = MyMsg & "- В заказе нет распределенных продуктов" & Chr(13)
                    MsgBox(MyMsg, MsgBoxStyle.Critical, "Внимание!")
                End If
            End If
            
            End If
    End Sub

    Private Sub DataGridView2_ColumnHeaderMouseClick(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellMouseEventArgs) Handles DataGridView2.ColumnHeaderMouseClick
        '/////////////////////////////////////////////////////////////////////////////////////
        '//
        '// Щелчок по заголовку таблицы 
        '// 
        '/////////////////////////////////////////////////////////////////////////////////////

        FormatDataGridView2()
    End Sub

    Private Sub DataGridView2_SelectionChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles DataGridView2.SelectionChanged
        '////////////////////////////////////////////////////////////////////////////////
        '//
        '// Смена выбранной строки в доступных заказах
        '//
        '////////////////////////////////////////////////////////////////////////////////

        CheckOrderButtonsState()
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
        '// Смена выбранной строки в отгрузках
        '//
        '////////////////////////////////////////////////////////////////////////////////

        CheckSHButtonsState()
    End Sub
End Class
