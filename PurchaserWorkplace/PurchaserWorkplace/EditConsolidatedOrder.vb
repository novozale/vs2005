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

        Label19.Text = My.Settings.BillPath

        If StartParam = "Create" Then '---создание нового заказа
            Declarations.MyOrderID = Microsoft.VisualBasic.Right("0000000000" & CStr(GetNewID()), 10)
            Label1.Text = Declarations.MyOrderID
            Label3.Text = Microsoft.VisualBasic.Right("00" & CStr(DatePart(DateInterval.Day, Now())), 2) & "/" & Microsoft.VisualBasic.Right("00" & CStr(DatePart(DateInterval.Month, Now())), 2) & "/" & CStr(DatePart(DateInterval.Year, Now()))
            Label4.Text = Declarations.MySupplierCode
            Label7.Text = Declarations.MyWH
            Label13.Text = ""
            Label10.Text = ""
            TextBox1.Text = ""
        ElseIf StartParam = "Copy" Then '---создание нового заказа копированием существующего
            Declarations.MyOrderID = Microsoft.VisualBasic.Right("0000000000" & CStr(GetNewID()), 10)
            Label1.Text = Declarations.MyOrderID
            Label3.Text = Microsoft.VisualBasic.Right("00" & CStr(DatePart(DateInterval.Day, Now())), 2) & "/" & Microsoft.VisualBasic.Right("00" & CStr(DatePart(DateInterval.Month, Now())), 2) & "/" & CStr(DatePart(DateInterval.Year, Now()))
            Label4.Text = Declarations.MySupplierCode
            Label7.Text = Declarations.MyWH
            Label13.Text = ""
            Label10.Text = ""
            TextBox1.Text = ""
            TextBox2.Text = MyConsolidatedOrders.DataGridView1.SelectedRows.Item(0).Cells(12).Value
            TextBox3.Text = MyConsolidatedOrders.DataGridView1.SelectedRows.Item(0).Cells(13).Value
        Else                          '---редактирование существующего
            MySQLStr = "SELECT  OrderDate, SupplierPlacedDate, ConfirmedDate, SupplierOrderNumber, "
            MySQLStr = MySQLStr & "InitialDeliveryDate, ControlDate, ContactInfo, Comments, FileName "
            MySQLStr = MySQLStr & "FROM tbl_PurchaseWorkplace_ConsolidatedOrders WITH(NOLOCK) "
            MySQLStr = MySQLStr & "WHERE (ID = N'" & Declarations.MyOrderID & "') "
            InitMyConn(False)
            InitMyRec(False, MySQLStr)
            If Declarations.MyRec.EOF = True And Declarations.MyRec.BOF = True Then
                trycloseMyRec()
                MsgBox("Данный заказ не существует, возможно, удален другим пользователем. Обновите данные в окне заказов.", MsgBoxStyle.Critical, "Внимание!")
                Me.Close()
            Else
                Declarations.MyRec.MoveFirst()
                Label1.Text = Declarations.MyOrderID
                Label3.Text = Microsoft.VisualBasic.Right("00" & CStr(DatePart(DateInterval.Day, Declarations.MyRec.Fields("OrderDate").Value)), 2) & "/" & Microsoft.VisualBasic.Right("00" & CStr(DatePart(DateInterval.Month, Declarations.MyRec.Fields("OrderDate").Value)), 2) & "/" & CStr(DatePart(DateInterval.Year, Declarations.MyRec.Fields("OrderDate").Value))
                Label4.Text = Declarations.MySupplierCode
                Label7.Text = Declarations.MyWH
                If IsDBNull(Declarations.MyRec.Fields("SupplierPlacedDate").Value) = True Then
                    Label13.Text = ""
                Else
                    Label13.Text = Microsoft.VisualBasic.Right("00" & CStr(DatePart(DateInterval.Day, Declarations.MyRec.Fields("SupplierPlacedDate").Value)), 2) & "/" & Microsoft.VisualBasic.Right("00" & CStr(DatePart(DateInterval.Month, Declarations.MyRec.Fields("SupplierPlacedDate").Value)), 2) & "/" & CStr(DatePart(DateInterval.Year, Declarations.MyRec.Fields("SupplierPlacedDate").Value))
                End If
                If IsDBNull(Declarations.MyRec.Fields("ConfirmedDate").Value) = True Then
                    Label10.Text = ""
                Else
                    Label10.Text = Microsoft.VisualBasic.Right("00" & CStr(DatePart(DateInterval.Day, Declarations.MyRec.Fields("ConfirmedDate").Value)), 2) & "/" & Microsoft.VisualBasic.Right("00" & CStr(DatePart(DateInterval.Month, Declarations.MyRec.Fields("ConfirmedDate").Value)), 2) & "/" & CStr(DatePart(DateInterval.Year, Declarations.MyRec.Fields("ConfirmedDate").Value))
                End If
                TextBox1.Text = Trim(Declarations.MyRec.Fields("SupplierOrderNumber").Value.ToString)
                If Trim(TextBox1.Text) <> "" Then
                    TextBox1.ReadOnly = True
                End If
                If IsDBNull(Declarations.MyRec.Fields("InitialDeliveryDate").Value) Then
                Else
                    DateTimePicker1.Value = Declarations.MyRec.Fields("InitialDeliveryDate").Value
                End If
                If IsDBNull(Declarations.MyRec.Fields("ControlDate").Value) Then
                Else
                    DateTimePicker2.Value = Declarations.MyRec.Fields("ControlDate").Value
                End If
                TextBox2.Text = Trim(Declarations.MyRec.Fields("ContactInfo").Value.ToString)
                TextBox3.Text = Trim(Declarations.MyRec.Fields("Comments").Value.ToString)
                TextBox4.Text = Trim(Declarations.MyRec.Fields("FileName").Value.ToString)

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
            If StartParam = "Create" Or StartParam = "Copy" Then
                SaveNewData()
            Else
                UpdateData()
            End If
            Me.Close()
        End If
    End Sub

    Private Function CheckData() As Boolean
        '////////////////////////////////////////////////////////////////////////////////
        '//
        '// Проверка внесенных данных
        '//
        '////////////////////////////////////////////////////////////////////////////////
        Dim MyDate As Date
        Dim CurrDate As Date

        'MyDate = DateTimePicker1.Value.Date
        'CurrDate = Now().Date
        'If DateTime.Compare(MyDate, CurrDate) < 0 Then
        '    MsgBox("Первоначальная дата поставки не должна быть меньше текущей даты. ", MsgBoxStyle.Critical, "Внимание!")
        '    DateTimePicker1.Select()
        '    CheckData = False
        '    Exit Function
        'End If

        'MyDate = DateTimePicker2.Value.Date
        'CurrDate = Now().Date
        'If DateTime.Compare(MyDate, CurrDate) < 0 Then
        '    MsgBox("Контрольная дата не должна быть меньше текущей даты. ", MsgBoxStyle.Critical, "Внимание!")
        '    DateTimePicker2.Select()
        '    CheckData = False
        '    Exit Function
        'End If

        If Trim(TextBox2.Text) = "" Then
            MsgBox("Поле ""Контактная информация"" должно быть заполнено. ", MsgBoxStyle.Critical, "Внимание!")
            TextBox2.Select()
            CheckData = False
            Exit Function
        End If

        If Trim(TextBox4.Text) <> "" Then
            If System.IO.Directory.Exists(My.Settings.BillPath + Trim(TextBox4.Text)) = False _
            And System.IO.File.Exists(My.Settings.BillPath + Trim(TextBox4.Text)) = False Then
                MsgBox("Введенный вами файл или каталог не найден в " & My.Settings.BillPath & ". ", MsgBoxStyle.Critical, "Внимание!")
                TextBox4.Select()
                CheckData = False
                Exit Function
            End If
        End If

        CheckData = True
    End Function

    Private Sub SaveNewData()
        '////////////////////////////////////////////////////////////////////////////////
        '//
        '// Сохранение нового заказа
        '//
        '////////////////////////////////////////////////////////////////////////////////
        Dim MySQLStr As String
        Dim MyRez As MsgBoxResult

        If Trim(TextBox1.Text) <> "" Then
            If TextBox1.ReadOnly = True Then
                MyRez = MsgBoxResult.Yes
            Else
                MyRez = MsgBox("Вы заполнили поле 'Номер заказа поставщика'. Это значит, что после сохранения вы не сможете редактировать этот заказ - добавлять в него заказы на закупку и удалять из него заказы на закупку. Будете сохранять?", MsgBoxStyle.YesNo, "Внимание!")
            End If
        Else
            MyRez = MsgBoxResult.Yes
        End If

        If MyRez = MsgBoxResult.Yes Then
            MySQLStr = "INSERT INTO tbl_PurchaseWorkplace_ConsolidatedOrders "
            MySQLStr = MySQLStr & "(ID, OrderDate, SupplierCode, WH, SupplierPlacedDate, ConfirmedDate, SupplierOrderNumber, UserID, "
            MySQLStr = MySQLStr & "InitialDeliveryDate, ControlDate, ContactInfo, Comments, FileName) "
            MySQLStr = MySQLStr & "VALUES (N'" & Declarations.MyOrderID & "', "
            MySQLStr = MySQLStr & "CONVERT(DATETIME, '" & Label3.Text & "', 103), "
            MySQLStr = MySQLStr & "N'" & Declarations.MySupplierCode & "', "
            MySQLStr = MySQLStr & "N'" & Declarations.MyWH & "', "
            If Trim(TextBox1.Text) <> "" Then
                MySQLStr = MySQLStr & "CONVERT(DATETIME, '" & Microsoft.VisualBasic.Right("00" & CStr(DatePart(DateInterval.Day, Now())), 2) & "/" & Microsoft.VisualBasic.Right("00" & CStr(DatePart(DateInterval.Month, Now())), 2) & "/" & CStr(DatePart(DateInterval.Year, Now())) & "', 103), "
            Else
                MySQLStr = MySQLStr & "NULL, "
            End If
            If Trim(Label10.Text) = "" Then
                MySQLStr = MySQLStr & "NULL, "
            Else
                MySQLStr = MySQLStr & "CONVERT(DATETIME, '" & Label10.Text & "', 103), "
            End If
            MySQLStr = MySQLStr & "N'" & Trim(TextBox1.Text) & "', "
            MySQLStr = MySQLStr & Declarations.UserID & ", "
            MySQLStr = MySQLStr & "CONVERT(DATETIME, '" & Microsoft.VisualBasic.Right("00" & CStr(DateTimePicker1.Value.Day), 2) & "/" & Microsoft.VisualBasic.Right("00" & CStr(DateTimePicker1.Value.Month), 2) & "/" & DateTimePicker1.Value.Year & "', 103), "
            MySQLStr = MySQLStr & "CONVERT(DATETIME, '" & Microsoft.VisualBasic.Right("00" & CStr(DateTimePicker2.Value.Day), 2) & "/" & Microsoft.VisualBasic.Right("00" & CStr(DateTimePicker2.Value.Month), 2) & "/" & DateTimePicker2.Value.Year & "', 103), "
            If Trim(TextBox2.Text) <> "" Then
                MySQLStr = MySQLStr & "N'" & Trim(TextBox2.Text) & "', "
            Else
                MySQLStr = MySQLStr & "NULL, "
            End If
            If Trim(TextBox3.Text) <> "" Then
                MySQLStr = MySQLStr & "N'" & Trim(TextBox3.Text) & "', "
            Else
                MySQLStr = MySQLStr & "NULL, "
            End If
            If Trim(TextBox4.Text) <> "" Then
                MySQLStr = MySQLStr & "N'" & Trim(TextBox4.Text) & "') "
            Else
                MySQLStr = MySQLStr & "NULL) "
            End If


            InitMyConn(False)
            Declarations.MyConn.Execute(MySQLStr)
        End If
    End Sub

    Private Sub UpdateData()
        '////////////////////////////////////////////////////////////////////////////////
        '//
        '// Обновление значений редактируемого заказа
        '//
        '////////////////////////////////////////////////////////////////////////////////
        Dim MySQLStr As String
        Dim MyRez As MsgBoxResult

        If Trim(TextBox1.Text) <> "" Then
            If TextBox1.ReadOnly = True Then
                MyRez = MsgBoxResult.Yes
            Else
                MyRez = MsgBox("Вы заполнили поле 'Номер заказа поставщика'. Это значит, что после сохранения вы не сможете редактировать этот заказ - добавлять в него заказы на закупку и удалять из него заказы на закупку. Будете сохранять?", MsgBoxStyle.YesNo, "Внимание!")
            End If
        Else
            MyRez = MsgBoxResult.Yes
        End If

        If MyRez = MsgBoxResult.Yes Then
            MySQLStr = "UPDATE tbl_PurchaseWorkplace_ConsolidatedOrders "
            If Trim(TextBox1.Text) <> "" Then
                MySQLStr = MySQLStr & "SET SupplierPlacedDate = CONVERT(DATETIME, '" & Microsoft.VisualBasic.Right("00" & CStr(DatePart(DateInterval.Day, Now())), 2) & "/" & Microsoft.VisualBasic.Right("00" & CStr(DatePart(DateInterval.Month, Now())), 2) & "/" & CStr(DatePart(DateInterval.Year, Now())) & "', 103), "
            Else
                MySQLStr = MySQLStr & "SET SupplierPlacedDate = NULL, "
            End If
            MySQLStr = MySQLStr & "SupplierOrderNumber = N'" & Trim(TextBox1.Text) & "', "
            MySQLStr = MySQLStr & "InitialDeliveryDate = CONVERT(DATETIME, '" & Microsoft.VisualBasic.Right("00" & CStr(DateTimePicker1.Value.Day), 2) & "/" & Microsoft.VisualBasic.Right("00" & CStr(DateTimePicker1.Value.Month), 2) & "/" & DateTimePicker1.Value.Year & "', 103), "
            MySQLStr = MySQLStr & "ControlDate = CONVERT(DATETIME, '" & Microsoft.VisualBasic.Right("00" & CStr(DateTimePicker2.Value.Day), 2) & "/" & Microsoft.VisualBasic.Right("00" & CStr(DateTimePicker2.Value.Month), 2) & "/" & DateTimePicker2.Value.Year & "', 103), "
            MySQLStr = MySQLStr & "ContactInfo = N'" & Trim(TextBox2.Text) & "', "
            If Trim(TextBox3.Text) <> "" Then
                MySQLStr = MySQLStr & "Comments = N'" & Trim(TextBox3.Text) & "', "
            Else
                MySQLStr = MySQLStr & "Comments = NULL, "
            End If
            If Trim(TextBox4.Text) <> "" Then
                MySQLStr = MySQLStr & "FileName = N'" & Trim(TextBox4.Text) & "' "
            Else
                MySQLStr = MySQLStr & "FileName = NULL "
            End If
            MySQLStr = MySQLStr & "WHERE (ID = N'" & Declarations.MyOrderID & "')"
            InitMyConn(False)
            Declarations.MyConn.Execute(MySQLStr)
        End If

    End Sub

    Private Sub Button2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button2.Click
        '////////////////////////////////////////////////////////////////////////////////
        '//
        '// Выбор файла с документом (счетом)
        '//
        '////////////////////////////////////////////////////////////////////////////////
        Dim MyDlg As OpenFileDialog

        MyDlg = New OpenFileDialog
        MyDlg.Filter = "Все файлы (*.*)|*.*"
        MyDlg.Multiselect = False
        MyDlg.SupportMultiDottedExtensions = True
        MyDlg.Title = "Выберите файл со счетом для этого заказа"
        MyDlg.InitialDirectory = My.Settings.BillPath
        If MyDlg.ShowDialog() <> DialogResult.OK Then
        Else
            If GetRestPart(MyDlg.FileName) = "" Then
                MsgBox("Выбранный вами файл " & Trim(MyDlg.FileName) & " не находится в каталоге " & My.Settings.BillPath & ". Выберите корректный файл.", MsgBoxStyle.Critical, "Внимание!")
            Else
                TextBox4.Text = GetRestPart(MyDlg.FileName)
            End If
        End If
    End Sub

    Private Function GetRestPart(ByVal MySrcPath As String) As String
        '////////////////////////////////////////////////////////////////////////////////
        '//
        '// Получение пути к файлу за вычетом пути к общей папке
        '//
        '////////////////////////////////////////////////////////////////////////////////
        Dim MyCodeLength As Integer

        MyCodeLength = InStr(MySrcPath, My.Settings.BillCommonPath)
        If MyCodeLength <= 0 Then
            GetRestPart = ""
        Else
            GetRestPart = Mid(MySrcPath, MyCodeLength + Len(My.Settings.BillCommonPath), Len(MySrcPath) - (MyCodeLength + Len(My.Settings.BillCommonPath) - 1))
        End If
    End Function

    Private Sub Button3_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button3.Click
        '////////////////////////////////////////////////////////////////////////////////
        '//
        '// Выбор каталога с документами (счетами)
        '//
        '////////////////////////////////////////////////////////////////////////////////
        Dim MyDlg As FolderBrowserDialog

        MyDlg = New FolderBrowserDialog
        MyDlg.SelectedPath = My.Settings.BillPath
        If MyDlg.ShowDialog <> Windows.Forms.DialogResult.OK Then
        Else
            If GetRestPart(MyDlg.SelectedPath) = "" Then
                MsgBox("Выбранный вами каталог " & Trim(MyDlg.SelectedPath) & " не находится в каталоге " & My.Settings.BillPath & ". Выберите корректный каталог.", MsgBoxStyle.Critical, "Внимание!")
            Else
                TextBox4.Text = GetRestPart(MyDlg.SelectedPath)
            End If
        End If
    End Sub
End Class