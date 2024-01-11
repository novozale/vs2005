Imports System.Net
Imports System.Xml

Public Class Shipment

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        '////////////////////////////////////////////////////////////////////////////////
        '//
        '// Выход
        '//
        '////////////////////////////////////////////////////////////////////////////////

        Me.Close()
    End Sub

    Private Sub Shipment_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        '////////////////////////////////////////////////////////////////////////////////
        '//
        '// Загрузка окна, выставление параметров
        '//
        '////////////////////////////////////////////////////////////////////////////////
        Dim MySQLStr As String
        Dim CustDelAddress As String

        CustDelAddress = ""

        '------Информация по покупателю------------------
        MySQLStr = "SELECT CASE WHEN Ltrim(Rtrim(tbl_CustomerCard0300.LongName)) <> '' THEN Ltrim(Rtrim(tbl_CustomerCard0300.LongName)) "
        MySQLStr = MySQLStr & "ELSE SL010300.SL01002 END AS CustomerName, CASE WHEN Ltrim(Rtrim(tbl_CustomerCard0300.LongAddress)) "
        MySQLStr = MySQLStr & "<> '' THEN Ltrim(Rtrim(tbl_CustomerCard0300.LongAddress)) ELSE Ltrim(Rtrim(Ltrim(Rtrim(SL010300.SL01003)) + ' ' + Ltrim(Rtrim(SL010300.SL01004)) "
        MySQLStr = MySQLStr & "+ ' ' + Ltrim(Rtrim(SL010300.SL01005)))) END AS CustomerAddress, SL010300.SL01021 AS CustomerINN, "
        MySQLStr = MySQLStr & "LTRIM(RTRIM(LTRIM(RTRIM(ISNULL(View_10.SL14004, N''))) + ' ' + LTRIM(RTRIM(ISNULL(View_10.SL14005, N''))) "
        MySQLStr = MySQLStr & "+ ' ' + LTRIM(RTRIM(ISNULL(View_10.SL14006, N''))))) AS CustomerDelAddress "
        MySQLStr = MySQLStr & "FROM SL010300 INNER JOIN "
        MySQLStr = MySQLStr & "tbl_CustomerCard0300 ON SL010300.SL01001 = tbl_CustomerCard0300.SL01001 LEFT OUTER JOIN "
        MySQLStr = MySQLStr & "(SELECT SL14001, SL14004, SL14005, SL14006 "
        MySQLStr = MySQLStr & "FROM SL140300 "
        MySQLStr = MySQLStr & "WHERE (SL14002 = N'00')) AS View_10 ON SL010300.SL01001 = View_10.SL14001 "
        MySQLStr = MySQLStr & "WHERE (SL010300.SL01001 = N'" & Trim(Declarations.MyCustomerCode) & "') "
        InitMyConn(False)
        InitMyRec(False, MySQLStr)
        If Declarations.MyRec.EOF = True And Declarations.MyRec.BOF = True Then
            MsgBox("Не найдена информация по клиенту " & Trim(Declarations.MyCustomerCode) & ". Обратитесь к администратору.", MsgBoxStyle.Critical, "Внимание!")
            trycloseMyRec()
            Me.Close()
        Else
            LblCustomerCode.Text = Declarations.MyCustomerCode
            LblCustomerName.Text = Declarations.MyRec.Fields("CustomerName").Value
            LblCustomerLegalAddress.Text = Declarations.MyRec.Fields("CustomerAddress").Value
            LblCustomerINN.Text = Declarations.MyRec.Fields("CustomerINN").Value
            CustDelAddress = Declarations.MyRec.Fields("CustomerDelAddress").Value
            trycloseMyRec()
        End If

        '------Информация по продавцу--------------------
        LblSalesmanCode.Text = Declarations.SalesmanCode
        LblSalesmanName.Text = Declarations.UserName
        LblWHCode.Text = Declarations.MyWH

        '------Информация по доставке--------------------
        MySQLStr = "SELECT tbl_OR010300.CAddr, tbl_OR010300.DeliveryAddr, tbl_OR010300.DeliveryDate, "
        MySQLStr = MySQLStr & "ISNULL(tbl_SalesHdr_AddInfo.DeliverySum, 0) AS DeliverySumm "
        MySQLStr = MySQLStr & "FROM tbl_SalesHdr_AddInfo RIGHT OUTER JOIN "
        MySQLStr = MySQLStr & "tbl_OR010300 ON tbl_SalesHdr_AddInfo.OrderID = tbl_OR010300.OrderN "
        MySQLStr = MySQLStr & "WHERE (tbl_OR010300.OrderN = N'" & Trim(Declarations.MyOrderNum) & "') "

        InitMyConn(False)
        InitMyRec(False, MySQLStr)
        If Declarations.MyRec.EOF = True And Declarations.MyRec.BOF = True Then
            ComboBox1.Text = "Самовывоз"
            TextBox1.Text = ""
            TextBox2.Text = ""
            TextBox3.Text = ""
            DateTimePicker1.Value = Now()
            trycloseMyRec()
        Else
            If Declarations.MyRec.Fields("DeliverySumm").Value = 0 Then
                ComboBox1.Text = "Самовывоз"
            Else
                ComboBox1.Text = "Доставка WEB"
            End If
            TextBox1.Text = Declarations.MyRec.Fields("CAddr").Value
            TextBox2.Text = Declarations.MyRec.Fields("DeliveryAddr").Value
            TextBox3.Text = ""
            If Declarations.MyRec.Fields("DeliveryDate").Value < Now() Then
                DateTimePicker1.Value = Now()
            Else
                DateTimePicker1.Value = Declarations.MyRec.Fields("DeliveryDate").Value
            End If
        End If
    End Sub

    Private Sub Button3_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button3.Click
        '////////////////////////////////////////////////////////////////////////////////
        '//
        '// Открытие окна контактов из CRM
        '//
        '////////////////////////////////////////////////////////////////////////////////

        Dim MyContactInfo = New ContactInfo
        MyContactInfo.StartParam = "Contact"
        MyContactInfo.ShowDialog()
    End Sub

    Private Sub Button4_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button4.Click
        '////////////////////////////////////////////////////////////////////////////////
        '//
        '// Открытие окна адресов из Scala
        '//
        '////////////////////////////////////////////////////////////////////////////////

        Dim MyDelAddresses = New DelAddresses
        MyDelAddresses.ShowDialog()
    End Sub

    Private Sub Button5_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button5.Click
        '////////////////////////////////////////////////////////////////////////////////
        '//
        '// Занесение пути к файлу, который будет приаттачен
        '//
        '////////////////////////////////////////////////////////////////////////////////
        Dim MyDlg As OpenFileDialog

        MyDlg = New OpenFileDialog
        MyDlg.Filter = "Все файлы (*.*)|*.*"
        MyDlg.Multiselect = False
        MyDlg.SupportMultiDottedExtensions = True
        MyDlg.Title = "Выберите файл"
        MyDlg.InitialDirectory = "::{20D04FE0-3AEA-1069-A2D8-08002B30309D}"
        MyDlg.RestoreDirectory = True
        If MyDlg.ShowDialog() <> DialogResult.OK Then
        Else
            TextBox4.Text = MyDlg.FileName
        End If
    End Sub

    Private Sub Button6_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button6.Click
        '////////////////////////////////////////////////////////////////////////////////
        '//
        '// очистка значения
        '//
        '////////////////////////////////////////////////////////////////////////////////

        TextBox4.Text = ""
    End Sub

    Private Sub Button2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button2.Click
        '////////////////////////////////////////////////////////////////////////////////
        '//
        '// Сохранение информации о доставке
        '//
        '////////////////////////////////////////////////////////////////////////////////

        If CheckFiields() = True Then
            Declarations.MyShipmentsID = SaveShipment()
            If Declarations.MyShipmentsID <> 0 Then
                AddOrderToShipment(Declarations.MyShipmentsID)
                MakeRequestToSharepoint(Declarations.MyShipmentsID)
            End If
            MyOperationFlag = 1
            Me.Close()
        End If
    End Sub

    Private Function CheckFiields() As Boolean
        '////////////////////////////////////////////////////////////////////////////////
        '//
        '// проверка сохранения полей
        '//
        '////////////////////////////////////////////////////////////////////////////////
        Dim DelDate As Date
        Dim CurrDate As Date

        If Trim(TextBox1.Text) = "" Then
            MsgBox("Поле ""Контактная информация"" должно быть заполнено. ", MsgBoxStyle.Critical, "Внимание!")
            TextBox1.Select()
            CheckFiields = False
            Exit Function
        End If

        If ComboBox1.Text = "Доставка" Or ComboBox1.Text = "Доставка с оплатой клиентом" Then
            If Trim(TextBox2.Text) = "" Then
                MsgBox("Поле ""Адрес доставки"" должно быть заполнено. ", MsgBoxStyle.Critical, "Внимание!")
                TextBox2.Select()
                CheckFiields = False
                Exit Function
            End If
        End If

        DelDate = DateTimePicker1.Value.Date
        CurrDate = Now().Date
        If DateTime.Compare(DelDate, CurrDate) < 0 Then
            MsgBox("Дата отгрузки со склада не должна быть меньше текущей даты. ", MsgBoxStyle.Critical, "Внимание!")
            DateTimePicker1.Select()
            CheckFiields = False
            Exit Function
        End If

        CheckFiields = True
    End Function

    Private Function SaveShipment() As Integer
        '////////////////////////////////////////////////////////////////////////////////
        '//
        '// создание записи об отгрузке
        '//
        '////////////////////////////////////////////////////////////////////////////////
        Dim MySQLStr As String

        Try
            MySQLStr = "INSERT INTO tbl_Shipments_SalesmanWP_Info "
            MySQLStr = MySQLStr & "(CustomerCode, CustomerName, CustomerINN, CustomerLegalAddress, SalesmanCode, WHCode, DeliveryOrNot, DeliverySumm, DeliveredSumm, "
            MySQLStr = MySQLStr & "ContactInfo, DeliveryAddress, Comment, PrintBillOrNot, PrintBillOrNot1, PrintFullBillOrNot, RequestedShipmentDate, IsRequestSend, IsReminderSend, AttFile) "
            MySQLStr = MySQLStr & "VALUES (N'" & Trim(LblCustomerCode.Text) & "', "
            MySQLStr = MySQLStr & "N'" & Trim(LblCustomerName.Text) & "', "
            MySQLStr = MySQLStr & "N'" & Trim(LblCustomerINN.Text) & "', "
            MySQLStr = MySQLStr & "N'" & Trim(LblCustomerLegalAddress.Text) & "', "
            MySQLStr = MySQLStr & "N'" & Trim(LblSalesmanCode.Text) & "', "
            MySQLStr = MySQLStr & "N'" & Trim(LblWHCode.Text) & "', "
            If ComboBox1.Text = "Доставка WEB" Then
                MySQLStr = MySQLStr & "3, "
            Else
                MySQLStr = MySQLStr & "0, "
            End If
            MySQLStr = MySQLStr & Replace(Replace(MainForm.DataGridView2.SelectedRows.Item(0).Cells(8).Value.ToString(), ",", "."), " ", "") + ", "
            MySQLStr = MySQLStr & Replace(Replace(MainForm.DataGridView2.SelectedRows.Item(0).Cells(10).Value.ToString(), ",", "."), " ", "") + ", "
            MySQLStr = MySQLStr & "N'" & Trim(TextBox1.Text) & "', "
            MySQLStr = MySQLStr & "N'" & Trim(TextBox2.Text) & "', "
            MySQLStr = MySQLStr & "N'" & Trim(TextBox3.Text) & "', "
            If CheckBox1.Checked = True Then
                MySQLStr = MySQLStr & "1, "
            Else
                MySQLStr = MySQLStr & "0, "
            End If
            If CheckBox2.Checked = True Then
                MySQLStr = MySQLStr & "1, "
            Else
                MySQLStr = MySQLStr & "0, "
            End If
            If CheckBox3.Checked = True Then
                MySQLStr = MySQLStr & "1, "
            Else
                MySQLStr = MySQLStr & "0, "
            End If
            MySQLStr = MySQLStr & "CONVERT(DATETIME, '" & Microsoft.VisualBasic.Right("00" & CStr(DateTimePicker1.Value.Day), 2) & "/" & Microsoft.VisualBasic.Right("00" & CStr(DateTimePicker1.Value.Month), 2) & "/" & DateTimePicker1.Value.Year & "', 103), "
            MySQLStr = MySQLStr & "1, "
            MySQLStr = MySQLStr & "0, "
            MySQLStr = MySQLStr & "N'" & Trim(TextBox4.Text) & "') "
            InitMyConn(False)
            Declarations.MyConn.Execute(MySQLStr)
            '--------ID записи
            MySQLStr = "SELECT MAX(ID) AS CC "
            MySQLStr = MySQLStr & "FROM tbl_Shipments_SalesmanWP_Info "
            InitMyConn(False)
            InitMyRec(False, MySQLStr)
            If Declarations.MyRec.EOF = True And Declarations.MyRec.BOF = True Then
                SaveShipment = 0
                trycloseMyRec()
                Exit Function
            Else
                SaveShipment = Declarations.MyRec.Fields("CC").Value
                trycloseMyRec()
                Exit Function
            End If

        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Внимание!")
            SaveShipment = 0
        End Try
    End Function

    Private Sub AddOrderToShipment(ByVal ShID As Integer)
        '////////////////////////////////////////////////////////////////////////////////
        '//
        '// добавление выбранного заказа в отгрузку
        '//
        '////////////////////////////////////////////////////////////////////////////////

        Dim MySQLStr As String

        MySQLStr = "INSERT INTO tbl_Shipments_SalesmanWP_Details "
        MySQLStr = MySQLStr & "(ShipmentsID, OrderNum, InvoiceNum, IsClosed) "
        MySQLStr = MySQLStr & "VALUES (" & ShID.ToString & ", "
        MySQLStr = MySQLStr & "N'" & MainForm.DataGridView2.SelectedRows.Item(0).Cells(0).Value.ToString() & "', "
        MySQLStr = MySQLStr & "'', 0)"
        InitMyConn(False)
        Declarations.MyConn.Execute(MySQLStr)
    End Sub

    Private Sub MakeRequestToSharepoint(ByVal ShID As Integer)
        '////////////////////////////////////////////////////////////////////////////////
        '//
        '// формирование запроса на портал
        '//
        '////////////////////////////////////////////////////////////////////////////////
        Dim MySQLStr As String

        MySQLStr = "SELECT tbl_Shipments_SalesmanWP_Info.CustomerCode, tbl_Shipments_SalesmanWP_Info.CustomerName, tbl_Shipments_SalesmanWP_Info.CustomerINN, "
        MySQLStr = MySQLStr & "tbl_Shipments_SalesmanWP_Info.CustomerLegalAddress, tbl_Shipments_SalesmanWP_Info.WHCode + ' ' + ISNULL(SC230300.SC23002, N'') "
        MySQLStr = MySQLStr & "AS WHCode, "
        MySQLStr = MySQLStr & "CASE DeliveryOrNot WHEN 0 THEN 'Самовывоз' WHEN 1 THEN 'Доставка' WHEN 2 THEN 'Доставка с оплатой клиентом' ELSE 'Доставка WEB' END "
        MySQLStr = MySQLStr & "AS DeliveryOrNot, tbl_Shipments_SalesmanWP_Info.DeliverySumm, tbl_Shipments_SalesmanWP_Info.ContactInfo, "
        MySQLStr = MySQLStr & "tbl_Shipments_SalesmanWP_Info.DeliveryAddress, tbl_Shipments_SalesmanWP_Info.Comment, tbl_Shipments_SalesmanWP_Info.PrintBillOrNot, "
        MySQLStr = MySQLStr & "tbl_Shipments_SalesmanWP_Info.PrintBillOrNot1, tbl_Shipments_SalesmanWP_Info.PrintFullBillOrNot, "
        MySQLStr = MySQLStr & "tbl_Shipments_SalesmanWP_Info.RequestedShipmentDate, tbl_Shipments_SalesmanWP_Details.OrderNum, "
        MySQLStr = MySQLStr & "tbl_Shipments_SalesmanWP_Info.AttFile "
        MySQLStr = MySQLStr & "FROM tbl_Shipments_SalesmanWP_Info INNER JOIN "
        MySQLStr = MySQLStr & "tbl_Shipments_SalesmanWP_Details ON tbl_Shipments_SalesmanWP_Info.ID = tbl_Shipments_SalesmanWP_Details.ShipmentsID LEFT OUTER JOIN "
        MySQLStr = MySQLStr & "SC230300 ON tbl_Shipments_SalesmanWP_Info.WHCode = SC230300.SC23001 "
        MySQLStr = MySQLStr & "WHERE (tbl_Shipments_SalesmanWP_Info.ID = " & ShID.ToString() & ") "
        InitMyConn(False)
        InitMyRec(False, MySQLStr)
        If Declarations.MyRec.EOF = True And Declarations.MyRec.BOF = True Then
        Else
            CreateRequest(Declarations.MyRec.Fields("CustomerCode").Value, _
                                Declarations.MyRec.Fields("CustomerName").Value, _
                                Declarations.MyRec.Fields("CustomerINN").Value, _
                                Declarations.MyRec.Fields("CustomerLegalAddress").Value, _
                                "ESKRU\" + Declarations.UserName, _
                                Declarations.MyRec.Fields("WHCode").Value, _
                                Declarations.MyRec.Fields("DeliveryOrNot").Value, _
                                Declarations.MyRec.Fields("DeliverySumm").Value, _
                                Declarations.MyRec.Fields("ContactInfo").Value, _
                                Declarations.MyRec.Fields("DeliveryAddress").Value, _
                                Declarations.MyRec.Fields("Comment").Value, _
                                Declarations.MyRec.Fields("PrintBillOrNot").Value, _
                                Declarations.MyRec.Fields("PrintBillOrNot1").Value, _
                                Declarations.MyRec.Fields("PrintFullBillOrNot").Value, _
                                Declarations.MyRec.Fields("RequestedShipmentDate").Value, _
                                Declarations.MyRec.Fields("OrderNum").Value, _
                                Declarations.MyRec.Fields("AttFile").Value)
            MsgBox("Заявка на портале создана.", MsgBoxStyle.Information, "Внимание!")
        End If

    End Sub

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
        listWebService.Credentials = New System.Net.NetworkCredential("developer", "!Devpass", "ESKRU")
        Dim listName = "{3e66b7ae-a55c-4e6f-92be-9f602e7d0417}"
        'Dim listName = "logistics/{3e66b7ae-a55c-4e6f-92be-9f602e7d0417}"

        Dim listView = ""
        Dim listItemId As String = ""
        Dim FileName As String = ""
        Dim MyAttachment As Byte()

        Dim strBatch As String = "<Method ID='1' Cmd='New'>"
        strBatch = strBatch + "<Field Name='ID'>New</Field>"
        strBatch = strBatch + "<Field Name='_x041a__x043e__x0434__x0020__x04'>" & MyCustomerCode & "</Field>"           '---Код клиента
        strBatch = strBatch + "<Field Name='Title'>" & MyCustomerName & "</Field>"                                      '---название клиента
        strBatch = strBatch + "<Field Name='_x0418__x041d__x041d__x0020__x04'>" & MyINN & "</Field>"                    '---ИНН клиента
        strBatch = strBatch + "<Field Name='_x042e__x0440__x0020__x0430__x04'>" & MyLegalAddress & "</Field>"           '---Юридический адрес клиента
        strBatch = strBatch + "<Field Name='_x041d__x043e__x043c__x0435__x04'>" & OrderList & "</Field>"                '---список заказов
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
End Class