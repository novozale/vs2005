Imports System.Net
Imports System.Xml

Public Class CreateSPTasks
    Public CommonPurchOrderNum As String

    Private Sub TestSub()
        'Dim listName = "{a737e822-163e-48c2-870d-eb45b2164e15}"
        'Dim listWebService As spbprd4.Lists = New spbprd4.Lists()
        'Dim listItems As System.Xml.XmlNode
        'Dim xmlDoc As XmlDocument = New XmlDocument()
        'Dim viewFields As XmlElement = xmlDoc.CreateElement("ViewFields")
        'Dim query As XmlElement = xmlDoc.CreateElement("Query")
        'Dim queryOptions As XmlElement = xmlDoc.CreateElement("QueryOptions")

        ''listWebService.Credentials = System.Net.CredentialCache.DefaultCredentials
        'listWebService.Credentials = New System.Net.NetworkCredential("novozhilov", "!531alexandr37", "ESKRU")
        'Try
        '    query.InnerXml = "<Where><Gt><FieldRef Name='ID'/><Value Type='Counter'>0</Value></Gt></Where>"
        '    viewFields.InnerXml = ""
        '    queryOptions.InnerXml = "<IncludeMandatoryColumns>TRUE</IncludeMandatoryColumns>"
        '    listItems = listWebService.GetListItems(listName, "", query, viewFields, "100", queryOptions, Nothing)
        '    For Each ff As System.Xml.XmlNode In listItems
        '        If ff.Name = "rs:data" Then
        '            For i As Integer = 0 To ff.ChildNodes.Count - 1
        '                If ff.ChildNodes(i).Name = "z:row" Then
        '                    MsgBox(ff.ChildNodes(i).Attributes("ows_Title").Value)
        '                End If
        '            Next
        '        End If
        '    Next
        'Catch ex As Exception
        '    MsgBox(ex.Message)
        'End Try

        'Dim listWebService As spbprd4.Lists = New spbprd4.Lists()
        ''listWebService.Credentials = System.Net.CredentialCache.DefaultCredentials
        'listWebService.Credentials = New System.Net.NetworkCredential("novozhilov", "!53alexandr37", "ESKRU")
        'Dim listName = "{4ad35c7c-ee5b-4c38-b9ab-775087bd73ae}"

        'Dim ndListView As System.Xml.XmlNode = listWebService.GetListAndView(listName, "")
        'Dim strViewID As String = ndListView.ChildNodes(1).Attributes("Name").Value


        'Dim listView = ""
        'Dim strBatch As String = "<Method ID='1' Cmd='New'>"
        'strBatch = strBatch + "<Field Name='ID'>New</Field>"
        'strBatch = strBatch + "<Field Name='Name'>Test</Field>"
        'strBatch = strBatch + "<Field Name='Город'>Санкт Петербург</Field>"
        'strBatch = strBatch + "<Field Name='Торговый код'>01 Small contractor</Field>"
        'strBatch = strBatch + "<Field Name='Группа клиентов Rexel'>C0110 Электромонтажная организация от 1 до 5 сотрудников</Field>"
        'strBatch = strBatch + "<Field Name='Сегмент рынка Rexel'>01 Гражданский</Field>"
        'strBatch = strBatch + "<Field Name='Отрасль'>1 Горнодобывающая и перерабатывающая промышленность</Field>"
        'strBatch = strBatch + "<Field Name='Тип IKA'>4 Не ключевой клиент</Field>"
        'strBatch = strBatch + "<Field Name='Примечание'>Тестовое примечание</Field>"
        'strBatch = strBatch + "<Field Name='Новый или нет'>Существовавший ранее</Field>"
        'strBatch = strBatch + "<Field Name='Cтарый код Scala и название клиента'></Field>"
        'strBatch = strBatch + "</Method>"
        'Dim xmlDoc As XmlDocument = New System.Xml.XmlDocument()
        'Dim elBatch As System.Xml.XmlElement = xmlDoc.CreateElement("Batch")
        'elBatch.SetAttribute("OnError", "Continue")
        'elBatch.SetAttribute("ListVersion", "1")
        'elBatch.SetAttribute("ViewName", strViewID)
        'elBatch.InnerXml = strBatch
        'Dim ndReturn As XmlNode = listWebService.UpdateListItems(listName, elBatch)

        'Dim listWebService As spbprd4.Lists = New spbprd4.Lists()
        'listWebService.Credentials = New System.Net.NetworkCredential("developer", "!Devpass", "ESKRU")
        'Dim listName = "{a737e822-163e-48c2-870d-eb45b2164e15}"
        'Dim listView = ""
        'Dim strBatch As String = "<Method ID='1' Cmd='New'>"
        'strBatch = strBatch + "<Field Name='ID'>New</Field>"
        'strBatch = strBatch + "<Field Name='Author'>eskru\\shutylev</Field>"                                    '---владелец

        'strBatch = strBatch + "<Field Name='Title'>0000022234</Field>"                                          '---N обобщенного заказа на закупку
        'strBatch = strBatch + "<Field Name='_x041f__x043e__x0441__x0442__x04'>3046 АББ</Field>"                 '---Поставщик
        'strBatch = strBatch + "<Field Name='N_x0020__x0437__x0430__x043a__x0'>0100545658</Field>"               '---N заказа на продажу
        'strBatch = strBatch + "<Field Name='_x041f__x043e__x043a__x0443__x04'>29347 Светлый путь</Field>"      '---покупатель
        'strBatch = strBatch + "<Field Name='_x041f__x0440__x043e__x0434__x04'>01 Шутылев</Field>"               '---продавец
        'strBatch = strBatch + "<Field Name='N_x0020__x0437__x0430__x043a__x00'>000033346</Field>"               '---N заказа на продажу
        'Dim ItemStr As String
        'ItemStr = "00000002 ДЕМОНСТРАЦИОННЫЙ ОБРАЗЕЦПРОБНИК НАПРЯЖЕНИЯ" + Chr(9) + " --> 02/11/2016" + Chr(9) + " (2)--> 02/11/2016" + Chr(13) + Chr(10)
        'ItemStr = ItemStr + "01050079 ПРОВОЛОКА HK16/100метровМЕДН.,PRYSMIAN упаковка" + Chr(9) + " --> 02/11/2016" + Chr(9) + " (2)--> 02/11/2016" + Chr(13) + Chr(10)
        'ItemStr = ItemStr + "01AH312035 КАБЕЛЬ AHXAMK-W3x120Al+35DRAKA NK     10kV" + Chr(9) + " --> 02/11/2016" + Chr(9) + " (2)--> 02/11/2016" + Chr(13) + Chr(10)
        'ItemStr = ItemStr + "02021512 КАБ.KLVMAAM2x(2x04+04)+04KLVMAAM 2x(2x0.4+0.4)+0.4" + Chr(9) + " --> 04/11/2016" + Chr(9) + " (2)--> 04/11/2016" + Chr(13) + Chr(10)
        'ItemStr = ItemStr + "02153426 КАБЕЛЬ VGA 15M/15M 10.0 МGEMBIRD/CABLEXPERT" + Chr(9) + " --> 02/11/2016" + Chr(9) + " (2)--> 05/11/2016" + Chr(13) + Chr(10)
        'ItemStr = ItemStr + "35S6037669 ДАТЧИК УЛЬТРОЗВУКОВОЙ6037669 UM30-211115 SICK" + Chr(9) + " --> 11/11/2016" + Chr(9) + " (2)--> 14/11/2016" + Chr(13) + Chr(10)
        'ItemStr = ItemStr + "520502010 ТРУБКА ТУТ CFM 19/6 ЧЕРНАЯ С КЛЕЕВЫМ СЛОЕМ 1,2м" + Chr(9) + " --> 02/11/2016" + Chr(9) + " (2)--> 05/11/2016" + Chr(13) + Chr(10)
        'ItemStr = ItemStr + "61ACS55055 ПР.ACS550-01-221А-2 55КВТ220В IP21 3AUA0000007126" + Chr(9) + " --> 11/11/2016" + Chr(9) + " (2)--> 13/11/2016" + Chr(13) + Chr(10)
        'ItemStr = ItemStr + "34R9344540 ПАНЕЛИ ЗАЩИТН. SV КЛЕММ.ПОДКЛ. 2ШТ 9344540 RITTAL" + Chr(9) + " --> 02/11/2016" + Chr(9) + " (2)--> 02/11/2016"
        'strBatch = strBatch + "<Field Name='_x0422__x043e__x0432__x0430__x04'>" & ItemStr & "</Field>"          '---товары
        'strBatch = strBatch + "<Field Name='_x041d__x0430__x0020__x0441__x04'>12</Field>"                       '---На сколько дней сдвигается
        'strBatch = strBatch + "<Field Name='_x041c__x0430__x043a__x0441__x00'>14.11.2016</Field>"               '---Максимальная новая дата поставки
        'strBatch = strBatch + "<Field Name='_x041f__x0440__x0438__x0447__x04'>разгильдяйство закупщика</Field>" '---Причина изменения

        'strBatch = strBatch + "</Method>"
        'Dim xmlDoc As XmlDocument = New System.Xml.XmlDocument()
        'Dim elBatch As System.Xml.XmlElement = xmlDoc.CreateElement("Batch")
        'elBatch.SetAttribute("OnError", "Continue")
        'elBatch.SetAttribute("ListVersion", "1")
        'elBatch.SetAttribute("ViewName", listView)
        'elBatch.InnerXml = strBatch
        'Dim ndReturn As XmlNode = listWebService.UpdateListItems(listName, elBatch)
        'MsgBox("Finished", MsgBoxStyle.OkOnly, "Внимание!")
    End Sub


    Private Sub Button4_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button4.Click
        '////////////////////////////////////////////////////////////////////////////////
        '//
        '// выход из окна
        '//
        '////////////////////////////////////////////////////////////////////////////////

        Me.Close()
    End Sub

    Private Sub CreateSPTasks_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        '////////////////////////////////////////////////////////////////////////////////
        '//
        '// Загрузка окна
        '//
        '////////////////////////////////////////////////////////////////////////////////
        Dim MySQLStr As String
        Dim MyAdapter As SqlClient.SqlDataAdapter
        Dim MyDs As New DataSet

        '---номер обобщенного заказа на закупку
        Label2.Text = Me.CommonPurchOrderNum
        '---предлагаемая максимальная дата поставки
        MySQLStr = "SELECT MAX(PC030300.PC03031) AS CC "
        MySQLStr = MySQLStr & "FROM PC030300 INNER JOIN "
        MySQLStr = MySQLStr & "PC010300 ON PC030300.PC03001 = PC010300.PC01001 "
        MySQLStr = MySQLStr & "WHERE (PC010300.PC01052 = N'" & Trim(Me.CommonPurchOrderNum) & "') "
        InitMyConn(False)
        InitMyRec(False, MySQLStr)
        If Declarations.MyRec.EOF = True And Declarations.MyRec.BOF = True Then
            DateTimePicker1.Value = Today()
        Else
            If Declarations.MyRec.Fields("CC").Value Is Nothing = True Then
                DateTimePicker1.Value = Today()
            Else
                If Declarations.MyRec.Fields("CC").Value < Today() Then
                    DateTimePicker1.Value = Today()
                Else
                    DateTimePicker1.Value = Declarations.MyRec.Fields("CC").Value
                End If
            End If
        End If


        '---Загрузка списка запасов в обобщенном заказе
        MySQLStr = "SELECT PC030300.PC03005 AS ItemNum, PC030300.PC03006 + PC030300.PC03007 AS ItemName "
        MySQLStr = MySQLStr & "FROM PC030300 INNER JOIN "
        MySQLStr = MySQLStr & "PC010300 ON PC030300.PC03001 = PC010300.PC01001 "
        MySQLStr = MySQLStr & "WHERE (PC010300.PC01052 = N'" & Trim(Me.CommonPurchOrderNum) & "') "
        MySQLStr = MySQLStr & "AND (PC030300.PC03010 > 0) "
        MySQLStr = MySQLStr & "GROUP BY PC030300.PC03005, PC030300.PC03006 + PC030300.PC03007 "
        MySQLStr = MySQLStr & "ORDER BY ItemNum "

        Try
            MyAdapter = New SqlClient.SqlDataAdapter(MySQLStr, Declarations.MyNETConnStr)
            MyAdapter.SelectCommand.CommandTimeout = 600
            MyAdapter.Fill(MyDs)
            DataGridView1.DataSource = MyDs.Tables(0)
        Catch ex As Exception
            MsgBox(ex.ToString)
        End Try

        DataGridView1.Columns(0).HeaderText = "Выбор"
        DataGridView1.Columns(0).ReadOnly = False
        DataGridView1.Columns(1).HeaderText = "Код товара"
        DataGridView1.Columns(1).Width = 150
        DataGridView1.Columns(1).ReadOnly = True
        DataGridView1.Columns(2).HeaderText = "Название товара"
        DataGridView1.Columns(2).Width = 500
        DataGridView1.Columns(2).ReadOnly = True
    End Sub

    Private Sub Button2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button2.Click
        '////////////////////////////////////////////////////////////////////////////////
        '//
        '// Выбрать все товары в списке
        '//
        '////////////////////////////////////////////////////////////////////////////////

        Windows.Forms.Cursor.Current = Cursors.WaitCursor
        For i As Integer = 0 To DataGridView1.Rows.Count - 1
            DataGridView1.Item(0, i).Value = -1
        Next
        Windows.Forms.Cursor.Current = Cursors.Default
    End Sub

    Private Sub Button3_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button3.Click
        '////////////////////////////////////////////////////////////////////////////////
        '//
        '// снятие выбора всех товаров в списке
        '//
        '////////////////////////////////////////////////////////////////////////////////

        Windows.Forms.Cursor.Current = Cursors.WaitCursor
        For i As Integer = 0 To DataGridView1.Rows.Count - 1
            DataGridView1.Item(0, i).Value = 0
        Next
        Windows.Forms.Cursor.Current = Cursors.Default
    End Sub

    Private Function CheckFilling() As Boolean
        '////////////////////////////////////////////////////////////////////////////////
        '//
        '// Проверка заполнения данных для создания задач
        '//
        '////////////////////////////////////////////////////////////////////////////////
        Dim MyRez As Double
        Dim MyCount As Integer

        If Trim(TextBox1.Text) = "" Then
            MsgBox("Необходимо заполнить, на сколько дней сдвигается поставка.", MsgBoxStyle.Critical, "Внимание!")
            TextBox1.Select()
            CheckFilling = False
            Exit Function
        End If

        If Trim(TextBox2.Text) = "" Then
            MsgBox("Необходимо заполнить причину, по которой сдвигается поставка.", MsgBoxStyle.Critical, "Внимание!")
            TextBox2.Select()
            CheckFilling = False
            Exit Function
        End If

        Try
            MyRez = CDbl(TextBox1.Text)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Внимание!")
            TextBox1.Select()
            CheckFilling = False
            Exit Function
        End Try

        If MyRez < 0 Then
            MsgBox("Количество дней, на сколько сдвигается поставка, должно быть больше нуля.", MsgBoxStyle.Critical, "Внимание!")
            TextBox1.Select()
            CheckFilling = False
            Exit Function
        End If

        If DateTimePicker1.Value <= Today() Then
            MsgBox("Максимальная ожидаемая дата поставки должна быть больше текущей даты.", MsgBoxStyle.Critical, "Внимание!")
            DateTimePicker1.Select()
            CheckFilling = False
            Exit Function
        End If


        MyCount = 0
        For i As Integer = 0 To DataGridView1.Rows.Count - 1
            If DataGridView1.Item(0, i).Value = -1 Then
                MyCount = MyCount + 1
            End If
        Next
        If MyCount = 0 Then
            MsgBox("Вы не выбрали ни одного товара, по которому изменится дата поставки.", MsgBoxStyle.Critical, "Внимание!")
            TextBox1.Select()
            CheckFilling = False
            Exit Function
        End If

        CheckFilling = True
    End Function

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        '////////////////////////////////////////////////////////////////////////////////
        '//
        '// Создание запросов на согласование переноса сроков поставки товаров
        '//
        '////////////////////////////////////////////////////////////////////////////////
        Dim MySQLStr As String
        Dim MyItemsList As String
        Dim MyOrderArray As New List(Of String)
        Dim MyFirstFlag As Integer

        If CheckFilling() = True Then
            MyItemsList = ""
            '---список продуктов, по которым сдвигаются сроки
            MyFirstFlag = 0
            For i As Integer = 0 To DataGridView1.Rows.Count - 1
                If DataGridView1.Item(0, i).Value = True Then
                    If MyFirstFlag = 0 Then
                        MyItemsList = "'" & DataGridView1.Item(1, i).Value & "'"
                        MyFirstFlag = 1
                    Else
                        MyItemsList = MyItemsList & ",'" & DataGridView1.Item(1, i).Value & "'"
                    End If
                End If
            Next

            '---список заказов на продажу для выбранного заказа на закупку с выбранными товарами
            MySQLStr = "SELECT DISTINCT PC010300.PC01060 "
            MySQLStr = MySQLStr & "FROM PC010300 INNER JOIN "
            MySQLStr = MySQLStr & "OR010300 ON PC010300.PC01060 = OR010300.OR01001 INNER JOIN "
            MySQLStr = MySQLStr & "OR030300 ON OR010300.OR01001 = OR030300.OR03001 "
            MySQLStr = MySQLStr & "WHERE (PC010300.PC01052 = N'" & Trim(Me.CommonPurchOrderNum) & "') "
            MySQLStr = MySQLStr & "AND (OR030300.OR03005 IN (" & MyItemsList & ")) "
            InitMyConn(False)
            InitMyRec(False, MySQLStr)
            If Declarations.MyRec.EOF = True And Declarations.MyRec.BOF = True Then
                trycloseMyRec()
                MsgBox("Не найдено ни одного неотгруженного заказа на продажу с выбранными товарами, относящегося к выбранному обобщенному заказу на закупку. ", MsgBoxStyle.OkOnly, "Внимание!")
            Else
                Declarations.MyRec.MoveFirst()
                While Declarations.MyRec.EOF = False
                    MyOrderArray.Add(Declarations.MyRec.Fields("PC01060").Value)
                    Declarations.MyRec.MoveNext()
                End While
                trycloseMyRec()

                For Each SalesOrder As String In MyOrderArray
                    'MsgBox(SalesOrder, MsgBoxStyle.OkOnly, "Заказ")
                    MySQLStr = "exec spp_PurchaseWorkplace_CreateConfirmRequest N'" & Trim(Me.CommonPurchOrderNum) & "', N'" & SalesOrder & "', N'" & Replace(MyItemsList, "'", "''") & "'"
                    InitMyConn(False)
                    InitMyRec(False, MySQLStr)
                    If Declarations.MyRec.EOF = True And Declarations.MyRec.BOF = True Then
                        trycloseMyRec()
                    Else
                        CreateConfirmRequest("ESKRU\" & Declarations.UserName, Trim(Me.CommonPurchOrderNum), Declarations.MyRec.Fields("Supplier").Value, _
                        Declarations.MyRec.Fields("SalesOrderNum").Value, Declarations.MyRec.Fields("Client").Value, Declarations.MyRec.Fields("Salesman").Value, _
                        Declarations.MyRec.Fields("PurchOrderNum").Value, Declarations.MyRec.Fields("Items").Value, Trim(TextBox1.Text), _
                        Format(DateTimePicker1.Value, "dd/MM/yyyy"), TextBox2.Text)
                        trycloseMyRec()
                    End If
                Next
                MsgBox("Процедура формирования заявок на согласование сроков переноса поставок завершена. ", MsgBoxStyle.OkOnly, "Внимание!")
                Me.Close()
            End If
        End If
    End Sub

    Public Function CreateConfirmRequest(ByVal MyOwner As String, ByVal MyCommonPurchOrderNum As String, ByVal MySupplier As String, ByVal MySalesOrderNum As String, _
    ByVal MyClient As String, ByVal MySalesman As String, ByVal MyPurchOrderNum As String, ByVal MyItems As String, ByVal MyDaysNumber As String, _
    ByVal MyNewData As String, ByVal MyReason As String) As String
        '////////////////////////////////////////////////////////////////////////////////
        '//
        '// Функция создания запроса на согласование переноса сроков поставки на шарепойнте
        '//
        '////////////////////////////////////////////////////////////////////////////////

        Dim listWebService As spbprd4.Lists = New spbprd4.Lists()
        listWebService.Credentials = New System.Net.NetworkCredential("developer", "!Devpass", "ESKRU")
        Dim listName = "{a737e822-163e-48c2-870d-eb45b2164e15}"
        Dim listView = ""
        Dim strBatch As String = "<Method ID='1' Cmd='New'>"

        strBatch = strBatch + "<Field Name='ID'>New</Field>"
        'strBatch = strBatch + "<Field Name='Author'>" & MyOwner & "</Field>"                                    '---владелец
        strBatch = strBatch + "<Field Name='_x0417__x0430__x043a__x0443__x04'>" & MyOwner & "</Field>"          '---Закупщик
        'strBatch = strBatch + "<Field Name='_x0417__x0430__x043a__x0443__x04'>ESKRU\Novozhilov</Field>"          '---Закупщик
        strBatch = strBatch + "<Field Name='Title'>" & MyCommonPurchOrderNum & "</Field>"                       '---N обобщенного заказа на закупку
        strBatch = strBatch + "<Field Name='_x041f__x043e__x0441__x0442__x04'>" & MySupplier & "</Field>"       '---Поставщик
        strBatch = strBatch + "<Field Name='N_x0020__x0437__x0430__x043a__x0'>" & MySalesOrderNum & "</Field>"  '---N заказа на продажу
        strBatch = strBatch + "<Field Name='_x041f__x043e__x043a__x0443__x04'>" & MyClient & "</Field>"         '---покупатель
        strBatch = strBatch + "<Field Name='_x041f__x0440__x043e__x0434__x04'>" & MySalesman & "</Field>"       '---продавец
        'strBatch = strBatch + "<Field Name='_x041f__x0440__x043e__x0434__x04'>ESKRU\Novozhilov</Field>"       '---продавец
        strBatch = strBatch + "<Field Name='N_x0020__x0437__x0430__x043a__x00'>" & MyPurchOrderNum & "</Field>" '---N заказа на закупку
        strBatch = strBatch + "<Field Name='_x0422__x043e__x0432__x0430__x04'>" & MyItems & "</Field>"          '---товары
        strBatch = strBatch + "<Field Name='_x041d__x0430__x0020__x0441__x04'>" & MyDaysNumber & "</Field>"     '---На сколько дней сдвигается
        strBatch = strBatch + "<Field Name='_x041c__x0430__x043a__x0441__x00'>" & MyNewData & "</Field>"        '---Максимальная новая дата поставки
        strBatch = strBatch + "<Field Name='_x041f__x0440__x0438__x0447__x04'>" & MyReason & "</Field>"         '---Причина изменения
        strBatch = strBatch + "</Method>"

        Dim xmlDoc As XmlDocument = New System.Xml.XmlDocument()
        Dim elBatch As System.Xml.XmlElement = xmlDoc.CreateElement("Batch")
        elBatch.SetAttribute("OnError", "Continue")
        elBatch.SetAttribute("ListVersion", "1")
        elBatch.SetAttribute("ViewName", listView)
        elBatch.InnerXml = strBatch

        Dim ndReturn As XmlNode = listWebService.UpdateListItems(listName, elBatch)

    End Function
End Class