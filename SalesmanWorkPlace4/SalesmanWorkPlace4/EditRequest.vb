Public Class EditRequest
    Public StartParam As String
    Public WindowFrom As String

    Private Sub Button5_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button5.Click
        '////////////////////////////////////////////////////////////////////////////////
        '//
        '// Закрытие окна без сохранения
        '//
        '////////////////////////////////////////////////////////////////////////////////

        Me.Close()
    End Sub

    Private Sub Button3_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button3.Click
        '////////////////////////////////////////////////////////////////////////////////
        '//
        '// открытие окна поиска покупателя в Скале
        '//
        '////////////////////////////////////////////////////////////////////////////////

        MyCustomerSelect = New CustomerSelect
        MyCustomerSelect.StartParam = "Search"
        MyCustomerSelect.ShowDialog()
    End Sub

    Private Sub Button4_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button4.Click
        '////////////////////////////////////////////////////////////////////////////////
        '//
        '// Закрытие окна с сохранением
        '//
        '////////////////////////////////////////////////////////////////////////////////

        If CheckFormFilling() = True Then
            '----сохранение результатов
            If SaveRequest() = True Then
                Me.Close()
            End If
        End If
    End Sub

    Private Function CheckFormFilling() As Boolean
        '////////////////////////////////////////////////////////////////////////////////
        '//
        '// Проверка заполнения полей формы
        '//
        '////////////////////////////////////////////////////////////////////////////////
        Dim MySQLStr As String

        If Trim(TextBox2.Text) = "" Then
            MsgBox("Поле ""Покупатель"" должно быть заполнено", MsgBoxStyle.Critical, "Внимание!")
            TextBox2.Select()
            CheckFormFilling = False
            Exit Function
        End If

        '-----дата больше или равна текущей
        If DateTimePicker1.Value < Now().AddDays(-1) Then
            MsgBox("Дата ""Срок представления КП"" должна быть больше или равна текущей", MsgBoxStyle.Critical, "Внимание!")
            DateTimePicker1.Select()
            CheckFormFilling = False
            Exit Function
        End If

        '-----после 15:00 дата только больше текущей
        If Hour(Now()) >= 15 Then
            If DateTimePicker1.Value < Now() Then
                MsgBox("После 15:00 Дата ""Срок представления КП"" должна быть больше текущей", MsgBoxStyle.Critical, "Внимание!")
                DateTimePicker1.Select()
                CheckFormFilling = False
                Exit Function
            End If
        End If

        '-----нельзя выбирать субботу и воскресенье
        If Weekday(DateTimePicker1.Value, 2) = 6 Or Weekday(DateTimePicker1.Value, 2) = 7 Then
            MsgBox("Дата ""Срок представления КП"" не должна быть субботой или воскресеньем.", MsgBoxStyle.Critical, "Внимание!")
            DateTimePicker1.Select()
            CheckFormFilling = False
            Exit Function
        End If

        '-----проверка правильности ввода номера КП
        If Trim(TextBox8.Text) <> "" Then
            MySQLStr = "SELECT COUNT(OR01001) AS CC "
            MySQLStr = MySQLStr & "FROM tbl_OR010300 "
            MySQLStr = MySQLStr & "WHERE (OR01001 = N'" & Strings.Right("0000000000" & Trim(TextBox8.Text), 10) & "') "
            InitMyConn(False)
            InitMyRec(False, MySQLStr)
            If Declarations.MyRec.EOF = True And Declarations.MyRec.BOF = True Then
                MsgBox("Невозможно проверить наличие КП с таким номером. Обратитесь к администратору.", MsgBoxStyle.Critical, "Внимание!")
                TextBox8.Select()
                CheckFormFilling = False
                trycloseMyRec()
                Exit Function
            Else
                If Declarations.MyRec.Fields("CC").Value = 0 Then
                    MsgBox("Коммерческого предложения с таким номером не существует. Введите корректноый номер КП.", MsgBoxStyle.Critical, "Внимание!")
                    TextBox8.Select()
                    CheckFormFilling = False
                    trycloseMyRec()
                    Exit Function
                Else
                    trycloseMyRec()
                End If
            End If
        End If


        If (WindowFrom = "OrderLines") Then
            '-----Если создаем из КП - сначала проверяем, что есть товары без цен
            MySQLStr = "SELECT COUNT(*) AS CC "
            MySQLStr = MySQLStr & "FROM tbl_OR030300 LEFT OUTER JOIN "
            MySQLStr = MySQLStr & "SC010300 ON tbl_OR030300.OR03005 = SC010300.SC01001 "
            MySQLStr = MySQLStr & "WHERE (tbl_OR030300.OR03001 = N'" & Trim(MyOrderLines.Label6.Text) & "') AND (SC010300.SC01001 IS NULL) OR "
            MySQLStr = MySQLStr & "(tbl_OR030300.OR03001 = N'" & Trim(MyOrderLines.Label6.Text) & "') AND (SC010300.SC01055 = 0) "
            InitMyConn(False)
            InitMyRec(False, MySQLStr)
            If Declarations.MyRec.EOF = True And Declarations.MyRec.BOF = True Then
                MsgBox("Невозможно проверить количество доступных для поиска строк. Обратитесь к администратору.", MsgBoxStyle.Critical, "Внимание!")
                TextBox1.Select()
                CheckFormFilling = False
                trycloseMyRec()
                Exit Function
            Else
                If Declarations.MyRec.Fields("CC").Value = 0 Then
                    MsgBox("В вашем коммерческом предложении нет ни одной строки без цены - нечего отправлять на поиск.", MsgBoxStyle.Critical, "Внимание!")
                    TextBox1.Select()
                    CheckFormFilling = False
                    trycloseMyRec()
                    Exit Function
                Else
                    trycloseMyRec()
                End If
            End If

            ''------Проверка уникальности кода товара поставщика в списке товаров без цен
            'MySQLStr = "SELECT COUNT(*) AS CC "
            'MySQLStr = MySQLStr & "FROM (SELECT tbl_OR030300.SuppItemCode "
            'MySQLStr = MySQLStr & "FROM tbl_OR030300 LEFT OUTER JOIN "
            'MySQLStr = MySQLStr & "SC010300 ON tbl_OR030300.OR03005 = SC010300.SC01001 "
            'MySQLStr = MySQLStr & "WHERE ((tbl_OR030300.OR03001 = N'" & Trim(MyOrderLines.Label6.Text) & "') AND (SC010300.SC01001 IS NULL) OR "
            'MySQLStr = MySQLStr & "(tbl_OR030300.OR03001 = N'" & Trim(MyOrderLines.Label6.Text) & "') AND (SC010300.SC01055 = 0)) "
            'MySQLStr = MySQLStr & "AND (ISNULL(tbl_OR030300.SuppItemCode, '') <> '') "
            'MySQLStr = MySQLStr & "GROUP BY tbl_OR030300.SuppItemCode "
            'MySQLStr = MySQLStr & "HAVING (COUNT(*) > 1)) AS View_1 "
            'InitMyConn(False)
            'InitMyRec(False, MySQLStr)
            'If Declarations.MyRec.EOF = True And Declarations.MyRec.BOF = True Then
            '    MsgBox("Невозможно проверить уникальность кода товара поставщика. Обратитесь к администратору.", MsgBoxStyle.Critical, "Внимание!")
            '    TextBox1.Select()
            '    CheckFormFilling = False
            '    trycloseMyRec()
            '    Exit Function
            'Else
            '    If Declarations.MyRec.Fields("CC").Value = 0 Then
            '        trycloseMyRec()
            '    Else
            '        MsgBox("В вашем коммерческом предложении есть товары без цены с одинаковым кодом товара поставщика. Такие товары невозможно отправить в рабочее место поисковика.", MsgBoxStyle.Critical, "Внимание!")
            '        TextBox1.Select()
            '        CheckFormFilling = False
            '        trycloseMyRec()
            '        Exit Function
            '    End If
            'End If

            ''------Проверка уникальности названия товара в списке товаров без цен
            'MySQLStr = "SELECT COUNT(*) AS CC "
            'MySQLStr = MySQLStr & "FROM (SELECT Ltrim(Rtrim(tbl_OR030300.OR03006 + tbl_OR030300.OR03007)) AS ItemName "
            'MySQLStr = MySQLStr & "FROM tbl_OR030300 LEFT OUTER JOIN "
            'MySQLStr = MySQLStr & "SC010300 ON tbl_OR030300.OR03005 = SC010300.SC01001 "
            'MySQLStr = MySQLStr & "WHERE ((tbl_OR030300.OR03001 = N'" & Trim(MyOrderLines.Label6.Text) & "') AND (SC010300.SC01001 IS NULL) OR "
            'MySQLStr = MySQLStr & "(tbl_OR030300.OR03001 = N'" & Trim(MyOrderLines.Label6.Text) & "') AND (SC010300.SC01055 = 0)) "
            'MySQLStr = MySQLStr & "and (Ltrim(Rtrim(tbl_OR030300.OR03006 + tbl_OR030300.OR03007)) <> '') "
            'MySQLStr = MySQLStr & "GROUP BY Ltrim(Rtrim(tbl_OR030300.OR03006 + tbl_OR030300.OR03007)) "
            'MySQLStr = MySQLStr & "HAVING (COUNT(*) > 1)) AS View_1 "
            'InitMyConn(False)
            'InitMyRec(False, MySQLStr)
            'If Declarations.MyRec.EOF = True And Declarations.MyRec.BOF = True Then
            '    MsgBox("Невозможно проверить уникальность названия товара. Обратитесь к администратору.", MsgBoxStyle.Critical, "Внимание!")
            '    TextBox1.Select()
            '    CheckFormFilling = False
            '    trycloseMyRec()
            '    Exit Function
            'Else
            '    If Declarations.MyRec.Fields("CC").Value = 0 Then
            '        trycloseMyRec()
            '    Else
            '        MsgBox("В вашем коммерческом предложении есть товары без цены с одинаковым названием. Такие товары невозможно отправить в рабочее место поисковика.", MsgBoxStyle.Critical, "Внимание!")
            '        TextBox1.Select()
            '        CheckFormFilling = False
            '        trycloseMyRec()
            '        Exit Function
            '    End If
            'End If

            '------Проверка уникальности кода товара поставщика + названия товара в запросе
            MySQLStr = "SELECT COUNT(*) AS CC "
            MySQLStr = MySQLStr & "FROM (SELECT LTRIM(RTRIM(tbl_OR030300.SuppItemCode)) + LTRIM(RTRIM(tbl_OR030300.OR03006 + tbl_OR030300.OR03007)) AS ItemName "
            MySQLStr = MySQLStr & "FROM tbl_OR030300 LEFT OUTER JOIN "
            MySQLStr = MySQLStr & "SC010300 ON tbl_OR030300.OR03005 = SC010300.SC01001 "
            MySQLStr = MySQLStr & "WHERE (tbl_OR030300.OR03001 = N'" & Trim(MyOrderLines.Label6.Text) & "') AND (SC010300.SC01001 IS NULL) AND "
            MySQLStr = MySQLStr & "(LTRIM(RTRIM(tbl_OR030300.OR03006 + tbl_OR030300.OR03007)) <> '') OR "
            MySQLStr = MySQLStr & "(tbl_OR030300.OR03001 = N'" & Trim(MyOrderLines.Label6.Text) & "') AND (LTRIM(RTRIM(tbl_OR030300.OR03006 + tbl_OR030300.OR03007)) <> '') AND (SC010300.SC01055 = 0) "
            MySQLStr = MySQLStr & "GROUP BY LTRIM(RTRIM(tbl_OR030300.SuppItemCode)) + LTRIM(RTRIM(tbl_OR030300.OR03006 + tbl_OR030300.OR03007)) "
            MySQLStr = MySQLStr & "HAVING (COUNT(*) > 1)) AS View_1 "
            InitMyConn(False)
            InitMyRec(False, MySQLStr)
            If Declarations.MyRec.BOF = True And Declarations.MyRec.EOF = True Then
                trycloseMyRec()
                MsgBox("Невозможно проверить уникальность кода поставщика + названия товара. Обратитесь к администратору.", MsgBoxStyle.Critical, "Внимание!")
                TextBox1.Select()
                CheckFormFilling = False
                trycloseMyRec()
                Exit Function
            Else
                If Declarations.MyRec.Fields("CC").Value = 0 Then
                    trycloseMyRec()
                Else
                    MsgBox("В вашем коммерческом предложении есть товары без цены с одинаковым кодом товара поставщика + название или одинаковым кодом товара в Scala. Такие товары невозможно отправить в рабочее место поисковика.", MsgBoxStyle.Critical, "Внимание!")
                    TextBox1.Select()
                    CheckFormFilling = False
                    trycloseMyRec()
                    Exit Function
                End If
            End If


            '------Проверка, что в списке товаров без цен обязательно заполнен или код товара поставщика, или название
            MySQLStr = "SELECT COUNT(*) AS CC "
            MySQLStr = MySQLStr & "FROM tbl_OR030300 LEFT OUTER JOIN "
            MySQLStr = MySQLStr & "SC010300 ON tbl_OR030300.OR03005 = SC010300.SC01001 "
            MySQLStr = MySQLStr & "WHERE ((tbl_OR030300.OR03001 = N'" & Trim(MyOrderLines.Label6.Text) & "') AND (SC010300.SC01001 IS NULL) OR "
            MySQLStr = MySQLStr & "(tbl_OR030300.OR03001 = N'" & Trim(MyOrderLines.Label6.Text) & "') AND (SC010300.SC01055 = 0)) "
            MySQLStr = MySQLStr & "AND ((tbl_OR030300.OR03006 + tbl_OR030300.OR03007 = '') AND (tbl_OR030300.SuppItemCode = '')) "
            InitMyConn(False)
            InitMyRec(False, MySQLStr)
            If Declarations.MyRec.EOF = True And Declarations.MyRec.BOF = True Then
                MsgBox("Невозможно проверить обязательность наличия или кода товара поставщика или названия товара. Обратитесь к администратору.", MsgBoxStyle.Critical, "Внимание!")
                TextBox1.Select()
                CheckFormFilling = False
                trycloseMyRec()
                Exit Function
            Else
                If Declarations.MyRec.Fields("CC").Value = 0 Then
                    trycloseMyRec()
                Else
                    MsgBox("В вашем коммерческом предложении есть товары без кода товара поставщика и названия. Такие товары невозможно отправить в рабочее место поисковика.", MsgBoxStyle.Critical, "Внимание!")
                    TextBox1.Select()
                    CheckFormFilling = False
                    trycloseMyRec()
                    Exit Function
                End If
            End If

        End If

        '-----проверка занесения условий оплаты
        If Trim(TextBox9.Text) = "" Then
            MsgBox("Поле ""Условия оплаты"" должно быть заполнено", MsgBoxStyle.Critical, "Внимание!")
            TextBox9.Select()
            CheckFormFilling = False
            Exit Function
        End If

        CheckFormFilling = True
    End Function

    Private Sub EditRequest_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles Me.KeyDown
        '////////////////////////////////////////////////////////////////////////////////
        '//
        '// Запрет закрытия окна по ALT+F4
        '//
        '////////////////////////////////////////////////////////////////////////////////

        If e.KeyData = Keys.Alt + Keys.F4 Then
            e.Handled = True
        End If
    End Sub

    Private Sub EditRequest_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        '////////////////////////////////////////////////////////////////////////////////
        '//
        '// Загрузка окна
        '//
        '////////////////////////////////////////////////////////////////////////////////
        Dim MySQLStr As String                        'рабочая строка

        If StartParam = "Create" Then
            If WindowFrom = "OrderLines" Then
                MySQLStr = "SELECT OR01003, CName, OR01001 "
                MySQLStr = MySQLStr & "FROM tbl_OR010300 "
                MySQLStr = MySQLStr & "WHERE (OR01001 = N'" & Trim(MyOrderLines.Label6.Text) & "')"
                InitMyConn(False)
                InitMyRec(False, MySQLStr)
                If Declarations.MyRec.EOF = True And Declarations.MyRec.BOF = True Then
                Else
                    TextBox1.Text = Declarations.MyRec.Fields("OR01003").Value.ToString
                    TextBox2.Text = Declarations.MyRec.Fields("CName").Value.ToString
                    TextBox8.Text = Declarations.MyRec.Fields("OR01001").Value.ToString
                End If
                trycloseMyRec()
                MySQLStr = "SELECT SL230300.SL23004 AS PT "
                MySQLStr = MySQLStr & "FROM SL010300 INNER JOIN "
                MySQLStr = MySQLStr & "SL230300 ON SL010300.SL01024 = SL230300.SL23003 "
                MySQLStr = MySQLStr & "WHERE (SL230300.SL23002 = N'RUS') "
                MySQLStr = MySQLStr & "AND (SL230300.SL23001 = N'0') "
                MySQLStr = MySQLStr & "AND (SL010300.SL01001 = N'" & Trim(TextBox1.Text) & "') "
                InitMyConn(False)
                InitMyRec(False, MySQLStr)
                If Declarations.MyRec.EOF = True And Declarations.MyRec.BOF = True Then
                    'TextBox9.Text = "Предоплата 100%"
                    TextBox9.Text = ""
                    TextBox9.Enabled = True
                Else
                    TextBox9.Text = Declarations.MyRec.Fields("PT").Value.ToString
                    TextBox9.Enabled = False
                End If
                trycloseMyRec()
            End If
            Declarations.MyRequestNum = 0
            Label3.Text = "New"

        Else
            Declarations.MyRequestNum = MySearchSupplier.DataGridView1.SelectedRows.Item(0).Cells(0).Value
            Label3.Text = Declarations.MyRequestNum.ToString

            MySQLStr = "SELECT ID, CustomerID, CustomerName, CustomerContactName, "
            MySQLStr = MySQLStr & "CustomerPhone, CustomerEmail, RequestDate, Comments, CustomerRequestNum, "
            MySQLStr = MySQLStr & "ISNULL(tbl_SupplSearch.CPNum, '') AS CPNum, "
            MySQLStr = MySQLStr & "ISNULL(tbl_SupplSearch.PaymentTerms, '') AS PaymentTerms "
            MySQLStr = MySQLStr & "FROM tbl_SupplSearch "
            MySQLStr = MySQLStr & "WHERE (ID = " & Declarations.MyRequestNum.ToString & ") "
            InitMyConn(False)
            InitMyRec(False, MySQLStr)
            If Declarations.MyRec.EOF = True And Declarations.MyRec.BOF = True Then
            Else
                TextBox1.Text = Declarations.MyRec.Fields("CustomerID").Value.ToString
                TextBox2.Text = Declarations.MyRec.Fields("CustomerName").Value.ToString
                TextBox3.Text = Declarations.MyRec.Fields("CustomerContactName").Value.ToString
                TextBox4.Text = Declarations.MyRec.Fields("CustomerPhone").Value.ToString
                TextBox5.Text = Declarations.MyRec.Fields("CustomerEmail").Value.ToString
                DateTimePicker1.Value = Declarations.MyRec.Fields("RequestDate").Value
                TextBox6.Text = Declarations.MyRec.Fields("Comments").Value.ToString
                TextBox7.Text = Declarations.MyRec.Fields("CustomerRequestNum").Value.ToString
                TextBox8.Text = Declarations.MyRec.Fields("CPNum").Value.ToString
                TextBox9.Text = Declarations.MyRec.Fields("PaymentTerms").Value.ToString
            End If
            trycloseMyRec()
            MySQLStr = "SELECT SL230300.SL23004 AS PT "
            MySQLStr = MySQLStr & "FROM SL010300 INNER JOIN "
            MySQLStr = MySQLStr & "SL230300 ON SL010300.SL01024 = SL230300.SL23003 "
            MySQLStr = MySQLStr & "WHERE (SL230300.SL23002 = N'RUS') "
            MySQLStr = MySQLStr & "AND (SL230300.SL23001 = N'0') "
            MySQLStr = MySQLStr & "AND (SL010300.SL01001 = N'" & Trim(TextBox1.Text) & "') "
            InitMyConn(False)
            InitMyRec(False, MySQLStr)
            If Declarations.MyRec.EOF = True And Declarations.MyRec.BOF = True Then
                'TextBox9.Text = ""
                'TextBox9.Enabled = True
            Else
                TextBox9.Text = Declarations.MyRec.Fields("PT").Value.ToString
                TextBox9.Enabled = False
            End If
            trycloseMyRec()
        End If
    End Sub

    Private Function SaveRequest() As Boolean
        '////////////////////////////////////////////////////////////////////////////////
        '//
        '// Сохранение данных введенных в форму
        '//
        '////////////////////////////////////////////////////////////////////////////////
        Dim MySQLStr As String                        'рабочая строка
        Dim MyRez As MsgBoxResult

        If StartParam = "Create" Then       '-----создание новой записи
            Try
                Declarations.MyRequestNum = 0
                MySQLStr = "exec spp_SupplSearch_SearchRequestCreate "
                MySQLStr = MySQLStr + "N'" + Trim(Declarations.SalesmanCode) + "', "
                MySQLStr = MySQLStr + "N'" + Trim(Declarations.SalesmanName) + "', "
                MySQLStr = MySQLStr + "N'" + Replace(Trim(TextBox1.Text), "'", "''") + "', "
                MySQLStr = MySQLStr + "N'" + Replace(Trim(TextBox2.Text), "'", "''") + "', "
                MySQLStr = MySQLStr + "N'" + Replace(Trim(TextBox3.Text), "'", "''") + "', "
                MySQLStr = MySQLStr + "N'" + Replace(Trim(TextBox4.Text), "'", "''") + "', "
                MySQLStr = MySQLStr + "N'" + Replace(Trim(TextBox5.Text), "'", "''") + "', "
                MySQLStr = MySQLStr + "N'" + Format(DateTimePicker1.Value, "dd/MM/yyyy") + "', "
                MySQLStr = MySQLStr + "N'--" + Format(Now, "dd/MM/yyyy HH:mm") + "-->" + Replace(Trim(TextBox6.Text), " '", "''") + "', "
                MySQLStr = MySQLStr + "N'" + Replace(Trim(TextBox7.Text), "'", "''") + "', "
                MySQLStr = MySQLStr + "N'" + Replace(Trim(TextBox8.Text), "'", "''") + "', "
                MySQLStr = MySQLStr + "N'" + Replace(Trim(TextBox9.Text), "'", "''") + "' "
                InitMyConn(False)
                InitMyRec(False, MySQLStr)
                If Declarations.MyRec.BOF = True And Declarations.MyRec.EOF = True Then
                    MsgBox("Ошибка создания записи")
                    SaveRequest = False
                    trycloseMyRec()
                    Exit Function
                Else
                    StartParam = "Edit"
                    Declarations.MyRequestNum = Declarations.MyRec.Fields("MyNewID").Value
                    Label3.Text = Declarations.MyRequestNum
                    trycloseMyRec()
                End If
                If (WindowFrom = "OrderLines") And (Declarations.MyRequestNum <> 0) Then
                    MySQLStr = "exec spp_SupplSearch_SearchRequestCreateFCP " & CStr(Declarations.MyRequestNum) & ", N'" & Trim(MyOrderLines.Label6.Text) & "'"
                    InitMyConn(False)
                    Declarations.MyConn.Execute(MySQLStr)
                    MySQLStr = "SELECT COUNT(*) AS CC "
                    MySQLStr = MySQLStr & "FROM tbl_SupplSearchItems "
                    MySQLStr = MySQLStr & "WHERE (SupplSearchID = " & CStr(Declarations.MyRequestNum) & ") "
                    InitMyConn(False)
                    InitMyRec(False, MySQLStr)
                    If Declarations.MyRec.BOF = True And Declarations.MyRec.EOF = True Then
                        trycloseMyRec()
                    Else
                        If Declarations.MyRec.Fields("CC").Value = 0 Then
                            trycloseMyRec()
                        Else
                            trycloseMyRec()
                            MyRez = MsgBox("Отправить запрос поисковикам без дополнительного редактирования строк?", MsgBoxStyle.YesNo, "Внимание!")
                            If (MyRez = MsgBoxResult.Yes) Then
                                '------отправка в работу и отправка почты об этом
                                MySQLStr = "UPDATE tbl_SupplSearch "
                                MySQLStr = MySQLStr + "SET SalesStatus = 0 "
                                MySQLStr = MySQLStr + "WHERE (ID = " & CStr(Declarations.MyRequestNum) & ") "
                                InitMyConn(False)
                                Declarations.MyConn.Execute(MySQLStr)

                                Dim EmailStr As String
                                EmailStr = GetSrchManagerEmailFromDB()
                                If EmailStr = "" Then
                                    MsgBox("Для руководителя поисковиков в БД не занесена почта или он не указан в таблице tbl_SupplSearch_Searchers. Обратитесь к администратору.", MsgBoxStyle.Critical, "Внимание!")
                                Else
                                    SendInfoByEmail(CStr(Declarations.MyRequestNum), Format(Now(), "dd/MM/yyyy"), _
                                       EmailStr, Trim(TextBox1.Text) + " " + Trim(TextBox2.Text), Trim(Declarations.SalesmanCode) + " " + Trim(Declarations.SalesmanName), _
                                       "Отправлен запрос")
                                End If
                            End If
                        End If
                    End If

                    MsgBox("Запрос на поиск товаров создан.", MsgBoxStyle.Information, "Внимание!")
                End If
            Catch ex As Exception
                MsgBox(ex.ToString)
                SaveRequest = False
                Exit Function
            End Try
        Else                                '-----редактирование существующей
            Try
                MySQLStr = "UPDATE tbl_SupplSearch "
                MySQLStr = MySQLStr & "SET CustomerID = N'" & Replace(Trim(TextBox1.Text), "'", "''") & "', "
                MySQLStr = MySQLStr & "CustomerName = N'" & Replace(Trim(TextBox2.Text), "'", "''") & "', "
                MySQLStr = MySQLStr & "CustomerContactName = N'" & Replace(Trim(TextBox3.Text), "'", "''") & "', "
                MySQLStr = MySQLStr & "CustomerPhone = N'" & Replace(Trim(TextBox4.Text), "'", "''") & "', "
                MySQLStr = MySQLStr & "CustomerEmail = N'" & Replace(Trim(TextBox5.Text), "'", "''") & "', "
                MySQLStr = MySQLStr & "RequestDate = CONVERT(DATETIME, '" & Format(DateTimePicker1.Value, "dd/MM/yyyy") & "', 103), "
                MySQLStr = MySQLStr & "Comments = N'" & Replace(Trim(TextBox6.Text), "'", "''") & "', "
                MySQLStr = MySQLStr & "CustomerRequestNum = N'" & Replace(Trim(TextBox7.Text), "'", "''") & "', "
                MySQLStr = MySQLStr & "CPNum = N'" & Replace(Trim(TextBox8.Text), "'", "''") & "', "
                MySQLStr = MySQLStr & "PaymentTerms = N'" & Replace(Trim(TextBox9.Text), "'", "''") & "' "
                MySQLStr = MySQLStr & "WHERE (ID = " & Declarations.MyRequestNum.ToString & ") "
                InitMyConn(False)
                Declarations.MyConn.Execute(MySQLStr)
            Catch ex As Exception
                MsgBox(ex.ToString)
                SaveRequest = False
                Exit Function
            End Try
        End If

        SaveRequest = True
    End Function

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        '////////////////////////////////////////////////////////////////////////////////
        '//
        '// открытие окна с аттачментами продавца
        '//
        '////////////////////////////////////////////////////////////////////////////////

        If CheckFormFilling() = True Then
            '----сохранение результатов
            If SaveRequest() = True Then
                MyAttachmentsList = New AttachmentsList
                MyAttachmentsList.AttType = "Sales"
                MyAttachmentsList.WhoStart = "Sales"
                MyAttachmentsList.MyPlace = "Dialog"
                MyAttachmentsList.ShowDialog()
            End If
        End If
    End Sub

    Private Sub Button2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button2.Click
        '////////////////////////////////////////////////////////////////////////////////
        '//
        '// открытие окна с аттачментами поисковика
        '//
        '////////////////////////////////////////////////////////////////////////////////

        If CheckFormFilling() = True Then
            '----сохранение результатов
            If SaveRequest() = True Then
                MyAttachmentsList = New AttachmentsList
                MyAttachmentsList.AttType = "Search"
                MyAttachmentsList.WhoStart = "Sales"
                MyAttachmentsList.MyPlace = "Dialog"
                MyAttachmentsList.ShowDialog()
            End If
        End If
    End Sub

    Private Sub Button6_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button6.Click
        '////////////////////////////////////////////////////////////////////////////////
        '//
        '// открытие окна с контактной информацией из CRM
        '//
        '////////////////////////////////////////////////////////////////////////////////

        If Trim(TextBox1.Text) = "" Then
            MsgBox("Для выбора контактов из CRM необходимо ввести код клиента в CRM.", MsgBoxStyle.Critical, "Внимание!")
            TextBox1.Select()
        Else
            Dim MyContactInfo = New ContactInfo
            MyContactInfo.ShowDialog()
        End If
    End Sub

    Private Sub TextBox1_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles TextBox1.KeyDown
        '////////////////////////////////////////////////////////////////////////////////
        '//
        '// переход на следующее поле по Enter
        '//
        '////////////////////////////////////////////////////////////////////////////////

        If e.KeyData = Keys.Enter Then
            Me.SelectNextControl(sender, True, True, True, False)
        End If
    End Sub

    Private Sub TextBox1_Validated(ByVal sender As Object, ByVal e As System.EventArgs) Handles TextBox1.Validated
        '////////////////////////////////////////////////////////////////////////////////
        '//
        '// занесение названия клиента из Scala
        '//
        '////////////////////////////////////////////////////////////////////////////////
        Dim MySQLStr As String

        If Trim(TextBox1.Text) = "" Then
        Else
            MySQLStr = "SELECT SL01002 "
            MySQLStr = MySQLStr & "FROM SL010300 "
            MySQLStr = MySQLStr & "WHERE (SL01001 = N'" & Trim(TextBox1.Text) & "') "
            InitMyConn(False)
            InitMyRec(False, MySQLStr)
            If Declarations.MyRec.EOF = True And Declarations.MyRec.BOF = True Then
                MsgBox("Ошибка получения информации о названии клиента из Sala. Обратитесь к администратору.", MsgBoxStyle.Critical, "Внимание!")
                TextBox2.Text = ""
            Else
                TextBox2.Text = Trim(Declarations.MyRec.Fields("SL01002").Value.ToString)
            End If
        End If

        MySQLStr = "SELECT SL230300.SL23004 AS PT "
        MySQLStr = MySQLStr & "FROM SL010300 INNER JOIN "
        MySQLStr = MySQLStr & "SL230300 ON SL010300.SL01024 = SL230300.SL23003 "
        MySQLStr = MySQLStr & "WHERE (SL230300.SL23002 = N'RUS') "
        MySQLStr = MySQLStr & "AND (SL230300.SL23001 = N'0') "
        MySQLStr = MySQLStr & "AND (SL010300.SL01001 = N'" & Trim(TextBox1.Text) & "') "
        InitMyConn(False)
        InitMyRec(False, MySQLStr)
        If Declarations.MyRec.EOF = True And Declarations.MyRec.BOF = True Then
            'TextBox9.Text = "Предоплата 100%"
            TextBox9.Text = ""
            TextBox9.Enabled = True
        Else
            TextBox9.Text = Declarations.MyRec.Fields("PT").Value.ToString
            TextBox9.Enabled = False
        End If
        trycloseMyRec()
    End Sub

    Private Sub TextBox1_Validating(ByVal sender As Object, ByVal e As System.ComponentModel.CancelEventArgs) Handles TextBox1.Validating
        '////////////////////////////////////////////////////////////////////////////////
        '//
        '// Проверка, что такой код клиента присутствует в Scala
        '//
        '////////////////////////////////////////////////////////////////////////////////
        Dim MySQLStr As String

        If Trim(TextBox1.Text) = "" Then
        Else
            MySQLStr = "SELECT COUNT(SL01001) AS CC "
            MySQLStr = MySQLStr & "FROM SL010300 "
            MySQLStr = MySQLStr & "WHERE (SL01001 = N'" & Trim(TextBox1.Text) & "')"
            InitMyConn(False)
            InitMyRec(False, MySQLStr)
            If Declarations.MyRec.EOF = True And Declarations.MyRec.BOF = True Then
                MsgBox("Ошибка проверки присутствия кода в Sala. Обратитесь к администратору.", MsgBoxStyle.Critical, "Внимание!")
                e.Cancel = True
            Else
                If Declarations.MyRec.Fields("CC").Value = 0 Then
                    MsgBox("Код клиента" & Trim(TextBox1.Text) & " не найден в Sala. Введите корректный код.", MsgBoxStyle.Critical, "Внимание!")
                    e.Cancel = True
                Else

                End If
            End If
            trycloseMyRec()

            
        End If
        
    End Sub
End Class