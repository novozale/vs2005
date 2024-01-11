Public Class AddItem
    Public StartParam As String

    Private Sub Button5_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button5.Click
        '////////////////////////////////////////////////////////////////////////////////
        '//
        '// Закрытие окна без сохранения
        '//
        '////////////////////////////////////////////////////////////////////////////////

        Me.Close()
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
        Dim MyRez As Double
        Dim MySQLStr

        If Trim(TextBox1.Text) = "" Then
            'MsgBox("""Код Скала"" должен быть заполнен", MsgBoxStyle.Critical, "Внимание!")
            'TextBox1.Select()
            'CheckFormFilling = False
            'Exit Function
        End If

        '---Проверка наличия кода Скала
        If Trim(TextBox1.Text) <> "" Then
            MySQLStr = "SELECT COUNT(SC01001) AS CC "
            MySQLStr = MySQLStr & "FROM SC010300 "
            MySQLStr = MySQLStr & "WHERE (SC01001 = N'" & Trim(TextBox1.Text) & "') "
            InitMyConn(False)
            InitMyRec(False, MySQLStr)
            If Declarations.MyRec.EOF = True And Declarations.MyRec.BOF = True Then
                MsgBox("Невозможно проверить наличие кода в Scala. Обратитесь к администратору.", MsgBoxStyle.Critical, "Внимание!")
                trycloseMyRec()
                TextBox1.Select()
                CheckFormFilling = False
                Exit Function
            Else
                If Declarations.MyRec.Fields("CC").Value = 0 Then
                    MsgBox("Такого кода товара в Скала нет.", MsgBoxStyle.Critical, "Внимание!")
                    trycloseMyRec()
                    TextBox1.Select()
                    CheckFormFilling = False
                    Exit Function
                Else
                    trycloseMyRec()
                End If
            End If
        End If


        If Trim(TextBox4.Text) = "" Then
            MsgBox("""Количество"" должно быть заполнено.", MsgBoxStyle.Critical, "Внимание!")
            TextBox4.Select()
            CheckFormFilling = False
            Exit Function
        End If

        Try
            MyRez = CDbl(TextBox4.Text)
            If MyRez <= 0 Then
                MsgBox("""Количество"" должно быть больше 0.", MsgBoxStyle.Critical, "Внимание!")
                TextBox4.Select()
                CheckFormFilling = False
                Exit Function
            End If
        Catch ex As Exception
            MsgBox("""Количество"" должно быть числом.", MsgBoxStyle.Critical, "Внимание!")
            TextBox4.Select()
            CheckFormFilling = False
            Exit Function
        End Try


        If Trim(TextBox5.Text) = "" Then
            MsgBox("""Срок поставки (нед)"" должно быть заполнено.", MsgBoxStyle.Critical, "Внимание!")
            TextBox5.Select()
            CheckFormFilling = False
            Exit Function
        End If

        Try
            MyRez = CDbl(TextBox5.Text)
            If MyRez <= 0 Then
                MsgBox("""Срок поставки (нед)"" должно быть больше 0.", MsgBoxStyle.Critical, "Внимание!")
                TextBox5.Select()
                CheckFormFilling = False
                Exit Function
            End If
        Catch ex As Exception
            MsgBox("""Срок поставки (нед)"" должно быть числом.", MsgBoxStyle.Critical, "Внимание!")
            TextBox5.Select()
            CheckFormFilling = False
            Exit Function
        End Try

        If Trim(TextBox2.Text) = "" Then
            MsgBox("""Закуп цена без НДС"" должна быть заполнено.", MsgBoxStyle.Critical, "Внимание!")
            TextBox2.Select()
            CheckFormFilling = False
            Exit Function
        End If

        Try
            MyRez = CDbl(TextBox2.Text)
            If MyRez <= 0 Then
                MsgBox("""Закуп цена без НДС"" должна быть больше 0.", MsgBoxStyle.Critical, "Внимание!")
                TextBox2.Select()
                CheckFormFilling = False
                Exit Function
            End If
        Catch ex As Exception
            MsgBox("""Закуп цена без НДС"" должна быть числом.", MsgBoxStyle.Critical, "Внимание!")
            TextBox2.Select()
            CheckFormFilling = False
            Exit Function
        End Try

        '-----Даты действия КП от поставщика
        If DateTimePicker1.Value = CDate("01/01/1900") Then
        Else
            If DateTimePicker1.Value < CDate("01/01/1900").AddDays(-1) Then
                MsgBox("Дата - ""Предложение от поставщика действует до"" - должна быть не меньше текущей.", MsgBoxStyle.Critical, "Внимание!")
                DateTimePicker1.Select()
                CheckFormFilling = False
                Exit Function
            End If
        End If

        CheckFormFilling = True
    End Function

    Private Sub AddItem_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles Me.KeyDown
        '////////////////////////////////////////////////////////////////////////////////
        '//
        '// Запрет закрытия окна по ALT+F4
        '//
        '////////////////////////////////////////////////////////////////////////////////

        If e.KeyData = Keys.Alt + Keys.F4 Then
            e.Handled = True
        End If
    End Sub

    Private Function SaveRequest() As Boolean
        '////////////////////////////////////////////////////////////////////////////////
        '//
        '// Сохранение данных введенных в форму
        '//
        '////////////////////////////////////////////////////////////////////////////////
        Dim MySQLStr As String                        'рабочая строка

        If StartParam = "Create" Then       '-----создание новой записи
        Else                                '-----редактирование существующей
            Try
                MySQLStr = "UPDATE tbl_SupplSearch_PropItems "
                MySQLStr = MySQLStr & "SET ItemCode = N'" + Replace(Trim(TextBox1.Text), "'", "''") + "', "
                MySQLStr = MySQLStr & "ItemSuppCode = N'" & Replace(Trim(TextBox6.Text), "'", "''") & "', "
                MySQLStr = MySQLStr & "UOM = " & ComboBox1.SelectedValue.ToString & ", "
                MySQLStr = MySQLStr & "ItemName = N'" + Replace(Trim(TextBox3.Text), "'", "''") + "', "
                MySQLStr = MySQLStr & "QTY = " + Trim(Replace(Replace(Replace(TextBox4.Text, ",", "."), " ", ""), Chr(160), "")) + ", "
                MySQLStr = MySQLStr & "Price = " + Trim(Replace(Replace(Replace(TextBox2.Text, ",", "."), " ", ""), Chr(160), "")) + ", "
                'If Trim(ComboBox3.Text) = "EUR" Then
                '    MySQLStr = MySQLStr & "CurrCode = 12, "
                'ElseIf Trim(ComboBox3.Text) = "USD" Then
                '    MySQLStr = MySQLStr & "CurrCode = 1, "
                'ElseIf Trim(ComboBox3.Text) = "CNY" Then
                '    MySQLStr = MySQLStr & "CurrCode = 6, "
                'Else
                '    MySQLStr = MySQLStr & "CurrCode = 0, "
                'End If
                MySQLStr = MySQLStr & "CurrCode = " & CStr(ComboBox3.SelectedValue) & ", "
                MySQLStr = MySQLStr & "LeadTimeWeek = " & IIf(Trim(TextBox5.Text) = "", "NULL", Trim(Replace(Replace(Replace(TextBox5.Text, ",", "."), " ", ""), Chr(160), ""))) & ", "
                MySQLStr = MySQLStr & "Comments = N'" & Replace(Trim(TextBox7.Text), "'", "''") & "', "
                MySQLStr = MySQLStr & "AlternateTo = N'" & Replace(Trim(TextBox8.Text), "'", "''") & "', "
                If DateTimePicker1.Value = CDate("01/01/1900") Then
                    MySQLStr = MySQLStr & "DueDate = NULL "
                Else
                    MySQLStr = MySQLStr & "DueDate = CONVERT(datetime, '" + Format(DateTimePicker1.Value, "dd/MM/yyyy") + "', 103) "
                End If
                MySQLStr = MySQLStr & "WHERE (ID = " & Declarations.MyItemSrchID & ") "
                InitMyConn(False)
                Declarations.MyConn.Execute(MySQLStr)

                '-----Обновление валюты для всех строк поставщика
                If CheckBox1.Checked = True Then
                    MySQLStr = "UPDATE tbl_SupplSearch_PropItems "
                    'If Trim(ComboBox3.Text) = "EUR" Then
                    '    MySQLStr = MySQLStr & "SET CurrCode = 12 "
                    'ElseIf Trim(ComboBox3.Text) = "USD" Then
                    '    MySQLStr = MySQLStr & "SET CurrCode = 1 "
                    'ElseIf Trim(ComboBox3.Text) = "CNY" Then
                    '    MySQLStr = MySQLStr & "SET CurrCode = 6, "
                    'Else
                    '    MySQLStr = MySQLStr & "SET CurrCode = 0 "
                    'End If
                    MySQLStr = MySQLStr & "SET CurrCode = " & CStr(ComboBox3.SelectedValue) & " "
                    MySQLStr = MySQLStr & "FROM tbl_SupplSearch_PropItems INNER JOIN "
                    MySQLStr = MySQLStr & "tbl_SupplSearch_PropItems AS tbl_SupplSearch_PropItems_1 ON "
                    MySQLStr = MySQLStr & "tbl_SupplSearch_PropItems.SupplSearchID = tbl_SupplSearch_PropItems_1.SupplSearchID AND "
                    MySQLStr = MySQLStr & "tbl_SupplSearch_PropItems.SupplierCode = tbl_SupplSearch_PropItems_1.SupplierCode AND "
                    MySQLStr = MySQLStr & "tbl_SupplSearch_PropItems.SupplierName = tbl_SupplSearch_PropItems_1.SupplierName "
                    MySQLStr = MySQLStr & "WHERE (tbl_SupplSearch_PropItems_1.ID = " & Declarations.MyItemSrchID & ") "
                    InitMyConn(False)
                    Declarations.MyConn.Execute(MySQLStr)
                End If

                '-----Обновление дат для всех строк поставщика
                If CheckBox2.Checked = True Then
                    MySQLStr = "UPDATE tbl_SupplSearch_PropItems "
                    If DateTimePicker1.Value = CDate("01/01/1900") Then
                        MySQLStr = MySQLStr & "SET DueDate = NULL "
                    Else
                        MySQLStr = MySQLStr & "SET DueDate = CONVERT(datetime, '" + Format(DateTimePicker1.Value, "dd/MM/yyyy") + "', 103) "
                    End If
                    MySQLStr = MySQLStr & "FROM tbl_SupplSearch_PropItems INNER JOIN "
                    MySQLStr = MySQLStr & "tbl_SupplSearch_PropItems AS tbl_SupplSearch_PropItems_1 ON "
                    MySQLStr = MySQLStr & "tbl_SupplSearch_PropItems.SupplSearchID = tbl_SupplSearch_PropItems_1.SupplSearchID AND "
                    MySQLStr = MySQLStr & "tbl_SupplSearch_PropItems.SupplierCode = tbl_SupplSearch_PropItems_1.SupplierCode AND "
                    MySQLStr = MySQLStr & "tbl_SupplSearch_PropItems.SupplierName = tbl_SupplSearch_PropItems_1.SupplierName "
                    MySQLStr = MySQLStr & "WHERE (tbl_SupplSearch_PropItems_1.ID = " & Declarations.MyItemSrchID & ") "
                    InitMyConn(False)
                    Declarations.MyConn.Execute(MySQLStr)
                End If

            Catch ex As Exception
                MsgBox(ex.ToString)
                SaveRequest = False
                Exit Function
            End Try
        End If

        SaveRequest = True
    End Function

    Private Sub AddItem_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        '////////////////////////////////////////////////////////////////////////////////
        '//
        '// Загрузка окна
        '//
        '////////////////////////////////////////////////////////////////////////////////
        Dim MySQLStr As String                        'рабочая строка
        Dim MyAdapter As SqlClient.SqlDataAdapter     '
        Dim MyDs As New DataSet
        Dim MyAdapter3 As SqlClient.SqlDataAdapter     '
        Dim MyDs3 As New DataSet

        '---единицы измерения
        MySQLStr = "SELECT 0 AS UMID, SC09002 AS UMName "
        MySQLStr = MySQLStr & "FROM SC090300 WITH (NOLOCK) "
        MySQLStr = MySQLStr & "WHERE (SC09001 = 'RUS') "
        MySQLStr = MySQLStr & "UNION "
        MySQLStr = MySQLStr & "SELECT 1 AS UMID, SC09003 "
        MySQLStr = MySQLStr & "FROM SC090300 AS SC090300_40 WITH (NOLOCK) "
        MySQLStr = MySQLStr & "WHERE (SC09001 = 'RUS') "
        MySQLStr = MySQLStr & "UNION "
        MySQLStr = MySQLStr & "SELECT 2 AS UMID, SC09004 "
        MySQLStr = MySQLStr & "FROM SC090300 AS SC090300_39 WITH (NOLOCK) "
        MySQLStr = MySQLStr & "WHERE (SC09001 = 'RUS') "
        MySQLStr = MySQLStr & "UNION "
        MySQLStr = MySQLStr & "SELECT 3 AS UMID, SC09005 "
        MySQLStr = MySQLStr & "FROM SC090300 AS SC090300_38 WITH (NOLOCK) "
        MySQLStr = MySQLStr & "WHERE (SC09001 = 'RUS') "
        MySQLStr = MySQLStr & "UNION "
        MySQLStr = MySQLStr & "SELECT 4 AS UMID, SC09006 "
        MySQLStr = MySQLStr & "FROM SC090300 AS SC090300_37 WITH (NOLOCK) "
        MySQLStr = MySQLStr & "WHERE (SC09001 = 'RUS') "
        MySQLStr = MySQLStr & "UNION "
        MySQLStr = MySQLStr & "SELECT 5 AS UMID, SC09007 "
        MySQLStr = MySQLStr & "FROM SC090300 AS SC090300_36 WITH (NOLOCK) "
        MySQLStr = MySQLStr & "WHERE (SC09001 = 'RUS') "
        MySQLStr = MySQLStr & "UNION "
        MySQLStr = MySQLStr & "SELECT 6 AS UMID, SC09008 "
        MySQLStr = MySQLStr & "FROM SC090300 AS SC090300_35 WITH (NOLOCK) "
        MySQLStr = MySQLStr & "WHERE (SC09001 = 'RUS') "
        MySQLStr = MySQLStr & "UNION "
        MySQLStr = MySQLStr & "SELECT 7 AS UMID, SC09009 "
        MySQLStr = MySQLStr & "FROM SC090300 AS SC090300_34 WITH (NOLOCK) "
        MySQLStr = MySQLStr & "WHERE (SC09001 = 'RUS') "
        MySQLStr = MySQLStr & "UNION "
        MySQLStr = MySQLStr & "SELECT 8 AS UMID, SC09010 "
        MySQLStr = MySQLStr & "FROM SC090300 AS SC090300_33 WITH (NOLOCK) "
        MySQLStr = MySQLStr & "WHERE (SC09001 = 'RUS') "
        MySQLStr = MySQLStr & "UNION "
        MySQLStr = MySQLStr & "SELECT 9 AS UMID, SC09011 "
        MySQLStr = MySQLStr & "FROM SC090300 AS SC090300_32 WITH (NOLOCK) "
        MySQLStr = MySQLStr & "WHERE (SC09001 = 'RUS') "
        MySQLStr = MySQLStr & "UNION "
        MySQLStr = MySQLStr & "SELECT 10 AS UMID, SC09012 "
        MySQLStr = MySQLStr & "FROM SC090300 AS SC090300_31 WITH (NOLOCK) "
        MySQLStr = MySQLStr & "WHERE (SC09001 = 'RUS') "
        MySQLStr = MySQLStr & "UNION "
        MySQLStr = MySQLStr & "SELECT 11 AS UMID, SC09013 "
        MySQLStr = MySQLStr & "FROM SC090300 AS SC090300_30 WITH (NOLOCK) "
        MySQLStr = MySQLStr & "WHERE (SC09001 = 'RUS') "
        MySQLStr = MySQLStr & "UNION "
        MySQLStr = MySQLStr & "SELECT 12 AS UMID, SC09014 "
        MySQLStr = MySQLStr & "FROM SC090300 AS SC090300_29 WITH (NOLOCK) "
        MySQLStr = MySQLStr & "WHERE (SC09001 = 'RUS') "
        MySQLStr = MySQLStr & "UNION "
        MySQLStr = MySQLStr & "SELECT 13 AS UMID, SC09015 "
        MySQLStr = MySQLStr & "FROM SC090300 AS SC090300_28 WITH (NOLOCK) "
        MySQLStr = MySQLStr & "WHERE (SC09001 = 'RUS') "
        MySQLStr = MySQLStr & "UNION "
        MySQLStr = MySQLStr & "SELECT 14 AS UMID, SC09016 "
        MySQLStr = MySQLStr & "FROM SC090300 AS SC090300_27 WITH (NOLOCK) "
        MySQLStr = MySQLStr & "WHERE (SC09001 = 'RUS') "
        MySQLStr = MySQLStr & "UNION "
        MySQLStr = MySQLStr & "SELECT 15 AS UMID, SC09017 "
        MySQLStr = MySQLStr & "FROM SC090300 AS SC090300_26 WITH (NOLOCK) "
        MySQLStr = MySQLStr & "WHERE (SC09001 = 'RUS') "
        MySQLStr = MySQLStr & "UNION "
        MySQLStr = MySQLStr & "SELECT 16 AS UMID, SC09018 "
        MySQLStr = MySQLStr & "FROM SC090300 AS SC090300_25 WITH (NOLOCK) "
        MySQLStr = MySQLStr & "WHERE (SC09001 = 'RUS') "
        MySQLStr = MySQLStr & "UNION "
        MySQLStr = MySQLStr & "SELECT 17 AS UMID, SC09019 "
        MySQLStr = MySQLStr & "FROM SC090300 AS SC090300_24 WITH (NOLOCK) "
        MySQLStr = MySQLStr & "WHERE (SC09001 = 'RUS') "
        MySQLStr = MySQLStr & "UNION "
        MySQLStr = MySQLStr & "SELECT 18 AS UMID, SC09020 "
        MySQLStr = MySQLStr & "FROM SC090300 AS SC090300_23 WITH (NOLOCK) "
        MySQLStr = MySQLStr & "WHERE (SC09001 = 'RUS') "
        MySQLStr = MySQLStr & "UNION "
        MySQLStr = MySQLStr & "SELECT 19 AS UMID, SC09021 "
        MySQLStr = MySQLStr & "FROM SC090300 AS SC090300_22 WITH (NOLOCK) "
        MySQLStr = MySQLStr & "WHERE (SC09001 = 'RUS') "
        MySQLStr = MySQLStr & "UNION "
        MySQLStr = MySQLStr & "SELECT 20 AS UMID, SC09022 "
        MySQLStr = MySQLStr & "FROM SC090300 AS SC090300_21 WITH (NOLOCK) "
        MySQLStr = MySQLStr & "WHERE (SC09001 = 'RUS') "
        MySQLStr = MySQLStr & "UNION "
        MySQLStr = MySQLStr & "SELECT 21 AS UMID, SC09023 "
        MySQLStr = MySQLStr & "FROM SC090300 AS SC090300_20 WITH (NOLOCK) "
        MySQLStr = MySQLStr & "WHERE (SC09001 = 'RUS') "
        MySQLStr = MySQLStr & "UNION "
        MySQLStr = MySQLStr & "SELECT 22 AS UMID, SC09024 "
        MySQLStr = MySQLStr & "FROM SC090300 AS SC090300_19 WITH (NOLOCK) "
        MySQLStr = MySQLStr & "WHERE (SC09001 = 'RUS') "
        MySQLStr = MySQLStr & "UNION "
        MySQLStr = MySQLStr & "SELECT 23 AS UMID, SC09025 "
        MySQLStr = MySQLStr & "FROM SC090300 AS SC090300_18 WITH (NOLOCK) "
        MySQLStr = MySQLStr & "WHERE (SC09001 = 'RUS') "
        MySQLStr = MySQLStr & "UNION "
        MySQLStr = MySQLStr & "SELECT 24 AS UMID, SC09026 "
        MySQLStr = MySQLStr & "FROM SC090300 AS SC090300_17 WITH (NOLOCK) "
        MySQLStr = MySQLStr & "WHERE (SC09001 = 'RUS') "
        MySQLStr = MySQLStr & "UNION "
        MySQLStr = MySQLStr & "SELECT 25 AS UMID, SC09027 "
        MySQLStr = MySQLStr & "FROM SC090300 AS SC090300_16 WITH (NOLOCK) "
        MySQLStr = MySQLStr & "WHERE (SC09001 = 'RUS') "
        MySQLStr = MySQLStr & "UNION "
        MySQLStr = MySQLStr & "SELECT 26 AS UMID, SC09028 "
        MySQLStr = MySQLStr & "FROM SC090300 AS SC090300_15 WITH (NOLOCK) "
        MySQLStr = MySQLStr & "WHERE (SC09001 = 'RUS') "
        MySQLStr = MySQLStr & "UNION "
        MySQLStr = MySQLStr & "SELECT 27 AS UMID, SC09029 "
        MySQLStr = MySQLStr & "FROM SC090300 AS SC090300_14 WITH (NOLOCK) "
        MySQLStr = MySQLStr & "WHERE (SC09001 = 'RUS') "
        MySQLStr = MySQLStr & "UNION "
        MySQLStr = MySQLStr & "SELECT 28 AS UMID, SC09030 "
        MySQLStr = MySQLStr & "FROM SC090300 AS SC090300_13 WITH (NOLOCK) "
        MySQLStr = MySQLStr & "WHERE (SC09001 = 'RUS') "
        MySQLStr = MySQLStr & "UNION "
        MySQLStr = MySQLStr & "SELECT 29 AS UMID, SC09031 "
        MySQLStr = MySQLStr & "FROM SC090300 AS SC090300_12 WITH (NOLOCK) "
        MySQLStr = MySQLStr & "WHERE (SC09001 = 'RUS') "
        MySQLStr = MySQLStr & "UNION "
        MySQLStr = MySQLStr & "SELECT 30 AS UMID, SC09032 "
        MySQLStr = MySQLStr & "FROM SC090300 AS SC090300_11 WITH (NOLOCK) "
        MySQLStr = MySQLStr & "WHERE (SC09001 = 'RUS') "
        MySQLStr = MySQLStr & "UNION "
        MySQLStr = MySQLStr & "SELECT 31 AS UMID, SC09033 "
        MySQLStr = MySQLStr & "FROM SC090300 AS SC090300_10 WITH (NOLOCK) "
        MySQLStr = MySQLStr & "WHERE (SC09001 = 'RUS') "
        MySQLStr = MySQLStr & "UNION "
        MySQLStr = MySQLStr & "SELECT 32 AS UMID, SC09034 "
        MySQLStr = MySQLStr & "FROM SC090300 AS SC090300_9 WITH (NOLOCK) "
        MySQLStr = MySQLStr & "WHERE (SC09001 = 'RUS') "
        MySQLStr = MySQLStr & "UNION "
        MySQLStr = MySQLStr & "SELECT 33 AS UMID, SC09035 "
        MySQLStr = MySQLStr & "FROM SC090300 AS SC090300_8 WITH (NOLOCK) "
        MySQLStr = MySQLStr & "WHERE (SC09001 = 'RUS') "
        MySQLStr = MySQLStr & "UNION "
        MySQLStr = MySQLStr & "SELECT 34 AS UMID, SC09036 "
        MySQLStr = MySQLStr & "FROM SC090300 AS SC090300_7 WITH (NOLOCK) "
        MySQLStr = MySQLStr & "WHERE (SC09001 = 'RUS') "
        MySQLStr = MySQLStr & "UNION "
        MySQLStr = MySQLStr & "SELECT 35 AS UMID, SC09037 "
        MySQLStr = MySQLStr & "FROM SC090300 AS SC090300_6 WITH (NOLOCK) "
        MySQLStr = MySQLStr & "WHERE (SC09001 = 'RUS') "
        MySQLStr = MySQLStr & "UNION "
        MySQLStr = MySQLStr & "SELECT 36 AS UMID, SC09038 "
        MySQLStr = MySQLStr & "FROM SC090300 AS SC090300_5 WITH (NOLOCK) "
        MySQLStr = MySQLStr & "WHERE (SC09001 = 'RUS') "
        MySQLStr = MySQLStr & "UNION "
        MySQLStr = MySQLStr & "SELECT 37 AS UMID, SC09039 "
        MySQLStr = MySQLStr & "FROM SC090300 AS SC090300_4 WITH (NOLOCK) "
        MySQLStr = MySQLStr & "WHERE (SC09001 = 'RUS') "
        MySQLStr = MySQLStr & "UNION "
        MySQLStr = MySQLStr & "SELECT 38 AS UMID, SC09040 "
        MySQLStr = MySQLStr & "FROM SC090300 AS SC090300_3 WITH (NOLOCK) "
        MySQLStr = MySQLStr & "WHERE (SC09001 = 'RUS') "
        MySQLStr = MySQLStr & "UNION "
        MySQLStr = MySQLStr & "SELECT 39 AS UMID, SC09041 "
        MySQLStr = MySQLStr & "FROM SC090300 AS SC090300_2 WITH (NOLOCK) "
        MySQLStr = MySQLStr & "WHERE (SC09001 = 'RUS') "
        MySQLStr = MySQLStr & "UNION "
        MySQLStr = MySQLStr & "SELECT 40 AS UMID, SC09042 "
        MySQLStr = MySQLStr & "FROM SC090300 AS SC090300_1 WITH (NOLOCK) "
        MySQLStr = MySQLStr & "WHERE (SC09001 = 'RUS') "
        Try
            MyAdapter = New SqlClient.SqlDataAdapter(MySQLStr, Declarations.MyNETConnStr)
            MyAdapter.SelectCommand.CommandTimeout = 600
            MyAdapter.Fill(MyDs)
            ComboBox1.DisplayMember = "UMName" 'Это то что будет отображаться
            ComboBox1.ValueMember = "UMID"   'это то что будет храниться
            ComboBox1.DataSource = MyDs.Tables(0).DefaultView
        Catch ex As Exception
            MsgBox(ex.ToString)
        End Try

        '-----валюты
        MySQLStr = "SELECT SYCD001, SYCD009 "
        MySQLStr = MySQLStr & "FROM SYCD0100 "
        MySQLStr = MySQLStr & "WHERE (SYCD009 <> N'') "
        MySQLStr = MySQLStr & "AND (SYCD009 NOT IN ('FIM', 'FRF', 'SEK', 'DK', 'DM', 'FI1', 'ROL')) "
        Try
            MyAdapter3 = New SqlClient.SqlDataAdapter(MySQLStr, Declarations.MyNETConnStr)
            MyAdapter3.Fill(MyDs3)
            ComboBox3.DisplayMember = "SYCD009" 'Это то что будет отображаться
            ComboBox3.ValueMember = "SYCD001"   'это то что будет храниться
            ComboBox3.DataSource = MyDs3.Tables(0).DefaultView
        Catch ex As Exception
            MsgBox(ex.ToString)
        End Try

        '---данные
        If StartParam = "Create" Then
            Declarations.MyItemSrchID = 0
            Label3.Text = "New"
        Else
            Declarations.MyItemSrchID = MainForm.DataGridView4.SelectedRows.Item(0).Cells("ID").Value
            Label3.Text = Declarations.MyItemSrchID.ToString

            MySQLStr = "SELECT ID, ISNULL(ItemCode, '') AS ItemCode, ItemSuppCode, ItemName, "
            MySQLStr = MySQLStr & "UOM, QTY, ISNULL(Price, 0) AS Price, ISNULL(LeadTimeWeek, 0) "
            MySQLStr = MySQLStr & "AS LeadTimeWeek, ISNULL(Comments, '') AS Comments, "
            MySQLStr = MySQLStr & "ISNULL(tbl_SupplSearch_PropItems.CurrCode, 0) AS Curr, "
            MySQLStr = MySQLStr & "ISNULL(AlternateTo, '') AS AlternateTo, "
            MySQLStr = MySQLStr & "ISNULL(DueDate, CONVERT(datetime, '01/01/1900', 103)) AS DueDate "
            MySQLStr = MySQLStr & "FROM tbl_SupplSearch_PropItems "
            MySQLStr = MySQLStr & "WHERE (ID = " & Declarations.MyItemSrchID.ToString & ") "
            InitMyConn(False)
            InitMyRec(False, MySQLStr)
            If Declarations.MyRec.EOF = True And Declarations.MyRec.BOF = True Then
            Else
                TextBox1.Text = Declarations.MyRec.Fields("ItemCode").Value.ToString
                TextBox6.Text = Declarations.MyRec.Fields("ItemSuppCode").Value.ToString
                ComboBox1.SelectedValue = Declarations.MyRec.Fields("UOM").Value
                TextBox3.Text = Declarations.MyRec.Fields("ItemName").Value.ToString
                TextBox4.Text = Format(Declarations.MyRec.Fields("QTY").Value, "n3")
                TextBox2.Text = Format(Declarations.MyRec.Fields("Price").Value, "n2")
                TextBox5.Text = IIf(Declarations.MyRec.Fields("LeadTimeWeek").Value = 0, "", Format(Declarations.MyRec.Fields("LeadTimeWeek").Value, "n2"))
                TextBox7.Text = Declarations.MyRec.Fields("Comments").Value
                TextBox8.Text = Declarations.MyRec.Fields("AlternateTo").Value
                ComboBox3.SelectedValue = Declarations.MyRec.Fields("Curr").Value
                DateTimePicker1.Value = Declarations.MyRec.Fields("DueDate").Value
            End If
            trycloseMyRec()
        End If

        '---блокировка единиц измерения и кода товара поставщика
        If Trim(TextBox1.Text).Equals("") Then
            ComboBox1.Enabled = True
            TextBox8.Enabled = True
        Else
            ComboBox1.Enabled = False
            TextBox8.Enabled = False
        End If
    End Sub

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        '////////////////////////////////////////////////////////////////////////////////
        '//
        '// Открытие окна с товарами Scala
        '//
        '////////////////////////////////////////////////////////////////////////////////

        MyItemSelect = New ItemSelect
        MyItemSelect.ShowDialog()
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
        '////////////////////////////////////////////////////////////////////////////////////////
        '//
        '// Введен код товара - блокируем единицы измерения
        '//
        '////////////////////////////////////////////////////////////////////////////////////////

        TextBox1Validated()
    End Sub

    Public Sub TextBox1Validated()
        '////////////////////////////////////////////////////////////////////////////////////////
        '//
        '// Введен код товара - блокируем единицы измерения и код товара поставщика
        '//
        '////////////////////////////////////////////////////////////////////////////////////////

        If Trim(TextBox1.Text).Equals("") Then
            ComboBox1.Enabled = True
            TextBox6.Enabled = True
        Else
            ComboBox1.Enabled = False
            TextBox6.Enabled = False
        End If
    End Sub

    Private Sub TextBox1_Validating(ByVal sender As Object, ByVal e As System.ComponentModel.CancelEventArgs) Handles TextBox1.Validating
        '////////////////////////////////////////////////////////////////////////////////////////
        '//
        '// Введен код товара - находим и подписываем его значения
        '//
        '////////////////////////////////////////////////////////////////////////////////////////

        If TextBox1Validating() = False Then
            e.Cancel = True
        End If

    End Sub

    Public Function TextBox1Validating() As Boolean
        '////////////////////////////////////////////////////////////////////////////////////////
        '//
        '// Введен код товара - находим и подписываем его значения
        '//
        '////////////////////////////////////////////////////////////////////////////////////////
        Dim MySQLStr As String

        If Trim(TextBox1.Text) = "" Then
        Else
            MySQLStr = "SELECT SC01135 AS UOM, LTRIM(RTRIM(LTRIM(RTRIM(SC01002)) + ' ' + LTRIM(RTRIM(SC01003)))) AS Name, SC01060 AS SuppItemID "
            MySQLStr = MySQLStr & "FROM SC010300 "
            MySQLStr = MySQLStr & "WHERE (SC01001 = N'" & Trim(TextBox1.Text) & "') "
            If Trim(MainForm.DataGridView4.SelectedRows.Item(0).Cells("SupplierCode").Value.ToString) = "" Then
                MySQLStr = MySQLStr & "AND (SC01058 = N'******') "
            Else
                MySQLStr = MySQLStr & "AND (SC01058 = N'" & Trim(MainForm.DataGridView4.SelectedRows.Item(0).Cells("SupplierCode").Value.ToString) & "') "
            End If

            InitMyConn(False)
            InitMyRec(False, MySQLStr)
            If Declarations.MyRec.EOF = True And Declarations.MyRec.BOF = True Then
                trycloseMyRec()
                MsgBox("Вы ввели неверный код товара для поставщика " & Trim(MainForm.DataGridView4.SelectedRows.Item(0).Cells("SupplierCode").Value.ToString) & _
                    ". Введите корректный или воспользуйтесь поиском.", vbCritical, "Внимание!")
                TextBox1Validating = False
                Exit Function
            Else
                ComboBox1.SelectedValue = Declarations.MyRec.Fields("UOM").Value
                TextBox3.Text = Declarations.MyRec.Fields("Name").Value
                TextBox6.Text = Declarations.MyRec.Fields("SuppItemID").Value
                trycloseMyRec()
            End If
        End If
        TextBox1Validating = True
    End Function

    Private Sub TextBox4_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles TextBox4.KeyDown
        '////////////////////////////////////////////////////////////////////////////////
        '//
        '// переход на следующее поле по Enter
        '//
        '////////////////////////////////////////////////////////////////////////////////

        If e.KeyData = Keys.Enter Then
            Me.SelectNextControl(sender, True, True, True, False)
        End If
    End Sub

    Private Sub TextBox2_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles TextBox2.KeyDown
        '////////////////////////////////////////////////////////////////////////////////
        '//
        '// переход на следующее поле по Enter
        '//
        '////////////////////////////////////////////////////////////////////////////////

        If e.KeyData = Keys.Enter Then
            Me.SelectNextControl(sender, True, True, True, False)
        End If
    End Sub

    Private Sub TextBox5_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles TextBox5.KeyDown
        '////////////////////////////////////////////////////////////////////////////////
        '//
        '// переход на следующее поле по Enter
        '//
        '////////////////////////////////////////////////////////////////////////////////

        If e.KeyData = Keys.Enter Then
            Me.SelectNextControl(sender, True, True, True, False)
        End If
    End Sub
End Class