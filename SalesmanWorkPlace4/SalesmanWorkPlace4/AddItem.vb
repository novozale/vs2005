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
        Dim MySQLStr As String

        If Trim(TextBox1.Text) = "" And Trim(TextBox2.Text) = "" And Trim(TextBox3.Text) = "" Then
            MsgBox("Должно быть заполнено или ""Код Скала"" или ""Код поставщика"" или ""Название товара""", MsgBoxStyle.Critical, "Внимание!")
            TextBox1.Select()
            CheckFormFilling = False
            Exit Function
        End If

        ''------Проверка уникальности кода товара поставщика в запросе
        'If Trim(TextBox2.Text).Equals("") Then
        'Else
        '    MySQLStr = "SELECT COUNT(ID) AS CC "
        '    MySQLStr = MySQLStr & "FROM tbl_SupplSearchItems "
        '    MySQLStr = MySQLStr & "WHERE (SupplSearchID = " & Trim(MySearchSupplier.DataGridView1.SelectedRows.Item(0).Cells(0).Value.ToString()) & ") "
        '    MySQLStr = MySQLStr & "AND (ID <> " & Declarations.MyItemSrchID.ToString & ") "
        '    MySQLStr = MySQLStr & "AND (ItemSuppID = N'" & Trim(TextBox2.Text) & "') "
        '    InitMyConn(False)
        '    InitMyRec(False, MySQLStr)
        '    If Declarations.MyRec.BOF = True And Declarations.MyRec.EOF = True Then
        '        trycloseMyRec()
        '        MsgBox("Ошибка проверки уникальности кода товара производителя в запросе. Обратитесь к администратору.", MsgBoxStyle.Critical, "Внимание!")
        '        TextBox2.Select()
        '        CheckFormFilling = False
        '        Exit Function
        '    Else
        '        If Declarations.MyRec.Fields("CC").Value = 0 Then
        '            trycloseMyRec()
        '        Else
        '            trycloseMyRec()
        '            MsgBox("Ошибка - товар с кодом товара производителя " & Trim(TextBox2.Text) & " уже присутствует в запросе на поиск поставщика.", MsgBoxStyle.Critical, "Внимание!")
        '            TextBox2.Select()
        '            CheckFormFilling = False
        '            Exit Function
        '        End If
        '    End If
        'End If

        ''------Проверка уникальности названия товара в запросе
        'If Trim(TextBox3.Text).Equals("") Then
        'Else
        '    MySQLStr = "SELECT COUNT(ID) AS CC "
        '    MySQLStr = MySQLStr & "FROM tbl_SupplSearchItems "
        '    MySQLStr = MySQLStr & "WHERE (SupplSearchID = " & Trim(MySearchSupplier.DataGridView1.SelectedRows.Item(0).Cells(0).Value.ToString()) & ") "
        '    MySQLStr = MySQLStr & "AND (ID <> " & Declarations.MyItemSrchID.ToString & ") "
        '    MySQLStr = MySQLStr & "AND (ItemName = N'" & Trim(TextBox3.Text) & "') "
        '    InitMyConn(False)
        '    InitMyRec(False, MySQLStr)
        '    If Declarations.MyRec.BOF = True And Declarations.MyRec.EOF = True Then
        '        trycloseMyRec()
        '        MsgBox("Ошибка проверки уникальности названия товара. Обратитесь к администратору.", MsgBoxStyle.Critical, "Внимание!")
        '        TextBox3.Select()
        '        CheckFormFilling = False
        '        Exit Function
        '    Else
        '        If Declarations.MyRec.Fields("CC").Value = 0 Then
        '            trycloseMyRec()
        '        Else
        '            trycloseMyRec()
        '            MsgBox("Ошибка - товар с названием " & Trim(TextBox3.Text) & " уже присутствует в запросе на поиск поставщика.", MsgBoxStyle.Critical, "Внимание!")
        '            TextBox3.Select()
        '            CheckFormFilling = False
        '            Exit Function
        '        End If
        '    End If
        'End If

        '------Проверка уникальности кода товара поставщика + названия товара в запросе
        If Trim(TextBox2.Text).Equals("") And Trim(TextBox3.Text).Equals("") Then
        Else
            MySQLStr = "SELECT COUNT(ID) AS CC "
            MySQLStr = MySQLStr & "FROM tbl_SupplSearchItems "
            MySQLStr = MySQLStr & "WHERE (SupplSearchID = " & Trim(MySearchSupplier.DataGridView1.SelectedRows.Item(0).Cells(0).Value.ToString()) & ") "
            MySQLStr = MySQLStr & "AND (ID <> " & Declarations.MyItemSrchID.ToString & ") "
            MySQLStr = MySQLStr & "AND (ItemSuppID = N'" & Trim(TextBox2.Text) & "') "
            MySQLStr = MySQLStr & "AND (ItemName = N'" & Trim(TextBox3.Text) & "') "
            InitMyConn(False)
            InitMyRec(False, MySQLStr)
            If Declarations.MyRec.BOF = True And Declarations.MyRec.EOF = True Then
                trycloseMyRec()
                MsgBox("Ошибка проверки уникальности кода товара производителя в запросе. Обратитесь к администратору.", MsgBoxStyle.Critical, "Внимание!")
                TextBox2.Select()
                CheckFormFilling = False
                Exit Function
            Else
                If Declarations.MyRec.Fields("CC").Value = 0 Then
                    trycloseMyRec()
                Else
                    trycloseMyRec()
                    MsgBox("Ошибка - товар с кодом товара производителя " & Trim(TextBox2.Text) & " уже присутствует в запросе на поиск поставщика.", MsgBoxStyle.Critical, "Внимание!")
                    TextBox2.Select()
                    CheckFormFilling = False
                    Exit Function
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

    Private Sub AddItem_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        '////////////////////////////////////////////////////////////////////////////////
        '//
        '// Загрузка окна
        '//
        '////////////////////////////////////////////////////////////////////////////////
        Dim MySQLStr As String                        'рабочая строка
        Dim MyAdapter As SqlClient.SqlDataAdapter     '
        Dim MyDs As New DataSet

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

        '---данные
        If StartParam = "Create" Then
            Declarations.MyItemSrchID = 0
            Label3.Text = "New"
        Else
            Declarations.MyItemSrchID = MySearchSupplier.DataGridView2.SelectedRows.Item(0).Cells(0).Value
            Label3.Text = Declarations.MyItemSrchID.ToString

            MySQLStr = "SELECT ID, ISNULL(ItemID, '') AS ItemID, ISNULL(ItemSuppID, '') "
            MySQLStr = MySQLStr & "AS ItemSuppID, ISNULL(ItemName, '') AS ItemName, UOM, QTY, ISNULL(LeadTimeWeek, 0) "
            MySQLStr = MySQLStr & "AS LeadTimeWeek, ISNULL(Comments, '') AS Comments "
            MySQLStr = MySQLStr & "FROM tbl_SupplSearchItems "
            MySQLStr = MySQLStr & "WHERE (ID = " & Declarations.MyItemSrchID.ToString & ") "
            InitMyConn(False)
            InitMyRec(False, MySQLStr)
            If Declarations.MyRec.EOF = True And Declarations.MyRec.BOF = True Then
            Else
                TextBox1.Text = Declarations.MyRec.Fields("ItemID").Value.ToString
                TextBox2.Text = Declarations.MyRec.Fields("ItemSuppID").Value.ToString
                TextBox3.Text = Declarations.MyRec.Fields("ItemName").Value.ToString
                ComboBox1.SelectedValue = Declarations.MyRec.Fields("UOM").Value
                TextBox4.Text = Format(Declarations.MyRec.Fields("QTY").Value, "n3")
                TextBox5.Text = IIf(Declarations.MyRec.Fields("LeadTimeWeek").Value = 0, "", Format(Declarations.MyRec.Fields("LeadTimeWeek").Value, "n2"))
                TextBox6.Text = Trim(Declarations.MyRec.Fields("Comments").Value.ToString)
            End If
            trycloseMyRec()
            TextBox1Validated()
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
            Try
                MySQLStr = "INSERT INTO tbl_SupplSearchItems "
                MySQLStr = MySQLStr + "(SupplSearchID, ItemID, ItemSuppID, ItemName, UOM, QTY, LeadTimeWeek, Comments) "
                MySQLStr = MySQLStr + "VALUES ("
                MySQLStr = MySQLStr + Trim(MySearchSupplier.DataGridView1.SelectedRows.Item(0).Cells(0).Value.ToString()) + ", "
                MySQLStr = MySQLStr + "N'" + Replace(Trim(TextBox1.Text), "'", "''") + "', "
                MySQLStr = MySQLStr + "N'" + Replace(Trim(TextBox2.Text), "'", "''") + "', "
                MySQLStr = MySQLStr + "N'" + Replace(Trim(TextBox3.Text), "'", "''") + "', "
                MySQLStr = MySQLStr + ComboBox1.SelectedValue.ToString + ", "
                MySQLStr = MySQLStr + Replace(Trim(TextBox4.Text), ",", ".") + ", "
                MySQLStr = MySQLStr + IIf(Trim(TextBox5.Text) = "", "NULL", Replace(Trim(TextBox5.Text), ",", ".")) + ", "
                MySQLStr = MySQLStr + "N'" + Replace(Trim(TextBox6.Text), "'", "''") + "')"
                InitMyConn(False)
                Declarations.MyConn.Execute(MySQLStr)
            Catch ex As Exception
                MsgBox(ex.ToString)
                SaveRequest = False
                Exit Function
            End Try
        Else                                '-----редактирование существующей
            Try
                MySQLStr = "UPDATE tbl_SupplSearchItems "
                MySQLStr = MySQLStr & "SET ItemID = N'" + Trim(TextBox1.Text) + "', "
                MySQLStr = MySQLStr & "ItemSuppID = N'" + Trim(TextBox2.Text) + "', "
                MySQLStr = MySQLStr & "ItemName = N'" + Trim(TextBox3.Text) + "', "
                MySQLStr = MySQLStr & "UOM = " & ComboBox1.SelectedValue.ToString & ", "
                MySQLStr = MySQLStr & "QTY = " + Trim(Replace(Replace(Replace(TextBox4.Text, ",", "."), " ", ""), Chr(160), "")) + ", "
                MySQLStr = MySQLStr & "LeadTimeWeek = " & IIf(Trim(TextBox5.Text) = "", "NULL", Trim(Replace(Replace(Replace(TextBox5.Text, ",", "."), " ", ""), Chr(160), ""))) & ", "
                MySQLStr = MySQLStr & "Comments = N'" + Replace(Trim(TextBox6.Text), "'", "''") + "' "
                MySQLStr = MySQLStr & "WHERE (ID = " & Declarations.MyItemSrchID & ") "
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
        '// открытие окна выбора кода товара в Scala
        '//
        '////////////////////////////////////////////////////////////////////////////////

        MyItemSelect = New ItemSelect
        MyItemSelect.MySrcWin = "AddItem"
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
        '// Введен код товара - блокируем единицы измерения, код товара поставщика и название
        '//
        '////////////////////////////////////////////////////////////////////////////////////////

        If Trim(TextBox1.Text).Equals("") Then
            ComboBox1.Enabled = True
            TextBox2.Enabled = True
            TextBox3.Enabled = True
        Else
            ComboBox1.Enabled = False
            TextBox2.Enabled = False
            TextBox3.Enabled = False
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

            InitMyConn(False)
            InitMyRec(False, MySQLStr)
            If Declarations.MyRec.EOF = True And Declarations.MyRec.BOF = True Then
                trycloseMyRec()
                MsgBox("Вы ввели неверный код товара Scala. Введите корректный или воспользуйтесь поиском.", vbCritical, "Внимание!")
                TextBox1Validating = False
                Exit Function
            Else
                ComboBox1.SelectedValue = Declarations.MyRec.Fields("UOM").Value
                TextBox3.Text = Declarations.MyRec.Fields("Name").Value
                TextBox2.Text = Declarations.MyRec.Fields("SuppItemID").Value
                trycloseMyRec()
            End If
        End If
        TextBox1Validating = True
    End Function
End Class