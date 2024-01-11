Public Class EditInOrder

    Public MyItem As String                           'Номер строки
    Public MyOrder As String                          'номер заказа


    Private Sub Button2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button2.Click
        '////////////////////////////////////////////////////////////////////////////////
        '//
        '// Закрытие окна без редактирования запаса в заказе
        '//
        '////////////////////////////////////////////////////////////////////////////////

        Declarations.MySuccess = False
        Me.Close()
    End Sub

    Private Sub EditInOrder_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        '////////////////////////////////////////////////////////////////////////////////
        '//
        '// Открытие окна и загрузка в него информации
        '//
        '////////////////////////////////////////////////////////////////////////////////
        Dim MySQLStr As String                        'рабочая строка
        Dim MyAdapter As SqlClient.SqlDataAdapter     '
        Dim MyDs As New DataSet                       '
        Dim ExistInScala As Double
        Dim SuppExistInScala As Double
        Dim aa As New System.Globalization.NumberFormatInfo

        '---Валюта
        Label8.Text = Declarations.CurrencyName
        Label9.Text = Declarations.CurrencyValue

        '---Список единиц измерения
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

        TextBox1.Text = MyOrderLines.DataGridView2.SelectedRows.Item(0).Cells(2).Value.ToString
        TextBox2.Text = MyOrderLines.DataGridView2.SelectedRows.Item(0).Cells(3).Value.ToString
        TextBox15.Text = MyOrderLines.DataGridView2.SelectedRows.Item(0).Cells(16).Value.ToString
        TextBox16.Text = MyOrderLines.DataGridView2.SelectedRows.Item(0).Cells(17).Value.ToString

        TextBox3.Text = Math.Round(CDbl(MyOrderLines.DataGridView2.SelectedRows.Item(0).Cells(7).Value.ToString), 3)
        TextBox4.Text = Math.Round(CDbl(MyOrderLines.DataGridView2.SelectedRows.Item(0).Cells(6).Value.ToString), 2)
        TextBox5.Text = Math.Round(CDbl(MyOrderLines.DataGridView2.SelectedRows.Item(0).Cells(5).Value.ToString), 2)
        TextBox6.Text = Math.Round(CDbl(MyOrderLines.DataGridView2.SelectedRows.Item(0).Cells(10).Value.ToString), 2)
        TextBox7.Text = Math.Round(CDbl(MyOrderLines.DataGridView2.SelectedRows.Item(0).Cells(12).Value.ToString), 2)
        '---Единица измерения
        MySQLStr = "Select OR03010 "
        MySQLStr = MySQLStr & "FROM tbl_OR030300 WITH (NOLOCK) "
        MySQLStr = MySQLStr & "WHERE (OR03001 = N'" & MyOrder & "') "
        MySQLStr = MySQLStr & "AND (OR03002 = N'" & Trim(MyOrderLines.DataGridView2.SelectedRows.Item(0).Cells(1).Value.ToString) & "') "
        InitMyConn(False)
        InitMyRec(False, MySQLStr)
        If Declarations.MyRec.EOF = True And Declarations.MyRec.BOF = True Then
            ComboBox1.SelectedValue = 0
        Else
            ComboBox1.SelectedValue = Declarations.MyRec.Fields("OR03010").Value
        End If
        trycloseMyRec()

        '---Есть ли в Scala
        MySQLStr = "Select COUNT(*) AS CC "
        MySQLStr = MySQLStr & "FROM SC010300 WITH (NOLOCK) "
        MySQLStr = MySQLStr & "WHERE SC01001 = N'" & Trim(TextBox1.Text) & "'"
        InitMyConn(False)
        InitMyRec(False, MySQLStr)
        ExistInScala = Declarations.MyRec.Fields("CC").Value
        trycloseMyRec()

        If ExistInScala = 0 Then '---В Scala нет
            TextBox2.ReadOnly = False
            ComboBox1.Enabled = True
            TextBox15.Enabled = False
            TextBox15.BackColor = Color.FromName("Window")
            Button9.Enabled = True
        Else
            TextBox2.ReadOnly = True
            ComboBox1.Enabled = False
            TextBox15.Enabled = True
            TextBox15.BackColor = Color.FromName("ButtonFace")
            Button9.Enabled = False
        End If

        '---Есть ли поставщик в Scala
        MySQLStr = "Select COUNT(*) AS CC "
        MySQLStr = MySQLStr & "FROM PL010300 WITH (NOLOCK) "
        MySQLStr = MySQLStr & "WHERE PL01001 = N'" & Trim(TextBox15.Text) & "'"
        InitMyConn(False)
        InitMyRec(False, MySQLStr)
        SuppExistInScala = Declarations.MyRec.Fields("CC").Value
        trycloseMyRec()
        If SuppExistInScala = 0 Then '---поставщика В Scala нет
            TextBox16.Enabled = False
            TextBox16.BackColor = Color.FromName("Window")
        Else
            TextBox16.Enabled = True
            TextBox16.BackColor = Color.FromName("ButtonFace")
        End If


        '---Есть ли расчетная себестоимость в Scala
        MySQLStr = "SELECT SC030300.SC03057 AS SS, "
        MySQLStr = MySQLStr & "SC010300.SC01053 AS CSS "
        MySQLStr = MySQLStr & "FROM SC030300 WITH (NOLOCK) INNER JOIN SC010300 ON "
        MySQLStr = MySQLStr & "SC030300.SC03001 = SC010300.SC01001 "
        MySQLStr = MySQLStr & "WHERE (SC030300.SC03001 = N'" & Trim(TextBox1.Text) & "') "
        MySQLStr = MySQLStr & "AND (SC030300.SC03002 = N'" & Declarations.WHNum & "')"
        InitMyConn(False)
        InitMyRec(False, MySQLStr)
        If Declarations.MyRec.EOF = True And Declarations.MyRec.BOF = True Then
            TextBox5.ReadOnly = False
            Label3.Text = "Рекомендованная цена и себестоимость  для этого запаса должны быть определены самостоятельно"
            Label3.ForeColor = Color.Red
        Else
            If Declarations.MyRec.Fields("CSS").Value = 0 Then
                TextBox5.ReadOnly = False
                Label3.Text = "Рекомендованная цена и себестоимость  для этого запаса должны быть определены самостоятельно"
                Label3.ForeColor = Color.Red
            Else
                TextBox5.ReadOnly = True
                Label3.Text = "Рекомендованная цена и себестоимость этого запаса на основе прайс - листа на закупку"
                Label3.ForeColor = Color.Green
            End If
        End If
        trycloseMyRec()

        '--------складской или нет на данном складе--------------------
        MySQLStr = "SELECT COUNT(SC010300_2.SC01001) AS CC "
        MySQLStr = MySQLStr & "FROM SC010300 AS SC010300_2 WITH (NOLOCK) CROSS JOIN "
        MySQLStr = MySQLStr & "(SELECT SC23001 AS WH, CHARINDEX('1', SC23007) AS WHPos "
        MySQLStr = MySQLStr & "FROM SC230300 WITH (NOLOCK)"
        MySQLStr = MySQLStr & "WHERE (SC23006 = N'1') AND (SC23001 = N'" & Declarations.WHNum & "')) AS View_2 "
        MySQLStr = MySQLStr & "WHERE (SUBSTRING(SC010300_2.SC01128, View_2.WHPos, 1) = N'1') AND "
        MySQLStr = MySQLStr & "(SC010300_2.SC01001 = N'" & Trim(TextBox1.Text) & "') "

        InitMyConn(False)
        InitMyRec(False, MySQLStr)
        If Declarations.MyRec.EOF = True And Declarations.MyRec.BOF = True Then
            Label25.Text = ""
        Else
            Declarations.MyRec.MoveFirst()
            If Declarations.MyRec.Fields("CC").Value = 0 Then '--не складской
                Label25.Text = ""
            Else
                Label25.Text = "Складской ассортимент"
            End If
        End If
        trycloseMyRec()

        '--------кратность в упаковке----------------------------------
        MySQLStr = "SELECT SC01072 AS CC "
        MySQLStr = MySQLStr & "FROM SC010300 WITH (NOLOCK) "
        MySQLStr = MySQLStr & "WHERE (SC01001 = N'" & Trim(TextBox1.Text) & "') "
        InitMyConn(False)
        InitMyRec(False, MySQLStr)
        If Declarations.MyRec.EOF = True And Declarations.MyRec.BOF = True Then
            Label23.Text = ""
        Else
            Declarations.MyRec.MoveFirst()
            Label23.Text = Math.Round(Declarations.MyRec.Fields("CC").Value, 2)
        End If
        trycloseMyRec()

        '--------минимальное количество в заказе на закупку------------
        MySQLStr = "SELECT tbl_PurchasePriceHistory.MinQTY AS CC, LT "
        MySQLStr = MySQLStr & "FROM SC010300 WITH (NOLOCK) INNER JOIN "
        MySQLStr = MySQLStr & "tbl_PurchasePriceHistory WITH (NOLOCK) ON SC010300.SC01001 = tbl_PurchasePriceHistory.SC01001 "
        MySQLStr = MySQLStr & "WHERE (SC010300.SC01001 = N'" & Trim(TextBox1.Text) & "') AND "
        MySQLStr = MySQLStr & "(tbl_PurchasePriceHistory.DateTo = CONVERT(DATETIME, '9999-12-31 00:00:00', 102)) "
        InitMyConn(False)
        InitMyRec(False, MySQLStr)
        If Declarations.MyRec.EOF = True And Declarations.MyRec.BOF = True Then
            Label19.Text = ""
            Label21.Text = ""
        Else
            Declarations.MyRec.MoveFirst()
            Label19.Text = Math.Round(Declarations.MyRec.Fields("CC").Value, 2)
            If Declarations.MyRec.Fields("LT").Value = 0 Then
                Label21.Text = ""
            Else
                Label21.Text = Math.Round(Declarations.MyRec.Fields("LT").Value, 0)
            End If
        End If
        trycloseMyRec()

        '---предполагаемый срок поставки строки
        MySQLStr = "SELECT ISNULL(WeekQTY,1) AS CC,  ISNULL(DelWeekQTY,0) AS CC1 "
        MySQLStr = MySQLStr & "FROM tbl_OR030300 "
        MySQLStr = MySQLStr & "WHERE (OR03003 = N'000000') AND (OR03004 = 0000) AND "
        MySQLStr = MySQLStr & "(OR03001 = N'" & Trim(MyOrderLines.Label6.Text) & "') AND "
        MySQLStr = MySQLStr & "(OR03002 = N'" & MyOrderLines.DataGridView2.SelectedRows.Item(0).Cells(1).Value.ToString & "') "
        InitMyRec(False, MySQLStr)
        If Declarations.MyRec.EOF = True And Declarations.MyRec.BOF = True Then
            TextBox10.Text = Format(1, "#####0.##")
            TextBox11.Text = Format(0, "#####0.##")
            If Declarations.IsSelfDelivery = 1 Then
                TextBox11.Enabled = False
            Else
                TextBox11.Enabled = True
            End If
        Else
            'TextBox10.Text = Format(Declarations.MyRec.Fields("CC").Value, "#####0" & aa.CurrentInfo.NumberDecimalSeparator & "##")
            TextBox10.Text = Format(Declarations.MyRec.Fields("CC").Value, "#####0.##")
            If Declarations.IsSelfDelivery = 1 Then
                TextBox11.Text = Format(0, "#####0.##")
                TextBox11.Enabled = False
            Else
                TextBox11.Text = Format(Declarations.MyRec.Fields("CC1").Value, "#####0.##")
                TextBox11.Enabled = True
            End If
        End If
        trycloseMyRec()

        If Declarations.IsSelfDelivery = 1 Then
            TextBox12.Text = Format(0, "#####0.##")
            TextBox12.Enabled = False
        Else
            TextBox12.Enabled = True
        End If

        '---Картинка, название для WEB, описание для WEB
        LoadItemAddInfo1(Trim(MyOrderLines.DataGridView2.SelectedRows.Item(0).Cells(2).Value.ToString))

        '---------Код товара поставщика
        MySQLStr = "SELECT ISNULL(SuppItemCode, '') AS SuppItemCode "
        MySQLStr = MySQLStr & "FROM tbl_OR030300 "
        MySQLStr = MySQLStr & "WHERE (OR03003 = N'000000') AND (OR03004 = 0000) AND "
        MySQLStr = MySQLStr & "(OR03001 = N'" & Trim(MyOrderLines.Label6.Text) & "') AND "
        MySQLStr = MySQLStr & "(OR03002 = N'" & MyOrderLines.DataGridView2.SelectedRows.Item(0).Cells(1).Value.ToString & "') "
        InitMyRec(False, MySQLStr)
        If Declarations.MyRec.EOF = True And Declarations.MyRec.BOF = True Then
            TextBox14.Text = ""
        Else
            TextBox14.Text = Declarations.MyRec.Fields("SuppItemCode").Value
        End If

        TextBox1.Select()
    End Sub

    Public Sub LoadItemAddInfo1(ByVal MyItemID As String)
        '////////////////////////////////////////////////////////////////////////////////
        '//
        '// Загрузка Картинки, названия для WEB, описания для WEB
        '// и выставление значений кнопок
        '//
        '////////////////////////////////////////////////////////////////////////////////
        Dim MySQLStr As String

        MySQLStr = "SELECT tbl_WEB_Pictures.Picture, ISNULL(tbl_WEB_Items.WEBName,'') AS WEBName, ISNULL(tbl_WEB_Items.Description, '') AS Description "
        MySQLStr = MySQLStr & "FROM SC010300 LEFT OUTER JOIN "
        MySQLStr = MySQLStr & "tbl_WEB_Pictures ON SC010300.SC01001 = tbl_WEB_Pictures.ScalaItemCode LEFT OUTER JOIN "
        MySQLStr = MySQLStr & "tbl_WEB_Items ON SC010300.SC01001 = tbl_WEB_Items.Code "
        MySQLStr = MySQLStr & "WHERE (SC010300.SC01001 = N'" & MyItemID & "') "
        InitMyConn(False)
        InitMyRec(False, MySQLStr)
        If Declarations.MyRec.EOF = True And Declarations.MyRec.BOF = True Then
            PictureBox1.Image = Nothing
            TextBox13.Text = ""
            RichTextBox1.Text = ""
            Button8.Enabled = False
            Button7.Enabled = False
            Button6.Enabled = False
            trycloseMyRec()
        Else
            Try
                Dim ms As New IO.MemoryStream(CType(Declarations.MyRec.Fields("Picture").Value, Byte()))
                Dim picture As Image

                picture = Image.FromStream(ms)
                PictureBox1.Image = picture
            Catch ex As Exception
            End Try
            TextBox13.Text = Declarations.MyRec.Fields("WEBName").Value
            RichTextBox1.Text = Declarations.MyRec.Fields("Description").Value

            If IsNothing(PictureBox1.Image) = True Then
                Button8.Enabled = False
            Else
                Button8.Enabled = True
            End If
            If TextBox13.Text = "" Then
                Button7.Enabled = False
            Else
                Button7.Enabled = True
            End If
            If RichTextBox1.Text = "" Then
                Button6.Enabled = False
            Else
                Button6.Enabled = True
            End If

            trycloseMyRec()
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

    Private Sub TextBox3_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles TextBox3.KeyDown
        '////////////////////////////////////////////////////////////////////////////////
        '//
        '// переход на следующее поле по Enter
        '//
        '////////////////////////////////////////////////////////////////////////////////

        If e.KeyData = Keys.Enter Then
            Me.SelectNextControl(sender, True, True, True, False)
        End If
    End Sub

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

    Private Sub TextBox6_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles TextBox6.KeyDown
        '////////////////////////////////////////////////////////////////////////////////
        '//
        '// переход на следующее поле по Enter
        '//
        '////////////////////////////////////////////////////////////////////////////////

        If e.KeyData = Keys.Enter Then
            Me.SelectNextControl(sender, True, True, True, False)
        End If
    End Sub

    Private Sub TextBox8_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles TextBox8.KeyDown
        '////////////////////////////////////////////////////////////////////////////////
        '//
        '// переход на следующее поле по Enter
        '//
        '////////////////////////////////////////////////////////////////////////////////

        If e.KeyData = Keys.Enter Then
            Me.SelectNextControl(sender, True, True, True, False)
        End If
    End Sub

    Private Sub ComboBox1_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles ComboBox1.KeyDown
        '////////////////////////////////////////////////////////////////////////////////
        '//
        '// переход на следующее поле по Enter
        '//
        '////////////////////////////////////////////////////////////////////////////////

        If e.KeyData = Keys.Enter Then
            Me.SelectNextControl(sender, True, True, True, False)
        End If
    End Sub

    Private Sub ComboBox1_SelectedValueChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles ComboBox1.SelectedValueChanged
        '////////////////////////////////////////////////////////////////////////////////
        '//
        '// переход на следующее поле по выбору значения
        '//
        '////////////////////////////////////////////////////////////////////////////////

        Me.SelectNextControl(ComboBox1, True, True, True, False)
    End Sub

    Private Sub TextBox1_Validating(ByVal sender As Object, ByVal e As System.ComponentModel.CancelEventArgs) Handles TextBox1.Validating
        '////////////////////////////////////////////////////////////////////////////////
        '//
        '// После ввода кода - проверка, есть ли такой запас в базе. Если есть - 
        '// подтягиваем его значения. Если нет - даем ввести свои
        '////////////////////////////////////////////////////////////////////////////////
        If TextBox1.Modified = True Then
            Me.TextBox1Validation()
            LoadItemAddInfo1(Trim(TextBox1.Text))
        End If
    End Sub
    Public Sub TextBox1Validation()
        '////////////////////////////////////////////////////////////////////////////////
        '//
        '// После ввода кода - проверка, есть ли такой запас в базе. Если есть - 
        '// подтягиваем его значения. Если нет - даем ввести свои
        '////////////////////////////////////////////////////////////////////////////////
        Dim MySQLStr As String                        'рабочая строка
        Dim ExistFlag As Integer
        Dim MyRez As Object

        MySQLStr = "SELECT SC010300.SC01001 AS ItemCode, "
        MySQLStr = MySQLStr & "LTRIM(RTRIM(SC010300.SC01002 + ' ' + SC010300.SC01003)) AS ItemName, "
        MySQLStr = MySQLStr & "SC010300.SC01135 AS UnitCode, "
        MySQLStr = MySQLStr & "ISNULL(View_1.UMName, '') AS UMName, "
        MySQLStr = MySQLStr & "ISNULL(View_2.SC39005, 0) / " & Replace(CStr(Declarations.CurrencyValue), ",", ".") & " AS Price, "
        MySQLStr = MySQLStr & "SC010300.SC01053 / " & Replace(CStr(Declarations.CurrencyValue), ",", ".") & " AS PriCost, "
        MySQLStr = MySQLStr & "SC010300.SC01058 AS SuppID, "
        MySQLStr = MySQLStr & "ISNULL(PL010300.PL01002, '') AS SuppName, "
        MySQLStr = MySQLStr & "SC010300.SC01060 AS SuppItemCode "
        MySQLStr = MySQLStr & "FROM SC010300 WITH (NOLOCK) LEFT OUTER JOIN "
        MySQLStr = MySQLStr & "PL010300 ON SC010300.SC01058 = PL010300.PL01001 LEFT OUTER JOIN "
        MySQLStr = MySQLStr & "(SELECT SC39001, SC39005 "
        MySQLStr = MySQLStr & "FROM SC390300 WITH (NOLOCK) "
        MySQLStr = MySQLStr & "WHERE (SC39002 = N'00')) AS View_2 ON "
        MySQLStr = MySQLStr & "SC010300.SC01001 = View_2.SC39001 LEFT OUTER JOIN "
        MySQLStr = MySQLStr & "(SELECT 0 AS UMID, SC09002 AS UMName "
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
        MySQLStr = MySQLStr & "FROM SC090300 AS SC090300_22 WITH (NOLOCK)"
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
        MySQLStr = MySQLStr & "WHERE (SC09001 = 'RUS')) AS View_1 ON "
        MySQLStr = MySQLStr & "SC010300.SC01135 = View_1.UMID "
        MySQLStr = MySQLStr & "WHERE (SC010300.SC01001 = N'" & TextBox1.Text & "')"
        InitMyConn(False)
        InitMyRec(False, MySQLStr)
        If Declarations.MyRec.EOF = True And Declarations.MyRec.BOF = True Then
            '---такого запаса нет у нас в БД
            Label3.Text = "Рекомендованная цена и себестоимость  для этого запаса должны быть определены самостоятельно"
            Label3.ForeColor = Color.Red
            MyRez = MsgBox("Очистить поля от предыдущих значений?", MsgBoxStyle.YesNo, "Внимание!")
            If MyRez = vbYes Then
                TextBox14.Text = ""
                TextBox2.Text = ""
                TextBox2.ReadOnly = False
                ComboBox1.Text = ""
                ComboBox1.Enabled = True
                TextBox4.Text = ""
                TextBox5.Text = ""
                TextBox5.ReadOnly = False
                TextBox15.Text = ""
                TextBox15.Enabled = True
                TextBox15.BackColor = Color.FromName("Window")
                Button9.Enabled = True
                TextBox16.Text = ""
                TextBox16.Enabled = False
                TextBox16.BackColor = Color.FromName("ButtonFace")
                Label25.Text = ""
                Label23.Text = ""
                Label19.Text = ""
                Label21.Text = ""
                TextBox10.Text = ""
                TextBox11.Text = "0"
            Else
                TextBox2.ReadOnly = False
                ComboBox1.Enabled = True
                TextBox5.ReadOnly = False
                TextBox15.Enabled = True
                TextBox15.BackColor = Color.FromName("Window")
                Button9.Enabled = True
                TextBox16.Enabled = False
                TextBox16.BackColor = Color.FromName("ButtonFace")
            End If
        Else
            '---Запас в БД есть
            If Declarations.MyRec.Fields("PriCost").Value = 0 Then
                TextBox5.Text = ""
                Label3.Text = "Рекомендованная цена и себестоимость  для этого запаса должны быть определены самостоятельно"
                Label3.ForeColor = Color.Red
                TextBox5.ReadOnly = False
            Else
                TextBox5.Text = Math.Round(Declarations.MyRec.Fields("PriCost").Value, 2)
                Label3.Text = "Рекомендованная цена и себестоимость этого запаса на основе прайс - листа на закупку"
                Label3.ForeColor = Color.Green
                TextBox5.ReadOnly = True
            End If
            TextBox14.Text = Declarations.MyRec.Fields("SuppItemCode").Value
            TextBox2.Text = Declarations.MyRec.Fields("ItemName").Value.ToString
            TextBox2.ReadOnly = True
            ComboBox1.Text = Declarations.MyRec.Fields("UMName").Value.ToString
            ComboBox1.Enabled = False
            If Declarations.MyRec.Fields("Price").Value = 0 Then
                TextBox4.Text = ""
            Else
                TextBox4.Text = Math.Round(Declarations.MyRec.Fields("Price").Value, 2)
            End If
            TextBox15.Text = Declarations.MyRec.Fields("SuppID").Value.ToString
            TextBox15.Enabled = False
            TextBox15.BackColor = Color.FromName("ButtonFace")
            Button9.Enabled = False
            TextBox16.Text = Declarations.MyRec.Fields("SuppName").Value.ToString
            TextBox16.Enabled = False
            TextBox16.BackColor = Color.FromName("ButtonFace")

            trycloseMyRec()

            System.Windows.Forms.Cursor.Current = Cursors.WaitCursor
            For i As Integer = 0 To MyOrderLines.DataGridView1.Rows.Count - 1
                If Trim(MyOrderLines.DataGridView1.Item(0, i).Value.ToString) = Trim(TextBox1.Text) Then
                    MyOrderLines.DataGridView1.CurrentCell = MyOrderLines.DataGridView1.Item(0, i)
                    System.Windows.Forms.Cursor.Current = Cursors.Default
                    Exit For
                End If
            Next
            System.Windows.Forms.Cursor.Current = Cursors.Default

            '--------складской или нет на данном складе--------------------
            MySQLStr = "SELECT COUNT(SC010300_2.SC01001) AS CC "
            MySQLStr = MySQLStr & "FROM SC010300 AS SC010300_2 WITH (NOLOCK) CROSS JOIN "
            MySQLStr = MySQLStr & "(SELECT SC23001 AS WH, CHARINDEX('1', SC23007) AS WHPos "
            MySQLStr = MySQLStr & "FROM SC230300 WITH (NOLOCK)"
            MySQLStr = MySQLStr & "WHERE (SC23006 = N'1') AND (SC23001 = N'" & Declarations.WHNum & "')) AS View_2 "
            MySQLStr = MySQLStr & "WHERE (SUBSTRING(SC010300_2.SC01128, View_2.WHPos, 1) = N'1') AND "
            MySQLStr = MySQLStr & "(SC010300_2.SC01001 = N'" & Trim(TextBox1.Text) & "') "

            InitMyConn(False)
            InitMyRec(False, MySQLStr)
            If Declarations.MyRec.EOF = True And Declarations.MyRec.BOF = True Then
                Label25.Text = ""
            Else
                Declarations.MyRec.MoveFirst()
                If Declarations.MyRec.Fields("CC").Value = 0 Then '--не складской
                    Label25.Text = ""
                Else
                    Label25.Text = "Складской ассортимент"
                End If
            End If
            trycloseMyRec()

            '--------кратность в упаковке----------------------------------
            MySQLStr = "SELECT SC01072 AS CC "
            MySQLStr = MySQLStr & "FROM SC010300 WITH (NOLOCK) "
            MySQLStr = MySQLStr & "WHERE (SC01001 = N'" & Trim(TextBox1.Text) & "') "
            InitMyConn(False)
            InitMyRec(False, MySQLStr)
            If Declarations.MyRec.EOF = True And Declarations.MyRec.BOF = True Then
                Label23.Text = ""
            Else
                Declarations.MyRec.MoveFirst()
                Label23.Text = Math.Round(Declarations.MyRec.Fields("CC").Value, 2)
            End If
            trycloseMyRec()

            '--------минимальное количество в заказе на закупку------------
            MySQLStr = "SELECT tbl_PurchasePriceHistory.MinQTY AS CC, LT "
            MySQLStr = MySQLStr & "FROM SC010300 WITH (NOLOCK) INNER JOIN "
            MySQLStr = MySQLStr & "tbl_PurchasePriceHistory WITH (NOLOCK) ON SC010300.SC01001 = tbl_PurchasePriceHistory.SC01001 "
            MySQLStr = MySQLStr & "WHERE (SC010300.SC01001 = N'" & Trim(TextBox1.Text) & "') AND "
            MySQLStr = MySQLStr & "(tbl_PurchasePriceHistory.DateTo = CONVERT(DATETIME, '9999-12-31 00:00:00', 102)) "
            InitMyConn(False)
            InitMyRec(False, MySQLStr)
            If Declarations.MyRec.EOF = True And Declarations.MyRec.BOF = True Then
                Label19.Text = ""
                Label21.Text = ""
            Else
                Declarations.MyRec.MoveFirst()
                Label19.Text = Math.Round(Declarations.MyRec.Fields("CC").Value, 2)
                If Declarations.MyRec.Fields("LT").Value = 0 Then
                    Label21.Text = ""
                Else
                    Label21.Text = Math.Round(Declarations.MyRec.Fields("LT").Value, 0)
                End If
            End If
            trycloseMyRec()
        End If
            MarginRecalc()
    End Sub

    Private Sub MarginRecalc()
        '////////////////////////////////////////////////////////////////////////////////////////
        '//
        '// пересчет маржи после изменения данных
        '//
        '////////////////////////////////////////////////////////////////////////////////////////
        Dim DiscCost As Double

        Try
            If CDbl(TextBox6.Text) = 0 Then
                TextBox7.Text = Math.Round((CDbl(TextBox4.Text) - CDbl(TextBox5.Text)) / CDbl(TextBox4.Text) * 100, 2)
            Else
                DiscCost = CDbl(TextBox4.Text) - CDbl(TextBox4.Text) * CDbl(TextBox6.Text) / 100
                TextBox7.Text = Math.Round((DiscCost - CDbl(TextBox5.Text)) / DiscCost * 100, 2)
            End If
        Catch ex As Exception
            'MsgBox(ex.ToString)
        End Try
    End Sub

    Private Sub TextBox3_Validated(ByVal sender As Object, ByVal e As System.EventArgs) Handles TextBox3.Validated
        '////////////////////////////////////////////////////////////////////////////////
        '//
        '// значение введено
        '//
        '////////////////////////////////////////////////////////////////////////////////

    End Sub

    Private Sub TextBox3_Validating(ByVal sender As Object, ByVal e As System.ComponentModel.CancelEventArgs) Handles TextBox3.Validating
        '////////////////////////////////////////////////////////////////////////////////
        '//
        '// проверка - числовое ли значение введено
        '//
        '////////////////////////////////////////////////////////////////////////////////
        Dim MyRez As Double
        Dim aa As New System.Globalization.NumberFormatInfo

        If Trim(TextBox3.Text) <> "" Then
            If InStr(TextBox3.Text, aa.CurrentInfo.NumberGroupSeparator) Then
                MsgBox("В поле ""Количество"" должно быть введено число", MsgBoxStyle.Critical, "Внимание!")
                e.Cancel = True
                Exit Sub
            Else
                Try
                    MyRez = TextBox3.Text
                Catch ex As Exception
                    MsgBox("В поле ""Количество"" должно быть введено число", MsgBoxStyle.Critical, "Внимание!")
                    e.Cancel = True
                    Exit Sub
                End Try
            End If
        End If
        e.Cancel = False
    End Sub

    Private Sub TextBox4_Validated(ByVal sender As Object, ByVal e As System.EventArgs) Handles TextBox4.Validated
        '////////////////////////////////////////////////////////////////////////////////
        '//
        '// значение введено
        '//
        '////////////////////////////////////////////////////////////////////////////////

        TextBox4Validation
    End Sub

    Public Sub TextBox4Validation()
        '////////////////////////////////////////////////////////////////////////////////
        '//
        '// значение введено
        '//
        '////////////////////////////////////////////////////////////////////////////////

        If TextBox4.Text <> "" Then
            TextBox4.Text = Math.Round(CDbl(TextBox4.Text), 2)
            If CheckDataFilingZ() = True Then
                MarginRecalc()
            End If
        End If
    End Sub

    Private Sub TextBox4_Validating(ByVal sender As Object, ByVal e As System.ComponentModel.CancelEventArgs) Handles TextBox4.Validating
        '////////////////////////////////////////////////////////////////////////////////
        '//
        '// проверка - числовое ли значение введено
        '//
        '////////////////////////////////////////////////////////////////////////////////
        Dim MyRez As Double
        Dim aa As New System.Globalization.NumberFormatInfo

        If Trim(TextBox4.Text) <> "" Then
            If InStr(TextBox4.Text, aa.CurrentInfo.NumberGroupSeparator) Then
                MsgBox("В поле ""Цена за единицу"" должно быть введено число", MsgBoxStyle.Critical, "Внимание!")
                e.Cancel = True
                Exit Sub
            Else
                Try
                    MyRez = TextBox4.Text
                Catch ex As Exception
                    MsgBox("В поле ""Цена за единицу"" должно быть введено число", MsgBoxStyle.Critical, "Внимание!")
                    e.Cancel = True
                    Exit Sub
                End Try
            End If
        End If
        e.Cancel = False
    End Sub

    Private Sub TextBox5_Validated(ByVal sender As Object, ByVal e As System.EventArgs) Handles TextBox5.Validated
        '////////////////////////////////////////////////////////////////////////////////
        '//
        '// значение введено
        '//
        '////////////////////////////////////////////////////////////////////////////////

        If TextBox5.Text <> "" Then
            TextBox5.Text = Math.Round(CDbl(TextBox5.Text), 2)
            If CheckDataFilingZ() = True Then
                MarginRecalc()
            End If
        End If
    End Sub


    Private Sub TextBox5_Validating(ByVal sender As Object, ByVal e As System.ComponentModel.CancelEventArgs) Handles TextBox5.Validating
        '////////////////////////////////////////////////////////////////////////////////
        '//
        '// проверка - числовое ли значение введено
        '//
        '////////////////////////////////////////////////////////////////////////////////
        Dim MyRez As Double
        Dim aa As New System.Globalization.NumberFormatInfo

        If Trim(TextBox5.Text) <> "" Then
            If InStr(TextBox5.Text, aa.CurrentInfo.NumberGroupSeparator) Then
                MsgBox("В поле ""Себестоимость единицы"" должно быть введено число", MsgBoxStyle.Critical, "Внимание!")
                e.Cancel = True
                Exit Sub
            Else
                Try
                    MyRez = TextBox5.Text
                Catch ex As Exception
                    MsgBox("В поле ""Себестоимость единицы"" должно быть введено число", MsgBoxStyle.Critical, "Внимание!")
                    e.Cancel = True
                    Exit Sub
                End Try
            End If
        End If
        e.Cancel = False
    End Sub

    Private Function CheckDataFiling() As Boolean
        '////////////////////////////////////////////////////////////////////////////////
        '//
        '// Проверка заполнения полей в окне
        '//
        '////////////////////////////////////////////////////////////////////////////////
        Dim MyRez As Object

        If Trim(TextBox1.Text) = "" Then
            MsgBox("Поле ""Код запаса"" должно быть заполнено", MsgBoxStyle.Critical, "Внимание")
            CheckDataFiling = False
            TextBox1.Select()
            Exit Function
        End If

        If Trim(TextBox2.Text) = "" Then
            MsgBox("Поле ""Название запаса"" должно быть заполнено", MsgBoxStyle.Critical, "Внимание")
            CheckDataFiling = False
            TextBox2.Select()
            Exit Function
        End If

        If Trim(TextBox3.Text) = "" Then
            MsgBox("Поле ""Количество"" должно быть заполнено", MsgBoxStyle.Critical, "Внимание")
            CheckDataFiling = False
            TextBox3.Select()
            Exit Function
        End If

        If Trim(TextBox4.Text) = "" Then
            MsgBox("Поле ""Цена за единицу"" должно быть заполнено", MsgBoxStyle.Critical, "Внимание")
            CheckDataFiling = False
            TextBox4.Select()
            Exit Function
        End If

        If Trim(TextBox5.Text) = "" Then
            MsgBox("Поле ""Себестоимость единицы"" должно быть заполнено", MsgBoxStyle.Critical, "Внимание")
            CheckDataFiling = False
            TextBox5.Select()
            Exit Function
        End If

        If Trim(TextBox6.Text) = "" Then
            MsgBox("Поле ""Скидка"" должно быть заполнено", MsgBoxStyle.Critical, "Внимание")
            CheckDataFiling = False
            TextBox6.Select()
            Exit Function
        End If

        If Trim(TextBox10.Text) = "" Then
            MsgBox("Поле ""Срок поставки (нед)"" должно быть заполнено", MsgBoxStyle.Critical, "Внимание")
            CheckDataFiling = False
            TextBox10.Select()
            Exit Function
        End If

        If Trim(TextBox11.Text) = "" Then
            MsgBox("Поле ""Срок доставки до клиента (нед)"" должно быть заполнено. Если товар есть в наличии на складе - то 0", MsgBoxStyle.Critical, "Внимание")
            CheckDataFiling = False
            TextBox11.Select()
            Exit Function
        End If

        If Math.Round(CDbl(TextBox11.Text), 1) > Math.Round(CDbl(TextBox10.Text), 1) Then
            MsgBox("Поле ""Срок доставки до клиента (нед)"" не может быть больше срока поставки клиенту", MsgBoxStyle.Critical, "Внимание")
            CheckDataFiling = False
            TextBox10.Select()
            Exit Function
        End If

        If Trim(TextBox15.Text) = "" Then
            MsgBox("Поле ""Код поставщика"" должно быть заполнено", MsgBoxStyle.Critical, "Внимание")
            CheckDataFiling = False
            Exit Function
        End If

        If Trim(TextBox16.Text) = "" Then
            MsgBox("Поле ""Поставщик"" должно быть заполнено", MsgBoxStyle.Critical, "Внимание")
            CheckDataFiling = False
            Exit Function
        End If

        'If DateTimePicker1.Value < Now() Then
        '    MsgBox("Дата поставки должна быть больше текущей", MsgBoxStyle.Critical, "Внимание!")
        '    DateTimePicker1.Select()
        '    CheckDataFiling = False
        '    Exit Function
        'End If

        'If Trim(TextBox8.Text) <> "" And CDbl(IIf(Trim(TextBox8.Text) = "", "0", TextBox8.Text)) < Declarations.MinMarginLevelManager Then
        'MyRez = MsgBox("Вы пытаетесь установить маржу меньшую, чем разрешено для данного клиента. Отменить ввод такой маржи?", vbYesNo, "Внимание!")
        'If MyRez = vbYes Then
        'TextBox8.Select()
        'CheckDataFiling = False
        'Exit Function
        'End If
        'End If

        'If Trim(TextBox8.Text) = "" And CDbl(TextBox7.Text) < Declarations.MinMarginLevelManager Then
        'MyRez = MsgBox("Вы пытаетесь ввести значения с маржой меньшей, чем разрешено для данного клиента. Отменить ввод таких данных?", vbYesNo, "Внимание!")
        'If MyRez = vbYes Then
        'TextBox4.Select()
        'CheckDataFiling = False
        'Exit Function
        'End If
        'End If

        CheckDataFiling = True
    End Function

    Private Function CheckDataFilingZ() As Boolean
        '////////////////////////////////////////////////////////////////////////////////
        '//
        '// Проверка заполнения полей в окне
        '//
        '////////////////////////////////////////////////////////////////////////////////


        If Trim(TextBox4.Text) = "" Then
            CheckDataFilingZ = False
            Exit Function
        End If

        If Trim(TextBox5.Text) = "" Then
            CheckDataFilingZ = False
            Exit Function
        End If

        If Trim(TextBox6.Text) = "" Then
            CheckDataFilingZ = False
            Exit Function
        End If

        CheckDataFilingZ = True
    End Function

    Private Sub TextBox6_Validated(ByVal sender As Object, ByVal e As System.EventArgs) Handles TextBox6.Validated
        '////////////////////////////////////////////////////////////////////////////////
        '//
        '// значение введено
        '//
        '////////////////////////////////////////////////////////////////////////////////

        If TextBox6.Text <> "" Then
            TextBox6.Text = Math.Round(CDbl(TextBox6.Text), 2)
            If CheckDataFilingZ() = True Then
                MarginRecalc()
            End If
        End If
    End Sub

    Private Sub TextBox6_Validating(ByVal sender As Object, ByVal e As System.ComponentModel.CancelEventArgs) Handles TextBox6.Validating
        '////////////////////////////////////////////////////////////////////////////////
        '//
        '// проверка - числовое ли значение введено
        '//
        '////////////////////////////////////////////////////////////////////////////////
        Dim MyRez As Double
        Dim aa As New System.Globalization.NumberFormatInfo

        If Trim(TextBox6.Text) <> "" Then
            If InStr(TextBox5.Text, aa.CurrentInfo.NumberGroupSeparator) Then
                MsgBox("В поле ""Скидка"" должно быть введено число", MsgBoxStyle.Critical, "Внимание!")
                e.Cancel = True
                Exit Sub
            Else
                Try
                    MyRez = TextBox6.Text
                Catch ex As Exception
                    MsgBox("В поле ""Скидка"" должно быть введено число", MsgBoxStyle.Critical, "Внимание!")
                    e.Cancel = True
                    Exit Sub
                End Try
                If MyRez >= 100 Or MyRez < 0 Then
                    MsgBox("В поле ""Скидка"" должно быть введено число больше или равно 0 и меньше 100", MsgBoxStyle.Critical, "Внимание!")
                    e.Cancel = True
                    Exit Sub
                End If
            End If
        End If
        e.Cancel = False
    End Sub

    Private Sub TextBox8_Validated(ByVal sender As Object, ByVal e As System.EventArgs) Handles TextBox8.Validated
        '////////////////////////////////////////////////////////////////////////////////
        '//
        '// значение введено
        '//
        '////////////////////////////////////////////////////////////////////////////////

        If Trim(TextBox8.Text) <> "" Then
            TextBox8.Text = Math.Round(CDbl(TextBox8.Text), 2)
        End If
    End Sub

    Private Sub TextBox8_Validating(ByVal sender As Object, ByVal e As System.ComponentModel.CancelEventArgs) Handles TextBox8.Validating
        '////////////////////////////////////////////////////////////////////////////////
        '//
        '// проверка - числовое ли значение введено
        '//
        '////////////////////////////////////////////////////////////////////////////////
        Dim MyRez As Double
        Dim aa As New System.Globalization.NumberFormatInfo

        If TextBox8.Text <> "" Then
            If InStr(TextBox5.Text, aa.CurrentInfo.NumberGroupSeparator) Then
                MsgBox("В поле ""Новая маржа %"" должно быть введено число", MsgBoxStyle.Critical, "Внимание!")
                e.Cancel = True
                Exit Sub
            Else
                Try
                    MyRez = TextBox8.Text
                Catch ex As Exception
                    MsgBox("В поле ""Новая маржа %"" должно быть введено число", MsgBoxStyle.Critical, "Внимание!")
                    e.Cancel = True
                    Exit Sub
                End Try
                If MyRez >= 100 Or MyRez < 0 Then
                    MsgBox("В поле ""Новая маржа %"" должно быть введено число больше или равно 0 и меньше 100", MsgBoxStyle.Critical, "Внимание!")
                    e.Cancel = True
                    Exit Sub
                End If
            End If
        End If
        e.Cancel = False
    End Sub

    Private Sub TextBox10_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles TextBox10.KeyDown
        '////////////////////////////////////////////////////////////////////////////////
        '//
        '// переход на следующее поле по Enter
        '//
        '////////////////////////////////////////////////////////////////////////////////

        If e.KeyData = Keys.Enter Then
            Me.SelectNextControl(sender, True, True, True, False)
        End If
    End Sub

    Private Sub TextBox10_Validating(ByVal sender As Object, ByVal e As System.ComponentModel.CancelEventArgs) Handles TextBox10.Validating
        '////////////////////////////////////////////////////////////////////////////////
        '//
        '// проверка - числовое ли знечение введено
        '//
        '////////////////////////////////////////////////////////////////////////////////
        Dim MyRez As Double
        Dim aa As New System.Globalization.NumberFormatInfo

        If Trim(TextBox10.Text) <> "" Then
            If InStr(TextBox10.Text, aa.CurrentInfo.NumberGroupSeparator) Then
                MsgBox("В поле ""Срок поставки (нед)"" должно быть введено число. Если товар есть в наличии на складе - то 0", MsgBoxStyle.Critical, "Внимание!")
                e.Cancel = True
                Exit Sub
            Else
                Try
                    MyRez = TextBox10.Text
                    If MyRez < 0 Then
                        MsgBox("В поле ""Срок поставки (нед)"" должно быть введено число, большее или равное 0", MsgBoxStyle.Critical, "Внимание!")
                        e.Cancel = True
                        Exit Sub
                    End If
                Catch ex As Exception
                    MsgBox("В поле ""Срок поставки (нед)"" должно быть введено число. Если товар есть в наличии на складе - то 0", MsgBoxStyle.Critical, "Внимание!")
                    e.Cancel = True
                    Exit Sub
                End Try
            End If
        End If
        e.Cancel = False
    End Sub

    Private Sub Button3_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button3.Click
        '////////////////////////////////////////////////////////////////////////////////
        '//
        '// Пересчет цены для подгона под указанное значение маржи
        '//
        '////////////////////////////////////////////////////////////////////////////////
        Dim MyDBL As Double
        Dim DiscCost As Double

        If Trim(TextBox8.Text) <> "" Then
            If CheckDataFiling() = True Then
                MyDBL = CDbl(TextBox8.Text)

                '---Обновление информации по продукту
                If CDbl(TextBox6.Text) = 0 Then '---пересчитываем только маржу
                    TextBox4.Text = Math.Round(CDbl(TextBox5.Text) * 100 / (100 - MyDBL), 3)
                    TextBox4Validation()
                Else '---пересчитываем скидку и маржу
                    DiscCost = Math.Round(CDbl(TextBox5.Text) * 100 / (100 - MyDBL), 3)
                    TextBox4.Text = Math.Round(DiscCost * 100 / (100 - CDbl(TextBox6.Text)), 2)
                    TextBox4Validation()
                End If
                TextBox8.Text = ""
            End If
        Else
            MsgBox("В поле ""Новая маржа %"" должно быть введено число больше или равно 0 и меньше 100", MsgBoxStyle.Critical, "Внимание!")
        End If
    End Sub

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        '////////////////////////////////////////////////////////////////////////////////
        '//
        '// Выход из окна с сохранением новых значений
        '//
        '////////////////////////////////////////////////////////////////////////////////

        If CheckDataFiling() = True Then
            Declarations.MySuccess = True                               'Успешность выполнения операции
            Declarations.MyItemID = Trim(TextBox1.Text)                 'код запаса
            Declarations.MyItemName = Trim(TextBox2.Text)               'имя запаса
            Declarations.MyQty = Trim(TextBox3.Text)                    'количество заказанного
            Declarations.MyUOM = ComboBox1.SelectedValue                'код единицы измерения
            Declarations.MySum = Trim(TextBox4.Text)                    'цена заказанного
            Declarations.MySS = Trim(TextBox5.Text) * Declarations.CurrencyValue 'себестоимость
            Declarations.MyDiscount = Trim(TextBox6.Text)               'Скидка
            Declarations.WeekQTY = Math.Round(CDbl(TextBox10.Text), 1)  'Срок поставки
            Declarations.DelWeekQTY = Math.Round(CDbl(TextBox11.Text), 1)  'Срок доставки до клиента
            Declarations.MyItemSuppID = Trim(TextBox14.Text)
            Declarations.MySuppID = Trim(TextBox15.Text)
            Declarations.MySuppName = Trim(TextBox16.Text)
            'If CheckBox1.Checked = True Then
            '    Declarations.DeliveryDateFlag = 1
            'Else
            '    Declarations.DeliveryDateFlag = 0
            'End If
            Me.Close()
        End If
    End Sub

    Private Sub Button4_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button4.Click
        '////////////////////////////////////////////////////////////////////////////////
        '//
        '// Открытие окна со списком запасов для выбора
        '//
        '////////////////////////////////////////////////////////////////////////////////

        MyItemSelect = New ItemSelect
        MyItemSelect.MySrcWin = "EditInOrder"
        MyItemSelect.ShowDialog()
    End Sub

    Private Sub Button5_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button5.Click
        '////////////////////////////////////////////////////////////////////////////////
        '//
        '// Выход из окна с сохранением новых значений Сроков поставки
        '//
        '////////////////////////////////////////////////////////////////////////////////
        Dim MySQLStr As String

        If CheckDataFilling1() = True Then
            MySQLStr = "UPDATE tbl_OR030300 "
            MySQLStr = MySQLStr & "Set WeekQTY = " & Replace(RTrim(LTrim(TextBox9.Text)), ",", ".") & ", "
            MySQLStr = MySQLStr & "DelWeekQTY = " & Replace(RTrim(LTrim(TextBox12.Text)), ",", ".") & " "
            MySQLStr = MySQLStr & "WHERE (OR03001 = N'" & Trim(MyOrderLines.Label6.Text) & "') "
            InitMyConn(False)
            Declarations.MyConn.Execute(MySQLStr)
            Me.Close()
        End If
    End Sub

    Private Function CheckDataFilling1() As Boolean
        '////////////////////////////////////////////////////////////////////////////////
        '//
        '// проверка ввода необходимых данных
        '//
        '////////////////////////////////////////////////////////////////////////////////

        If Trim(TextBox9.Text) = "" Then
            MsgBox("Поле ""Выставить всему заказу срок поставки"" должно быть заполнено", MsgBoxStyle.Critical, "Внимание")
            TextBox9.Select()
            CheckDataFilling1 = False
            Exit Function
        End If

        If Trim(TextBox12.Text) = "" Then
            MsgBox("Поле ""Срок доставки до клиента (нед)"" должно быть заполнено. Если товар есть в наличии на складе - то 0", MsgBoxStyle.Critical, "Внимание")
            CheckDataFilling1 = False
            TextBox12.Select()
            Exit Function
        End If

        If Math.Round(CDbl(TextBox12.Text), 1) > Math.Round(CDbl(TextBox9.Text), 1) Then
            MsgBox("Поле ""Срок доставки до клиента (нед)"" не может быть больше срока поставки клиенту", MsgBoxStyle.Critical, "Внимание")
            CheckDataFilling1 = False
            TextBox9.Select()
            Exit Function
        End If

        CheckDataFilling1 = True
    End Function

    Private Sub TextBox9_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles TextBox9.KeyDown
        '////////////////////////////////////////////////////////////////////////////////
        '//
        '// переход на следующее поле по Enter
        '//
        '////////////////////////////////////////////////////////////////////////////////

        If e.KeyData = Keys.Enter Then
            Me.SelectNextControl(sender, True, True, True, False)
        End If
    End Sub

    Private Sub TextBox9_Validating(ByVal sender As Object, ByVal e As System.ComponentModel.CancelEventArgs) Handles TextBox9.Validating
        '////////////////////////////////////////////////////////////////////////////////
        '//
        '// проверка - числовое ли знечение введено
        '//
        '////////////////////////////////////////////////////////////////////////////////
        Dim MyRez As Double
        Dim aa As New System.Globalization.NumberFormatInfo

        If Trim(TextBox9.Text) <> "" Then
            If InStr(TextBox9.Text, aa.CurrentInfo.NumberGroupSeparator) Then
                MsgBox("В поле ""Выставить всему заказу срок поставки"" должно быть введено число. Если товар есть в наличии на складе - то срок доставки до клиента; Если товар есть в наличии на складе и доставки нет - то 0", MsgBoxStyle.Critical, "Внимание!")
                e.Cancel = True
                Exit Sub
            Else
                Try
                    MyRez = TextBox9.Text
                    If MyRez < 0 Then
                        MsgBox("В поле ""Выставить всему заказу срок поставки"" должно быть введено число, большее или равное 0", MsgBoxStyle.Critical, "Внимание!")
                        e.Cancel = True
                        Exit Sub
                    End If
                Catch ex As Exception
                    MsgBox("В поле ""Выставить всему заказу срок поставки"" должно быть введено число. Если товар есть в наличии на складе - то срок доставки до клиента; Если товар есть в наличии на складе и доставки нет - то 0", MsgBoxStyle.Critical, "Внимание!")
                    e.Cancel = True
                    Exit Sub
                End Try
            End If
        End If
        e.Cancel = False
    End Sub

    Private Sub TextBox11_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles TextBox11.KeyDown
        '////////////////////////////////////////////////////////////////////////////////
        '//
        '// переход на следующее поле по Enter
        '//
        '////////////////////////////////////////////////////////////////////////////////

        If e.KeyData = Keys.Enter Then
            Me.SelectNextControl(sender, True, True, True, False)
        End If
    End Sub

    Private Sub TextBox12_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles TextBox12.KeyDown
        '////////////////////////////////////////////////////////////////////////////////
        '//
        '// переход на следующее поле по Enter
        '//
        '////////////////////////////////////////////////////////////////////////////////

        If e.KeyData = Keys.Enter Then
            Me.SelectNextControl(sender, True, True, True, False)
        End If
    End Sub

    Private Sub TextBox11_Validating(ByVal sender As Object, ByVal e As System.ComponentModel.CancelEventArgs) Handles TextBox11.Validating
        '////////////////////////////////////////////////////////////////////////////////
        '//
        '// проверка - числовое ли знечение введено
        '//
        '////////////////////////////////////////////////////////////////////////////////
        Dim MyRez As Double
        Dim aa As New System.Globalization.NumberFormatInfo

        If Trim(TextBox11.Text) <> "" Then
            If InStr(TextBox11.Text, aa.CurrentInfo.NumberGroupSeparator) Then
                MsgBox("В поле ""Срок доставки до клиента (нед)"" должно быть введено число. Если доставки до клиента нет (самовывоз) - то 0 - то 0", MsgBoxStyle.Critical, "Внимание!")
                e.Cancel = True
                Exit Sub
            Else
                Try
                    MyRez = TextBox11.Text
                    If MyRez < 0 Then
                        MsgBox("В поле ""Срок доставки до клиента (нед)"" должно быть введено число, большее или равное 0", MsgBoxStyle.Critical, "Внимание!")
                        e.Cancel = True
                        Exit Sub
                    End If
                Catch ex As Exception
                    MsgBox("В поле ""Срок доставки до клиента (нед)"" должно быть введено число. Если доставки до клиента нет (самовывоз) - то 0 - то 0", MsgBoxStyle.Critical, "Внимание!")
                    e.Cancel = True
                    Exit Sub
                End Try
            End If
        End If
        e.Cancel = False
    End Sub

    Private Sub TextBox12_Validating(ByVal sender As Object, ByVal e As System.ComponentModel.CancelEventArgs) Handles TextBox12.Validating
        '////////////////////////////////////////////////////////////////////////////////
        '//
        '// проверка - числовое ли знечение введено
        '//
        '////////////////////////////////////////////////////////////////////////////////
        Dim MyRez As Double
        Dim aa As New System.Globalization.NumberFormatInfo

        If Trim(TextBox12.Text) <> "" Then
            If InStr(TextBox12.Text, aa.CurrentInfo.NumberGroupSeparator) Then
                MsgBox("В поле ""Срок доставки до клиента (нед)"" должно быть введено число. Если доставки до клиента нет (самовывоз) - то 0 - то 0", MsgBoxStyle.Critical, "Внимание!")
                e.Cancel = True
                Exit Sub
            Else
                Try
                    MyRez = TextBox12.Text
                    If MyRez < 0 Then
                        MsgBox("В поле ""Срок доставки до клиента (нед)"" должно быть введено число, большее или равное 0", MsgBoxStyle.Critical, "Внимание!")
                        e.Cancel = True
                        Exit Sub
                    End If
                Catch ex As Exception
                    MsgBox("В поле ""Срок доставки до клиента (нед)"" должно быть введено число. Если доставки до клиента нет (самовывоз) - то 0 - то 0", MsgBoxStyle.Critical, "Внимание!")
                    e.Cancel = True
                    Exit Sub
                End Try
            End If
        End If
        e.Cancel = False
    End Sub

    Private Sub Button8_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button8.Click
        '////////////////////////////////////////////////////////////////////////////////
        '//
        '// Уведомление о неверной картинке
        '//
        '////////////////////////////////////////////////////////////////////////////////
        Dim myRez As MsgBoxResult

        myRez = MsgBox("Вы уверены?", MsgBoxStyle.YesNo, "Внимание!")
        If myRez = MsgBoxResult.Yes Then
            SendAddInfoReminder(Trim(TextBox1.Text), Declarations.SalesmanName, 0)
            MsgBox("Уведомление о неверной картинке отправлено", MsgBoxStyle.Information, "Внимание!")
        End If
    End Sub

    Private Sub Button7_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button7.Click
        '////////////////////////////////////////////////////////////////////////////////
        '//
        '// Уведомление о неверном названии
        '//
        '////////////////////////////////////////////////////////////////////////////////
        Dim myRez As MsgBoxResult

        myRez = MsgBox("Вы уверены?", MsgBoxStyle.YesNo, "Внимание!")
        If myRez = MsgBoxResult.Yes Then
            SendAddInfoReminder(Trim(TextBox1.Text), Declarations.SalesmanName, 1)
            MsgBox("Уведомление о неверном названии отправлено", MsgBoxStyle.Information, "Внимание!")
        End If
    End Sub

    Private Sub Button6_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button6.Click
        '////////////////////////////////////////////////////////////////////////////////
        '//
        '// Уведомление о неверном описании
        '//
        '////////////////////////////////////////////////////////////////////////////////
        Dim myRez As MsgBoxResult

        myRez = MsgBox("Вы уверены?", MsgBoxStyle.YesNo, "Внимание!")
        If myRez = MsgBoxResult.Yes Then
            SendAddInfoReminder(Trim(TextBox1.Text), Declarations.SalesmanName, 2)
            MsgBox("Уведомление о неверном описании отправлено", MsgBoxStyle.Information, "Внимание!")
        End If
    End Sub

    Private Sub TextBox14_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles TextBox14.KeyDown
        '////////////////////////////////////////////////////////////////////////////////
        '//
        '// переход на следующее поле по Enter
        '//
        '////////////////////////////////////////////////////////////////////////////////

        If e.KeyData = Keys.Enter Then
            Me.SelectNextControl(sender, True, True, True, False)
        End If
    End Sub

    Private Sub TextBox14_Validated(ByVal sender As Object, ByVal e As System.EventArgs) Handles TextBox14.Validated
        '////////////////////////////////////////////////////////////////////////////////
        '//
        '// После ввода кода - проверка, есть ли такой запас в базе. Если есть - 
        '// подтягиваем его значения. Если нет - даем ввести свои
        '////////////////////////////////////////////////////////////////////////////////
        If TextBox14.Modified = True Then
            TextBox14Validation()
        End If
    End Sub

    Public Sub TextBox14Validation()
        '////////////////////////////////////////////////////////////////////////////////
        '//
        '// После ввода кода товара поставщика- проверка, есть ли такой запас в базе. Если есть - 
        '// подтягиваем его значения. Если нет - даем ввести свои
        '////////////////////////////////////////////////////////////////////////////////
        Dim MySQLStr As String                          'рабочая строка
        Dim RecQTY As Double                            '--Количество записей
        Dim MyRez As Object
        'Dim FoundFlag As Integer                        'флаг - сколько найдено

        If Trim(TextBox14.Text) <> "" Then
            MySQLStr = "exec spp_SalesWorkplace4_SpecificationFilling "
            MySQLStr = MySQLStr & "N'', "                                   '--Код товара в Scala
            MySQLStr = MySQLStr & "N'" & Trim(TextBox14.Text) & "'  "        '--Код товара поставщика

            InitMyConn(False)
            InitMyRec(False, MySQLStr)

            If Declarations.MyRec.EOF = True And Declarations.MyRec.BOF = True Then
                '--ничего не найдено
                '---очищаем все поля 
                Label3.Text = "Рекомендованная цена и себестоимость этого запаса на основе прайс - листа на закупку"
                Label3.ForeColor = Color.Green
                MyRez = MsgBox("Очистить поля от предыдущих значений?", MsgBoxStyle.YesNo, "Внимание!")
                If MyRez = vbYes Then
                    TextBox1.Text = "Unknown"                      '--код Scala
                    TextBox2.Text = ""
                    TextBox2.ReadOnly = False
                    ComboBox1.Text = ""
                    ComboBox1.Enabled = True
                    TextBox3.Text = "1"
                    TextBox4.Text = ""
                    TextBox5.Text = ""
                    TextBox5.ReadOnly = False
                    TextBox15.Text = ""
                    TextBox15.Enabled = True
                    TextBox15.BackColor = Color.FromName("Window")
                    Button9.Enabled = True
                    TextBox16.Text = ""
                    TextBox16.Enabled = False
                    TextBox16.BackColor = Color.FromName("ButtonFace")
                    Label25.Text = ""
                    Label23.Text = ""
                    Label19.Text = ""
                    Label21.Text = ""
                    TextBox10.Text = ""
                    TextBox11.Text = "0"
                Else
                    TextBox2.ReadOnly = False
                    ComboBox1.Enabled = True
                    TextBox5.ReadOnly = False
                    TextBox15.Enabled = True
                    TextBox15.BackColor = Color.FromName("Window")
                    Button9.Enabled = True
                    TextBox16.Enabled = False
                    TextBox16.BackColor = Color.FromName("ButtonFace")
                End If
            Else
                Declarations.MyRec.MoveLast()
                RecQTY = Declarations.MyRec.RecordCount
                Declarations.MyRec.MoveFirst()
                If RecQTY = 1 Then      '--только одна запись, сразу заносим
                    '---Заносим найденные значения полей в форму
                    If Trim(Declarations.MyRec.Fields("SC01001").Value) = "" Then       '--из прайс листа поставщика
                        '---Заносим найденные значения полей в форму
                        Label3.Text = "Рекомендованная цена и себестоимость этого запаса на основе прайс - листа на закупку"
                        Label3.ForeColor = Color.Green
                        If Declarations.MyRec.Fields("SC01001").Value.trim() = "" Then
                            TextBox1.Text = "Unknown"
                        Else
                            TextBox1.Text = Declarations.MyRec.Fields("SC01001").Value
                        End If
                        TextBox2.Text = Declarations.MyRec.Fields("Name").Value
                        TextBox2.ReadOnly = False
                        ComboBox1.Text = Declarations.MyRec.Fields("SC01135").Value
                        ComboBox1.Enabled = True
                        TextBox4.Text = Declarations.MyRec.Fields("Price").Value / Declarations.CurrencyValue
                        TextBox5.Text = Declarations.MyRec.Fields("PriCost").Value / Declarations.CurrencyValue
                        TextBox5.ReadOnly = False
                        TextBox15.Text = Declarations.MyRec.Fields("SuppID").Value
                        TextBox15.Enabled = True
                        TextBox15.BackColor = Color.FromName("Window")
                        Button9.Enabled = True
                        TextBox16.Text = Declarations.MyRec.Fields("SuppName").Value
                        TextBox16.Enabled = False
                        TextBox16.BackColor = Color.FromName("ButtonFace")
                        Label25.Text = ""
                        Label23.Text = ""
                        Label19.Text = Math.Round(Declarations.MyRec.Fields("MinQTY").Value, 2)
                        Label21.Text = Math.Round(Declarations.MyRec.Fields("LT").Value, 0)
                        TextBox10.Text = ""
                        TextBox11.Text = "1"
                        TextBox6.Text = 0
                        TextBox3.Text = 1
                        TextBox2.Text = ""
                        LoadItemAddInfo1(Trim(TextBox1.Text))
                        trycloseMyRec()
                    Else                                                '--из Scala
                        TextBox1.Text = Declarations.MyRec.Fields("SC01001").Value
                        TextBox1Validation()
                        LoadItemAddInfo1(Trim(TextBox1.Text))
                        trycloseMyRec()
                    End If
                Else
                    MySelectItemBySuppCode = New SelectItemBySuppCode
                    MySelectItemBySuppCode.MyItemSuppCode = Trim(Trim(TextBox14.Text))
                    MySelectItemBySuppCode.MyWindowFrom = "EditInOrder"
                    MySelectItemBySuppCode.ShowDialog()
                End If
            End If
            Else
                '---очищаем все поля 
                TextBox1.Text = "Unknown"                      '--код Scala
                TextBox2.Text = ""
                TextBox2.ReadOnly = False
                ComboBox1.Text = ""
                ComboBox1.Enabled = True
                TextBox3.Text = "1"
                TextBox4.Text = ""
                TextBox5.Text = ""
                TextBox5.ReadOnly = False
                TextBox15.Text = ""
                TextBox15.Enabled = True
                TextBox15.BackColor = Color.FromName("Window")
                Button9.Enabled = True
                TextBox16.Text = ""
                TextBox16.Enabled = False
                TextBox16.BackColor = Color.FromName("ButtonFace")
                Label25.Text = ""
                Label23.Text = ""
                Label19.Text = ""
                Label21.Text = ""
                TextBox10.Text = ""
                TextBox11.Text = "0"
            End If

    End Sub

    Private Sub Button9_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button9.Click
        '////////////////////////////////////////////////////////////////////////////////
        '//
        '// Выбор кода и названия поставщика из списка существующих в Scala
        '//
        '////////////////////////////////////////////////////////////////////////////////

        MySupplierSelect = New SupplierSelect
        MySupplierSelect.MySrcWin = "EditInOrder"
        MySupplierSelect.ShowDialog()
    End Sub

    Private Sub TextBox15_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles TextBox15.KeyDown
        '////////////////////////////////////////////////////////////////////////////////
        '//
        '// переход на следующее поле по Enter
        '//
        '////////////////////////////////////////////////////////////////////////////////

        If e.KeyData = Keys.Enter Then
            Me.SelectNextControl(sender, True, True, True, False)
        End If
    End Sub

    Private Sub TextBox15_Validated(ByVal sender As Object, ByVal e As System.EventArgs) Handles TextBox15.Validated
        '////////////////////////////////////////////////////////////////////////////////
        '//
        '// После ввода кода - проверка, есть ли такой поставщик в базе. Если есть - 
        '// подтягиваем его значения. Если нет - даем ввести свои
        '////////////////////////////////////////////////////////////////////////////////

        If TextBox15.Modified = True Then
            TextBox15Validation()
        End If
    End Sub

    Public Sub TextBox15Validation()
        '////////////////////////////////////////////////////////////////////////////////
        '//
        '// После ввода кода - проверка, есть ли такой поставщик в базе. Если есть - 
        '// подтягиваем его значения. Если нет - даем ввести свои
        '////////////////////////////////////////////////////////////////////////////////
        Dim MySQLStr As String                        'рабочая строка

        MySQLStr = "SELECT PL01001, PL01002 "
        MySQLStr = MySQLStr & "FROM PL010300 "
        MySQLStr = MySQLStr & "WHERE (PL01001 = N'" & TextBox15.Text & "')"
        InitMyConn(False)
        InitMyRec(False, MySQLStr)
        If Declarations.MyRec.EOF = True And Declarations.MyRec.BOF = True Then
            '---такого поставщика нет в БД
            TextBox16.Text = ""
            TextBox16.Enabled = True
            TextBox16.BackColor = Color.FromName("Window")
        Else
            '---такой поставщик в БД есть
            TextBox16.Text = Declarations.MyRec.Fields("PL01002").Value
            TextBox16.Enabled = False
            TextBox16.BackColor = Color.FromName("ButtonFace")
        End If

    End Sub
End Class