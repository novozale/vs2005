Public Class AddToOrder

    Private Sub Button2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button2.Click
        '////////////////////////////////////////////////////////////////////////////////
        '//
        '// Закрытие окна без добавления запаса в заказ
        '//
        '////////////////////////////////////////////////////////////////////////////////

        Declarations.MySuccess = False
        Me.Close()
    End Sub

    Private Sub AddToOrder_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        '////////////////////////////////////////////////////////////////////////////////
        '//
        '// Загрузка данных в форму
        '//
        '////////////////////////////////////////////////////////////////////////////////
        Dim MySQLStr As String                        'рабочая строка
        Dim MyAdapter As SqlClient.SqlDataAdapter     '
        Dim MyDs As New DataSet                       '
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


        '---ID запаса, имя запаса, кол - во, ед. измерения, цена (прайс), код поставщика, поставщик
        TextBox2.Text = MyOrderLines.DataGridView1.SelectedRows.Item(0).Cells(1).Value.ToString
        TextBox1.Text = MyOrderLines.DataGridView1.SelectedRows.Item(0).Cells(2).Value.ToString
        TextBox1.ReadOnly = True
        TextBox3.Text = 1
        ComboBox1.Text = MyOrderLines.DataGridView1.SelectedRows.Item(0).Cells(5).Value.ToString
        ComboBox1.Enabled = False
        If MyOrderLines.DataGridView1.SelectedRows.Item(0).Cells(3).Value.ToString = 0 Then
            TextBox4.Text = ""
        Else
            TextBox4.Text = CStr(Math.Round(CDbl(MyOrderLines.DataGridView1.SelectedRows.Item(0).Cells(3).Value.ToString), 2))
        End If
        TextBox13.Text = MyOrderLines.DataGridView1.SelectedRows.Item(0).Cells(9).Value.ToString
        TextBox13.Enabled = False
        TextBox13.BackColor = Color.FromName("ButtonFace")
        Button7.Enabled = False
        TextBox14.Text = MyOrderLines.DataGridView1.SelectedRows.Item(0).Cells(10).Value.ToString
        TextBox14.Enabled = False
        TextBox14.BackColor = Color.FromName("ButtonFace")

        '--- себестоимость
        'MySQLStr = "SELECT SC030300.SC03057 AS SS, "
        'MySQLStr = MySQLStr & "SC010300.SC01053 AS CSS "
        'MySQLStr = MySQLStr & "FROM SC030300 WITH (NOLOCK) INNER JOIN "
        'MySQLStr = MySQLStr & "SC010300 ON SC030300.SC03001 = SC010300.SC01001 "
        'MySQLStr = MySQLStr & "WHERE (SC030300.SC03001 = N'" & Trim(MyOrderLines.DataGridView1.SelectedRows.Item(0).Cells(0).Value.ToString) & "') AND "
        'MySQLStr = MySQLStr & "(SC030300.SC03002 = N'" & Declarations.WHNum & "')"

        MySQLStr = "SELECT SC01053 AS CSS "
        MySQLStr = MySQLStr & "FROM SC010300 "
        MySQLStr = MySQLStr & "WHERE (SC01001 = N'" & Trim(MyOrderLines.DataGridView1.SelectedRows.Item(0).Cells(1).Value.ToString) & "') "

        InitMyConn(False)
        InitMyRec(False, MySQLStr)
        If Declarations.MyRec.EOF = True And Declarations.MyRec.BOF = True Then
            TextBox5.Text = ""
            TextBox5.ReadOnly = False
            Label3.Text = "Рекомендованная цена и себестоимость  для этого запаса должны быть определены самостоятельно"
            Label3.ForeColor = Color.Red
        Else
            Declarations.MyRec.MoveFirst()
            Declarations.MySS = Declarations.MyRec.Fields("CSS").Value
            trycloseMyRec()
            If Declarations.MySS <> 0 Then
                TextBox5.Text = Math.Round(Declarations.MySS / Declarations.CurrencyValue, 2)
                TextBox5.ReadOnly = True
                Label3.Text = "Рекомендованная цена и себестоимость этого запаса на основе прайс - листа на закупку"
                Label3.ForeColor = Color.Green
            Else
                TextBox5.Text = ""
                TextBox5.ReadOnly = False
                Label3.Text = "Рекомендованная цена и себестоимость  для этого запаса должны быть определены самостоятельно"
                Label3.ForeColor = Color.Red
            End If
        End If

        '--------складской или нет на данном складе--------------------
        MySQLStr = "SELECT COUNT(SC010300_2.SC01001) AS CC "
        MySQLStr = MySQLStr & "FROM SC010300 AS SC010300_2 WITH (NOLOCK) CROSS JOIN "
        MySQLStr = MySQLStr & "(SELECT SC23001 AS WH, CHARINDEX('1', SC23007) AS WHPos "
        MySQLStr = MySQLStr & "FROM SC230300 WITH (NOLOCK)"
        MySQLStr = MySQLStr & "WHERE (SC23006 = N'1') AND (SC23001 = N'" & Declarations.WHNum & "')) AS View_2 "
        MySQLStr = MySQLStr & "WHERE (SUBSTRING(SC010300_2.SC01128, View_2.WHPos, 1) = N'1') AND "
        MySQLStr = MySQLStr & "(SC010300_2.SC01001 = N'" & Trim(MyOrderLines.DataGridView1.SelectedRows.Item(0).Cells(1).Value.ToString) & "') "

        InitMyConn(False)
        InitMyRec(False, MySQLStr)
        If Declarations.MyRec.EOF = True And Declarations.MyRec.BOF = True Then
            Label16.Text = ""
        Else
            Declarations.MyRec.MoveFirst()
            If Declarations.MyRec.Fields("CC").Value = 0 Then '--не складской
                Label16.Text = ""
            Else
                Label16.Text = "Складской ассортимент"
            End If
        End If
        trycloseMyRec()

        '--------кратность в упаковке----------------------------------
        MySQLStr = "SELECT SC01072 AS CC "
        MySQLStr = MySQLStr & "FROM SC010300 WITH (NOLOCK) "
        MySQLStr = MySQLStr & "WHERE (SC01001 = N'" & Trim(MyOrderLines.DataGridView1.SelectedRows.Item(0).Cells(1).Value.ToString) & "') "
        InitMyConn(False)
        InitMyRec(False, MySQLStr)
        If Declarations.MyRec.EOF = True And Declarations.MyRec.BOF = True Then
            Label17.Text = ""
        Else
            Declarations.MyRec.MoveFirst()
            Label17.Text = Math.Round(Declarations.MyRec.Fields("CC").Value, 2)
        End If
        trycloseMyRec()

        '--------минимальное количество в заказе на закупку------------
        MySQLStr = "SELECT tbl_PurchasePriceHistory.MinQTY AS CC, LT "
        MySQLStr = MySQLStr & "FROM SC010300 WITH (NOLOCK) INNER JOIN "
        MySQLStr = MySQLStr & "tbl_PurchasePriceHistory WITH (NOLOCK) ON SC010300.SC01001 = tbl_PurchasePriceHistory.SC01001 "
        MySQLStr = MySQLStr & "WHERE (SC010300.SC01001 = N'" & Trim(MyOrderLines.DataGridView1.SelectedRows.Item(0).Cells(1).Value.ToString) & "') AND "
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
        MySQLStr = "SELECT ISNULL(MAX(WeekQTY), 1) AS CC "
        MySQLStr = MySQLStr & "FROM tbl_OR030300 "
        MySQLStr = MySQLStr & "WHERE (OR03001 = N'" & Trim(MyOrderLines.Label6.Text) & "') "
        InitMyConn(False)
        InitMyRec(False, MySQLStr)
        If Declarations.MyRec.EOF = True And Declarations.MyRec.BOF = True Then
            TextBox6.Text = 1
        Else
            'TextBox6.Text = Format(Declarations.MyRec.Fields("CC").Value, "#####0" & aa.CurrentInfo.NumberDecimalSeparator & "##")
            TextBox6.Text = Format(Declarations.MyRec.Fields("CC").Value, "#####0.##")
        End If
        trycloseMyRec()

        '---предполагаемый срок доставки строки до клиента
        TextBox7.Text = Format(0, "#####0.##")
        If Declarations.IsSelfDelivery = 1 Then
            TextBox7.Enabled = False
        Else
            TextBox7.Enabled = True
        End If

        '---Картинка, название для WEB, описание для WEB
        LoadItemAddInfo(Trim(MyOrderLines.DataGridView1.SelectedRows.Item(0).Cells(1).Value.ToString))

        TextBox3.Select()
    End Sub

    Public Sub LoadItemAddInfo(ByVal MyItemID As String)
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
            TextBox8.Text = ""
            RichTextBox1.Text = ""
            Button3.Enabled = False
            Button4.Enabled = False
            Button5.Enabled = False
            trycloseMyRec()
        Else
            Try
                Dim ms As New IO.MemoryStream(CType(Declarations.MyRec.Fields("Picture").Value, Byte()))
                Dim picture As Image

                picture = Image.FromStream(ms)
                PictureBox1.Image = picture
            Catch ex As Exception
                PictureBox1.Image = Nothing
            End Try
            TextBox8.Text = Declarations.MyRec.Fields("WEBName").Value
            RichTextBox1.Text = Declarations.MyRec.Fields("Description").Value

            If IsNothing(PictureBox1.Image) = True Then
                Button3.Enabled = False
            Else
                Button3.Enabled = True
            End If
            If TextBox8.Text = "" Then
                Button4.Enabled = False
            Else
                Button4.Enabled = True
            End If
            If RichTextBox1.Text = "" Then
                Button5.Enabled = False
            Else
                Button5.Enabled = True
            End If

            trycloseMyRec()
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

    Private Sub TextBox2_Validated(ByVal sender As Object, ByVal e As System.EventArgs) Handles TextBox2.Validated
        '////////////////////////////////////////////////////////////////////////////////
        '//
        '// После ввода кода - проверка, есть ли такой запас в базе. Если есть - 
        '// подтягиваем его значения. Если нет - даем ввести свои
        '////////////////////////////////////////////////////////////////////////////////
        If TextBox2.Modified = True Then
            TextBox2Validation()
            LoadItemAddInfo(Trim(TextBox2.Text))
        End If
    End Sub

    Public Sub TextBox2Validation()
        '////////////////////////////////////////////////////////////////////////////////
        '//
        '// После ввода кода - проверка, есть ли такой запас в базе. Если есть - 
        '// подтягиваем его значения. Если нет - даем ввести свои
        '////////////////////////////////////////////////////////////////////////////////
        Dim MySQLStr As String                        'рабочая строка
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
        MySQLStr = MySQLStr & "WHERE (SC09001 = 'RUS')) AS View_1 ON "
        MySQLStr = MySQLStr & "SC010300.SC01135 = View_1.UMID "
        MySQLStr = MySQLStr & "WHERE (SC010300.SC01001 = N'" & TextBox2.Text & "')"
        InitMyConn(False)
        InitMyRec(False, MySQLStr)
        If Declarations.MyRec.EOF = True And Declarations.MyRec.BOF = True Then
            '---такого запаса нет у нас в БД
            Label3.Text = "Рекомендованная цена и себестоимость  для этого запаса должны быть определены самостоятельно"
            Label3.ForeColor = Color.Red
            MyRez = MsgBox("Очистить поля от предыдущих значений?", MsgBoxStyle.YesNo, "Внимание!")
            If MyRez = vbYes Then
                TextBox9.Text = ""
                TextBox1.Text = ""
                TextBox1.ReadOnly = False
                ComboBox1.Text = ""
                ComboBox1.Enabled = True
                TextBox4.Text = ""
                TextBox5.Text = ""
                TextBox5.ReadOnly = False
                TextBox13.Text = ""
                TextBox13.Enabled = True
                TextBox13.BackColor = Color.FromName("Window")
                Button7.Enabled = True
                TextBox14.Text = ""
                TextBox14.Enabled = False
                TextBox14.BackColor = Color.FromName("ButtonFace")
                Label16.Text = ""
                Label17.Text = ""
                Label19.Text = ""
                Label21.Text = ""
                TextBox6.Text = ""
                TextBox7.Text = "0"
            Else
                TextBox1.ReadOnly = False
                ComboBox1.Enabled = True
                TextBox5.ReadOnly = False
                TextBox13.Enabled = True
                TextBox13.BackColor = Color.FromName("Window")
                Button7.Enabled = True
                TextBox14.Enabled = False
                TextBox14.BackColor = Color.FromName("ButtonFace")
            End If
        Else
            '---Запас в БД есть
            If Declarations.MyRec.Fields("PriCost").Value = 0 Then
                TextBox5.Text = ""
                TextBox5.ReadOnly = False
                Label3.Text = "Рекомендованная цена и себестоимость  для этого запаса должны быть определены самостоятельно"
                Label3.ForeColor = Color.Red
            Else
                TextBox5.Text = Math.Round(Declarations.MyRec.Fields("PriCost").Value, 2)
                TextBox5.ReadOnly = True
                Label3.Text = "Рекомендованная цена и себестоимость этого запаса на основе прайс - листа на закупку"
                Label3.ForeColor = Color.Green
            End If
            TextBox9.Text = Declarations.MyRec.Fields("SuppItemCode").Value
            TextBox1.Text = Declarations.MyRec.Fields("ItemName").Value.ToString
            TextBox1.ReadOnly = True
            ComboBox1.Text = Declarations.MyRec.Fields("UMName").Value.ToString
            ComboBox1.Enabled = False
            If Declarations.MyRec.Fields("Price").Value = 0 Then
                TextBox4.Text = ""
            Else
                TextBox4.Text = Math.Round(Declarations.MyRec.Fields("Price").Value, 2)
            End If
            TextBox13.Text = Declarations.MyRec.Fields("SuppID").Value.ToString
            TextBox13.Enabled = False
            TextBox13.BackColor = Color.FromName("ButtonFace")
            Button7.Enabled = False
            TextBox14.Text = Declarations.MyRec.Fields("SuppName").Value.ToString
            TextBox14.Enabled = False
            TextBox14.BackColor = Color.FromName("ButtonFace")

            trycloseMyRec()

            System.Windows.Forms.Cursor.Current = Cursors.WaitCursor
            For i As Integer = 0 To MyOrderLines.DataGridView1.Rows.Count - 1
                If Trim(MyOrderLines.DataGridView1.Item(0, i).Value.ToString) = Trim(TextBox2.Text) Then
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
            MySQLStr = MySQLStr & "(SC010300_2.SC01001 = N'" & TextBox2.Text & "') "

            InitMyConn(False)
            InitMyRec(False, MySQLStr)
            If Declarations.MyRec.EOF = True And Declarations.MyRec.BOF = True Then
                Label16.Text = ""
            Else
                Declarations.MyRec.MoveFirst()
                If Declarations.MyRec.Fields("CC").Value = 0 Then '--не складской
                    Label16.Text = ""
                Else
                    Label16.Text = "Складской ассортимент"
                End If
            End If
            trycloseMyRec()

            '--------кратность в упаковке----------------------------------
            MySQLStr = "SELECT SC01072 AS CC "
            MySQLStr = MySQLStr & "FROM SC010300 WITH (NOLOCK) "
            MySQLStr = MySQLStr & "WHERE (SC01001 = N'" & TextBox2.Text & "') "
            InitMyConn(False)
            InitMyRec(False, MySQLStr)
            If Declarations.MyRec.EOF = True And Declarations.MyRec.BOF = True Then
                Label17.Text = ""
            Else
                Declarations.MyRec.MoveFirst()
                Label17.Text = Math.Round(Declarations.MyRec.Fields("CC").Value, 2)
            End If
            trycloseMyRec()

            '--------минимальное количество в заказе на закупку------------
            MySQLStr = "SELECT tbl_PurchasePriceHistory.MinQTY AS CC, LT "
            MySQLStr = MySQLStr & "FROM SC010300 WITH (NOLOCK) INNER JOIN "
            MySQLStr = MySQLStr & "tbl_PurchasePriceHistory WITH (NOLOCK) ON SC010300.SC01001 = tbl_PurchasePriceHistory.SC01001 "
            MySQLStr = MySQLStr & "WHERE (SC010300.SC01001 = N'" & TextBox2.Text & "') AND "
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

    Private Sub TextBox3_Validated(ByVal sender As Object, ByVal e As System.EventArgs) Handles TextBox3.Validated
        '////////////////////////////////////////////////////////////////////////////////
        '//
        '// Проверки корректности выбранного количества
        '//
        '////////////////////////////////////////////////////////////////////////////////
        Dim MyRez As Object
        Dim MySQLStr As String
        Dim MyQTYAction As String               '--акция по количеству
        Dim MyActionStopQTY As Double           '--количество в акции по количеству
        Dim MyTimeAction As String              '--акция по времени
        Dim MyDateStart As DateTime             '--дата начала акции
        Dim MyDateFinish As DateTime            '--дата окончания акции
        Dim MyActionOrSales As String           '--признак - акция или распродажа
        Dim MyAvlQTY As Double                  '--доступное к заказу количество
        Dim MyRezQTY As Double                  '--зарезервированное количество
        Dim MySoldQTY As Double                 '--проданное количество
        Dim MyMessage As String

        If TextBox3.Text <> "" And TextBox2.Text <> "" Then
            '------------------проверка количеств по акциям / распродажам
            MyActionOrSales = ""
            MyQTYAction = "нет"
            MyActionStopQTY = 999999999
            MyTimeAction = "нет"
            MyDateStart = New DateTime(1900, 1, 1, 0, 0, 0)
            MyDateFinish = New DateTime(1900, 1, 1, 0, 0, 0)
            MyAvlQTY = 0
            MyRezQTY = 0
            MySoldQTY = 0
            MyMessage = ""

            '---классические неликвиды
            MySQLStr = "SELECT COUNT(*) AS CC "
            MySQLStr = MySQLStr & "FROM SC010300 "
            MySQLStr = MySQLStr & "WHERE (SC01001 = N'" & Trim(TextBox2.Text) & "') "
            MySQLStr = MySQLStr & "AND (SC01058 = N'FIN001_U')"
            InitMyConn(False)
            InitMyRec(False, MySQLStr)
            If Declarations.MyRec.EOF = True And Declarations.MyRec.BOF = True Then
            Else
                If Declarations.MyRec.Fields("CC").Value = 0 Then
                Else
                    MyActionOrSales = "неликвид"
                    MyQTYAction = "да"
                End If
            End If
            trycloseMyRec()

            '---новые неликвиды или распродажа
            If MyActionOrSales.Equals("") Then
                MySQLStr = "SELECT QTYAction, ActionStopQTY, TimeAction, DateStart, DateFinish, ActionOrSales "
                MySQLStr = MySQLStr & "FROM tbl_ActionsAndSales "
                MySQLStr = MySQLStr & "WHERE (ScalaCode = N'" & Trim(TextBox2.Text) & "') "
                MySQLStr = MySQLStr & "AND (ActionFinished = 0) "
                MySQLStr = MySQLStr & "AND (DateStart <= dateadd(day, datediff(day, 0, GETDATE()), 0)) "
                MySQLStr = MySQLStr & "AND (DateFinish >= dateadd(day, datediff(day, 0, GETDATE()), 0))"
                InitMyConn(False)
                InitMyRec(False, MySQLStr)
                If Declarations.MyRec.EOF = True And Declarations.MyRec.BOF = True Then
                Else
                    MyActionOrSales = Declarations.MyRec.Fields("ActionOrSales").Value
                    MyQTYAction = Declarations.MyRec.Fields("QTYAction").Value
                    MyActionStopQTY = Declarations.MyRec.Fields("ActionStopQTY").Value
                    MyTimeAction = Declarations.MyRec.Fields("TimeAction").Value
                    MyDateStart = Declarations.MyRec.Fields("DateStart").Value
                    MyDateFinish = Declarations.MyRec.Fields("DateFinish").Value
                End If
                trycloseMyRec()
            End If

            '---доступное к заказу количество на всех складах
            MySQLStr = "SELECT SC33001, SUM(SC33005 - SC33006) AS QTY, SUM(SC33006) as RezQTY "
            MySQLStr = MySQLStr & "FROM SC330300 "
            MySQLStr = MySQLStr & "WHERE (SC33001 = N'" & Trim(TextBox2.Text) & "') "
            MySQLStr = MySQLStr & "GROUP BY SC33001 "
            InitMyConn(False)
            InitMyRec(False, MySQLStr)
            If Declarations.MyRec.EOF = True And Declarations.MyRec.BOF = True Then
                MyAvlQTY = 0
                MyRezQTY = 0
            Else
                MyAvlQTY = Declarations.MyRec.Fields("QTY").Value
                MyRezQTY = Declarations.MyRec.Fields("RezQTY").Value
            End If
            trycloseMyRec()

            '---Уже отгруженное в рамках акции количество
            If MyActionStopQTY <> 999999999 Then
                MySQLStr = "SELECT ISNULL(SUM(SC07004), 0) AS SoldQTY "
                MySQLStr = MySQLStr & "FROM SC070300 "
                MySQLStr = MySQLStr & "WHERE (SC07002 >= CONVERT(datetime, '" & Strings.Right("00" & MyDateStart.Day.ToString, 2) & "/" & Strings.Right("00" & MyDateStart.Month.ToString, 2) & "/" & MyDateStart.Year.ToString & "', 103)) "
                MySQLStr = MySQLStr & "AND (SC07003 = N'" & Trim(TextBox2.Text) & "')"
                InitMyConn(False)
                InitMyRec(False, MySQLStr)
                If Declarations.MyRec.EOF = True And Declarations.MyRec.BOF = True Then
                    MySoldQTY = 0
                Else
                    MySoldQTY = Declarations.MyRec.Fields("SoldQTY").Value
                End If
                trycloseMyRec()
            End If

            '--Формирование строки сообщения
            '----------------------неликвид
            If MyActionOrSales = "неликвид" Then
                '-----только по количеству
                If MyQTYAction = "да" And MyTimeAction = "нет" Then
                    '---неограниченное кол-во до ухода в 0
                    If MyActionStopQTY = 999999999 Then
                        If CDbl(TextBox3.Text) > MyAvlQTY Then
                            MyMessage = "Данный товар является неликвидом и в настоящий момент " & Chr(13) & Chr(10)
                            MyMessage = MyMessage & "идет акция по его полной распродаже " & Chr(13) & Chr(10)
                            MyMessage = MyMessage & "Вы включаете в заказ количество большее, чем есть на складе " & Chr(13) & Chr(10)
                            MyMessage = MyMessage & "В этом случае недостающее количество должно быть " & Chr(13) & Chr(10)
                            MyMessage = MyMessage & "продано по обычной цене. " & Chr(13) & Chr(10) & Chr(13) & Chr(10)
                            MyMessage = MyMessage & "отменить ввод данного количества? " & Chr(13) & Chr(10)
                        End If
                    Else    '---кол-во ограничено
                        If CDbl(TextBox3.Text) > MyActionStopQTY - MySoldQTY - MyRezQTY Then
                            MyMessage = "Данный товар является неликвидом и в настоящий момент " & Chr(13) & Chr(10)
                            MyMessage = MyMessage & "идет акция по распродаже следующего количества: " & MyActionStopQTY.ToString & Chr(13) & Chr(10)
                            MyMessage = MyMessage & "При этом по данной распродаже уже продано: " & MySoldQTY.ToString & Chr(13) & Chr(10)
                            MyMessage = MyMessage & "Также в заказах на продажу ненулевого типа зарезервировано: " & MyRezQTY.ToString & Chr(13) & Chr(10)
                            MyMessage = MyMessage & "Вы включаете в заказ количество большее, чем должно быть продано по данной распродаже " & Chr(13) & Chr(10)
                            MyMessage = MyMessage & "В этом случае количество, превышающее количество по распродаже, должно быть " & Chr(13) & Chr(10)
                            MyMessage = MyMessage & "продано по обычной цене. " & Chr(13) & Chr(10) & Chr(13) & Chr(10)
                            MyMessage = MyMessage & "отменить ввод данного количества? " & Chr(13) & Chr(10)
                        End If
                    End If
                End If
                '-----только по срокам
                If MyQTYAction = "нет" And MyTimeAction = "да" Then
                    MyMessage = "Распродажа данного неликвида идет по времени и заканчивается " & Chr(13) & Chr(10)
                    MyMessage = MyMessage & MyDateFinish.Day.ToString & "/" & MyDateFinish.Month.ToString & "/" & MyDateFinish.Year.ToString & ". Чтобы данная продажа фиксировалась как продажа неликвида, " & Chr(13) & Chr(10)
                    MyMessage = MyMessage & "заказ должен быть переведен в 1 тип до этой даты. " & Chr(13) & Chr(10)
                    MyMessage = MyMessage & "В противном случае это будет обычная продажа, которую надо будет " & Chr(13) & Chr(10)
                    MyMessage = MyMessage & "осуществить по обычной цене " & Chr(13) & Chr(10) & Chr(13) & Chr(10)
                    MyMessage = MyMessage & "отменить ввод данного количества? " & Chr(13) & Chr(10)
                End If
                '-----по количеству и по срокам
                If MyQTYAction = "да" And MyTimeAction = "да" Then
                    '---неограниченное кол-во до ухода в 0 или окончания срока акции
                    If MyActionStopQTY = 999999999 Then
                        If CDbl(TextBox3.Text) > MyAvlQTY Then
                            MyMessage = "Данный товар является неликвидом и в настоящий момент " & Chr(13) & Chr(10)
                            MyMessage = MyMessage & "идет акция по его полной распродаже " & Chr(13) & Chr(10)
                            MyMessage = MyMessage & "Вы включаете в заказ количество большее, чем есть на складе " & Chr(13) & Chr(10)
                            MyMessage = MyMessage & "В этом случае недостающее количество должно быть " & Chr(13) & Chr(10)
                            MyMessage = MyMessage & "продано по обычной цене. " & Chr(13) & Chr(10) & Chr(13) & Chr(10)
                            MyMessage = MyMessage & "Также распродажа данного неликвида идет по времени и заканчивается " & Chr(13) & Chr(10)
                            MyMessage = MyMessage & MyDateFinish.Day.ToString & "/" & MyDateFinish.Month.ToString & "/" & MyDateFinish.Year.ToString & ". Чтобы данная продажа фиксировалась как продажа неликвида, " & Chr(13) & Chr(10)
                            MyMessage = MyMessage & "заказ должен быть переведен в 1 тип до этой даты. " & Chr(13) & Chr(10)
                            MyMessage = MyMessage & "В противном случае это будет обычная продажа, которую надо будет " & Chr(13) & Chr(10)
                            MyMessage = MyMessage & "осуществить по обычной цене " & Chr(13) & Chr(10) & Chr(13) & Chr(10)
                            MyMessage = MyMessage & "отменить ввод данного количества? " & Chr(13) & Chr(10)
                        Else
                            MyMessage = "Распродажа данного неликвида идет по времени и заканчивается " & Chr(13) & Chr(10)
                            MyMessage = MyMessage & MyDateFinish.Day.ToString & "/" & MyDateFinish.Month.ToString & "/" & MyDateFinish.Year.ToString & ". Чтобы данная продажа фиксировалась как продажа неликвида, " & Chr(13) & Chr(10)
                            MyMessage = MyMessage & "заказ должен быть переведен в 1 тип до этой даты. " & Chr(13) & Chr(10)
                            MyMessage = MyMessage & "В противном случае это будет обычная продажа, которую надо будет " & Chr(13) & Chr(10)
                            MyMessage = MyMessage & "осуществить по обычной цене " & Chr(13) & Chr(10) & Chr(13) & Chr(10)
                            MyMessage = MyMessage & "отменить ввод данного количества? " & Chr(13) & Chr(10)
                        End If
                    Else    '---кол-во ограничено и есть срок акции
                        If CDbl(TextBox3.Text) > MyActionStopQTY - MySoldQTY - MyRezQTY Then
                            MyMessage = "Данный товар является неликвидом и в настоящий момент " & Chr(13) & Chr(10)
                            MyMessage = MyMessage & "идет акция по распродаже следующего количества: " & MyActionStopQTY.ToString & Chr(13) & Chr(10)
                            MyMessage = MyMessage & "При этом по данной распродаже уже продано: " & MySoldQTY.ToString & Chr(13) & Chr(10)
                            MyMessage = MyMessage & "Также в заказах на продажу ненулевого типа зарезервировано: " & MyRezQTY.ToString & Chr(13) & Chr(10)
                            MyMessage = MyMessage & "Вы включаете в заказ количество большее, чем должно быть продано по данной распродаже " & Chr(13) & Chr(10)
                            MyMessage = MyMessage & "В этом случае количество, превышающее количество по распродаже, должно быть " & Chr(13) & Chr(10)
                            MyMessage = MyMessage & "продано по обычной цене. " & Chr(13) & Chr(10) & Chr(13) & Chr(10)
                            MyMessage = MyMessage & "Также распродажа данного неликвида идет по времени и заканчивается " & Chr(13) & Chr(10)
                            MyMessage = MyMessage & MyDateFinish.Day.ToString & "/" & MyDateFinish.Month.ToString & "/" & MyDateFinish.Year.ToString & ". Чтобы данная продажа фиксировалась как продажа неликвида, " & Chr(13) & Chr(10)
                            MyMessage = MyMessage & "заказ должен быть переведен в 1 тип до этой даты. " & Chr(13) & Chr(10)
                            MyMessage = MyMessage & "В противном случае это будет обычная продажа, которую надо будет " & Chr(13) & Chr(10)
                            MyMessage = MyMessage & "осуществить по обычной цене " & Chr(13) & Chr(10) & Chr(13) & Chr(10)
                            MyMessage = MyMessage & "отменить ввод данного количества? " & Chr(13) & Chr(10)
                        Else
                            MyMessage = "Распродажа данного неликвида идет по времени и заканчивается " & Chr(13) & Chr(10)
                            MyMessage = MyMessage & MyDateFinish.Day.ToString & "/" & MyDateFinish.Month.ToString & "/" & MyDateFinish.Year.ToString & ". Чтобы данная продажа фиксировалась как продажа неликвида, " & Chr(13) & Chr(10)
                            MyMessage = MyMessage & "заказ должен быть переведен в 1 тип до этой даты. " & Chr(13) & Chr(10)
                            MyMessage = MyMessage & "В противном случае это будет обычная продажа, которую надо будет " & Chr(13) & Chr(10)
                            MyMessage = MyMessage & "осуществить по обычной цене " & Chr(13) & Chr(10) & Chr(13) & Chr(10)
                            MyMessage = MyMessage & "отменить ввод данного количества? " & Chr(13) & Chr(10)
                        End If
                    End If
                End If
            End If
            '----------------------акция
            If MyActionOrSales = "акция" Then
                '-----только по количеству
                If MyQTYAction = "да" And MyTimeAction = "нет" Then
                    '---неограниченное кол-во до ухода в 0
                    If MyActionStopQTY = 999999999 Then
                        If CDbl(TextBox3.Text) > MyAvlQTY Then
                            MyMessage = "Данный товар является акционным и в настоящий момент " & Chr(13) & Chr(10)
                            MyMessage = MyMessage & "идет акция по его полной распродаже " & Chr(13) & Chr(10)
                            MyMessage = MyMessage & "Вы включаете в заказ количество большее, чем есть на складе " & Chr(13) & Chr(10)
                            MyMessage = MyMessage & "В этом случае недостающее количество должно быть " & Chr(13) & Chr(10)
                            MyMessage = MyMessage & "продано по обычной цене. " & Chr(13) & Chr(10) & Chr(13) & Chr(10)
                            MyMessage = MyMessage & "отменить ввод данного количества? " & Chr(13) & Chr(10)
                        End If
                    Else    '---кол-во ограничено
                        If CDbl(TextBox3.Text) > MyActionStopQTY - MySoldQTY - MyRezQTY Then
                            MyMessage = "Данный товар является акционным и в настоящий момент " & Chr(13) & Chr(10)
                            MyMessage = MyMessage & "идет акция по продаже следующего количества: " & MyActionStopQTY.ToString & Chr(13) & Chr(10)
                            MyMessage = MyMessage & "При этом по данной акции уже продано: " & MySoldQTY.ToString & Chr(13) & Chr(10)
                            MyMessage = MyMessage & "Также в заказах на продажу ненулевого типа зарезервировано: " & MyRezQTY.ToString & Chr(13) & Chr(10)
                            MyMessage = MyMessage & "Вы включаете в заказ количество большее, чем должно быть продано по данной акции " & Chr(13) & Chr(10)
                            MyMessage = MyMessage & "В этом случае количество, превышающее количество по акции, должно быть " & Chr(13) & Chr(10)
                            MyMessage = MyMessage & "продано по обычной цене. " & Chr(13) & Chr(10) & Chr(13) & Chr(10)
                            MyMessage = MyMessage & "отменить ввод данного количества? " & Chr(13) & Chr(10)
                        End If
                    End If
                End If
                '-----только по срокам
                If MyQTYAction = "нет" And MyTimeAction = "да" Then
                    If MyQTYAction = "нет" And MyTimeAction = "да" Then
                        MyMessage = "Акция по данному товару идет по времени и заканчивается " & Chr(13) & Chr(10)
                        MyMessage = MyMessage & MyDateFinish.Day.ToString & "/" & MyDateFinish.Month.ToString & "/" & MyDateFinish.Year.ToString & ". Чтобы данная продажа фиксировалась как продажа по акции, " & Chr(13) & Chr(10)
                        MyMessage = MyMessage & "заказ должен быть переведен в 1 тип до этой даты. " & Chr(13) & Chr(10)
                        MyMessage = MyMessage & "В противном случае это будет обычная продажа, которую надо будет " & Chr(13) & Chr(10)
                        MyMessage = MyMessage & "осуществить по обычной цене " & Chr(13) & Chr(10) & Chr(13) & Chr(10)
                        MyMessage = MyMessage & "отменить ввод данного количества? " & Chr(13) & Chr(10)
                    End If
                End If
                '-----по количеству и по срокам
                If MyQTYAction = "да" And MyTimeAction = "да" Then
                    '---неограниченное кол-во до ухода в 0 или окончания срока акции
                    If MyActionStopQTY = 999999999 Then
                        If CDbl(TextBox3.Text) > MyAvlQTY Then
                            MyMessage = "Данный товар является акционным и в настоящий момент " & Chr(13) & Chr(10)
                            MyMessage = MyMessage & "идет акция по его полной распродаже " & Chr(13) & Chr(10)
                            MyMessage = MyMessage & "Вы включаете в заказ количество большее, чем есть на складе " & Chr(13) & Chr(10)
                            MyMessage = MyMessage & "В этом случае недостающее количество должно быть " & Chr(13) & Chr(10)
                            MyMessage = MyMessage & "продано по обычной цене. " & Chr(13) & Chr(10) & Chr(13) & Chr(10)
                            MyMessage = MyMessage & "Также акция по продаже данного товара идет по времени и заканчивается " & Chr(13) & Chr(10)
                            MyMessage = MyMessage & MyDateFinish.Day.ToString & "/" & MyDateFinish.Month.ToString & "/" & MyDateFinish.Year.ToString & ". Чтобы данная продажа фиксировалась как продажа по акции, " & Chr(13) & Chr(10)
                            MyMessage = MyMessage & "заказ должен быть переведен в 1 тип до этой даты. " & Chr(13) & Chr(10)
                            MyMessage = MyMessage & "В противном случае это будет обычная продажа, которую надо будет " & Chr(13) & Chr(10)
                            MyMessage = MyMessage & "осуществить по обычной цене " & Chr(13) & Chr(10) & Chr(13) & Chr(10)
                            MyMessage = MyMessage & "отменить ввод данного количества? " & Chr(13) & Chr(10)
                        Else
                            MyMessage = "Акция по данному товару идет по времени и заканчивается " & Chr(13) & Chr(10)
                            MyMessage = MyMessage & MyDateFinish.Day.ToString & "/" & MyDateFinish.Month.ToString & "/" & MyDateFinish.Year.ToString & ". Чтобы данная продажа фиксировалась как продажа по акции, " & Chr(13) & Chr(10)
                            MyMessage = MyMessage & "заказ должен быть переведен в 1 тип до этой даты. " & Chr(13) & Chr(10)
                            MyMessage = MyMessage & "В противном случае это будет обычная продажа, которую надо будет " & Chr(13) & Chr(10)
                            MyMessage = MyMessage & "осуществить по обычной цене " & Chr(13) & Chr(10) & Chr(13) & Chr(10)
                            MyMessage = MyMessage & "отменить ввод данного количества? " & Chr(13) & Chr(10)
                        End If
                    Else    '---кол-во ограничено и есть срок акции
                        If CDbl(TextBox3.Text) > MyActionStopQTY - MySoldQTY - MyRezQTY Then
                            MyMessage = "Данный товар является акционным и в настоящий момент " & Chr(13) & Chr(10)
                            MyMessage = MyMessage & "идет акция по продаже следующего количества: " & MyActionStopQTY.ToString & Chr(13) & Chr(10)
                            MyMessage = MyMessage & "При этом по данной акции уже продано: " & MySoldQTY.ToString & Chr(13) & Chr(10)
                            MyMessage = MyMessage & "Также в заказах на продажу ненулевого типа зарезервировано: " & MyRezQTY.ToString & Chr(13) & Chr(10)
                            MyMessage = MyMessage & "Вы включаете в заказ количество большее, чем должно быть продано по данной акции " & Chr(13) & Chr(10)
                            MyMessage = MyMessage & "В этом случае количество, превышающее количество по акции, должно быть " & Chr(13) & Chr(10)
                            MyMessage = MyMessage & "продано по обычной цене. " & Chr(13) & Chr(10) & Chr(13) & Chr(10)
                            MyMessage = MyMessage & "Также акция по продаже данного товара идет по времени и заканчивается " & Chr(13) & Chr(10)
                            MyMessage = MyMessage & MyDateFinish.Day.ToString & "/" & MyDateFinish.Month.ToString & "/" & MyDateFinish.Year.ToString & ". Чтобы данная продажа фиксировалась как продажа по акции, " & Chr(13) & Chr(10)
                            MyMessage = MyMessage & "заказ должен быть переведен в 1 тип до этой даты. " & Chr(13) & Chr(10)
                            MyMessage = MyMessage & "В противном случае это будет обычная продажа, которую надо будет " & Chr(13) & Chr(10)
                            MyMessage = MyMessage & "осуществить по обычной цене " & Chr(13) & Chr(10) & Chr(13) & Chr(10)
                            MyMessage = MyMessage & "отменить ввод данного количества? " & Chr(13) & Chr(10)
                        Else
                            MyMessage = "Акция по данному товару идет по времени и заканчивается " & Chr(13) & Chr(10)
                            MyMessage = MyMessage & MyDateFinish.Day.ToString & "/" & MyDateFinish.Month.ToString & "/" & MyDateFinish.Year.ToString & ". Чтобы данная продажа фиксировалась как продажа по акции, " & Chr(13) & Chr(10)
                            MyMessage = MyMessage & "заказ должен быть переведен в 1 тип до этой даты. " & Chr(13) & Chr(10)
                            MyMessage = MyMessage & "В противном случае это будет обычная продажа, которую надо будет " & Chr(13) & Chr(10)
                            MyMessage = MyMessage & "осуществить по обычной цене " & Chr(13) & Chr(10) & Chr(13) & Chr(10)
                            MyMessage = MyMessage & "отменить ввод данного количества? " & Chr(13) & Chr(10)
                        End If
                    End If
                End If
            End If

            '---предупреждение продавцу
            If MyMessage <> "" Then
                MyRez = MsgBox(MyMessage, vbYesNo, "Внимание!")
                If MyRez = vbYes Then
                    TextBox3.Select()
                    Exit Sub
                End If
            End If
        End If

        If TextBox3.Text <> "" And TextBox2.Text <> "" Then
            '---Вывод окна с альтернативными продуктами при определенных условиях
            If Trim(TextBox2.Text) = Trim(MyOrderLines.DataGridView1.SelectedRows.Item(0).Cells(1).Value.ToString) And _
                CDbl(TextBox3.Text) > CDbl(MyOrderLines.DataGridView1.SelectedRows.Item(0).Cells(8).Value.ToString) Then '---Заказываем больше чем есть на складе отгрузки
                MySQLStr = "SELECT ISNULL(SUM(WhAvl), 0) AS AVL "
                MySQLStr = MySQLStr & "FROM (SELECT TOP (100) PERCENT "
                MySQLStr = MySQLStr & "tbl_AlternativeProducts.ALTCode AS ProductCode, "
                MySQLStr = MySQLStr & "ISNULL(View_1_1.WhAvl, 0) AS WhAvl "
                MySQLStr = MySQLStr & "FROM SC010300 WITH (NOLOCK) INNER JOIN "
                MySQLStr = MySQLStr & "(SELECT SC03001, SC03003 AS WhQty, "
                MySQLStr = MySQLStr & "SC03003 - SC03004 - SC03005 AS WhAvl "
                MySQLStr = MySQLStr & "FROM SC030300 WITH (NOLOCK) "
                MySQLStr = MySQLStr & "WHERE (SC03002 = N'" & Declarations.WHNum & "')) "
                MySQLStr = MySQLStr & "AS View_1_1 ON SC010300.SC01001 = View_1_1.SC03001 RIGHT OUTER JOIN "
                MySQLStr = MySQLStr & "tbl_AlternativeProducts ON SC010300.SC01001 = tbl_AlternativeProducts.ALTCode "
                MySQLStr = MySQLStr & "WHERE (tbl_AlternativeProducts.ProductCode = N'" & Trim(TextBox2.Text) & "') "
                MySQLStr = MySQLStr & "ORDER BY ProductCode) AS View_1 "
                InitMyConn(False)
                InitMyRec(False, MySQLStr)
                If Declarations.MyRec.EOF = True And Declarations.MyRec.BOF = True Then
                Else
                    If Declarations.MyRec.Fields("AVL").Value > 0 Then '---Есть доступные альтернативные продукты
                        trycloseMyRec()
                        MyALTItems = New ALTItems
                        MyALTItems.MyItem = Trim(TextBox2.Text)
                        MyALTItems.MySrcWin = "AddToOrder"
                        MyALTItems.ShowDialog()
                    Else
                        trycloseMyRec()
                    End If
                End If
            End If
        End If
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

    Private Sub TextBox4_Validating(ByVal sender As Object, ByVal e As System.ComponentModel.CancelEventArgs) Handles TextBox4.Validating
        '////////////////////////////////////////////////////////////////////////////////
        '//
        '// проверка - числовое ли знечение введено
        '//
        '////////////////////////////////////////////////////////////////////////////////
        Dim MyRez As Double
        Dim aa As New System.Globalization.NumberFormatInfo

        If Trim(TextBox4.Text) <> "" Then
            If InStr(TextBox4.Text, aa.CurrentInfo.NumberGroupSeparator) Then
                MsgBox("В поле ""Рекоменд. цена"" должно быть введено число", MsgBoxStyle.Critical, "Внимание!")
                e.Cancel = True
                Exit Sub
            Else
                Try
                    MyRez = TextBox4.Text
                Catch ex As Exception
                    MsgBox("В поле ""Рекомендов. цена"" должно быть введено число", MsgBoxStyle.Critical, "Внимание!")
                    e.Cancel = True
                    Exit Sub
                End Try
            End If
        End If
        e.Cancel = False
    End Sub

    Private Sub TextBox5_Validating(ByVal sender As Object, ByVal e As System.ComponentModel.CancelEventArgs) Handles TextBox5.Validating
        '////////////////////////////////////////////////////////////////////////////////
        '//
        '// проверка - числовое ли знечение введено
        '//
        '////////////////////////////////////////////////////////////////////////////////
        Dim MyRez As Double
        Dim aa As New System.Globalization.NumberFormatInfo

        If Trim(TextBox5.Text) <> "" Then
            If InStr(TextBox5.Text, aa.CurrentInfo.NumberGroupSeparator) Then
                MsgBox("В поле ""Себестоимость"" должно быть введено число", MsgBoxStyle.Critical, "Внимание!")
                e.Cancel = True
                Exit Sub
            Else
                Try
                    MyRez = TextBox5.Text
                Catch ex As Exception
                    MsgBox("В поле ""Себестоимость"" должно быть введено число", MsgBoxStyle.Critical, "Внимание!")
                    e.Cancel = True
                    Exit Sub
                End Try
            End If
        End If
        e.Cancel = False
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

    Private Sub TextBox6_Validating(ByVal sender As Object, ByVal e As System.ComponentModel.CancelEventArgs) Handles TextBox6.Validating
        '////////////////////////////////////////////////////////////////////////////////
        '//
        '// проверка - числовое ли знечение введено
        '//
        '////////////////////////////////////////////////////////////////////////////////
        Dim MyRez As Double
        Dim aa As New System.Globalization.NumberFormatInfo

        If Trim(TextBox6.Text) <> "" Then
            If InStr(TextBox6.Text, aa.CurrentInfo.NumberGroupSeparator) Then
                MsgBox("В поле ""Срок поставки (нед)"" должно быть введено число. Если товар есть в наличии на складе - то срок доставки до клиента; Если товар есть в наличии на складе и доставки нет - то 0", MsgBoxStyle.Critical, "Внимание!")
                e.Cancel = True
                Exit Sub
            Else
                Try
                    MyRez = TextBox6.Text
                    If MyRez < 0 Then
                        MsgBox("В поле ""Срок поставки (нед)"" должно быть введено число, большее или равное 0", MsgBoxStyle.Critical, "Внимание!")
                        e.Cancel = True
                        Exit Sub
                    End If
                Catch ex As Exception
                    MsgBox("В поле ""Срок поставки (нед)"" должно быть введено число. Если товар есть в наличии на складе - то срок доставки до клиента; Если товар есть в наличии на складе и доставки нет - то 0", MsgBoxStyle.Critical, "Внимание!")
                    e.Cancel = True
                    Exit Sub
                End Try
            End If
        End If
        e.Cancel = False
    End Sub

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        '////////////////////////////////////////////////////////////////////////////////
        '//
        '// Занесение данных в строку заказа
        '//
        '////////////////////////////////////////////////////////////////////////////////

        If CheckDataFiling() = True Then
            Declarations.MySuccess = True                               'Успешность выполнения операции
            Declarations.MyItemID = Trim(TextBox2.Text)                 'код запаса
            Declarations.MyItemName = Trim(TextBox1.Text)               'имя запаса
            Declarations.MyQty = Trim(TextBox3.Text)                    'количество заказанного
            Declarations.MyUOM = ComboBox1.SelectedValue                'код единицы измерения
            Declarations.MySum = Trim(TextBox4.Text)                    'сумма заказанного
            Declarations.MySS = Trim(TextBox5.Text)                     'себестоимость
            Declarations.WeekQTY = Math.Round(CDbl(Trim(TextBox6.Text)), 1)  'Срок поставки
            Declarations.DelWeekQTY = Math.Round(CDbl(Trim(TextBox7.Text)), 1)  'Срок доставки до клиента
            Declarations.MyItemSuppID = Trim(TextBox9.Text)
            Declarations.MySuppID = Trim(TextBox13.Text)
            Declarations.MySuppName = Trim(TextBox14.Text)
            Me.Close()
        End If
    End Sub

    Private Function CheckDataFiling() As Boolean
        '////////////////////////////////////////////////////////////////////////////////
        '//
        '// Проверка заполнения полей в окне
        '//
        '////////////////////////////////////////////////////////////////////////////////
        Dim MyMargin As Double
        Dim MyRez As Object

        If Trim(TextBox2.Text) = "" Then
            MsgBox("Поле ""Код запаса"" должно быть заполнено", MsgBoxStyle.Critical, "Внимание")
            CheckDataFiling = False
            Exit Function
        End If

        If Trim(TextBox1.Text) = "" Then
            MsgBox("Поле ""Название запаса"" должно быть заполнено", MsgBoxStyle.Critical, "Внимание")
            CheckDataFiling = False
            Exit Function
        End If

        If Trim(TextBox3.Text) = "" Then
            MsgBox("Поле ""Количество"" должно быть заполнено", MsgBoxStyle.Critical, "Внимание")
            CheckDataFiling = False
            Exit Function
        End If

        If Trim(TextBox4.Text) = "" Then
            MsgBox("Поле ""Рекоменд. цена"" должно быть заполнено", MsgBoxStyle.Critical, "Внимание")
            CheckDataFiling = False
            Exit Function
        End If

        If Trim(TextBox5.Text) = "" Then
            MsgBox("Поле ""Себестоимость"" должно быть заполнено", MsgBoxStyle.Critical, "Внимание")
            CheckDataFiling = False
            Exit Function
        End If

        If Trim(TextBox6.Text) = "" Then
            MsgBox("Поле ""Срок поставки (нед)"" должно быть заполнено", MsgBoxStyle.Critical, "Внимание")
            CheckDataFiling = False
            Exit Function
        End If

        If Trim(TextBox7.Text) = "" Then
            MsgBox("Поле ""Срок доставки до клиента (нед)"" должно быть заполнено. Если доставки до клиента нет (самовывоз) - то 0", MsgBoxStyle.Critical, "Внимание")
            CheckDataFiling = False
            Exit Function
        End If

        If Math.Round(CDbl(Trim(TextBox7.Text)), 1) > Math.Round(CDbl(Trim(TextBox6.Text)), 1) Then
            MsgBox("Поле ""Срок доставки до клиента (нед)"" не может быть больше срока поставки клиенту", MsgBoxStyle.Critical, "Внимание")
            CheckDataFiling = False
            Exit Function
        End If

        If Trim(TextBox13.Text) = "" Then
            MsgBox("Поле ""Код поставщика"" должно быть заполнено", MsgBoxStyle.Critical, "Внимание")
            CheckDataFiling = False
            Exit Function
        End If

        If Trim(TextBox14.Text) = "" Then
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

        'If CDbl(TextBox4.Text) = 0 Then
        'MyMargin = 0
        'Else
        'MyMargin = (CDbl(TextBox4.Text) - CDbl(TextBox5.Text)) / (TextBox4.Text) * 100
        'End If
        'If MyMargin < Declarations.MinMarginLevelManager Then
        'MyRez = MsgBox("Вы пытаетесь внести в заказ строку с маржой меньшей, чем разрешено для данного клиента. Отменить ввод такого запаса?", vbYesNo, "Внимание!")
        'If MyRez = vbYes Then
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
        Dim MyMargin As Double
        Dim MyRez As Object

        If Trim(TextBox5.Text) = "" Then
            MsgBox("Поле ""Себестоимость"" должно быть заполнено", MsgBoxStyle.Critical, "Внимание")
            CheckDataFilingZ = False
            Exit Function
        End If

        CheckDataFilingZ = True
    End Function

    Private Sub TextBox7_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles TextBox7.KeyDown
        '////////////////////////////////////////////////////////////////////////////////
        '//
        '// переход на следующее поле по Enter
        '//
        '////////////////////////////////////////////////////////////////////////////////

        If e.KeyData = Keys.Enter Then
            Me.SelectNextControl(sender, True, True, True, False)
        End If
    End Sub

    Private Sub TextBox7_Validating(ByVal sender As Object, ByVal e As System.ComponentModel.CancelEventArgs) Handles TextBox7.Validating
        '////////////////////////////////////////////////////////////////////////////////
        '//
        '// проверка - числовое ли знечение введено
        '//
        '////////////////////////////////////////////////////////////////////////////////
        Dim MyRez As Double
        Dim aa As New System.Globalization.NumberFormatInfo

        If Trim(TextBox7.Text) <> "" Then
            If InStr(TextBox7.Text, aa.CurrentInfo.NumberGroupSeparator) Then
                MsgBox("В поле ""Срок доставки до клиента (нед)"" должно быть введено число. Если доставки до клиента нет (самовывоз) - то 0", MsgBoxStyle.Critical, "Внимание!")
                e.Cancel = True
                Exit Sub
            Else
                Try
                    MyRez = TextBox7.Text
                    If MyRez < 0 Then
                        MsgBox("В поле ""Срок доставки до клиента (нед)"" должно быть введено число, большее или равное 0", MsgBoxStyle.Critical, "Внимание!")
                        e.Cancel = True
                        Exit Sub
                    End If
                Catch ex As Exception
                    MsgBox("В поле ""Срок доставки до клиента (нед)"" должно быть введено число. Если доставки до клиента нет (самовывоз) - то 0", MsgBoxStyle.Critical, "Внимание!")
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
        '// Уведомление о неверной картинке
        '//
        '////////////////////////////////////////////////////////////////////////////////
        Dim myRez As MsgBoxResult

        myRez = MsgBox("Вы уверены?", MsgBoxStyle.YesNo, "Внимание!")
        If myRez = MsgBoxResult.Yes Then
            SendAddInfoReminder(Trim(TextBox2.Text), Declarations.SalesmanName, 0)
            MsgBox("Уведомление о неверной картинке отправлено", MsgBoxStyle.Information, "Внимание!")
        End If
    End Sub

    Private Sub Button4_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button4.Click
        '////////////////////////////////////////////////////////////////////////////////
        '//
        '// Уведомление о неверном названии
        '//
        '////////////////////////////////////////////////////////////////////////////////
        Dim myRez As MsgBoxResult

        myRez = MsgBox("Вы уверены?", MsgBoxStyle.YesNo, "Внимание!")
        If myRez = MsgBoxResult.Yes Then
            SendAddInfoReminder(Trim(TextBox2.Text), Declarations.SalesmanName, 1)
            MsgBox("Уведомление о неверном названии отправлено", MsgBoxStyle.Information, "Внимание!")
        End If
    End Sub

    Private Sub Button5_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button5.Click
        '////////////////////////////////////////////////////////////////////////////////
        '//
        '// Уведомление о неверном описании
        '//
        '////////////////////////////////////////////////////////////////////////////////
        Dim myRez As MsgBoxResult

        myRez = MsgBox("Вы уверены?", MsgBoxStyle.YesNo, "Внимание!")
        If myRez = MsgBoxResult.Yes Then
            SendAddInfoReminder(Trim(TextBox2.Text), Declarations.SalesmanName, 2)
            MsgBox("Уведомление о неверном описании отправлено", MsgBoxStyle.Information, "Внимание!")
        End If
    End Sub

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

    Private Sub TextBox9_Validated(ByVal sender As Object, ByVal e As System.EventArgs) Handles TextBox9.Validated
        '////////////////////////////////////////////////////////////////////////////////
        '//
        '// После ввода кода - проверка, есть ли такой запас в базе. Если есть - 
        '// подтягиваем его значения. Если нет - даем ввести свои
        '////////////////////////////////////////////////////////////////////////////////
        If TextBox9.Modified = True Then
            TextBox9Validation()
        End If
    End Sub

    Public Sub TextBox9Validation()
        '////////////////////////////////////////////////////////////////////////////////
        '//
        '// После ввода кода товара поставщика- проверка, есть ли такой запас в базе. Если есть - 
        '// подтягиваем его значения. Если нет - даем ввести свои
        '////////////////////////////////////////////////////////////////////////////////
        Dim MySQLStr As String                          'рабочая строка
        Dim RecQTY As Double                            '--Количество записей
        Dim MyRez As Object
        'Dim FoundFlag As Integer                        'флаг - сколько найдено

        If Trim(TextBox9.Text) <> "" Then
            MySQLStr = "exec spp_SalesWorkplace4_SpecificationFilling "
            MySQLStr = MySQLStr & "N'', "                                   '--Код товара в Scala
            MySQLStr = MySQLStr & "N'" & Trim(TextBox9.Text) & "'  "        '--Код товара поставщика

            InitMyConn(False)
            InitMyRec(False, MySQLStr)

            If Declarations.MyRec.EOF = True And Declarations.MyRec.BOF = True Then
                '--ничего не найдено
                '---очищаем все поля 
                Label3.Text = "Рекомендованная цена и себестоимость  для этого запаса должны быть определены самостоятельно"
                Label3.ForeColor = Color.Red
                MyRez = MsgBox("Очистить поля от предыдущих значений?", MsgBoxStyle.YesNo, "Внимание!")
                If MyRez = vbYes Then
                    TextBox2.Text = "Unknown"                      '--код Scala
                    TextBox1.Text = ""
                    TextBox1.ReadOnly = False
                    ComboBox1.Text = ""
                    ComboBox1.Enabled = True
                    TextBox3.Text = "1"
                    TextBox4.Text = ""
                    TextBox5.Text = ""
                    TextBox5.ReadOnly = False
                    TextBox13.Text = ""
                    TextBox13.Enabled = True
                    TextBox13.BackColor = Color.FromName("Window")
                    Button7.Enabled = True
                    TextBox14.Text = ""
                    TextBox14.Enabled = False
                    TextBox14.BackColor = Color.FromName("ButtonFace")
                    Label16.Text = ""
                    Label17.Text = ""
                    Label19.Text = ""
                    Label21.Text = ""
                    TextBox6.Text = ""
                    TextBox7.Text = "0"
                    LoadItemAddInfo(Trim(TextBox2.Text))
                Else
                    TextBox1.ReadOnly = False
                    ComboBox1.Enabled = True
                    TextBox5.ReadOnly = False
                    TextBox13.Enabled = True
                    TextBox13.BackColor = Color.FromName("Window")
                    Button7.Enabled = True
                    TextBox14.Enabled = False
                    TextBox14.BackColor = Color.FromName("ButtonFace")
                    LoadItemAddInfo(Trim(TextBox2.Text))
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
                            TextBox2.Text = "Unknown"
                        Else
                            TextBox2.Text = Declarations.MyRec.Fields("SC01001").Value
                        End If
                        TextBox1.Text = Declarations.MyRec.Fields("Name").Value
                        TextBox1.ReadOnly = False
                        ComboBox1.Text = Declarations.MyRec.Fields("SC01135").Value
                        ComboBox1.Enabled = True
                        TextBox4.Text = Declarations.MyRec.Fields("Price").Value / Declarations.CurrencyValue
                        TextBox5.Text = Declarations.MyRec.Fields("PriCost").Value / Declarations.CurrencyValue
                        TextBox5.ReadOnly = False
                        TextBox13.Text = Declarations.MyRec.Fields("SuppID").Value
                        TextBox13.Enabled = True
                        TextBox13.BackColor = Color.FromName("Window")
                        Button7.Enabled = True
                        TextBox14.Text = Declarations.MyRec.Fields("SuppName").Value
                        TextBox14.Enabled = False
                        TextBox14.BackColor = Color.FromName("ButtonFace")
                        Label16.Text = ""
                        Label17.Text = ""
                        Label19.Text = Math.Round(Declarations.MyRec.Fields("MinQTY").Value, 2)
                        Label21.Text = Math.Round(Declarations.MyRec.Fields("LT").Value, 0)
                        TextBox6.Text = 1
                        TextBox3.Text = 1
                        LoadItemAddInfo(Trim(TextBox2.Text))
                        trycloseMyRec()
                    Else                                                '--из Scala
                        TextBox2.Text = Declarations.MyRec.Fields("SC01001").Value
                        TextBox2Validation()
                        LoadItemAddInfo(Trim(TextBox2.Text))
                        trycloseMyRec()
                    End If
                Else
                    MySelectItemBySuppCode = New SelectItemBySuppCode
                    MySelectItemBySuppCode.MyItemSuppCode = Trim(Trim(TextBox9.Text))
                    MySelectItemBySuppCode.MyWindowFrom = "AddToOrder"
                    MySelectItemBySuppCode.ShowDialog()
                End If
            End If
            Else
                '---очищаем все поля 
                TextBox2.Text = ""                      '--код Scala
                TextBox1.Text = ""
                TextBox1.ReadOnly = False
                ComboBox1.Text = ""
                ComboBox1.Enabled = True
                TextBox3.Text = "1"
                TextBox4.Text = ""
                TextBox5.Text = ""
                TextBox5.ReadOnly = False
                TextBox13.Text = ""
                TextBox13.Enabled = True
                TextBox13.BackColor = Color.FromName("Window")
                Button7.Enabled = True
                TextBox14.Text = ""
                TextBox14.Enabled = False
                TextBox14.BackColor = Color.FromName("ButtonFace")
                Label17.Text = ""
                Label19.Text = ""
                Label21.Text = ""
                TextBox6.Text = ""
                TextBox7.Text = "0"
                LoadItemAddInfo(Trim(TextBox2.Text))
            End If

    End Sub

    Private Sub Button6_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button6.Click
        '////////////////////////////////////////////////////////////////////////////////
        '//
        '// Пересчет цены для подгона под указанное значение маржи
        '//
        '////////////////////////////////////////////////////////////////////////////////
        Dim MyDBL As Double

        If Trim(TextBox10.Text) <> "" Then
            If CheckDataFilingZ() = True Then
                Try
                    MyDBL = CDbl(TextBox10.Text)

                    TextBox4.Text = Math.Round(CDbl(TextBox5.Text) * 100 / (100 - MyDBL), 3)
                    TextBox10.Text = ""
                Catch ex As Exception
                    MsgBox("В поле ""Новая маржа %"" должно быть введено число больше или равно 0 и меньше 100", MsgBoxStyle.Critical, "Внимание!")
                End Try
            End If
        Else
            MsgBox("В поле ""Новая маржа %"" должно быть введено число больше или равно 0 и меньше 100", MsgBoxStyle.Critical, "Внимание!")
        End If
    End Sub

    
    Private Sub Button7_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button7.Click
        '////////////////////////////////////////////////////////////////////////////////
        '//
        '// Выбор кода и названия поставщика из списка существующих в Scala
        '//
        '////////////////////////////////////////////////////////////////////////////////

        MySupplierSelect = New SupplierSelect
        MySupplierSelect.MySrcWin = "AddToOrder"
        MySupplierSelect.ShowDialog()
    End Sub

    Private Sub TextBox13_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles TextBox13.KeyDown
        '////////////////////////////////////////////////////////////////////////////////
        '//
        '// переход на следующее поле по Enter
        '//
        '////////////////////////////////////////////////////////////////////////////////

        If e.KeyData = Keys.Enter Then
            Me.SelectNextControl(sender, True, True, True, False)
        End If
    End Sub

    Private Sub TextBox13_Validated(ByVal sender As Object, ByVal e As System.EventArgs) Handles TextBox13.Validated
        '////////////////////////////////////////////////////////////////////////////////
        '//
        '// После ввода кода - проверка, есть ли такой поставщик в базе. Если есть - 
        '// подтягиваем его значения. Если нет - даем ввести свои
        '////////////////////////////////////////////////////////////////////////////////
        If TextBox13.Modified = True Then
            TextBox13Validation()
        End If
    End Sub

    Public Sub TextBox13Validation()
        '////////////////////////////////////////////////////////////////////////////////
        '//
        '// После ввода кода - проверка, есть ли такой поставщик в базе. Если есть - 
        '// подтягиваем его значения. Если нет - даем ввести свои
        '////////////////////////////////////////////////////////////////////////////////
        Dim MySQLStr As String                        'рабочая строка

        MySQLStr = "SELECT PL01001, PL01002 "
        MySQLStr = MySQLStr & "FROM PL010300 "
        MySQLStr = MySQLStr & "WHERE (PL01001 = N'" & TextBox13.Text & "')"
        InitMyConn(False)
        InitMyRec(False, MySQLStr)
        If Declarations.MyRec.EOF = True And Declarations.MyRec.BOF = True Then
            '---такого поставщика нет в БД
            TextBox14.Text = ""
            TextBox14.Enabled = True
            TextBox14.BackColor = Color.FromName("Window")
        Else
            '---такой поставщик в БД есть
            TextBox14.Text = Declarations.MyRec.Fields("PL01002").Value
            TextBox14.Enabled = False
            TextBox14.BackColor = Color.FromName("ButtonFace")
        End If

    End Sub

End Class