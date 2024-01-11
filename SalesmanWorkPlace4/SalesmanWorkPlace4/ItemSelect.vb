Public Class ItemSelect

    Public MySrcWin As String                         'окно, из которого вызвано

    Private Sub Button5_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button5.Click
        '////////////////////////////////////////////////////////////////////////////////
        '//
        '// Выход из окна с выбором запаса
        '//
        '////////////////////////////////////////////////////////////////////////////////

        If MySrcWin = "EditInOrder" Then
            MyEditInOrder.TextBox1.Text = Trim(Me.DataGridView1.SelectedRows.Item(0).Cells(1).Value.ToString)
            MyEditInOrder.TextBox1Validation()
        ElseIf MySrcWin = "AddItem" Then
            MyAddItem.TextBox1.Text = Trim(Me.DataGridView1.SelectedRows.Item(0).Cells(1).Value.ToString())
            MyAddItem.TextBox1Validating()
            MyAddItem.TextBox1Validated()
        End If
        Me.Close()
    End Sub

    Private Sub Button7_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button7.Click
        '////////////////////////////////////////////////////////////////////////////////
        '//
        '// Выход из окна без выбора запаса
        '//
        '////////////////////////////////////////////////////////////////////////////////

        Me.Close()
    End Sub

    Private Sub ItemSelect_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        '////////////////////////////////////////////////////////////////////////////////
        '//
        '// Загрузка данных в форму
        '//
        '////////////////////////////////////////////////////////////////////////////////

        DataPreparation(0)
        Label3.Text = ""
        ChangeButtonsStatus()
    End Sub

    Private Function DataPreparation(ByVal MyParam As Integer)
        '////////////////////////////////////////////////////////////////////////////////////////
        '//
        '// Формирование списка продуктов
        '//
        '////////////////////////////////////////////////////////////////////////////////////////
        Dim MySQLStr As String                   'рабочая строка
        Dim PriceNum As String                   'прайс - лист
        Dim MyAdapter As SqlClient.SqlDataAdapter
        Dim MyDs As New DataSet
        Dim WhName As String                     'название склада

        '---прайс - лист
        If MySrcWin = "EditInOrder" Then
            MySQLStr = "SELECT OR01051 "
            MySQLStr = MySQLStr & "FROM tbl_OR010300 WITH (NOLOCK) "
            MySQLStr = MySQLStr & "WHERE (OR01001 = N'" & Declarations.MyOrderNum & "')"
            InitMyConn(False)
            InitMyRec(False, MySQLStr)
            PriceNum = Declarations.MyRec.Fields("OR01051").Value
            trycloseMyRec()
        ElseIf MySrcWin = "AddItem" Then
            PriceNum = "00"
        End If

        '---номер склада
        If MySrcWin = "EditInOrder" Then
            MySQLStr = "SELECT SC23001, SC23002 "
            MySQLStr = MySQLStr & "FROM SC230300 WITH (NOLOCK) "
            MySQLStr = MySQLStr & "WHERE (SC23001 IN "
            MySQLStr = MySQLStr & "(SELECT OR01050 "
            MySQLStr = MySQLStr & "FROM tbl_OR010300 "
            MySQLStr = MySQLStr & "WHERE (OR01001 = N'" & Declarations.MyOrderNum & "')))"
            InitMyConn(False)
            InitMyRec(False, MySQLStr)
            If Declarations.MyRec.BOF = True And Declarations.MyRec.EOF = True Then
                trycloseMyRec()
                MySQLStr = "SELECT SC23001, SC23002 "
                MySQLStr = MySQLStr & "FROM SC230300 WITH (NOLOCK) "
                MySQLStr = MySQLStr & "WHERE (SC23001 = '01') "
                InitMyConn(False)
                InitMyRec(False, MySQLStr)
            End If
            Declarations.WHNum = Declarations.MyRec.Fields("SC23001").Value
            WhName = Declarations.MyRec.Fields("SC23002").Value
            trycloseMyRec()
        ElseIf MySrcWin = "AddItem" Then
            MySQLStr = "SELECT SC23001, SC23002 "
            MySQLStr = MySQLStr & "FROM SC230300 WITH (NOLOCK) "
            MySQLStr = MySQLStr & "WHERE (SC23001 = '01') "
            InitMyConn(False)
            InitMyRec(False, MySQLStr)
            Declarations.WHNum = Declarations.MyRec.Fields("SC23001").Value
            WhName = Declarations.MyRec.Fields("SC23002").Value
            trycloseMyRec()
        End If

        'MySQLStr = "SELECT tbl_WEB_Pictures.PictureSmall, SC010300.SC01001 AS ID, "
        MySQLStr = "SELECT NULL, SC010300.SC01001 AS ID, "
        MySQLStr = MySQLStr & "SC010300.SC01002 + ' ' + SC010300.SC01003 AS Name, "
        If MySrcWin = "EditInOrder" Then
            MySQLStr = MySQLStr & "Round(ISNULL(t2.SC39005, 0) / " & Replace(CStr(Declarations.CurrencyValue), ",", ".") & ",2) AS Price, "
        ElseIf MySrcWin = "AddItem" Then
            MySQLStr = MySQLStr & "Round(ISNULL(t2.SC39005, 0),2) AS Price, "
        End If
        MySQLStr = MySQLStr & "SC010300.SC01060 AS SuppID, "
        MySQLStr = MySQLStr & "View_1.txt AS UnitName, "
        MySQLStr = MySQLStr & "SC010300.SC01042 AS TotalQty, "
        MySQLStr = MySQLStr & "ISNULL(t1.WhQty, 0) AS WhQty, "
        MySQLStr = MySQLStr & "ISNULL(t1.WhAvl, 0) AS WhAvl, "
        MySQLStr = MySQLStr & "SC010300.SC01058 AS SuppCode, "
        MySQLStr = MySQLStr & "ISNULL(PL010300.PL01002, N'') + ' ' + ISNULL(PL010300.PL01003, N'') AS SuppName, "
        '---поставщик "неликвидов"
        MySQLStr = MySQLStr & "CASE WHEN Ltrim(Rtrim(SC010300.SC01058)) = 'FIN001_U' THEN '+' ELSE '' END AS Dead, "
        MySQLStr = MySQLStr & "LTRIM(RTRIM(tbl_ItemCard0300.ManufacturerItemCode)) AS ManufacturerItemCode "
        MySQLStr = MySQLStr & "FROM SC010300 WITH (NOLOCK) INNER JOIN "
        MySQLStr = MySQLStr & "(SELECT 0 AS Expr1, SC09002 AS txt "
        MySQLStr = MySQLStr & "FROM SC090300 WITH (NOLOCK) "
        MySQLStr = MySQLStr & "WHERE (SC09001 = 'RUS') "
        MySQLStr = MySQLStr & "UNION "
        MySQLStr = MySQLStr & "SELECT 1 AS Expr1, SC09003 "
        MySQLStr = MySQLStr & "FROM SC090300 AS SC090300_40 WITH (NOLOCK) "
        MySQLStr = MySQLStr & "WHERE (SC09001 = 'RUS') "
        MySQLStr = MySQLStr & "UNION "
        MySQLStr = MySQLStr & "SELECT 2 AS Expr1, SC09004 "
        MySQLStr = MySQLStr & "FROM SC090300 AS SC090300_39 WITH (NOLOCK) "
        MySQLStr = MySQLStr & "WHERE (SC09001 = 'RUS') "
        MySQLStr = MySQLStr & "UNION "
        MySQLStr = MySQLStr & "SELECT 3 AS Expr1, SC09005 "
        MySQLStr = MySQLStr & "FROM SC090300 AS SC090300_38 WITH (NOLOCK) "
        MySQLStr = MySQLStr & "WHERE (SC09001 = 'RUS') "
        MySQLStr = MySQLStr & "UNION "
        MySQLStr = MySQLStr & "SELECT 4 AS Expr1, SC09006 "
        MySQLStr = MySQLStr & "FROM SC090300 AS SC090300_37 WITH (NOLOCK) "
        MySQLStr = MySQLStr & "WHERE (SC09001 = 'RUS') "
        MySQLStr = MySQLStr & "UNION "
        MySQLStr = MySQLStr & "SELECT 5 AS Expr1, SC09007 "
        MySQLStr = MySQLStr & "FROM SC090300 AS SC090300_36 WITH (NOLOCK) "
        MySQLStr = MySQLStr & "WHERE (SC09001 = 'RUS') "
        MySQLStr = MySQLStr & "UNION "
        MySQLStr = MySQLStr & "SELECT 6 AS Expr1, SC09008 "
        MySQLStr = MySQLStr & "FROM SC090300 AS SC090300_35 WITH (NOLOCK) "
        MySQLStr = MySQLStr & "WHERE (SC09001 = 'RUS') "
        MySQLStr = MySQLStr & "UNION "
        MySQLStr = MySQLStr & "SELECT 7 AS Expr1, SC09009 "
        MySQLStr = MySQLStr & "FROM SC090300 AS SC090300_34 WITH (NOLOCK) "
        MySQLStr = MySQLStr & "WHERE (SC09001 = 'RUS') "
        MySQLStr = MySQLStr & "UNION "
        MySQLStr = MySQLStr & "SELECT 8 AS Expr1, SC09010 "
        MySQLStr = MySQLStr & "FROM SC090300 AS SC090300_33 WITH (NOLOCK) "
        MySQLStr = MySQLStr & "WHERE (SC09001 = 'RUS') "
        MySQLStr = MySQLStr & "UNION "
        MySQLStr = MySQLStr & "SELECT 9 AS Expr1, SC09011 "
        MySQLStr = MySQLStr & "FROM SC090300 AS SC090300_32 WITH (NOLOCK) "
        MySQLStr = MySQLStr & "WHERE (SC09001 = 'RUS') "
        MySQLStr = MySQLStr & "UNION "
        MySQLStr = MySQLStr & "SELECT 10 AS Expr1, SC09012 "
        MySQLStr = MySQLStr & "FROM SC090300 AS SC090300_31 WITH (NOLOCK) "
        MySQLStr = MySQLStr & "WHERE (SC09001 = 'RUS') "
        MySQLStr = MySQLStr & "UNION "
        MySQLStr = MySQLStr & "SELECT 11 AS Expr1, SC09013 "
        MySQLStr = MySQLStr & "FROM SC090300 AS SC090300_30 WITH (NOLOCK) "
        MySQLStr = MySQLStr & "WHERE (SC09001 = 'RUS') "
        MySQLStr = MySQLStr & "UNION "
        MySQLStr = MySQLStr & "SELECT 12 AS Expr1, SC09014 "
        MySQLStr = MySQLStr & "FROM SC090300 AS SC090300_29 WITH (NOLOCK) "
        MySQLStr = MySQLStr & "WHERE (SC09001 = 'RUS') "
        MySQLStr = MySQLStr & "UNION "
        MySQLStr = MySQLStr & "SELECT 13 AS Expr1, SC09015 "
        MySQLStr = MySQLStr & "FROM SC090300 AS SC090300_28 WITH (NOLOCK) "
        MySQLStr = MySQLStr & "WHERE (SC09001 = 'RUS') "
        MySQLStr = MySQLStr & "UNION "
        MySQLStr = MySQLStr & "SELECT 14 AS Expr1, SC09016 "
        MySQLStr = MySQLStr & "FROM SC090300 AS SC090300_27 WITH (NOLOCK) "
        MySQLStr = MySQLStr & "WHERE (SC09001 = 'RUS') "
        MySQLStr = MySQLStr & "UNION "
        MySQLStr = MySQLStr & "SELECT 15 AS Expr1, SC09017 "
        MySQLStr = MySQLStr & "FROM SC090300 AS SC090300_26 WITH (NOLOCK) "
        MySQLStr = MySQLStr & "WHERE (SC09001 = 'RUS') "
        MySQLStr = MySQLStr & "UNION "
        MySQLStr = MySQLStr & "SELECT 16 AS Expr1, SC09018 "
        MySQLStr = MySQLStr & "FROM SC090300 AS SC090300_25 WITH (NOLOCK) "
        MySQLStr = MySQLStr & "WHERE (SC09001 = 'RUS') "
        MySQLStr = MySQLStr & "UNION "
        MySQLStr = MySQLStr & "SELECT 17 AS Expr1, SC09019 "
        MySQLStr = MySQLStr & "FROM SC090300 AS SC090300_24 WITH (NOLOCK) "
        MySQLStr = MySQLStr & "WHERE (SC09001 = 'RUS') "
        MySQLStr = MySQLStr & "UNION "
        MySQLStr = MySQLStr & "SELECT 18 AS Expr1, SC09020 "
        MySQLStr = MySQLStr & "FROM SC090300 AS SC090300_23 WITH (NOLOCK) "
        MySQLStr = MySQLStr & "WHERE (SC09001 = 'RUS') "
        MySQLStr = MySQLStr & "UNION "
        MySQLStr = MySQLStr & "SELECT 19 AS Expr1, SC09021 "
        MySQLStr = MySQLStr & "FROM SC090300 AS SC090300_22 WITH (NOLOCK) "
        MySQLStr = MySQLStr & "WHERE (SC09001 = 'RUS') "
        MySQLStr = MySQLStr & "UNION "
        MySQLStr = MySQLStr & "SELECT 20 AS Expr1, SC09022 "
        MySQLStr = MySQLStr & "FROM SC090300 AS SC090300_21 WITH (NOLOCK) "
        MySQLStr = MySQLStr & "WHERE (SC09001 = 'RUS') "
        MySQLStr = MySQLStr & "UNION "
        MySQLStr = MySQLStr & "SELECT 21 AS Expr1, SC09023 "
        MySQLStr = MySQLStr & "FROM SC090300 AS SC090300_20 WITH (NOLOCK) "
        MySQLStr = MySQLStr & "WHERE (SC09001 = 'RUS') "
        MySQLStr = MySQLStr & "UNION "
        MySQLStr = MySQLStr & "SELECT 22 AS Expr1, SC09024 "
        MySQLStr = MySQLStr & "FROM SC090300 AS SC090300_19 WITH (NOLOCK) "
        MySQLStr = MySQLStr & "WHERE (SC09001 = 'RUS') "
        MySQLStr = MySQLStr & "UNION "
        MySQLStr = MySQLStr & "SELECT 23 AS Expr1, SC09025 "
        MySQLStr = MySQLStr & "FROM SC090300 AS SC090300_18 WITH (NOLOCK) "
        MySQLStr = MySQLStr & "WHERE (SC09001 = 'RUS') "
        MySQLStr = MySQLStr & "UNION "
        MySQLStr = MySQLStr & "SELECT 24 AS Expr1, SC09026 "
        MySQLStr = MySQLStr & "FROM SC090300 AS SC090300_17 WITH (NOLOCK) "
        MySQLStr = MySQLStr & "WHERE (SC09001 = 'RUS') "
        MySQLStr = MySQLStr & "UNION "
        MySQLStr = MySQLStr & "SELECT 25 AS Expr1, SC09027 "
        MySQLStr = MySQLStr & "FROM SC090300 AS SC090300_16 WITH (NOLOCK) "
        MySQLStr = MySQLStr & "WHERE (SC09001 = 'RUS') "
        MySQLStr = MySQLStr & "UNION "
        MySQLStr = MySQLStr & "SELECT 26 AS Expr1, SC09028 "
        MySQLStr = MySQLStr & "FROM SC090300 AS SC090300_15 WITH (NOLOCK) "
        MySQLStr = MySQLStr & "WHERE (SC09001 = 'RUS') "
        MySQLStr = MySQLStr & "UNION "
        MySQLStr = MySQLStr & "SELECT 27 AS Expr1, SC09029 "
        MySQLStr = MySQLStr & "FROM SC090300 AS SC090300_14 WITH (NOLOCK) "
        MySQLStr = MySQLStr & "WHERE (SC09001 = 'RUS') "
        MySQLStr = MySQLStr & "UNION "
        MySQLStr = MySQLStr & "SELECT 28 AS Expr1, SC09030 "
        MySQLStr = MySQLStr & "FROM SC090300 AS SC090300_13 WITH (NOLOCK)"
        MySQLStr = MySQLStr & "WHERE (SC09001 = 'RUS') "
        MySQLStr = MySQLStr & "UNION "
        MySQLStr = MySQLStr & "SELECT 29 AS Expr1, SC09031 "
        MySQLStr = MySQLStr & "FROM SC090300 AS SC090300_12 WITH (NOLOCK) "
        MySQLStr = MySQLStr & "WHERE (SC09001 = 'RUS') "
        MySQLStr = MySQLStr & "UNION "
        MySQLStr = MySQLStr & "SELECT 30 AS Expr1, SC09032 "
        MySQLStr = MySQLStr & "FROM SC090300 AS SC090300_11 WITH (NOLOCK) "
        MySQLStr = MySQLStr & "WHERE (SC09001 = 'RUS') "
        MySQLStr = MySQLStr & "UNION "
        MySQLStr = MySQLStr & "SELECT 31 AS Expr1, SC09033 "
        MySQLStr = MySQLStr & "FROM SC090300 AS SC090300_10 WITH (NOLOCK) "
        MySQLStr = MySQLStr & "WHERE (SC09001 = 'RUS') "
        MySQLStr = MySQLStr & "UNION "
        MySQLStr = MySQLStr & "SELECT 32 AS Expr1, SC09034 "
        MySQLStr = MySQLStr & "FROM SC090300 AS SC090300_9 WITH (NOLOCK) "
        MySQLStr = MySQLStr & "WHERE (SC09001 = 'RUS') "
        MySQLStr = MySQLStr & "UNION "
        MySQLStr = MySQLStr & "SELECT 33 AS Expr1, SC09035 "
        MySQLStr = MySQLStr & "FROM SC090300 AS SC090300_8 WITH (NOLOCK) "
        MySQLStr = MySQLStr & "WHERE (SC09001 = 'RUS') "
        MySQLStr = MySQLStr & "UNION "
        MySQLStr = MySQLStr & "SELECT 34 AS Expr1, SC09036 "
        MySQLStr = MySQLStr & "FROM SC090300 AS SC090300_7 WITH (NOLOCK) "
        MySQLStr = MySQLStr & "WHERE (SC09001 = 'RUS') "
        MySQLStr = MySQLStr & "UNION "
        MySQLStr = MySQLStr & "SELECT 35 AS Expr1, SC09037 "
        MySQLStr = MySQLStr & "FROM SC090300 AS SC090300_6 WITH (NOLOCK) "
        MySQLStr = MySQLStr & "WHERE (SC09001 = 'RUS') "
        MySQLStr = MySQLStr & "UNION "
        MySQLStr = MySQLStr & "SELECT 36 AS Expr1, SC09038 "
        MySQLStr = MySQLStr & "FROM SC090300 AS SC090300_5 WITH (NOLOCK) "
        MySQLStr = MySQLStr & "WHERE (SC09001 = 'RUS') "
        MySQLStr = MySQLStr & "UNION "
        MySQLStr = MySQLStr & "SELECT 37 AS Expr1, SC09039 "
        MySQLStr = MySQLStr & "FROM SC090300 AS SC090300_4 WITH (NOLOCK) "
        MySQLStr = MySQLStr & "WHERE (SC09001 = 'RUS') "
        MySQLStr = MySQLStr & "UNION "
        MySQLStr = MySQLStr & "SELECT 38 AS Expr1, SC09040 "
        MySQLStr = MySQLStr & "FROM SC090300 AS SC090300_3 WITH (NOLOCK) "
        MySQLStr = MySQLStr & "WHERE (SC09001 = 'RUS') "
        MySQLStr = MySQLStr & "UNION "
        MySQLStr = MySQLStr & "SELECT 39 AS Expr1, SC09041 "
        MySQLStr = MySQLStr & "FROM SC090300 AS SC090300_2 WITH (NOLOCK) "
        MySQLStr = MySQLStr & "WHERE (SC09001 = 'RUS') "
        MySQLStr = MySQLStr & "UNION "
        MySQLStr = MySQLStr & "SELECT 40 AS Expr1, SC09042 "
        MySQLStr = MySQLStr & "FROM SC090300 AS SC090300_1 WITH (NOLOCK) ) AS View_1 "
        MySQLStr = MySQLStr & "ON SC010300.SC01135 = View_1.Expr1 LEFT OUTER JOIN "
        MySQLStr = MySQLStr & "tbl_WEB_Pictures ON SC010300.SC01001 = tbl_WEB_Pictures.ScalaItemCode LEFT OUTER JOIN "
        MySQLStr = MySQLStr & "PL010300 ON SC010300.SC01058 = PL010300.PL01001 LEFT OUTER JOIN "
        MySQLStr = MySQLStr & "tbl_ItemCard0300 ON SC010300.SC01001 = tbl_ItemCard0300.SC01001 LEFT OUTER JOIN "
        MySQLStr = MySQLStr & "(SELECT SC03001, SC03003 AS WhQty, "
        MySQLStr = MySQLStr & "SC03003 - SC03004 - SC03005 AS WhAvl "
        MySQLStr = MySQLStr & "FROM SC030300 WITH (NOLOCK) "
        MySQLStr = MySQLStr & "WHERE (SC03002 = N'" & Declarations.WHNum & "')) AS t1 ON "
        MySQLStr = MySQLStr & "SC010300.SC01001 = t1.SC03001 LEFT OUTER JOIN "
        MySQLStr = MySQLStr & "(SELECT SC39001, SC39005 "
        MySQLStr = MySQLStr & "FROM SC390300 WITH (NOLOCK) "
        MySQLStr = MySQLStr & "WHERE (SC39002 = N'00')) AS t2 ON "
        MySQLStr = MySQLStr & "SC010300.SC01001 = t2.SC39001 "
        MySQLStr = MySQLStr & "WHERE (SC010300.SC01001 <> N'00000000') AND "
        MySQLStr = MySQLStr & "(tbl_ItemCard0300.IsBlocked = N'0') AND "
        MySQLStr = MySQLStr & "(Ltrim(Rtrim(SC010300.SC01066)) <> N'8')"
        If Trim(TextBox1.Text) = "" Then
        Else
            MySQLStr = MySQLStr & "AND (SC010300.SC01058 = N'" & Trim(TextBox1.Text) & "') "
        End If
        If MyParam = 1 Then '---выводятся только доступные продукты
            MySQLStr = MySQLStr & " AND (WhAvl > 0) "
        End If
        MySQLStr = MySQLStr & "ORDER BY dbo.SC010300.SC01001  "

        DataGridView1.RowTemplate.MinimumHeight = 35

        Try
            MyAdapter = New SqlClient.SqlDataAdapter(MySQLStr, Declarations.MyNETConnStr)
            MyAdapter.SelectCommand.CommandTimeout = 600
            MyAdapter.Fill(MyDs)
            DataGridView1.DataSource = MyDs.Tables(0)
        Catch ex As Exception
            MsgBox(ex.ToString)
        End Try

        DataGridView1.Columns(0).HeaderText = "Кар тин ка"
        DataGridView1.Columns(0).Width = 35
        DataGridView1.Columns(0).Visible = False
        DataGridView1.Columns(1).HeaderText = "ID продукта"
        DataGridView1.Columns(1).Width = 110
        DataGridView1.Columns(2).HeaderText = "Имя продукта"
        DataGridView1.Columns(2).Width = 300
        DataGridView1.Columns(3).HeaderText = "Прайс"
        DataGridView1.Columns(3).Width = 80
        DataGridView1.Columns(4).HeaderText = "Код тов постав"
        DataGridView1.Columns(4).Width = 110
        DataGridView1.Columns(5).HeaderText = "Ед изм"
        DataGridView1.Columns(5).Width = 40
        DataGridView1.Columns(6).HeaderText = "Всего на складах"
        DataGridView1.Columns(6).Width = 115
        DataGridView1.Columns(7).HeaderText = "Ост" + " " + Declarations.WHNum + " " + WhName
        DataGridView1.Columns(7).Width = 115
        DataGridView1.Columns(8).HeaderText = "Дост" + " " + Declarations.WHNum + " " + WhName
        DataGridView1.Columns(8).Width = 115
        DataGridView1.Columns(9).HeaderText = "Поставщик ID"
        DataGridView1.Columns(9).Width = 70
        DataGridView1.Columns(10).HeaderText = "Поставщик"
        DataGridView1.Columns(10).Width = 200
        DataGridView1.Columns(11).HeaderText = "Неликвид"
        DataGridView1.Columns(11).Width = 70
        DataGridView1.Columns(12).HeaderText = "Код тов производителя"
        DataGridView1.Columns(12).Width = 150

    End Function

    Private Sub ChangeButtonsStatus()
        '////////////////////////////////////////////////////////////////////////////////////////
        '//
        '// смена состояния кнопки "Выбрать запас" 
        '//
        '////////////////////////////////////////////////////////////////////////////////////////

        If DataGridView1.SelectedRows.Count = 0 Then
            Button5.Enabled = False
        Else
            Button5.Enabled = True
        End If
    End Sub

    Private Sub DataGridView1_CellFormatting(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellFormattingEventArgs) Handles DataGridView1.CellFormatting
        '////////////////////////////////////////////////////////////////////////////////////////
        '//
        '// смена выделения неликвидов
        '//
        '////////////////////////////////////////////////////////////////////////////////////////
        Dim row As DataGridViewRow

        row = DataGridView1.Rows(e.RowIndex)
        If row.Cells(11).Value.ToString = "+" Then
            row.DefaultCellStyle.Font = New Font(row.InheritedStyle.Font, FontStyle.Bold)
            row.DefaultCellStyle.ForeColor = Color.Red
        End If
    End Sub

    Private Sub Button3_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button3.Click
        '////////////////////////////////////////////////////////////////////////////////////////
        '//
        '// Открытие окна со списком поставщиков
        '//
        '////////////////////////////////////////////////////////////////////////////////////////

        MySupplierSelect = New SupplierSelect
        MySupplierSelect.MySrcWin = "ItemSelect"
        MySupplierSelect.ShowDialog()
    End Sub

    Public Function RefreshProductList()
        '////////////////////////////////////////////////////////////////////////////////////////
        '//
        '// ф-ция обновляет список продуктов Электроскандии для последующего выбора
        '//
        '////////////////////////////////////////////////////////////////////////////////////////

        DataPreparation(0)
        ChangeButtonsStatus()
    End Function

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
        '// Введен код поставщика - Обновляем список продуктов
        '//
        '////////////////////////////////////////////////////////////////////////////////////////

        If TextBox1.Modified = True Then
            TextBox2.Text = ""
            TextBox3.Text = ""
            System.Windows.Forms.Cursor.Current = Cursors.WaitCursor
            RefreshProductList()
            System.Windows.Forms.Cursor.Current = Cursors.Default
        End If
    End Sub

    Private Sub TextBox1_Validating(ByVal sender As Object, ByVal e As System.ComponentModel.CancelEventArgs) Handles TextBox1.Validating
        '////////////////////////////////////////////////////////////////////////////////////////
        '//
        '// Введен код поставщика - находим и подписываем его название
        '//
        '////////////////////////////////////////////////////////////////////////////////////////
        Dim MySQLStr As String

        If TextBox1.Modified = True Then
            If Trim(TextBox1.Text) = "" Then
                Label3.Text = ""
            Else
                MySQLStr = "SELECT PL01002, PL01003 + ' ' + PL01004 + ' ' + PL01005 AS PL01003 "
                MySQLStr = MySQLStr & "FROM PL010300 WITH (NOLOCK) "
                MySQLStr = MySQLStr & "WHERE (PL01001 = N'" & Trim(MyItemSelect.TextBox1.Text) & "') "
                InitMyConn(False)
                InitMyRec(False, MySQLStr)
                If Declarations.MyRec.EOF = True And Declarations.MyRec.BOF = True Then
                    trycloseMyRec()
                    MsgBox("Вы ввели неверный код поставщика. Введите корректный или воспользуйтесь поиском.", vbCritical, "Внимание!")
                    Label3.Text = ""
                    e.Cancel = True
                    Exit Sub
                Else
                    Label3.Text = Declarations.MyRec.Fields("PL01002").Value & " " & Declarations.MyRec.Fields("PL01003").Value
                    trycloseMyRec()
                End If
            End If
        End If
    End Sub

    Private Sub Button4_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button4.Click
        '////////////////////////////////////////////////////////////////////////////////
        '//
        '// Поиск первого подходящего по критерию запаса
        '//
        '////////////////////////////////////////////////////////////////////////////////

        If Trim(TextBox2.Text) = "" And Trim(TextBox3.Text) = "" Then
        Else
            System.Windows.Forms.Cursor.Current = Cursors.WaitCursor
            For i As Integer = 0 To DataGridView1.Rows.Count - 1
                If InStr(UCase(Trim(DataGridView1.Item(1, i).Value.ToString)), UCase(Trim(TextBox2.Text))) <> 0 And InStr(UCase(Trim(DataGridView1.Item(1, i).Value.ToString)), UCase(Trim(TextBox3.Text))) <> 0 _
                    Or InStr(UCase(Trim(DataGridView1.Item(2, i).Value.ToString)), UCase(Trim(TextBox2.Text))) <> 0 And InStr(UCase(Trim(DataGridView1.Item(2, i).Value.ToString)), UCase(Trim(TextBox3.Text))) <> 0 _
                    Or InStr(UCase(Trim(DataGridView1.Item(4, i).Value.ToString)), UCase(Trim(TextBox2.Text))) <> 0 And InStr(UCase(Trim(DataGridView1.Item(4, i).Value.ToString)), UCase(Trim(TextBox3.Text))) <> 0 _
                    Or InStr(UCase(Trim(DataGridView1.Item(12, i).Value.ToString)), UCase(Trim(TextBox2.Text))) <> 0 And InStr(UCase(Trim(DataGridView1.Item(12, i).Value.ToString)), UCase(Trim(TextBox3.Text))) <> 0 Then
                    DataGridView1.CurrentCell = DataGridView1.Item(1, i)
                    System.Windows.Forms.Cursor.Current = Cursors.Default
                    Exit Sub
                End If
            Next
            System.Windows.Forms.Cursor.Current = Cursors.Default
        End If
    End Sub

    Private Sub Button2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button2.Click
        '////////////////////////////////////////////////////////////////////////////////
        '//
        '// Поиск следующего подходящего по критерию поставщика
        '//
        '////////////////////////////////////////////////////////////////////////////////
        Dim MyRez As Object

        If Trim(TextBox1.Text) = "" And Trim(TextBox2.Text) = "" Then
        Else
            System.Windows.Forms.Cursor.Current = Cursors.WaitCursor
            For i As Integer = DataGridView1.CurrentCellAddress.Y + 1 To DataGridView1.Rows.Count
                If i = DataGridView1.Rows.Count Then
                    MyRez = MsgBox("Поиск дошел до конца списка. Начать сначала?", MsgBoxStyle.YesNo, "Внимание!")
                    If MyRez = 6 Then
                        i = 0
                    Else
                        System.Windows.Forms.Cursor.Current = Cursors.Default
                        Exit Sub
                    End If
                End If
                If DataGridView1.Rows.Count = 0 Then
                Else
                    If InStr(UCase(Trim(DataGridView1.Item(1, i).Value.ToString)), UCase(Trim(TextBox2.Text))) <> 0 And InStr(UCase(Trim(DataGridView1.Item(1, i).Value.ToString)), UCase(Trim(TextBox3.Text))) <> 0 _
                        Or InStr(UCase(Trim(DataGridView1.Item(2, i).Value.ToString)), UCase(Trim(TextBox2.Text))) <> 0 And InStr(UCase(Trim(DataGridView1.Item(2, i).Value.ToString)), UCase(Trim(TextBox3.Text))) <> 0 _
                        Or InStr(UCase(Trim(DataGridView1.Item(4, i).Value.ToString)), UCase(Trim(TextBox2.Text))) <> 0 And InStr(UCase(Trim(DataGridView1.Item(4, i).Value.ToString)), UCase(Trim(TextBox3.Text))) <> 0 _
                        Or InStr(UCase(Trim(DataGridView1.Item(12, i).Value.ToString)), UCase(Trim(TextBox2.Text))) <> 0 And InStr(UCase(Trim(DataGridView1.Item(12, i).Value.ToString)), UCase(Trim(TextBox3.Text))) <> 0 Then
                        DataGridView1.CurrentCell = DataGridView1.Item(1, i)
                        System.Windows.Forms.Cursor.Current = Cursors.Default
                        Exit Sub
                    End If
                End If
            Next i
            System.Windows.Forms.Cursor.Current = Cursors.Default
        End If
    End Sub

    Private Sub Button6_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button6.Click
        '////////////////////////////////////////////////////////////////////////////////
        '//
        '// Подсвечивание всех подходящих по критерию запасов
        '//
        '////////////////////////////////////////////////////////////////////////////////

        If Button6.Text = "Подсветить все" Then
            System.Windows.Forms.Cursor.Current = Cursors.WaitCursor
            For i As Integer = 0 To DataGridView1.Rows.Count - 1
                If InStr(UCase(Trim(DataGridView1.Item(1, i).Value.ToString)), UCase(Trim(TextBox2.Text))) <> 0 And InStr(UCase(Trim(DataGridView1.Item(1, i).Value.ToString)), UCase(Trim(TextBox3.Text))) <> 0 _
                    Or InStr(UCase(Trim(DataGridView1.Item(2, i).Value.ToString)), UCase(Trim(TextBox2.Text))) <> 0 And InStr(UCase(Trim(DataGridView1.Item(2, i).Value.ToString)), UCase(Trim(TextBox3.Text))) <> 0 _
                    Or InStr(UCase(Trim(DataGridView1.Item(4, i).Value.ToString)), UCase(Trim(TextBox2.Text))) <> 0 And InStr(UCase(Trim(DataGridView1.Item(4, i).Value.ToString)), UCase(Trim(TextBox3.Text))) <> 0 _
                    Or InStr(UCase(Trim(DataGridView1.Item(12, i).Value.ToString)), UCase(Trim(TextBox2.Text))) <> 0 And InStr(UCase(Trim(DataGridView1.Item(12, i).Value.ToString)), UCase(Trim(TextBox3.Text))) <> 0 Then
                    DataGridView1.Rows(i).DefaultCellStyle.BackColor = Color.Yellow
                Else
                    DataGridView1.Rows(i).DefaultCellStyle.BackColor = Color.Empty
                End If
            Next
            System.Windows.Forms.Cursor.Current = Cursors.Default
            Button6.Text = "Снять выдел."
        Else
            System.Windows.Forms.Cursor.Current = Cursors.WaitCursor
            For i As Integer = 0 To DataGridView1.Rows.Count - 1
                DataGridView1.Rows(i).DefaultCellStyle.BackColor = Color.Empty
            Next
            System.Windows.Forms.Cursor.Current = Cursors.Default
            Button6.Text = "Подсветить все"
        End If
    End Sub

    Private Sub DataGridView1_ColumnHeaderMouseClick(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellMouseEventArgs) Handles DataGridView1.ColumnHeaderMouseClick
        '/////////////////////////////////////////////////////////////////////////////////////
        '//
        '// Щелчок по заголовку таблицы 
        '// 
        '/////////////////////////////////////////////////////////////////////////////////////

        Button6.Text = "Подсветить все"
    End Sub

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        '////////////////////////////////////////////////////////////////////////////////
        '//
        '// Выбор всех подходящих по критерию запасов в отдельное окно
        '//
        '////////////////////////////////////////////////////////////////////////////////

        If Trim(TextBox2.Text) = "" And Trim(TextBox3.Text) = "" Then
            MsgBox("Необходимо ввести критерий поиска", MsgBoxStyle.OkOnly, "Внимание!")
            TextBox2.Select()
        Else
            MyItemSelectList = New ItemSelectList
            MyItemSelectList.MySrcWin = "ItemSelect"
            MyItemSelectList.ShowDialog()
        End If
    End Sub

    Private Sub DataGridView1_SelectionChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles DataGridView1.SelectionChanged
        '////////////////////////////////////////////////////////////////////////////////////////
        '//
        '// смена состояния кнопки "Выбрать запас" 
        '//
        '////////////////////////////////////////////////////////////////////////////////////////

        ChangeButtonsStatus()
    End Sub
End Class