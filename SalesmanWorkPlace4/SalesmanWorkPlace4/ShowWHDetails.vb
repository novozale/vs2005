Public Class ShowWHDetails
    Public MyItem As String                           'код запаса


    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        '////////////////////////////////////////////////////////////////////////////////
        '//
        '// Выход из окна 
        '//
        '////////////////////////////////////////////////////////////////////////////////

        Me.Close()
    End Sub

    Private Sub Button2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button2.Click
        '////////////////////////////////////////////////////////////////////////////////////////
        '//
        '// Открытие окна с детальной информацией по складам
        '//
        '////////////////////////////////////////////////////////////////////////////////////////

        MyShowBatchInfo = New ShowBatchInfo
        MyShowBatchInfo.MyItem = Trim(Me.MyItem)
        MyShowBatchInfo.ShowDialog()
    End Sub

    Private Sub ShowWHDetails_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        '////////////////////////////////////////////////////////////////////////////////////////
        '//
        '// Загрузка информации в окно
        '//
        '////////////////////////////////////////////////////////////////////////////////////////
        Dim MySQLStr As String
        Dim MyAdapter As SqlClient.SqlDataAdapter
        Dim MyDs As New DataSet

        Label8.Text = Declarations.CurrencyName
        Label4.Text = Declarations.CurrencyValue

        MySQLStr = "SELECT SC01001 + ' ' + SC01002 as Name, SC01053 "
        MySQLStr = MySQLStr & "FROM SC010300 WITH (NOLOCK) "
        MySQLStr = MySQLStr & "where SC01001 = N'" & Trim(MyItem) & "'"
        InitMyConn(False)
        InitMyRec(False, MySQLStr)
        Label2.Text = Declarations.MyRec.Fields("Name").Value
        If Declarations.MyRec.Fields("SC01053").Value = 0 Then
            Label3.Text = "Рекомендованная цена и себестоимость  для этого запаса должны быть определены самостоятельно"
            Label3.ForeColor = Color.Red
        Else
            Label3.Text = "Рекомендованная цена и себестоимость этого запаса на основе прайс - листа на закупку"
            Label3.ForeColor = Color.DarkGreen
        End If
        trycloseMyRec()

        '---информация о транспортных расходах за рубежом, таможенных пошлинах и закупочной цене
        MySQLStr = "SELECT SC010300.SC01057 AS CustomTax, "
        MySQLStr = MySQLStr & "ISNULL(tbl_SupplierCard0300.ShippingCost, 0) AS ShippingCost, "
        '--Закупочные цены в единицах измерения продаж
        'MySQLStr = MySQLStr & "CONVERT(nvarchar, CONVERT(numeric(28, 2), SC010300.SC01055 / CASE WHEN SC010300.SC01140 = 0 THEN 1 ELSE SC010300.SC01140 END * CASE WHEN SC010300.SC01141 = 0 THEN 1 ELSE SC010300.SC01141 END)) + ' ' + SYCD0100.SYCD009 AS PP "
        MySQLStr = MySQLStr & "CONVERT(nvarchar, CONVERT(numeric(28, 2), CONVERT(Float,SC010300.SC01094) / CASE WHEN SC010300.SC01140 = 0 THEN 1 ELSE SC010300.SC01140 END * CASE WHEN SC010300.SC01141 = 0 THEN 1 ELSE SC010300.SC01141 END)) + ' ' + SYCD0100.SYCD009 AS PP "
        MySQLStr = MySQLStr & "FROM SYCD0100 WITH (NOLOCK) RIGHT OUTER JOIN "
        MySQLStr = MySQLStr & "SC010300 ON SYCD0100.SYCD001 = SC010300.SC01056 LEFT OUTER JOIN "
        MySQLStr = MySQLStr & "tbl_SupplierCard0300 RIGHT OUTER JOIN "
        MySQLStr = MySQLStr & "PL010300 ON tbl_SupplierCard0300.PL01001 = PL010300.PL01001 ON SC010300.SC01058 = PL010300.PL01001 "
        MySQLStr = MySQLStr & "WHERE (SC010300.SC01001 = N'" & Trim(MyItem) & "')"
        InitMyConn(False)
        InitMyRec(False, MySQLStr)
        Label6.Text = Math.Round(Declarations.MyRec.Fields("ShippingCost").Value, 2)
        Label5.Text = Math.Round(Declarations.MyRec.Fields("CustomTax").Value, 2)
        Label7.Text = Declarations.MyRec.Fields("PP").Value
        trycloseMyRec()

        MySQLStr = "SELECT SC230300.SC23001 + ' ' + SC230300.SC23002 AS WH, "
        MySQLStr = MySQLStr & "SC030300.SC03003 AS Qty, "
        MySQLStr = MySQLStr & "SC030300.SC03003 - SUM(SC030300.SC03004) - SUM(SC030300.SC03005) AS Avl, "
        MySQLStr = MySQLStr & "SUM(SC030300.SC03004) AS Rsrv, "
        MySQLStr = MySQLStr & "SUM(SC030300.SC03005) AS Zadl, "
        MySQLStr = MySQLStr & "SUM(SC030300.SC03016) AS Allocated, "
        MySQLStr = MySQLStr & "SC010300.SC01053 / " & Replace(CStr(Declarations.CurrencyValue), ",", ".") & " AS CalcPriCost, "
        'MySQLStr = MySQLStr & "CASE SC030300.SC03003 WHEN 0 THEN 0 ELSE SC030300.SC03057 END / " & Replace(CStr(Declarations.CurrencyValue), ",", ".") & " AS PriCost, "
        MySQLStr = MySQLStr & "CASE SUBSTRING(SC010300.SC01128, CHARINDEX('1', "
        MySQLStr = MySQLStr & "SC230300.SC23007, 1), 1) WHEN '1' THEN '+' ELSE '' END AS Wa "
        MySQLStr = MySQLStr & "FROM SC010300 WITH (NOLOCK) INNER JOIN "
        MySQLStr = MySQLStr & "SC030300 ON SC010300.SC01001 = SC030300.SC03001 RIGHT OUTER JOIN "
        MySQLStr = MySQLStr & "SC230300 ON SC030300.SC03002 = SC230300.SC23001 "
        MySQLStr = MySQLStr & "WHERE (SC230300.SC23006 = N'1') AND (SC030300.SC03001 = N'" & Trim(MyItem) & "') "
        MySQLStr = MySQLStr & "GROUP BY SC230300.SC23001 + ' ' + SC230300.SC23002, "
        MySQLStr = MySQLStr & "SC030300.SC03003, SC030300.SC03057, "
        MySQLStr = MySQLStr & "SUBSTRING(SC010300.SC01128, CHARINDEX('1', SC230300.SC23007, 1), 1), "
        MySQLStr = MySQLStr & "SC010300.SC01053 "

        Try
            MyAdapter = New SqlClient.SqlDataAdapter(MySQLStr, Declarations.MyNETConnStr)
            MyAdapter.SelectCommand.CommandTimeout = 600
            MyAdapter.Fill(MyDs)
            DataGridView1.DataSource = MyDs.Tables(0)
        Catch ex As Exception
            MsgBox(ex.ToString)
        End Try
        DataGridView1.Columns(0).HeaderText = "Склад"
        DataGridView1.Columns(0).Width = 150
        DataGridView1.Columns(1).HeaderText = "Баланс"
        DataGridView1.Columns(2).HeaderText = "Доступно"
        DataGridView1.Columns(3).HeaderText = "Резерв"
        DataGridView1.Columns(4).HeaderText = "Задолж"
        DataGridView1.Columns(5).HeaderText = "Распределено"
        DataGridView1.Columns(6).HeaderText = "Расчетн.Себестоимость"
        'DataGridView1.Columns(7).HeaderText = "Себестоимость"
        DataGridView1.Columns(7).HeaderText = "Складской"

    End Sub

    Private Sub Button3_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button3.Click
        '////////////////////////////////////////////////////////////////////////////////////////
        '//
        '// Открытие окна с информацией об ожидаемом приходе
        '//
        '////////////////////////////////////////////////////////////////////////////////////////

        MyEstimatedIncome = New EstimatedIncome
        MyEstimatedIncome.MyItem = Trim(Me.MyItem)
        MyEstimatedIncome.ShowDialog()
    End Sub
End Class