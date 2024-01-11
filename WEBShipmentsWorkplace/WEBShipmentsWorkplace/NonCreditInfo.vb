Public Class NonCreditInfo

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        '////////////////////////////////////////////////////////////////////////////////////////
        '// Выход из формы
        '//
        '////////////////////////////////////////////////////////////////////////////////////////

        Me.Close()
    End Sub

    Private Sub NonCreditInfo_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        '////////////////////////////////////////////////////////////////////////////////////////
        '// Загрузка информации в форму
        '//
        '////////////////////////////////////////////////////////////////////////////////////////
        Dim MySQLStr As String
        Dim MyAdapter As SqlClient.SqlDataAdapter     '
        Dim MyDs As New DataSet                       '
        Dim MyAdapter1 As SqlClient.SqlDataAdapter     '
        Dim MyDs1 As New DataSet                       '

        '---Незакрытые заказы
        MySQLStr = " EXEC spp_NonCloseOrdersInfoPrepare N'" & Declarations.CustomerID & "'"
        InitMyConn(False)

        Try
            MyAdapter = New SqlClient.SqlDataAdapter(MySQLStr, Declarations.MyNETConnStr)
            MyAdapter.SelectCommand.CommandTimeout = 600
            MyAdapter.Fill(MyDs)
            DataGridView1.DataSource = MyDs.Tables(0)
        Catch ex As Exception
            MsgBox(ex.ToString)
        End Try

        DataGridView1.Columns(0).HeaderText = "Номер заказа"
        DataGridView1.Columns(0).Width = 60
        DataGridView1.Columns(1).HeaderText = "Тип заказа"
        DataGridView1.Columns(1).Width = 60
        DataGridView1.Columns(2).HeaderText = "Код покупателя"
        DataGridView1.Columns(2).Width = 60
        DataGridView1.Columns(3).HeaderText = "Имя покупателя"
        DataGridView1.Columns(3).Width = 150
        DataGridView1.Columns(4).HeaderText = "Код продавца"
        DataGridView1.Columns(4).Width = 60
        DataGridView1.Columns(5).HeaderText = "Имя продавца"
        DataGridView1.Columns(5).Width = 150
        DataGridView1.Columns(6).HeaderText = "Валюта"
        DataGridView1.Columns(6).Width = 50
        DataGridView1.Columns(7).HeaderText = "Сумма заказа"
        DataGridView1.Columns(7).Width = 80
        DataGridView1.Columns(8).HeaderText = "Оплачено"
        DataGridView1.Columns(8).Width = 80
        DataGridView1.Columns(9).HeaderText = "Аванс 1 типа"
        DataGridView1.Columns(9).Width = 80
        DataGridView1.Columns(10).HeaderText = "Аванс 2 типа"
        DataGridView1.Columns(10).Width = 80

        '---Дебиторская задолженность
        MySQLStr = "SELECT View_0.InvoiceNumber, View_0.DeliveryDate, SL030300_3.SL03036 AS OrderNum, "
        MySQLStr = MySQLStr & "SL030300_3.SL03041 AS SalesmanCode, ST010300.ST01002 AS SalesmanName, "
        MySQLStr = MySQLStr & "ISNULL(View_0.InvoiceSum, 0) - ISNULL(View_2.InvoicePayed, 0) AS Debt "
        MySQLStr = MySQLStr & "FROM (SELECT DISTINCT SL03001 AS CustomerCode, "
        MySQLStr = MySQLStr & "LEFT(LTRIM(RTRIM(SL03002)), dbo.fnc_GetFirstNotDigitPos(LTRIM(RTRIM(SL03002))) - 1) AS InvoiceNumber "
        MySQLStr = MySQLStr & "FROM SL030300 WITH (NOLOCK) "
        MySQLStr = MySQLStr & "WHERE (LEFT(SL03017, 6) = '621010') OR "
        MySQLStr = MySQLStr & "(LEFT(SL03017, 6) = '621410')) AS View_5 INNER JOIN "
        MySQLStr = MySQLStr & "SL030300 AS SL030300_3 ON View_5.CustomerCode = SL030300_3.SL03001 "
        MySQLStr = MySQLStr & "AND View_5.InvoiceNumber = SL030300_3.SL03002 LEFT OUTER JOIN "
        MySQLStr = MySQLStr & "ST010300 ON SL030300_3.SL03041 = ST010300.ST01001 LEFT OUTER JOIN "
        MySQLStr = MySQLStr & "(SELECT LEFT(LTRIM(RTRIM(SL030300_2.SL03002)), "
        MySQLStr = MySQLStr & "dbo.fnc_GetFirstNotDigitPos(LTRIM(RTRIM(SL030300_2.SL03002))) - 1) AS invoiceNumber, "
        MySQLStr = MySQLStr & "SUM(SL210300.SL21007) As InvoicePayed "
        MySQLStr = MySQLStr & "FROM SL210300 WITH (NOLOCK) INNER JOIN "
        MySQLStr = MySQLStr & "SL030300 AS SL030300_2 ON SL210300.SL21002 = SL030300_2.SL03002 "
        MySQLStr = MySQLStr & "WHERE (LEFT(SL030300_2.SL03017, 6) = '621010' OR "
        MySQLStr = MySQLStr & "LEFT(SL030300_2.SL03017, 6) = '621410') AND (SL210300.SL21005 <= GETDATE()) "
        MySQLStr = MySQLStr & "GROUP BY LEFT(LTRIM(RTRIM(SL030300_2.SL03002)), "
        MySQLStr = MySQLStr & "dbo.fnc_GetFirstNotDigitPos(LTRIM(RTRIM(SL030300_2.SL03002))) - 1)) AS View_2 ON "
        MySQLStr = MySQLStr & "View_5.InvoiceNumber = View_2.invoiceNumber LEFT OUTER JOIN "
        MySQLStr = MySQLStr & "(SELECT LEFT(LTRIM(RTRIM(SL03002)), "
        MySQLStr = MySQLStr & "dbo.fnc_GetFirstNotDigitPos(LTRIM(RTRIM(SL03002))) - 1) AS InvoiceNumber, "
        MySQLStr = MySQLStr & "MIN(SL03004) AS DeliveryDate, SUM(SL03013) As InvoiceSum "
        MySQLStr = MySQLStr & "FROM SL030300 AS SL030300_1 WITH (NOLOCK) "
        MySQLStr = MySQLStr & "WHERE (LEFT(SL03017, 6) = '621010' OR "
        MySQLStr = MySQLStr & "LEFT(SL03017, 6) = '621410') AND (SL03004 <= GETDATE()) "
        MySQLStr = MySQLStr & "GROUP BY LEFT(LTRIM(RTRIM(SL03002)), "
        MySQLStr = MySQLStr & "dbo.fnc_GetFirstNotDigitPos(LTRIM(RTRIM(SL03002))) - 1)) AS View_0 ON "
        MySQLStr = MySQLStr & "View_5.InvoiceNumber = View_0.InvoiceNumber "
        MySQLStr = MySQLStr & "WHERE (View_0.DeliveryDate <= GETDATE()) AND "
        MySQLStr = MySQLStr & "(ABS(ISNULL(View_0.InvoiceSum, 0) - ISNULL(View_2.InvoicePayed, 0)) > 10) AND "
        MySQLStr = MySQLStr & "(ISNULL(View_0.InvoiceSum, 0)- ISNULL(View_2.InvoicePayed, 0) > 0) AND "
        MySQLStr = MySQLStr & "(View_5.CustomerCode = N'" & Declarations.CustomerID & "') "
        InitMyConn(False)

        Try
            MyAdapter1 = New SqlClient.SqlDataAdapter(MySQLStr, Declarations.MyNETConnStr)
            MyAdapter1.SelectCommand.CommandTimeout = 600
            MyAdapter1.Fill(MyDs1)
            DataGridView2.DataSource = MyDs1.Tables(0)
        Catch ex As Exception
            MsgBox(ex.ToString)
        End Try

        DataGridView2.Columns(0).HeaderText = "Номер СФ"
        DataGridView2.Columns(0).Width = 90
        DataGridView2.Columns(1).HeaderText = "Дата СФ"
        DataGridView2.Columns(1).Width = 90
        DataGridView2.Columns(2).HeaderText = "Номер заказа"
        DataGridView2.Columns(2).Width = 90
        DataGridView2.Columns(3).HeaderText = "Код продавца"
        DataGridView2.Columns(3).Width = 90
        DataGridView2.Columns(4).HeaderText = "Имя продавца"
        DataGridView2.Columns(4).Width = 140
        DataGridView2.Columns(5).HeaderText = "Сумма задолженности"
        DataGridView2.Columns(5).Width = 90
    End Sub
End Class