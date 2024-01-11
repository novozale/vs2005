Public Class CreditInfo
    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        '////////////////////////////////////////////////////////////////////////////////////////
        '// Выход из формы
        '//
        '////////////////////////////////////////////////////////////////////////////////////////

        Me.Close()
    End Sub

    Private Sub CreditInfo_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        '////////////////////////////////////////////////////////////////////////////////////////
        '// Загрузка информации в форму
        '//
        '////////////////////////////////////////////////////////////////////////////////////////
        Dim MySQLStr As String
        Dim MyAdapter As SqlClient.SqlDataAdapter      '
        Dim MyDs As New DataSet                        '
        Dim MyAdapter1 As SqlClient.SqlDataAdapter     '
        Dim MyDs1 As New DataSet                       '

        MySQLStr = "SELECT TOP 100 PERCENT t11.CustomerCode, "
        MySQLStr = MySQLStr & "dbo.SL010300.SL01002 AS CustomerName, "
        MySQLStr = MySQLStr & "t11.SalesmanCode, dbo.ST010300.ST01002 AS SalesmanName, "
        MySQLStr = MySQLStr & "t11.OrderNum, t11.InvoiceNum, t11.InvoiceData, "
        MySQLStr = MySQLStr & "DATEDIFF(day, t11.InvoiceData, GETDATE()) AS OverdueDays, "
        MySQLStr = MySQLStr & "ROUND(CONVERT(float,t11.InvoiceSum),2) AS InvoiceSum, ROUND(CONVERT(float,ISNULL(t12.PayedSum, 0)),2) AS PayedSum "
        MySQLStr = MySQLStr & "FROM (SELECT TOP 100 PERCENT SL03001 AS CustomerCode, "
        MySQLStr = MySQLStr & "SL03041 AS SalesmanCode, SL03036 AS OrderNum, SL03002 AS InvoiceNum, "
        MySQLStr = MySQLStr & "SUM(SL03013) AS InvoiceSum, SL03004 AS InvoiceData "
        MySQLStr = MySQLStr & "FROM dbo.SL030300 WITH (NOLOCK) "
        MySQLStr = MySQLStr & "WHERE (SL03005 <= GETDATE()) AND (LEFT(SL03017, 6) = '621010') "
        MySQLStr = MySQLStr & "GROUP BY SL03001, SL03002, SL03004, SL03041, SL03036 "
        MySQLStr = MySQLStr & "HAVING (SL03001 = N'" & Declarations.CustomerID & "')) AS t11 LEFT OUTER JOIN "
        MySQLStr = MySQLStr & "dbo.SL010300 ON t11.CustomerCode = dbo.SL010300.SL01001 LEFT OUTER JOIN "
        MySQLStr = MySQLStr & "dbo.ST010300 ON t11.SalesmanCode = dbo.ST010300.ST01001 LEFT OUTER JOIN "
        MySQLStr = MySQLStr & "(SELECT TOP 100 PERCENT SL21001 AS CustomerCode, SL21002 AS InvoiceNum, "
        MySQLStr = MySQLStr & "SUM(SL21007) AS PayedSum "
        MySQLStr = MySQLStr & "FROM dbo.SL210300 WITH (NOLOCK) "
        MySQLStr = MySQLStr & "WHERE (SL21006 <= GETDATE()) AND (SL21002 IN "
        MySQLStr = MySQLStr & "(SELECT DISTINCT SL03002 "
        MySQLStr = MySQLStr & "FROM dbo.SL030300 AS SL030300_1 WITH (NOLOCK) "
        MySQLStr = MySQLStr & "WHERE (SL03005 <= GETDATE()) AND (LEFT(SL03017, 6) = '621010'))) "
        MySQLStr = MySQLStr & "GROUP BY SL21001, SL21002 "
        MySQLStr = MySQLStr & "HAVING (SL21001 = N'" & Declarations.CustomerID & "')) AS t12 ON t11.InvoiceNum = t12.InvoiceNum AND "
        MySQLStr = MySQLStr & "t11.CustomerCode = t12.CustomerCode "
        'MySQLStr = MySQLStr & "WHERE (DateDiff(Day, t11.InvoiceData, GETDATE()) > " & Declarations.CreditInDays & ") And "
        'MySQLStr = MySQLStr & "(IsNull(t12.PayedSum, 0) < t11.InvoiceSum - 1) "
        MySQLStr = MySQLStr & "WHERE (IsNull(t12.PayedSum, 0) < t11.InvoiceSum - 1) "
        MySQLStr = MySQLStr & "ORDER BY t11.InvoiceData "

        InitMyConn(False)

        Try
            MyAdapter = New SqlClient.SqlDataAdapter(MySQLStr, Declarations.MyNETConnStr)
            MyAdapter.SelectCommand.CommandTimeout = 600
            MyAdapter.Fill(MyDs)
            DataGridView1.DataSource = MyDs.Tables(0)
        Catch ex As Exception
            MsgBox(ex.ToString)
        End Try

        DataGridView1.Columns(0).HeaderText = "Код клиента"
        DataGridView1.Columns(1).HeaderText = "Имя клиента"
        DataGridView1.Columns(2).HeaderText = "Код продавца"
        DataGridView1.Columns(3).HeaderText = "Имя продавца"
        DataGridView1.Columns(4).HeaderText = "Номер заказа"
        DataGridView1.Columns(5).HeaderText = "Номер СФ"
        DataGridView1.Columns(6).HeaderText = "Дата СФ"
        DataGridView1.Columns(7).HeaderText = "Срок в днях"
        DataGridView1.Columns(8).HeaderText = "Сумма СФ"
        DataGridView1.Columns(9).HeaderText = "Оплачено"


        MySQLStr = "SELECT OR010300.OR01003 AS CustomerCode, "
        MySQLStr = MySQLStr & "SL010300.SL01002 AS CustomerName, "
        MySQLStr = MySQLStr & "OR010300.OR01019 AS SalesmanCode, "
        MySQLStr = MySQLStr & "ST010300.ST01002 AS SalesmanName, "
        MySQLStr = MySQLStr & "OR010300.OR01001 AS OrderNUmber, "
        MySQLStr = MySQLStr & "OR010300.OR01015 AS OrderDate, "
        MySQLStr = MySQLStr & "OR010300.OR01002 AS OrderType, "
        MySQLStr = MySQLStr & "View_1.OrderSum "
        MySQLStr = MySQLStr & "FROM OR010300 WITH (NOLOCK) INNER JOIN "
        MySQLStr = MySQLStr & "(SELECT OR030300.OR03001, "
        MySQLStr = MySQLStr & "ROUND(SUM(ROUND(ROUND((OR030300.OR03008 * "
        MySQLStr = MySQLStr & "CASE WHEN OR010300_1.OR01067 = 0 THEN 1 ELSE OR010300_1.OR01067 "
        MySQLStr = MySQLStr & "END) * (100 - CONVERT(float,OR030300.OR03018) - CONVERT(float,OR030300.OR03017)) / 100, 2) * "
        MySQLStr = MySQLStr & "OR030300.OR03011 / OR030300.OR03022, 2)) "
        MySQLStr = MySQLStr & "+ SUM(ROUND((OR030300.OR03008 * "
        MySQLStr = MySQLStr & "CASE WHEN OR010300_1.OR01067 = 0 THEN 1 ELSE OR010300_1.OR01067 "
        MySQLStr = MySQLStr & "END) * (100 - CONVERT(float,OR030300.OR03018) - CONVERT(float,OR030300.OR03017)) / 100, 2) * "
        MySQLStr = MySQLStr & "OR030300.OR03011 * CONVERT(float, SY290300.SY29003) / 100 / OR030300.OR03022), 2) "
        MySQLStr = MySQLStr & "AS OrderSum "
        MySQLStr = MySQLStr & "FROM OR030300 WITH (NOLOCK) INNER JOIN "
        MySQLStr = MySQLStr & "OR010300 AS OR010300_1 ON OR030300.OR03001 = "
        MySQLStr = MySQLStr & "OR010300_1.OR01001 LEFT OUTER JOIN "
        MySQLStr = MySQLStr & "SY290300 ON OR030300.OR03061 = SY290300.SY29001 "
        MySQLStr = MySQLStr & "WHERE (OR030300.OR03003 = N'000000') "
        MySQLStr = MySQLStr & "GROUP BY OR030300.OR03001) AS View_1 ON "
        MySQLStr = MySQLStr & "OR010300.OR01001 = View_1.OR03001 LEFT OUTER JOIN "
        MySQLStr = MySQLStr & "ST010300 ON OR010300.OR01019 = ST010300.ST01001 LEFT OUTER JOIN "
        MySQLStr = MySQLStr & "SL010300 ON OR010300.OR01003 = SL010300.SL01001 "
        MySQLStr = MySQLStr & "WHERE (OR010300.OR01002 = 1) "
        MySQLStr = MySQLStr & "AND (OR010300.OR01008 = 3) OR "
        MySQLStr = MySQLStr & "(OR010300.OR01002 = 4) AND (OR010300.OR01008 = 3) "
        MySQLStr = MySQLStr & "GROUP BY OR010300.OR01003, OR010300.OR01028, SL010300.SL01002, "
        MySQLStr = MySQLStr & "OR010300.OR01019, ST010300.ST01002, OR010300.OR01001, OR010300.OR01002, "
        MySQLStr = MySQLStr & "OR010300.OR01015 , View_1.OrderSum "
        MySQLStr = MySQLStr & "HAVING (OR010300.OR01003 = N'" & Declarations.CustomerID & "') "

        InitMyConn(False)

        Try
            MyAdapter1 = New SqlClient.SqlDataAdapter(MySQLStr, Declarations.MyNETConnStr)
            MyAdapter1.SelectCommand.CommandTimeout = 600
            MyAdapter1.Fill(MyDs1)
            DataGridView2.DataSource = MyDs1.Tables(0)
        Catch ex As Exception
            MsgBox(ex.ToString)
        End Try

        DataGridView2.Columns(0).HeaderText = "Код клиента"
        DataGridView2.Columns(1).HeaderText = "Имя клиента"
        DataGridView2.Columns(2).HeaderText = "Код продавца"
        DataGridView2.Columns(3).HeaderText = "Имя продавца"
        DataGridView2.Columns(4).HeaderText = "Номер заказа"
        DataGridView2.Columns(5).HeaderText = "Дата заказа"
        DataGridView2.Columns(6).HeaderText = "Тип заказа"
        DataGridView2.Columns(7).HeaderText = "Сумма заказа"
    End Sub

    Private Sub DataGridView1_CellFormatting(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellFormattingEventArgs) Handles DataGridView1.CellFormatting
        '////////////////////////////////////////////////////////////////////////////////
        '//
        '// подсветка информации по просроченным СФ
        '//
        '////////////////////////////////////////////////////////////////////////////////
        Dim row As DataGridViewRow

        row = DataGridView1.Rows(e.RowIndex)
        If row.Cells(7).Value > Declarations.CreditInDays Then
            row.DefaultCellStyle.BackColor = Color.LightPink
        Else
            row.DefaultCellStyle.BackColor = Color.Empty
        End If
    End Sub
End Class