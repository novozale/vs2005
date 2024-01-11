Public Class ALTItems
    Public MyItem As String                           'код запаса
    Public MySrcWin As String                         'окно, из которого вызвано


    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        '////////////////////////////////////////////////////////////////////////////////
        '//
        '// Выход из окна 
        '//
        '////////////////////////////////////////////////////////////////////////////////

        Me.Close()
    End Sub

    Private Sub ALTItems_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        '////////////////////////////////////////////////////////////////////////////////////////
        '//
        '// Загрузка информации в окно
        '//
        '////////////////////////////////////////////////////////////////////////////////////////
        Dim MySQLStr As String
        Dim MyAdapter As SqlClient.SqlDataAdapter
        Dim MyDs As New DataSet

        MySQLStr = "SELECT SC01001 + ' ' + SC01002 as Name "
        MySQLStr = MySQLStr & "FROM SC010300 WITH (NOLOCK) "
        MySQLStr = MySQLStr & "where SC01001 = N'" & Trim(MyItem) & "'"
        InitMyConn(False)
        InitMyRec(False, MySQLStr)
        Label2.Text = Declarations.MyRec.Fields("Name").Value
        trycloseMyRec()

        'MySQLStr = "SELECT tbl_AlternativeProducts.ALTCode AS ProductCode, "
        'MySQLStr = MySQLStr & "ISNULL(SC010300.SC01002, N'') + ' ' + ISNULL(SC010300.SC01003, N'') AS ProductName, "
        'MySQLStr = MySQLStr & "ISNULL(SC010300.SC01060, N'') AS SuppProductCode, "
        'MySQLStr = MySQLStr & "ISNULL(SC010300.SC01042, 0) AS TotalQTY, "
        'MySQLStr = MySQLStr & "ISNULL(View_1.WhQty, 0) AS WhQty, "
        'MySQLStr = MySQLStr & "ISNULL(View_1.WhAvl, 0) AS WhAvl, "
        'MySQLStr = MySQLStr & "ISNULL(SC010300.SC01058, N'') AS SuppCode, "
        'MySQLStr = MySQLStr & "ISNULL(PL010300.PL01002, N'') AS SuppName "
        'MySQLStr = MySQLStr & "FROM SC010300 WITH (NOLOCK) INNER JOIN "
        'MySQLStr = MySQLStr & "(SELECT SC03001, SC03003 AS WhQty, SC03003 - SC03004 - SC03005 AS WhAvl "
        'MySQLStr = MySQLStr & "FROM SC030300 WITH (NOLOCK) "
        'MySQLStr = MySQLStr & "WHERE (SC03002 = N'" & Declarations.WHNum & "')) "
        'MySQLStr = MySQLStr & "AS View_1 ON SC010300.SC01001 = View_1.SC03001 RIGHT OUTER JOIN "
        'MySQLStr = MySQLStr & "tbl_AlternativeProducts ON "
        'MySQLStr = MySQLStr & "SC010300.SC01001 = tbl_AlternativeProducts.ALTCode LEFT OUTER JOIN "
        'MySQLStr = MySQLStr & "PL010300 ON SC010300.SC01058 = PL010300.PL01001 "
        'MySQLStr = MySQLStr & "WHERE (tbl_AlternativeProducts.ProductCode = N'" & Trim(MyItem) & "') "
        'MySQLStr = MySQLStr & "ORDER BY ProductCode"

        MySQLStr = "SELECT tbl_AlternativeProducts.ALTCode AS ProdCode, "
        MySQLStr = MySQLStr & "ISNULL(SC010300.SC01002, N'') + ' ' + ISNULL(SC010300.SC01003, N'') AS ProductName, "
        MySQLStr = MySQLStr & "ISNULL(SC010300.SC01060, N'') AS SuppProductCode, "
        MySQLStr = MySQLStr & "ISNULL(SC010300.SC01042, 0) AS TotalQTY, "
        MySQLStr = MySQLStr & "ISNULL(View_1.WhQty, 0) AS WhQty, "
        MySQLStr = MySQLStr & "ISNULL(View_1.WhAvl, 0) AS WhAvl, "
        MySQLStr = MySQLStr & "ISNULL(SC010300.SC01058, N'') AS SuppCode, "
        MySQLStr = MySQLStr & "ISNULL(PL010300.PL01002, N'') AS SuppName, "
        MySQLStr = MySQLStr & "CASE WHEN Ltrim(Rtrim(SC010300.SC01058)) = 'FIN001_U' OR "
        MySQLStr = MySQLStr & "Ltrim(Rtrim(ISNULL(View_4.ActionOrSales, ''))) = 'неликвид' THEN '+' ELSE '' END AS Dead, "
        MySQLStr = MySQLStr & "CASE WHEN Ltrim(Rtrim(ISNULL(View_4.ActionOrSales, ''))) = 'акция' THEN '+' ELSE '' END AS Action "
        MySQLStr = MySQLStr & "FROM SC010300 WITH (NOLOCK) INNER JOIN "
        MySQLStr = MySQLStr & "(SELECT SC03001, SC03003 AS WhQty, SC03003 - SC03004 - SC03005 AS WhAvl "
        MySQLStr = MySQLStr & "FROM SC030300 WITH (NOLOCK) "
        MySQLStr = MySQLStr & "WHERE (SC03002 = N'" & Declarations.WHNum & "')) AS View_1 ON SC010300.SC01001 = View_1.SC03001 LEFT OUTER JOIN "
        MySQLStr = MySQLStr & "(SELECT ScalaCode, ActionOrSales "
        MySQLStr = MySQLStr & "FROM tbl_ActionsAndSales "
        MySQLStr = MySQLStr & "WHERE (ActionFinished = 0) AND (DateStart <= DATEADD(day, DATEDIFF(day, 0, GETDATE()), 0)) "
        MySQLStr = MySQLStr & "AND (DateFinish >= DATEADD(day, DATEDIFF(day, 0, GETDATE()), 0))) AS View_4 ON "
        MySQLStr = MySQLStr & "SC010300.SC01001 = View_4.ScalaCode RIGHT OUTER JOIN "
        MySQLStr = MySQLStr & "tbl_AlternativeProducts ON SC010300.SC01001 = tbl_AlternativeProducts.ALTCode LEFT OUTER JOIN "
        MySQLStr = MySQLStr & "PL010300 ON SC010300.SC01058 = PL010300.PL01001 "
        MySQLStr = MySQLStr & "WHERE (tbl_AlternativeProducts.ProductCode = N'" & Trim(MyItem) & "') "
        MySQLStr = MySQLStr & "ORDER BY ProdCode "

        Try
            MyAdapter = New SqlClient.SqlDataAdapter(MySQLStr, Declarations.MyNETConnStr)
            MyAdapter.SelectCommand.CommandTimeout = 600
            MyAdapter.Fill(MyDs)
            DataGridView1.DataSource = MyDs.Tables(0)
        Catch ex As Exception
            MsgBox(ex.ToString)
        End Try
        DataGridView1.Columns(0).HeaderText = "ID запаса"
        DataGridView1.Columns(1).HeaderText = "Имя запаса"
        DataGridView1.Columns(1).Width = 200
        DataGridView1.Columns(2).HeaderText = "Код зап. поставщика"
        DataGridView1.Columns(3).HeaderText = "Всего на складах"
        DataGridView1.Columns(4).HeaderText = "Ост. " & Declarations.WHNum
        DataGridView1.Columns(5).HeaderText = "Дост. " & Declarations.WHNum
        DataGridView1.Columns(6).HeaderText = "Поставщик ID"
        DataGridView1.Columns(6).Width = 90
        DataGridView1.Columns(7).HeaderText = "Поставщик"
        DataGridView1.Columns(8).Visible = False
        DataGridView1.Columns(9).Visible = False

        If DataGridView1.Rows.Count = 0 Then
            Button2.Enabled = False
        Else
            Button2.Enabled = True
        End If
    End Sub

    Private Sub Button2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button2.Click
        '////////////////////////////////////////////////////////////////////////////////////////
        '//
        '// Выход из формы с выбором альтернативного продукта
        '//
        '////////////////////////////////////////////////////////////////////////////////////////

        SelectAlternateProduct()
    End Sub

    Private Sub DataGridView1_CellFormatting(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellFormattingEventArgs) Handles DataGridView1.CellFormatting
        '////////////////////////////////////////////////////////////////////////////////////////
        '//
        '// смена выделения неликвидов и акций
        '//
        '////////////////////////////////////////////////////////////////////////////////////////
        Dim row As DataGridViewRow

        row = DataGridView1.Rows(e.RowIndex)
        If row.Cells(9).Value.ToString = "+" Then
            row.DefaultCellStyle.Font = New Font(row.InheritedStyle.Font, FontStyle.Bold)
            row.DefaultCellStyle.ForeColor = Color.DarkGreen
        End If
        If row.Cells(8).Value.ToString = "+" Then
            row.DefaultCellStyle.Font = New Font(row.InheritedStyle.Font, FontStyle.Bold)
            row.DefaultCellStyle.ForeColor = Color.Red
        End If
    End Sub

    Private Sub DataGridView1_DoubleClick(ByVal sender As Object, ByVal e As System.EventArgs) Handles DataGridView1.DoubleClick
        '////////////////////////////////////////////////////////////////////////////////////////
        '//
        '// Выход из формы с выбором альтернативного продукта
        '//
        '////////////////////////////////////////////////////////////////////////////////////////

        If DataGridView1.Rows.Count > 0 Then
            SelectAlternateProduct()
        End If
    End Sub

    Private Sub SelectAlternateProduct()
        '////////////////////////////////////////////////////////////////////////////////////////
        '//
        '// Выход из формы с выбором альтернативного продукта
        '//
        '////////////////////////////////////////////////////////////////////////////////////////

        If MySrcWin = "OrderLines" Then
            For i As Integer = 0 To MyOrderLines.DataGridView1.Rows.Count - 1
                If Trim(MyOrderLines.DataGridView1.Item(1, i).Value.ToString) = Trim(Me.DataGridView1.SelectedRows.Item(0).Cells(0).Value.ToString()) Then
                    MyOrderLines.DataGridView1.CurrentCell = MyOrderLines.DataGridView1.Item(1, i)
                    Me.Close()
                    Exit Sub
                End If
            Next
        End If
        If MySrcWin = "AddToOrder" Then
            MyAddToOrder.TextBox2.Text = Trim(Me.DataGridView1.SelectedRows.Item(0).Cells(0).Value.ToString())
            MyAddToOrder.TextBox2Validation()
        End If
        Me.Close()
    End Sub
End Class