Public Class SearcherList

    Private Sub Button5_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button5.Click
        '/////////////////////////////////////////////////////////////////////////////////////
        '//
        '// Выход без выбора поисковика
        '// 
        '/////////////////////////////////////////////////////////////////////////////////////

        Me.Close()
    End Sub

    Private Sub SearcherList_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles Me.KeyDown
        '/////////////////////////////////////////////////////////////////////////////////////
        '//
        '// выход без сохранения по Escape
        '// 
        '/////////////////////////////////////////////////////////////////////////////////////

        If e.KeyData = Keys.Escape Then
            Me.Close()
        End If
    End Sub

    Private Sub Button4_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button4.Click
        '/////////////////////////////////////////////////////////////////////////////////////
        '//
        '// Выход с выбором поисковика
        '// 
        '/////////////////////////////////////////////////////////////////////////////////////

        SearcherSelect()
    End Sub

    Private Sub DataGridView1_DoubleClick(ByVal sender As Object, ByVal e As System.EventArgs) Handles DataGridView1.DoubleClick
        '/////////////////////////////////////////////////////////////////////////////////////
        '//
        '// Выход с выбором поисковика
        '// 
        '/////////////////////////////////////////////////////////////////////////////////////

        SearcherSelect()
    End Sub

    Private Sub SearcherSelect()
        '/////////////////////////////////////////////////////////////////////////////////////
        '//
        '// Выход с выбором поисковика
        '// 
        '/////////////////////////////////////////////////////////////////////////////////////
        Dim MySQLStr As String

        MySQLStr = "UPDATE tbl_SupplSearch "
        MySQLStr = MySQLStr & "SET SearcherID = N'" & DataGridView1.SelectedRows.Item(0).Cells(0).Value.ToString() & "', "
        MySQLStr = MySQLStr & "SearcherName = N'" & DataGridView1.SelectedRows.Item(0).Cells(1).Value.ToString() & "', "
        MySQLStr = MySQLStr & "SearchStatus = -1 "
        MySQLStr = MySQLStr & "WHERE (ID = " & MainForm.DataGridView1.SelectedRows.Item(0).Cells(0).Value.ToString() & ") "
        InitMyConn(False)
        Declarations.MyConn.Execute(MySQLStr)

        Me.Close()
    End Sub

    Private Sub SearcherList_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        '/////////////////////////////////////////////////////////////////////////////////////
        '//
        '// При запуске загружаем список поисковиков
        '// 
        '/////////////////////////////////////////////////////////////////////////////////////
        Dim MySQLStr As String                        'рабочая строка
        Dim MyAdapter As SqlClient.SqlDataAdapter     '
        Dim MyDs As New DataSet

        MySQLStr = "SELECT tbl_SupplSearch_Searchers.PurchID, View_17.SYPD003, ISNULL(View_9.CC, 0) AS QTY, ISNULL(View_6.CC, 0) AS Started, ISNULL(View_7.CC, 0) AS Proposed, ISNULL(View_8.CC, 0) "
        MySQLStr = MySQLStr & "AS Confirmed "
        MySQLStr = MySQLStr & "FROM tbl_SupplSearch_Searchers INNER JOIN "
        MySQLStr = MySQLStr & "(SELECT SYPD001, SYPD003 "
        MySQLStr = MySQLStr & "FROM SYPD0300 "
        MySQLStr = MySQLStr & "WHERE (SYPD002 = N'ENG')) AS View_17 ON tbl_SupplSearch_Searchers.PurchID = View_17.SYPD001 LEFT OUTER JOIN "
        MySQLStr = MySQLStr & "(SELECT tbl_SupplSearch.SearcherID, COUNT(tbl_SupplSearchItems.ItemID) AS CC "
        MySQLStr = MySQLStr & "FROM tbl_SupplSearch INNER JOIN "
        MySQLStr = MySQLStr & "tbl_SupplSearchItems ON tbl_SupplSearch.ID = tbl_SupplSearchItems.SupplSearchID "
        MySQLStr = MySQLStr & "WHERE (tbl_SupplSearch.StartDate > DATEADD(mm, - 1, GETDATE())) AND (tbl_SupplSearch.SearcherID IS NOT NULL) AND (tbl_SupplSearch.SearchStatus = 2) AND "
        MySQLStr = MySQLStr & "(tbl_SupplSearch.SalesStatus = 0) "
        MySQLStr = MySQLStr & "GROUP BY tbl_SupplSearch.SearcherID) AS View_8 ON tbl_SupplSearch_Searchers.PurchID = View_8.SearcherID LEFT OUTER JOIN "
        MySQLStr = MySQLStr & "(SELECT tbl_SupplSearch_3.SearcherID, COUNT(tbl_SupplSearchItems_3.ItemID) AS CC "
        MySQLStr = MySQLStr & "FROM dbo.tbl_SupplSearch AS tbl_SupplSearch_3 INNER JOIN "
        MySQLStr = MySQLStr & "dbo.tbl_SupplSearchItems AS tbl_SupplSearchItems_3 ON tbl_SupplSearch_3.ID = tbl_SupplSearchItems_3.SupplSearchID "
        MySQLStr = MySQLStr & "WHERE (tbl_SupplSearch_3.StartDate > DATEADD(mm, - 1, GETDATE())) AND (tbl_SupplSearch_3.SearcherID IS NOT NULL) AND (tbl_SupplSearch_3.SearchStatus = 2) "
        MySQLStr = MySQLStr & "AND (tbl_SupplSearch_3.SalesStatus = 0 OR tbl_SupplSearch_3.SalesStatus = 1 OR tbl_SupplSearch_3.SalesStatus = 2) "
        MySQLStr = MySQLStr & "GROUP BY tbl_SupplSearch_3.SearcherID) AS View_7 ON tbl_SupplSearch_Searchers.PurchID = View_7.SearcherID LEFT OUTER JOIN "
        MySQLStr = MySQLStr & "(SELECT tbl_SupplSearch_2.SearcherID, COUNT(tbl_SupplSearchItems_2.ItemID) AS CC "
        MySQLStr = MySQLStr & "FROM dbo.tbl_SupplSearch AS tbl_SupplSearch_2 INNER JOIN "
        MySQLStr = MySQLStr & "dbo.tbl_SupplSearchItems AS tbl_SupplSearchItems_2 ON tbl_SupplSearch_2.ID = tbl_SupplSearchItems_2.SupplSearchID "
        MySQLStr = MySQLStr & "WHERE (tbl_SupplSearch_2.StartDate > DATEADD(mm, - 1, GETDATE()))  AND (tbl_SupplSearch_2.SearcherID IS NOT NULL) AND (tbl_SupplSearch_2.SearchStatus = 1 OR "
        MySQLStr = MySQLStr & "tbl_SupplSearch_2.SearchStatus = 2) AND (tbl_SupplSearch_2.SalesStatus = 0 or tbl_SupplSearch_2.SalesStatus = 1) "
        MySQLStr = MySQLStr & "GROUP BY tbl_SupplSearch_2.SearcherID) AS View_6 ON tbl_SupplSearch_Searchers.PurchID = View_6.SearcherID LEFT OUTER JOIN "
        MySQLStr = MySQLStr & "(SELECT tbl_SupplSearch_1.SearcherID, COUNT(tbl_SupplSearchItems_1.ItemID) AS CC "
        MySQLStr = MySQLStr & "FROM tbl_SupplSearch AS tbl_SupplSearch_1 INNER JOIN "
        MySQLStr = MySQLStr & "tbl_SupplSearchItems AS tbl_SupplSearchItems_1 ON tbl_SupplSearch_1.ID = tbl_SupplSearchItems_1.SupplSearchID "
        MySQLStr = MySQLStr & "WHERE (tbl_SupplSearch_1.StartDate > DateAdd(mm, -1, GETDATE())) "
        MySQLStr = MySQLStr & "GROUP BY tbl_SupplSearch_1.SearcherID "
        MySQLStr = MySQLStr & "HAVING (tbl_SupplSearch_1.SearcherID IS NOT NULL)) AS View_9 ON tbl_SupplSearch_Searchers.PurchID = View_9.SearcherID "

        Try
            MyAdapter = New SqlClient.SqlDataAdapter(MySQLStr, Declarations.MyNETConnStr)
            MyAdapter.SelectCommand.CommandTimeout = 600
            MyAdapter.Fill(MyDs)
            DataGridView1.DataSource = MyDs.Tables(0)
        Catch ex As Exception
            MsgBox(ex.ToString)
        End Try

        DataGridView1.Columns(0).HeaderText = "Код поисковика"
        DataGridView1.Columns(0).Width = 110
        DataGridView1.Columns(1).HeaderText = "Имя поисковика"
        DataGridView1.Columns(1).Width = 300
        DataGridView1.Columns(2).HeaderText = "Товаров в поиске за последний месяц"
        DataGridView1.Columns(2).Width = 150
        DataGridView1.Columns(3).HeaderText = "Товаров в поиске сейчас у поисковика"
        DataGridView1.Columns(3).Width = 150
        DataGridView1.Columns(4).HeaderText = "Товаров предложено"
        DataGridView1.Columns(4).Width = 150
        DataGridView1.Columns(5).HeaderText = "Товаров подтверждено"
        DataGridView1.Columns(5).Width = 150
    End Sub
End Class