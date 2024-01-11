Public Class MatchPictAndScalaCode

    Private Sub Button3_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button3.Click
        '/////////////////////////////////////////////////////////////////////////////////////
        '//
        '// выход из окна
        '// 
        '/////////////////////////////////////////////////////////////////////////////////////

        Me.Close()
    End Sub

    Private Sub MatchPictAndScalaCode_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        '/////////////////////////////////////////////////////////////////////////////////////
        '//
        '// Загрузка формы
        '// 
        '/////////////////////////////////////////////////////////////////////////////////////
        Dim MySQLStr As String                        'рабочая строка
        Dim MyAdapter As SqlClient.SqlDataAdapter     '
        Dim MyDs As New DataSet                       '

        '---тут выводим картинки и товары, которые могут им соответствовать
        'MySQLStr = "SELECT tbl_WEB_Pictures.ID, tbl_WEB_Pictures.SupplierItemCode, tbl_WEB_Pictures.PictureMedium, CONVERT(bit, CASE WHEN ISNULL(View_1.TotalQTY, 0) "
        'MySQLStr = MySQLStr & "- ISNULL(View_3.MatchedQTY, 0) = 1 THEN 1 ELSE 0 END) AS ToMatch, CONVERT(bit, 0) AS ToDelete, ISNULL(SC010300.SC01001, N'') "
        'MySQLStr = MySQLStr & "AS ScalaItemCode, LTRIM(RTRIM(LTRIM(RTRIM(ISNULL(SC010300.SC01002, N''))) + ' ' + LTRIM(RTRIM(ISNULL(SC010300.SC01003, N''))))) "
        'MySQLStr = MySQLStr & "AS ItemName, ISNULL(SC010300.SC01058, N'') AS SupplierCode, ISNULL(PL010300.PL01002, N'') AS SupplierName, CONVERT(bit, CASE WHEN ISNULL(View_1.TotalQTY, 0) "
        'MySQLStr = MySQLStr & "- ISNULL(View_3.MatchedQTY, 0) = 1 THEN 1 ELSE 0 END) AS ToMatchSrc "
        'MySQLStr = MySQLStr & "FROM tbl_WEB_Pictures INNER JOIN "
        'MySQLStr = MySQLStr & "SC010300 ON tbl_WEB_Pictures.SupplierItemCode = SC010300.SC01060 LEFT OUTER JOIN "
        'MySQLStr = MySQLStr & "(SELECT SupplierItemCode, COUNT(SupplierItemCode) AS MatchedQTY "
        'MySQLStr = MySQLStr & "FROM tbl_WEB_Pictures AS tbl_WEB_Pictures_1 "
        'MySQLStr = MySQLStr & "WHERE(Not (ScalaItemCode Is NULL)) "
        'MySQLStr = MySQLStr & "GROUP BY SupplierItemCode) AS View_3 ON tbl_WEB_Pictures.SupplierItemCode = View_3.SupplierItemCode LEFT OUTER JOIN "
        'MySQLStr = MySQLStr & "PL010300 ON SC010300.SC01058 = PL010300.PL01001 LEFT OUTER JOIN "
        'MySQLStr = MySQLStr & "(SELECT SC01060, COUNT(SC01060) AS TotalQTY "
        'MySQLStr = MySQLStr & "FROM SC010300 AS SC010300_1 "
        'MySQLStr = MySQLStr & "GROUP BY SC01060) AS View_1 ON tbl_WEB_Pictures.SupplierItemCode = View_1.SC01060 "
        'MySQLStr = MySQLStr & "WHERE (tbl_WEB_Pictures.ScalaItemCode IS NULL) AND (SC010300.SC01001 NOT IN "
        'MySQLStr = MySQLStr & "(SELECT ScalaItemCode "
        'MySQLStr = MySQLStr & "FROM tbl_WEB_Pictures AS tbl_WEB_Pictures_2 "
        'MySQLStr = MySQLStr & "WHERE (ScalaItemCode IS NOT NULL))) "
        'MySQLStr = MySQLStr & "ORDER BY tbl_WEB_Pictures.SupplierItemCode, tbl_WEB_Pictures.ID "

        MySQLStr = "SELECT tbl_WEB_Pictures.ID, tbl_WEB_Pictures.SupplierItemCode, tbl_WEB_Pictures.PictureMedium, CONVERT(bit, CASE WHEN ISNULL(View_1.TotalQTY, 0) "
        MySQLStr = MySQLStr & "- ISNULL(View_3.MatchedQTY, 0) = 1 THEN 1 ELSE 0 END) AS ToMatch, CONVERT(bit, 0) AS ToDelete, ISNULL(SC010300.SC01001, N'') "
        MySQLStr = MySQLStr & "AS ScalaItemCode, LTRIM(RTRIM(LTRIM(RTRIM(ISNULL(SC010300.SC01002, N''))) + ' ' + LTRIM(RTRIM(ISNULL(SC010300.SC01003, N''))))) "
        MySQLStr = MySQLStr & "AS ItemName, ISNULL(SC010300.SC01058, N'') AS SupplierCode, ISNULL(PL010300.PL01002, N'') AS SupplierName, CONVERT(bit, "
        MySQLStr = MySQLStr & "CASE WHEN ISNULL(View_1.TotalQTY, 0) - ISNULL(View_3.MatchedQTY, 0) = 1 THEN 1 ELSE 0 END) AS ToMatchSrc, ISNULL(tbl_ItemCard0300.Manufacturer, '') AS Manufacturer, "
        MySQLStr = MySQLStr & "ISNULL(tbl_Manufacturers.Name, '') AS ManufacturerName "
        MySQLStr = MySQLStr & "FROM tbl_Manufacturers RIGHT OUTER JOIN "
        MySQLStr = MySQLStr & "tbl_ItemCard0300 ON tbl_Manufacturers.ID = tbl_ItemCard0300.Manufacturer RIGHT OUTER JOIN "
        MySQLStr = MySQLStr & "tbl_WEB_Pictures INNER JOIN "
        MySQLStr = MySQLStr & "SC010300 ON tbl_WEB_Pictures.SupplierItemCode = SC010300.SC01060 ON tbl_ItemCard0300.SC01001 = SC010300.SC01001 LEFT OUTER JOIN "
        MySQLStr = MySQLStr & "(SELECT SupplierItemCode, COUNT(SupplierItemCode) AS MatchedQTY "
        MySQLStr = MySQLStr & "FROM tbl_WEB_Pictures AS tbl_WEB_Pictures_1 "
        MySQLStr = MySQLStr & "WHERE (Not (ScalaItemCode Is NULL)) "
        MySQLStr = MySQLStr & "GROUP BY SupplierItemCode) AS View_3 ON tbl_WEB_Pictures.SupplierItemCode = View_3.SupplierItemCode LEFT OUTER JOIN "
        MySQLStr = MySQLStr & "PL010300 ON SC010300.SC01058 = PL010300.PL01001 LEFT OUTER JOIN "
        MySQLStr = MySQLStr & "(SELECT SC01060, COUNT(SC01060) AS TotalQTY "
        MySQLStr = MySQLStr & "FROM SC010300 AS SC010300_1 "
        MySQLStr = MySQLStr & "GROUP BY SC01060) AS View_1 ON tbl_WEB_Pictures.SupplierItemCode = View_1.SC01060 "
        MySQLStr = MySQLStr & "WHERE (tbl_WEB_Pictures.ScalaItemCode IS NULL) AND (SC010300.SC01001 NOT IN "
        MySQLStr = MySQLStr & "(SELECT ScalaItemCode "
        MySQLStr = MySQLStr & "FROM tbl_WEB_Pictures AS tbl_WEB_Pictures_2 "
        MySQLStr = MySQLStr & "WHERE (ScalaItemCode IS NOT NULL))) "
        MySQLStr = MySQLStr & "ORDER BY tbl_WEB_Pictures.SupplierItemCode, tbl_WEB_Pictures.ID "

        DataGridView12.RowTemplate.MinimumHeight = 100

        InitMyConn(False)
        Try
            MyAdapter = New SqlClient.SqlDataAdapter(MySQLStr, Declarations.MyNETConnStr)
            MyAdapter.SelectCommand.CommandTimeout = 600
            MyAdapter.Fill(MyDs)
            DataGridView12.DataSource = MyDs.Tables(0)
        Catch ex As Exception
            MsgBox(ex.ToString)
        End Try

        DataGridView12.Columns(0).HeaderText = "ID"
        DataGridView12.Columns(0).Width = 100
        DataGridView12.Columns(0).Visible = False
        DataGridView12.Columns(0).ReadOnly = True
        DataGridView12.Columns(1).HeaderText = "Код товара поставщика"
        DataGridView12.Columns(1).Width = 150
        DataGridView12.Columns(1).ReadOnly = True
        DataGridView12.Columns(2).HeaderText = "Картинка"
        DataGridView12.Columns(2).Width = 100
        DataGridView12.Columns(2).ReadOnly = True
        DataGridView12.Columns(3).HeaderText = "Связать"
        DataGridView12.Columns(3).Width = 50
        DataGridView12.Columns(4).HeaderText = "Удалить"
        DataGridView12.Columns(4).Width = 50
        DataGridView12.Columns(5).HeaderText = "Код товара в Scala"
        DataGridView12.Columns(5).Width = 100
        DataGridView12.Columns(5).ReadOnly = True
        DataGridView12.Columns(6).HeaderText = "Название товара в Scala"
        DataGridView12.Columns(6).Width = 300
        DataGridView12.Columns(6).ReadOnly = True
        DataGridView12.Columns(7).HeaderText = "Код поставщика"
        DataGridView12.Columns(7).Width = 100
        DataGridView12.Columns(7).ReadOnly = True
        DataGridView12.Columns(8).HeaderText = "Название поставщика"
        DataGridView12.Columns(8).Width = 300
        DataGridView12.Columns(8).ReadOnly = True
        DataGridView12.Columns(9).HeaderText = "Состояние"
        DataGridView12.Columns(9).Width = 50
        DataGridView12.Columns(9).ReadOnly = True
        DataGridView12.Columns(9).Visible = False
        DataGridView12.Columns(10).HeaderText = "Код производителя"
        DataGridView12.Columns(10).Width = 100
        DataGridView12.Columns(10).ReadOnly = True
        DataGridView12.Columns(11).HeaderText = "Название производителя"
        DataGridView12.Columns(11).Width = 300
        DataGridView12.Columns(11).ReadOnly = True



    End Sub

    Private Sub Button2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button2.Click
        '/////////////////////////////////////////////////////////////////////////////////////
        '//
        '// Связывание картинок с кодом Скала или удаление и выход
        '// 
        '/////////////////////////////////////////////////////////////////////////////////////

        Me.Cursor = Cursors.WaitCursor
        MakePictMatchingOrDel()
        Me.Cursor = Cursors.Default
        Me.Close()
    End Sub

    Private Sub MakePictMatchingOrDel()
        '/////////////////////////////////////////////////////////////////////////////////////
        '//
        '// Связывание картинок с кодом Скала или удаление 
        '// 
        '/////////////////////////////////////////////////////////////////////////////////////
        Dim MySQLStr As String

        Try
            For i As Integer = 0 To DataGridView12.Rows.Count - 1
                If DataGridView12.Item(4, i).Value = True Then  '---удаление
                    DeletePicture(Trim(DataGridView12.Item(0, i).Value.ToString))
                Else
                    If DataGridView12.Item(3, i).Value = True Then '---связывание
                        MatchPicture(Trim(DataGridView12.Item(0, i).Value.ToString), Trim(DataGridView12.Item(5, i).Value.ToString))
                        '---помечаем товар измененным
                        MySQLStr = "UPDATE tbl_WEB_Items "
                        MySQLStr = MySQLStr & "SET RMStatus = CASE WHEN tbl_WEB_Items.RMStatus = 1 THEN 1 ELSE CASE WHEN tbl_WEB_Items.RMStatus = 2 THEN 2 ELSE 3 END END, "
                        MySQLStr = MySQLStr & "WEBStatus = CASE WHEN tbl_WEB_Items.WEBStatus = 1 THEN 1 ELSE CASE WHEN tbl_WEB_Items.WEBStatus = 2 THEN 2 ELSE 3 END END "
                        MySQLStr = MySQLStr & "WHERE (Code = N'" & Trim(DataGridView12.Item(5, i).Value.ToString) & "') "
                        InitMyConn(False)
                        Declarations.MyConn.Execute(MySQLStr)
                    End If
                End If
            Next i
        Catch ex As Exception
        End Try
    End Sub

    Private Sub DeletePicture(ByVal PictureID As String)
        '/////////////////////////////////////////////////////////////////////////////////////
        '//
        '// Удаление картинки из БД 
        '// 
        '/////////////////////////////////////////////////////////////////////////////////////
        Dim MySQLStr As String

        MySQLStr = "DELETE FROM tbl_WEB_Pictures "
        MySQLStr = MySQLStr & "WHERE (ID = N'" & PictureID & "') "

        InitMyConn(False)
        Declarations.MyConn.Execute(MySQLStr)
    End Sub

    Private Sub MatchPicture(ByVal PictureID As String, ByVal ScalaCode As String)
        '/////////////////////////////////////////////////////////////////////////////////////
        '//
        '// Удаление картинки из БД 
        '// 
        '/////////////////////////////////////////////////////////////////////////////////////
        Dim MySQLStr As String

        MySQLStr = "UPDATE tbl_WEB_Pictures "
        MySQLStr = MySQLStr & "SET ScalaItemCode = N'" & ScalaCode & "' "
        MySQLStr = MySQLStr & "WHERE (ID = N'" & PictureID & "') "

        InitMyConn(False)
        Declarations.MyConn.Execute(MySQLStr)
    End Sub

    Private Sub DataGridView12_CellFormatting(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellFormattingEventArgs) Handles DataGridView12.CellFormatting
        '////////////////////////////////////////////////////////////////////////////////
        '//
        '// Подсветка строк продуктов в зависимости от статуса
        '//
        '////////////////////////////////////////////////////////////////////////////////
        Dim row As DataGridViewRow

        row = DataGridView12.Rows(e.RowIndex)
        If row.Cells(9).Value = False Then
            row.DefaultCellStyle.BackColor = Color.LightGray
        End If
    End Sub

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        '////////////////////////////////////////////////////////////////////////////////
        '//
        '// Снятие всех галочек связывания картинок 
        '//
        '////////////////////////////////////////////////////////////////////////////////

        Try
            For i As Integer = 0 To DataGridView12.Rows.Count - 1
                DataGridView12.Item(3, i).Value = False
            Next i
        Catch ex As Exception
        End Try
    End Sub

    Private Sub Button4_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button4.Click
        '////////////////////////////////////////////////////////////////////////////////
        '//
        '// Выставление всех галочек связывания картинок 
        '//
        '////////////////////////////////////////////////////////////////////////////////

        Try
            For i As Integer = 0 To DataGridView12.Rows.Count - 1
                DataGridView12.Item(3, i).Value = True
            Next i
        Catch ex As Exception
        End Try
    End Sub
End Class