Public Class DeletePictureFromDB

    Private Sub Button3_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button3.Click
        '/////////////////////////////////////////////////////////////////////////////////////
        '//
        '// выход из окна
        '// 
        '/////////////////////////////////////////////////////////////////////////////////////

        Me.Close()
    End Sub

    Private Sub DeletePictureFromDB_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        '/////////////////////////////////////////////////////////////////////////////////////
        '//
        '// Загрузка формы
        '// 
        '/////////////////////////////////////////////////////////////////////////////////////
        Dim MySQLStr As String                        'рабочая строка
        Dim MyAdapter As SqlClient.SqlDataAdapter     '
        Dim MyDs As New DataSet                       '

        '---тут выводим группы, а сами подгруппы - по изменению выбора группы
        MySQLStr = "SELECT tbl_WEB_Pictures.ID, tbl_WEB_Pictures.PictureMedium, CONVERT(bit, 0) AS ToDelete, ISNULL(tbl_WEB_Pictures.ScalaItemCode, N'') "
        MySQLStr = MySQLStr & "AS ScalaItemCode, LTRIM(RTRIM(LTRIM(RTRIM(ISNULL(SC010300.SC01002, N''))) + ' ' + LTRIM(RTRIM(ISNULL(SC010300.SC01003, N''))))) "
        MySQLStr = MySQLStr & "AS ScalaName, tbl_WEB_Pictures.SupplierItemCode, ISNULL(SC010300.SC01058, N'') AS SuppCode, ISNULL(PL010300.PL01002, N'') "
        MySQLStr = MySQLStr & "AS SuppName "
        MySQLStr = MySQLStr & "FROM PL010300 RIGHT OUTER JOIN "
        MySQLStr = MySQLStr & "SC010300 ON PL010300.PL01001 = SC010300.SC01058 RIGHT OUTER JOIN "
        MySQLStr = MySQLStr & "tbl_WEB_Pictures ON SC010300.SC01001 = tbl_WEB_Pictures.ScalaItemCode "
        MySQLStr = MySQLStr & "ORDER BY ScalaItemCode "

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
        DataGridView12.Columns(1).HeaderText = "Картинка"
        DataGridView12.Columns(1).Width = 100
        DataGridView12.Columns(1).ReadOnly = True
        DataGridView12.Columns(2).HeaderText = "Удалить"
        DataGridView12.Columns(2).Width = 50
        DataGridView12.Columns(3).HeaderText = "Код товара в Scala"
        DataGridView12.Columns(3).Width = 150
        DataGridView12.Columns(3).ReadOnly = True
        DataGridView12.Columns(4).HeaderText = "Название товара в Scala"
        DataGridView12.Columns(4).Width = 400
        DataGridView12.Columns(4).ReadOnly = True
        DataGridView12.Columns(5).HeaderText = "Код товара поставщика"
        DataGridView12.Columns(5).Width = 200
        DataGridView12.Columns(5).ReadOnly = True
        DataGridView12.Columns(6).HeaderText = "Код поставщика"
        DataGridView12.Columns(6).Width = 100
        DataGridView12.Columns(6).ReadOnly = True
        DataGridView12.Columns(7).HeaderText = "Название поставщика"
        DataGridView12.Columns(7).Width = 400
        DataGridView12.Columns(7).ReadOnly = True
    End Sub

    Private Sub Button2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button2.Click
        '/////////////////////////////////////////////////////////////////////////////////////
        '//
        '// Связывание картинок с кодом Скала или удаление и выход
        '// 
        '/////////////////////////////////////////////////////////////////////////////////////

        Me.Cursor = Cursors.WaitCursor
        DeleteSelPictures()
        Me.Cursor = Cursors.Default
        Me.Close()
    End Sub

    Private Sub DeleteSelPictures()
        '/////////////////////////////////////////////////////////////////////////////////////
        '//
        '// Удаление выделенных картинок
        '// 
        '/////////////////////////////////////////////////////////////////////////////////////

        Try
            For i As Integer = 0 To DataGridView12.Rows.Count - 1
                If DataGridView12.Item(2, i).Value = True Then  '---удаление
                    DeletePicture(Trim(DataGridView12.Item(0, i).Value.ToString))
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
End Class