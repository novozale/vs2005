Public Class CheckUpdateNamesDescr
    Public MyItemPart As String         '----флаг - Названия или описания
    Public MySelection As Integer       '----флаг - перезаписывать старые названия (описания) или нет
    Public MyCompanyCode As Integer     '----код компании, с сайта которой берутся названия / описания

    

    Private Sub Button3_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button3.Click
        '/////////////////////////////////////////////////////////////////////////////////////
        '//
        '// выход из окна
        '// 
        '/////////////////////////////////////////////////////////////////////////////////////

        Me.Close()
    End Sub

    Private Sub CheckUpdateNamesDescr_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        '/////////////////////////////////////////////////////////////////////////////////////
        '//
        '// Загрузка данных в окно
        '// 
        '/////////////////////////////////////////////////////////////////////////////////////

        Dim MySQLStr As String                        'рабочая строка
        Dim MyAdapter As SqlClient.SqlDataAdapter     '
        Dim MyDs As New DataSet                       '

        '---тут выводим список товаров и предлагаемые названия / описания
        If MyItemPart = "Названия" Then         '------Перенос названий
            MySQLStr = "SELECT tbl_WEB_Pictures.PictureMedium, CONVERT(bit, CASE WHEN ISNULL(View_1.TotalQTY, 0) = 1 THEN 1 ELSE 0 END) AS ToMatch, "
            MySQLStr = MySQLStr & "tbl_WEB_Items.Code, tbl_WEB_Items.Name, tbl_WEBSearchScrapping_Result.ItemName , CASE WHEN ISNULL(View_1.TotalQTY, 0) = 1 THEN 1 ELSE 0 END AS ToMatchSrc, "
            MySQLStr = MySQLStr & "tbl_WEBSearchScrapping_Result.ID "
            MySQLStr = MySQLStr & "FROM tbl_WEB_Items INNER JOIN "
            MySQLStr = MySQLStr & "tbl_WEBSearchScrapping_Result ON tbl_WEB_Items.ManufacturerItemCode = tbl_WEBSearchScrapping_Result.SuppItemID LEFT OUTER JOIN "
            MySQLStr = MySQLStr & "(SELECT SC01060, COUNT(SC01060) AS TotalQTY "
            MySQLStr = MySQLStr & "FROM SC010300 "
            MySQLStr = MySQLStr & "GROUP BY SC01060) AS View_1 ON tbl_WEB_Items.ManufacturerItemCode = View_1.SC01060 LEFT OUTER JOIN "
            MySQLStr = MySQLStr & "tbl_WEB_Pictures ON tbl_WEB_Items.Code = tbl_WEB_Pictures.ScalaItemCode "
            MySQLStr = MySQLStr & "WHERE (tbl_WEBSearchScrapping_Result.CompanyID = " & MyCompanyCode & ") "
            MySQLStr = MySQLStr & "AND (tbl_WEBSearchScrapping_Result.ItemName <> N'') AND "
            MySQLStr = MySQLStr & "(tbl_WEBSearchScrapping_Result.SuppItemID NOT IN "
            MySQLStr = MySQLStr & "(SELECT SC010300_1.SC01060 "
            MySQLStr = MySQLStr & "FROM tbl_WEBSearchScrapping_NotCorrectSuppl INNER JOIN "
            MySQLStr = MySQLStr & "SC010300 AS SC010300_1 ON tbl_WEBSearchScrapping_NotCorrectSuppl.PL01001 = SC010300_1.SC01058 "
            MySQLStr = MySQLStr & "WHERE (tbl_WEBSearchScrapping_NotCorrectSuppl.CompanyID = " & MyCompanyCode & "))) "
            If MySelection = 0 Then
                MySQLStr = MySQLStr & "AND (tbl_WEB_Items.WEBName = N'') "
            End If
        Else                                   '------Перенос описаний
            MySQLStr = "SELECT tbl_WEB_Pictures.PictureMedium, CONVERT(bit, CASE WHEN ISNULL(View_1.TotalQTY, 0) = 1 THEN 1 ELSE 0 END) AS ToMatch, "
            MySQLStr = MySQLStr & "tbl_WEB_Items.Code, tbl_WEB_Items.Name, tbl_WEBSearchScrapping_Result.ItemDescription, CASE WHEN ISNULL(View_1.TotalQTY, 0) = 1 THEN 1 ELSE 0 END AS ToMatchSrc, "
            MySQLStr = MySQLStr & "tbl_WEBSearchScrapping_Result.ID "
            MySQLStr = MySQLStr & "FROM tbl_WEB_Items INNER JOIN "
            MySQLStr = MySQLStr & "tbl_WEBSearchScrapping_Result ON tbl_WEB_Items.ManufacturerItemCode = tbl_WEBSearchScrapping_Result.SuppItemID LEFT OUTER JOIN "
            MySQLStr = MySQLStr & "(SELECT SC01060, COUNT(SC01060) AS TotalQTY "
            MySQLStr = MySQLStr & "FROM SC010300 "
            MySQLStr = MySQLStr & "GROUP BY SC01060) AS View_1 ON tbl_WEB_Items.ManufacturerItemCode = View_1.SC01060 LEFT OUTER JOIN "
            MySQLStr = MySQLStr & "tbl_WEB_Pictures ON tbl_WEB_Items.Code = tbl_WEB_Pictures.ScalaItemCode "
            MySQLStr = MySQLStr & "WHERE (tbl_WEBSearchScrapping_Result.CompanyID = " & MyCompanyCode & ") "
            MySQLStr = MySQLStr & "AND (tbl_WEBSearchScrapping_Result.ItemDescription NOT LIKE N'') AND "
            MySQLStr = MySQLStr & "(tbl_WEBSearchScrapping_Result.SuppItemID NOT IN "
            MySQLStr = MySQLStr & "(SELECT SC010300_1.SC01060 "
            MySQLStr = MySQLStr & "FROM tbl_WEBSearchScrapping_NotCorrectSuppl INNER JOIN "
            MySQLStr = MySQLStr & "SC010300 AS SC010300_1 ON tbl_WEBSearchScrapping_NotCorrectSuppl.PL01001 = SC010300_1.SC01058 "
            MySQLStr = MySQLStr & "WHERE (tbl_WEBSearchScrapping_NotCorrectSuppl.CompanyID = " & MyCompanyCode & "))) "
            If MySelection = 0 Then
                MySQLStr = MySQLStr & "AND (tbl_WEB_Items.Description = N'') "
            End If
        End If
        
        If MyItemPart = "Названия" Then         '------Перенос названий
            DataGridView12.RowTemplate.MinimumHeight = 100
        Else                                   '------Перенос описаний
            DataGridView12.RowTemplate.MinimumHeight = 300
        End If


        InitMyConn(False)
        Try
            MyAdapter = New SqlClient.SqlDataAdapter(MySQLStr, Declarations.MyNETConnStr)
            MyAdapter.SelectCommand.CommandTimeout = 600
            MyAdapter.Fill(MyDs)
            DataGridView12.DataSource = MyDs.Tables(0)
        Catch ex As Exception
            MsgBox(ex.ToString)
        End Try

        DataGridView12.Columns(0).HeaderText = "Картинка"
        DataGridView12.Columns(0).Width = 100
        DataGridView12.Columns(0).ReadOnly = True
        DataGridView12.Columns(1).HeaderText = "Загру зить"
        DataGridView12.Columns(1).Width = 50
        DataGridView12.Columns(2).HeaderText = "Код товара в Scala"
        DataGridView12.Columns(2).Width = 100
        DataGridView12.Columns(2).ReadOnly = True
        DataGridView12.Columns(3).HeaderText = "Название товара в Scala"
        DataGridView12.Columns(3).Width = 300
        DataGridView12.Columns(3).ReadOnly = True
        If MyItemPart = "Названия" Then         '------Перенос названий
            DataGridView12.Columns(4).HeaderText = "Предлагаемое название"
            DataGridView12.Columns(4).Width = 600
            DataGridView12.Columns(4).ReadOnly = True
        Else
            DataGridView12.Columns(4).HeaderText = "Предлагаемое описание"
            DataGridView12.Columns(4).Width = 600
            DataGridView12.Columns(4).ReadOnly = True
        End If
        DataGridView12.Columns(5).HeaderText = "Состояние"
        DataGridView12.Columns(5).Width = 50
        DataGridView12.Columns(5).Visible = False
        DataGridView12.Columns(6).HeaderText = "ID"
        DataGridView12.Columns(6).Width = 50
        DataGridView12.Columns(6).Visible = False
        
    End Sub

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        '////////////////////////////////////////////////////////////////////////////////
        '//
        '// Снятие всех галочек связывания названий / описаний 
        '//
        '////////////////////////////////////////////////////////////////////////////////

        Try
            For i As Integer = 0 To DataGridView12.Rows.Count - 1
                DataGridView12.Item(1, i).Value = False
            Next i
        Catch ex As Exception
        End Try
    End Sub

    Private Sub Button4_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button4.Click
        '////////////////////////////////////////////////////////////////////////////////
        '//
        '// Выставление всех галочек связывания названий / описаний 
        '//
        '////////////////////////////////////////////////////////////////////////////////

        Try
            For i As Integer = 0 To DataGridView12.Rows.Count - 1
                DataGridView12.Item(1, i).Value = True
            Next i
        Catch ex As Exception
        End Try
    End Sub

    Private Sub DataGridView12_CellFormatting(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellFormattingEventArgs) Handles DataGridView12.CellFormatting
        '////////////////////////////////////////////////////////////////////////////////
        '//
        '// Подсветка строк продуктов в зависимости от статуса
        '//
        '////////////////////////////////////////////////////////////////////////////////
        Dim row As DataGridViewRow

        row = DataGridView12.Rows(e.RowIndex)
        If row.Cells(5).Value = False Then
            row.DefaultCellStyle.BackColor = Color.LightGray
        End If
    End Sub

    Private Sub Button2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button2.Click
        '/////////////////////////////////////////////////////////////////////////////////////
        '//
        '// Перенос названий или описаний 
        '// 
        '/////////////////////////////////////////////////////////////////////////////////////

        Me.Cursor = Cursors.WaitCursor
        UpdateNamesDescr()
        Me.Cursor = Cursors.Default
        Me.Close()
    End Sub

    Private Sub UpdateNamesDescr()
        '/////////////////////////////////////////////////////////////////////////////////////
        '//
        '// Процедура переноса названий или описаний 
        '// 
        '/////////////////////////////////////////////////////////////////////////////////////

        Try
            For i As Integer = 0 To DataGridView12.Rows.Count - 1
                If DataGridView12.Item(1, i).Value = True Then  '---Обновление
                    UpdateOneNameDescr(DataGridView12.Item(2, i).Value, DataGridView12.Item(6, i).Value.ToString)
                End If
            Next i
        Catch ex As Exception
        End Try
    End Sub

    Private Sub UpdateOneNameDescr(ByVal ScalaCode As String, ByVal UpdateStr As String)
        '/////////////////////////////////////////////////////////////////////////////////////
        '//
        '// Процедура обновления названия или описания для одного товара 
        '// 
        '/////////////////////////////////////////////////////////////////////////////////////
        Dim MySQLStr As String

        If MyItemPart = "Названия" Then         '------Перенос названий
            MySQLStr = "UPDATE tbl_WEB_Items "
            MySQLStr = MySQLStr & "SET WEBName = tbl_WEBSearchScrapping_Result.ItemName, "
            MySQLStr = MySQLStr & "RMStatus = CASE WHEN RMStatus = 1 THEN 1 ELSE CASE WHEN RMStatus = 2 THEN 2 ELSE 3 END END, "
            MySQLStr = MySQLStr & "WEBStatus = CASE WHEN WEBStatus = 1 THEN 1 ELSE CASE WHEN WEBStatus = 2 THEN 2 ELSE 3 END END "
            MySQLStr = MySQLStr & "FROM tbl_WEB_Items CROSS JOIN "
            MySQLStr = MySQLStr & "tbl_WEBSearchScrapping_Result "
            MySQLStr = MySQLStr & "WHERE (tbl_WEB_Items.Code = N'" & Trim(ScalaCode) & "') "
            MySQLStr = MySQLStr & "AND (tbl_WEBSearchScrapping_Result.ID = N'" & Trim(UpdateStr) & "') "
        Else                                    '------Перенос описаний
            MySQLStr = "UPDATE tbl_WEB_Items "
            MySQLStr = MySQLStr & "SET Description = tbl_WEBSearchScrapping_Result.ItemDescription, "
            MySQLStr = MySQLStr & "RMStatus = CASE WHEN RMStatus = 1 THEN 1 ELSE CASE WHEN RMStatus = 2 THEN 2 ELSE 3 END END, "
            MySQLStr = MySQLStr & "WEBStatus = CASE WHEN WEBStatus = 1 THEN 1 ELSE CASE WHEN WEBStatus = 2 THEN 2 ELSE 3 END END "
            MySQLStr = MySQLStr & "FROM tbl_WEB_Items CROSS JOIN "
            MySQLStr = MySQLStr & "tbl_WEBSearchScrapping_Result "
            MySQLStr = MySQLStr & "WHERE (tbl_WEB_Items.Code = N'" & Trim(ScalaCode) & "') "
            MySQLStr = MySQLStr & "AND (tbl_WEBSearchScrapping_Result.ID = N'" & Trim(UpdateStr) & "') "
        End If
        InitMyConn(False)
        Declarations.MyConn.Execute(MySQLStr)
    End Sub
End Class