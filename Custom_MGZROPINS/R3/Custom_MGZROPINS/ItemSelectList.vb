Public Class ItemSelectList

    Private Sub Button5_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button5.Click
        '/////////////////////////////////////////////////////////////////////////////////////
        '//
        '// Выход без выбора запаса
        '// 
        '/////////////////////////////////////////////////////////////////////////////////////

        Me.Close()
    End Sub

    Private Sub Button4_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button4.Click
        '/////////////////////////////////////////////////////////////////////////////////////
        '//
        '// Выход с выбором запаса
        '// 
        '/////////////////////////////////////////////////////////////////////////////////////

        For i As Integer = 0 To MainForm.DataGridView1.Rows.Count - 1
            If Trim(MainForm.DataGridView1.Item(0, i).Value.ToString) = Trim(Me.DataGridView1.SelectedRows.Item(0).Cells(0).Value.ToString()) Then
                MainForm.DataGridView1.CurrentCell = MainForm.DataGridView1.Item(0, i)
                Me.Close()
                Exit Sub
            End If
        Next
        Me.Close()
    End Sub

    Private Sub ItemSelectList_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        '/////////////////////////////////////////////////////////////////////////////////////
        '//
        '// При запуске загружаем список запасов
        '// 
        '/////////////////////////////////////////////////////////////////////////////////////
        Dim MySQLStr As String                        'рабочая строка
        Dim MyAdapter As SqlClient.SqlDataAdapter     '
        Dim MyDs As New DataSet                       '

        MySQLStr = "SELECT tbl_ForecastOrderR3_Main.Code, "
        MySQLStr = MySQLStr & "tbl_ForecastOrderR3_Main.Name, "
        MySQLStr = MySQLStr & "tbl_ForecastOrderR3_Main.ABC, "
        MySQLStr = MySQLStr & "tbl_ForecastOrderR3_Main.XYZ, "
        MySQLStr = MySQLStr & "tbl_ForecastOrderR3_Main.LT, "
        MySQLStr = MySQLStr & "tbl_ForecastOrderR3_Main.OI, "
        MySQLStr = MySQLStr & "ROUND(CONVERT(float, tbl_ForecastOrderR3_Main.MGZ), 3) AS MGZ, "
        MySQLStr = MySQLStr & "ROUND(CONVERT(float, ISNULL(View_1.MGZ, tbl_ForecastOrderR3_Main.MGZ)), 3) AS MGZ_OLD, "
        MySQLStr = MySQLStr & "ROUND(CONVERT(float, tbl_ForecastOrderR3_Main.ROP), 3) AS ROP, "
        MySQLStr = MySQLStr & "ROUND(CONVERT(float, ISNULL(View_1.ROP, tbl_ForecastOrderR3_Main.ROP)), 3) AS ROP_OLD, "
        MySQLStr = MySQLStr & "ROUND(CONVERT(float, tbl_ForecastOrderR3_Main.InshuranceLVL), 3) AS InshuranceLVL, "
        MySQLStr = MySQLStr & "ROUND(CONVERT(float, ISNULL(View_1.InshuranceLVL, tbl_ForecastOrderR3_Main.InshuranceLVL)), 3) AS InshuranceLVL_OLD "
        MySQLStr = MySQLStr & "FROM tbl_ForecastOrderR3_Main WITH(NOLOCK) LEFT OUTER JOIN "
        MySQLStr = MySQLStr & "(SELECT tbl_ForecastOrderR3_Main_History.Code, "
        MySQLStr = MySQLStr & "tbl_ForecastOrderR3_Main_History.MGZ, "
        MySQLStr = MySQLStr & "tbl_ForecastOrderR3_Main_History.ROP, "
        MySQLStr = MySQLStr & "tbl_ForecastOrderR3_Main_History.InshuranceLVL "
        MySQLStr = MySQLStr & "FROM tbl_ForecastOrderR3_Main_History WITH(NOLOCK) INNER JOIN "
        MySQLStr = MySQLStr & "(SELECT Code, WarNo, MAX(Date) AS Expr1 "
        MySQLStr = MySQLStr & "FROM (SELECT tbl_ForecastOrderR3_Main_History_2.Code, "
        MySQLStr = MySQLStr & "tbl_ForecastOrderR3_Main_History_2.WarNo, "
        MySQLStr = MySQLStr & "tbl_ForecastOrderR3_Main_History_2.Date "
        MySQLStr = MySQLStr & "FROM tbl_ForecastOrderR3_Main_History AS tbl_ForecastOrderR3_Main_History_2 WITH(NOLOCK) LEFT OUTER JOIN "
        MySQLStr = MySQLStr & "(SELECT Code, WarNo, MAX(Date) AS Expr1 "
        MySQLStr = MySQLStr & "FROM  tbl_ForecastOrderR3_Main_History AS tbl_ForecastOrderR3_Main_History_1 WITH(NOLOCK) "
        MySQLStr = MySQLStr & "WHERE (WarNo = N'" & MainForm.ComboBox1.SelectedValue & "') "
        MySQLStr = MySQLStr & "GROUP BY Code, WarNo) AS View_2 ON tbl_ForecastOrderR3_Main_History_2.Code = View_2.Code AND "
        MySQLStr = MySQLStr & "tbl_ForecastOrderR3_Main_History_2.WarNo = View_2.WarNo AND "
        MySQLStr = MySQLStr & "tbl_ForecastOrderR3_Main_History_2.Date = View_2.Expr1 "
        MySQLStr = MySQLStr & "WHERE (tbl_ForecastOrderR3_Main_History_2.WarNo = N'" & MainForm.ComboBox1.SelectedValue & "') AND "
        MySQLStr = MySQLStr & "(View_2.Expr1 IS NULL)) AS View_3 "
        MySQLStr = MySQLStr & "GROUP BY Code, WarNo) AS View_4 ON "
        MySQLStr = MySQLStr & "tbl_ForecastOrderR3_Main_History.Code = View_4.Code AND "
        MySQLStr = MySQLStr & "tbl_ForecastOrderR3_Main_History.WarNo = View_4.WarNo AND "
        MySQLStr = MySQLStr & "tbl_ForecastOrderR3_Main_History.Date = View_4.Expr1) AS View_1 ON "
        MySQLStr = MySQLStr & "tbl_ForecastOrderR3_Main.Code = View_1.Code "
        MySQLStr = MySQLStr & "WHERE (tbl_ForecastOrderR3_Main.WarNo = N'" & MainForm.ComboBox1.SelectedValue & "') AND "
        MySQLStr = MySQLStr & "(tbl_ForecastOrderR3_Main.WHass = - 1) AND "
        MySQLStr = MySQLStr & "(tbl_ForecastOrderR3_Main.Code NOT IN "
        MySQLStr = MySQLStr & "(SELECT Code "
        MySQLStr = MySQLStr & "FROM tbl_ForecastOrderR2_CustomMGZROPINS WITH(NOLOCK) "
        MySQLStr = MySQLStr & "WHERE (WH = N'" & MainForm.ComboBox1.SelectedValue & "'))) "
        If Trim(MainForm.TextBox2.Text) = "" Then
            '----В первое окно условие не введено - считаем, что во второе введено
            MySQLStr = MySQLStr & "AND ((UPPER(tbl_ForecastOrderR3_Main.Code) LIKE N'%" & UCase(MainForm.TextBox3.Text) & "%') "
            MySQLStr = MySQLStr & " OR (UPPER(tbl_ForecastOrderR3_Main.Name) LIKE N'%" & UCase(MainForm.TextBox3.Text) & "%')) "
        Else
            If Trim(MainForm.TextBox3.Text) = "" Then
                '----Во второе окно условие не введено
                MySQLStr = MySQLStr & "AND ((UPPER(tbl_ForecastOrderR3_Main.Code) LIKE N'%" & UCase(MainForm.TextBox2.Text) & "%') "
                MySQLStr = MySQLStr & " OR (UPPER(tbl_ForecastOrderR3_Main.Name) LIKE N'%" & UCase(MainForm.TextBox2.Text) & "%')) "
            Else
                '----Условия введены в оба окна
                MySQLStr = MySQLStr & "AND (((UPPER(tbl_ForecastOrderR3_Main.Code) LIKE N'%" & UCase(MainForm.TextBox2.Text) & "%') "
                MySQLStr = MySQLStr & "AND (UPPER(tbl_ForecastOrderR3_Main.Code) LIKE N'%" & UCase(MainForm.TextBox3.Text) & "%')) "
                MySQLStr = MySQLStr & " OR ((UPPER(tbl_ForecastOrderR3_Main.Name) LIKE N'%" & UCase(MainForm.TextBox2.Text) & "%') "
                MySQLStr = MySQLStr & " AND (UPPER(tbl_ForecastOrderR3_Main.Name) LIKE N'%" & UCase(MainForm.TextBox3.Text) & "%'))) "
            End If
        End If
        MySQLStr = MySQLStr & "ORDER BY tbl_ForecastOrderR3_Main.Code "

        InitMyConn(False)

        Try
            MyAdapter = New SqlClient.SqlDataAdapter(MySQLStr, Declarations.MyNETConnStr)
            MyAdapter.SelectCommand.CommandTimeout = 600
            MyAdapter.Fill(MyDs)
            DataGridView1.DataSource = MyDs.Tables(0)

            DataGridView1.Columns(0).HeaderText = "ID"
            DataGridView1.Columns(0).Width = 80
            DataGridView1.Columns(1).HeaderText = "Запас"
            DataGridView1.Columns(1).Width = 200
            DataGridView1.Columns(2).HeaderText = "ABC"
            DataGridView1.Columns(2).Width = 50
            DataGridView1.Columns(3).HeaderText = "XYZ"
            DataGridView1.Columns(3).Width = 50
            DataGridView1.Columns(4).HeaderText = "LT"
            DataGridView1.Columns(4).Width = 50
            DataGridView1.Columns(5).HeaderText = "OI"
            DataGridView1.Columns(5).Width = 50
            DataGridView1.Columns(6).HeaderText = "МЖЗ"
            DataGridView1.Columns(6).Width = 60
            DataGridView1.Columns(7).HeaderText = "МЖЗ старый"
            DataGridView1.Columns(7).Width = 60
            DataGridView1.Columns(8).HeaderText = "ROP"
            DataGridView1.Columns(8).Width = 60
            DataGridView1.Columns(9).HeaderText = "ROP старый"
            DataGridView1.Columns(9).Width = 60
            DataGridView1.Columns(10).HeaderText = "Страх уровень"
            DataGridView1.Columns(10).Width = 60
            DataGridView1.Columns(11).HeaderText = "Страх уровень старый"
            DataGridView1.Columns(11).Width = 60

            DataGridView1.SelectionMode = DataGridViewSelectionMode.FullRowSelect

            If DataGridView1.Rows.Count > 0 Then
                Button4.Enabled = True
            Else
                Button4.Enabled = False
            End If
        Catch ex As Exception
            MsgBox(ex.ToString)
        End Try
    End Sub

    Private Sub DataGridView1_CellFormatting(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellFormattingEventArgs) Handles DataGridView1.CellFormatting
        '////////////////////////////////////////////////////////////////////////////////
        '//
        '// подсветка Запасов, у которых МЖЗ 0
        '//
        '////////////////////////////////////////////////////////////////////////////////
        Dim row As DataGridViewRow

        row = DataGridView1.Rows(e.RowIndex)
        If row.Cells(6).Value = 0 Or row.Cells(8).Value = 0 Or row.Cells(10).Value = 0 Then
            row.DefaultCellStyle.BackColor = Color.Yellow
        End If
    End Sub
End Class