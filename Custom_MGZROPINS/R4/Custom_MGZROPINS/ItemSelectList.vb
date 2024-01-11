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

        If Declarations.MyWorkLevel = 0 Then          '---Работаем на уровне компании
            MySQLStr = "SELECT tbl_ForecastOrderR4_Main_DC.Code, "
            MySQLStr = MySQLStr & "tbl_ForecastOrderR4_Main_DC.Name, "
            MySQLStr = MySQLStr & "tbl_ForecastOrderR4_Main_DC.DC, "
            MySQLStr = MySQLStr & "tbl_ForecastOrderR4_Main_DC.ABC, "
            MySQLStr = MySQLStr & "tbl_ForecastOrderR4_Main_DC.XYZ, "
            MySQLStr = MySQLStr & "tbl_ForecastOrderR4_Main_DC.LT, "
            MySQLStr = MySQLStr & "tbl_ForecastOrderR4_Main_DC.OI, "
            MySQLStr = MySQLStr & "tbl_ForecastOrderR4_Main_DC.MGZ, "
            MySQLStr = MySQLStr & "tbl_ForecastOrderR4_Main_DC.ROP, "
            MySQLStr = MySQLStr & "tbl_ForecastOrderR4_Main_DC.InshuranceLVL "
            MySQLStr = MySQLStr & "FROM tbl_ForecastOrderR4_Main_DC LEFT OUTER JOIN "
            MySQLStr = MySQLStr & "(SELECT Code FROM  dbo.tbl_ForecastOrderR4_CustomMGZROPINS_DC) AS View_1 ON tbl_ForecastOrderR4_Main_DC.Code = View_1.Code "
            MySQLStr = MySQLStr & "WHERE (View_1.Code IS NULL) AND "
            MySQLStr = MySQLStr & "(tbl_ForecastOrderR4_Main_DC.WHass = 1) "
            If Trim(MainForm.TextBox2.Text) = "" Then
                '----В первое окно условие не введено - считаем, что во второе введено
                MySQLStr = MySQLStr & "AND ((UPPER(tbl_ForecastOrderR4_Main_DC.Code) LIKE N'%" & UCase(MainForm.TextBox3.Text) & "%') "
                MySQLStr = MySQLStr & " OR (UPPER(tbl_ForecastOrderR4_Main_DC.Name) LIKE N'%" & UCase(MainForm.TextBox3.Text) & "%')) "
            Else
                If Trim(MainForm.TextBox3.Text) = "" Then
                    '----Во второе окно условие не введено
                    MySQLStr = MySQLStr & "AND ((UPPER(tbl_ForecastOrderR4_Main_DC.Code) LIKE N'%" & UCase(MainForm.TextBox2.Text) & "%') "
                    MySQLStr = MySQLStr & " OR (UPPER(tbl_ForecastOrderR4_Main_DC.Name) LIKE N'%" & UCase(MainForm.TextBox2.Text) & "%')) "
                Else
                    '----Условия введены в оба окна
                    MySQLStr = MySQLStr & "AND (((UPPER(tbl_ForecastOrderR4_Main_DC.Code) LIKE N'%" & UCase(MainForm.TextBox2.Text) & "%') "
                    MySQLStr = MySQLStr & "AND (UPPER(tbl_ForecastOrderR4_Main_DC.Code) LIKE N'%" & UCase(MainForm.TextBox3.Text) & "%')) "
                    MySQLStr = MySQLStr & " OR ((UPPER(tbl_ForecastOrderR4_Main_DC.Name) LIKE N'%" & UCase(MainForm.TextBox2.Text) & "%') "
                    MySQLStr = MySQLStr & " AND (UPPER(tbl_ForecastOrderR4_Main_DC.Name) LIKE N'%" & UCase(MainForm.TextBox3.Text) & "%'))) "
                End If
            End If
            MySQLStr = MySQLStr & "Order BY tbl_ForecastOrderR4_Main_DC.Code "
        Else
            MySQLStr = "SELECT View_2.Code, "
            MySQLStr = MySQLStr & "View_2.Name, "
            MySQLStr = MySQLStr & "View_2.DC, "
            MySQLStr = MySQLStr & "View_2.ABC, "
            MySQLStr = MySQLStr & "View_2.XYZ, "
            MySQLStr = MySQLStr & "View_2.LT, "
            MySQLStr = MySQLStr & "View_2.OI, "
            MySQLStr = MySQLStr & "View_2.MGZ, "
            MySQLStr = MySQLStr & "View_2.ROP, "
            MySQLStr = MySQLStr & "View_2.InshuranceLVL "
            MySQLStr = MySQLStr & "FROM (SELECT Code, WH "
            MySQLStr = MySQLStr & "FROM tbl_ForecastOrderR4_CustomMGZROPINS_RWH "
            MySQLStr = MySQLStr & "WHERE (WH = N'" & MainForm.ComboBox1.SelectedValue & "')) AS View_1 RIGHT OUTER JOIN "
            MySQLStr = MySQLStr & "(SELECT Code, Name, DC, ABC, XYZ, LT, OI, MGZ, ROP, InshuranceLVL, WarNo "
            MySQLStr = MySQLStr & "FROM tbl_ForecastOrderR4_Main_RWH "
            MySQLStr = MySQLStr & "WHERE (WarNo = N'" & MainForm.ComboBox1.SelectedValue & "') AND ("
            MySQLStr = MySQLStr & "WHass = 1) AND "
            MySQLStr = MySQLStr & "(DC <> WarNo)) AS View_2 ON View_1.Code = View_2.Code AND View_1.WH = View_2.WarNo "
            MySQLStr = MySQLStr & "WHERE (View_1.Code Is NULL) AND (View_1.WH Is NULL) "
            If Trim(MainForm.TextBox2.Text) = "" Then
                '----В первое окно условие не введено - считаем, что во второе введено
                MySQLStr = MySQLStr & "AND ((UPPER(View_2.Code) LIKE N'%" & UCase(MainForm.TextBox3.Text) & "%') "
                MySQLStr = MySQLStr & " OR (UPPER(View_2.Name) LIKE N'%" & UCase(MainForm.TextBox3.Text) & "%')) "
            Else
                If Trim(MainForm.TextBox3.Text) = "" Then
                    '----Во второе окно условие не введено
                    MySQLStr = MySQLStr & "AND ((UPPER(View_2.Code) LIKE N'%" & UCase(MainForm.TextBox2.Text) & "%') "
                    MySQLStr = MySQLStr & " OR (UPPER(View_2.Name) LIKE N'%" & UCase(MainForm.TextBox2.Text) & "%')) "
                Else
                    '----Условия введены в оба окна
                    MySQLStr = MySQLStr & "AND (((UPPER(View_2.Code) LIKE N'%" & UCase(MainForm.TextBox2.Text) & "%') "
                    MySQLStr = MySQLStr & "AND (UPPER(View_2.Code) LIKE N'%" & UCase(MainForm.TextBox3.Text) & "%')) "
                    MySQLStr = MySQLStr & " OR ((UPPER(View_2.Name) LIKE N'%" & UCase(MainForm.TextBox2.Text) & "%') "
                    MySQLStr = MySQLStr & " AND (UPPER(View_2.Name) LIKE N'%" & UCase(MainForm.TextBox3.Text) & "%'))) "
                End If
            End If
            MySQLStr = MySQLStr & "Order BY View_2.Code "
        End If
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
            DataGridView1.Columns(2).HeaderText = "DC"
            DataGridView1.Columns(2).Width = 50
            DataGridView1.Columns(3).HeaderText = "ABC"
            DataGridView1.Columns(3).Width = 50
            DataGridView1.Columns(4).HeaderText = "XYZ"
            DataGridView1.Columns(4).Width = 50
            DataGridView1.Columns(5).HeaderText = "LT"
            DataGridView1.Columns(5).Width = 50
            DataGridView1.Columns(6).HeaderText = "OI"
            DataGridView1.Columns(6).Width = 50
            DataGridView1.Columns(7).HeaderText = "МЖЗ"
            DataGridView1.Columns(7).Width = 60
            DataGridView1.Columns(8).HeaderText = "ROP"
            DataGridView1.Columns(8).Width = 60
            DataGridView1.Columns(9).HeaderText = "Страх уровень"
            DataGridView1.Columns(9).Width = 60

            DataGridView1.SelectionMode = DataGridViewSelectionMode.FullRowSelect

            DataGridView1.Columns(7).DefaultCellStyle.Format = "### ##0.00"
            DataGridView1.Columns(8).DefaultCellStyle.Format = "### ##0.00"
            DataGridView1.Columns(9).DefaultCellStyle.Format = "### ##0.00"

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
        If row.Cells(6).Value = 0 Or row.Cells(7).Value = 0 Or row.Cells(8).Value = 0 Then
            row.DefaultCellStyle.BackColor = Color.Yellow
        End If
    End Sub
End Class