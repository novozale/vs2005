Public Class ItemSelectList2

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

        For i As Integer = 0 To MainForm.DataGridView2.Rows.Count - 1
            If Trim(MainForm.DataGridView2.Item(0, i).Value.ToString) = Trim(Me.DataGridView1.SelectedRows.Item(0).Cells(0).Value.ToString()) Then
                MainForm.DataGridView2.CurrentCell = MainForm.DataGridView2.Item(0, i)
                Me.Close()
                Exit Sub
            End If
        Next
        Me.Close()
    End Sub

    Private Sub ItemSelectList2_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
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
            MySQLStr = MySQLStr & "tbl_ForecastOrderR4_Main_DC.MGZ AS AMGZ, "
            MySQLStr = MySQLStr & "tbl_ForecastOrderR4_Main_DC.ROP AS AROP, "
            MySQLStr = MySQLStr & "tbl_ForecastOrderR4_Main_DC.InshuranceLVL AS AIshuranceLVL, "
            MySQLStr = MySQLStr & "tbl_ForecastOrderR4_CustomMGZROPINS_DC.MGZ, "
            MySQLStr = MySQLStr & "tbl_ForecastOrderR4_CustomMGZROPINS_DC.ROP, "
            MySQLStr = MySQLStr & "tbl_ForecastOrderR4_CustomMGZROPINS_DC.IshuranceLVL, "
            MySQLStr = MySQLStr & "tbl_ForecastOrderR4_CustomMGZROPINS_DC.DueDate "
            MySQLStr = MySQLStr & "FROM  tbl_ForecastOrderR4_Main_DC INNER JOIN "
            MySQLStr = MySQLStr & "tbl_ForecastOrderR4_CustomMGZROPINS_DC ON tbl_ForecastOrderR4_Main_DC.Code = tbl_ForecastOrderR4_CustomMGZROPINS_DC.Code "
            MySQLStr = MySQLStr & "WHERE (1 = 1) "
            If Trim(MainForm.TextBox4.Text) = "" Then
                '----В первое окно условие не введено - считаем, что во второе введено
                MySQLStr = MySQLStr & "AND ((UPPER(tbl_ForecastOrderR4_Main_DC.Code) LIKE N'%" & UCase(MainForm.TextBox1.Text) & "%') "
                MySQLStr = MySQLStr & " OR (UPPER(tbl_ForecastOrderR4_Main_DC.Name) LIKE N'%" & UCase(MainForm.TextBox1.Text) & "%')) "
            Else
                If Trim(MainForm.TextBox1.Text) = "" Then
                    '----Во второе окно условие не введено
                    MySQLStr = MySQLStr & "AND ((UPPER(tbl_ForecastOrderR4_Main_DC.Code) LIKE N'%" & UCase(MainForm.TextBox4.Text) & "%') "
                    MySQLStr = MySQLStr & " OR (UPPER(tbl_ForecastOrderR4_Main_DC.Name) LIKE N'%" & UCase(MainForm.TextBox4.Text) & "%')) "
                Else
                    '----Условия введены в оба окна
                    MySQLStr = MySQLStr & "AND (((UPPER(tbl_ForecastOrderR4_Main_DC.Code) LIKE N'%" & UCase(MainForm.TextBox4.Text) & "%') "
                    MySQLStr = MySQLStr & "AND (UPPER(tbl_ForecastOrderR4_Main_DC.Code) LIKE N'%" & UCase(MainForm.TextBox1.Text) & "%')) "
                    MySQLStr = MySQLStr & " OR ((UPPER(tbl_ForecastOrderR4_Main_DC.Name) LIKE N'%" & UCase(MainForm.TextBox4.Text) & "%') "
                    MySQLStr = MySQLStr & " AND (UPPER(tbl_ForecastOrderR4_Main_DC.Name) LIKE N'%" & UCase(MainForm.TextBox1.Text) & "%'))) "
                End If
            End If
            MySQLStr = MySQLStr & "Order By tbl_ForecastOrderR4_Main_DC.Code "
        Else
            MySQLStr = "SELECT tbl_ForecastOrderR4_Main_RWH.Code, "
            MySQLStr = MySQLStr & "tbl_ForecastOrderR4_Main_RWH.Name, "
            MySQLStr = MySQLStr & "tbl_ForecastOrderR4_Main_RWH.DC, "
            MySQLStr = MySQLStr & "tbl_ForecastOrderR4_Main_RWH.ABC, "
            MySQLStr = MySQLStr & "tbl_ForecastOrderR4_Main_RWH.XYZ, "
            MySQLStr = MySQLStr & "tbl_ForecastOrderR4_Main_RWH.LT, "
            MySQLStr = MySQLStr & "tbl_ForecastOrderR4_Main_RWH.OI, "
            MySQLStr = MySQLStr & "tbl_ForecastOrderR4_Main_RWH.MGZ AS AMGZ, "
            MySQLStr = MySQLStr & "tbl_ForecastOrderR4_Main_RWH.ROP AS AROP, "
            MySQLStr = MySQLStr & "tbl_ForecastOrderR4_Main_RWH.InshuranceLVL AS AIshuranceLVL, "
            MySQLStr = MySQLStr & "tbl_ForecastOrderR4_CustomMGZROPINS_RWH.MGZ, "
            MySQLStr = MySQLStr & "tbl_ForecastOrderR4_CustomMGZROPINS_RWH.ROP, "
            MySQLStr = MySQLStr & "tbl_ForecastOrderR4_CustomMGZROPINS_RWH.IshuranceLVL, "
            MySQLStr = MySQLStr & "tbl_ForecastOrderR4_CustomMGZROPINS_RWH.DueDate "
            MySQLStr = MySQLStr & "FROM  tbl_ForecastOrderR4_Main_RWH INNER JOIN "
            MySQLStr = MySQLStr & "tbl_ForecastOrderR4_CustomMGZROPINS_RWH ON tbl_ForecastOrderR4_Main_RWH.Code = tbl_ForecastOrderR4_CustomMGZROPINS_RWH.Code AND "
            MySQLStr = MySQLStr & "tbl_ForecastOrderR4_Main_RWH.WarNo = tbl_ForecastOrderR4_CustomMGZROPINS_RWH.WH "
            MySQLStr = MySQLStr & "WHERE (tbl_ForecastOrderR4_Main_RWH.WarNo = N'" & MainForm.ComboBox1.SelectedValue & "') "
            If Trim(MainForm.TextBox4.Text) = "" Then
                '----В первое окно условие не введено - считаем, что во второе введено
                MySQLStr = MySQLStr & "AND ((UPPER(tbl_ForecastOrderR4_Main_RWH.Code) LIKE N'%" & UCase(MainForm.TextBox1.Text) & "%') "
                MySQLStr = MySQLStr & " OR (UPPER(tbl_ForecastOrderR4_Main_RWH.Name) LIKE N'%" & UCase(MainForm.TextBox1.Text) & "%')) "
            Else
                If Trim(MainForm.TextBox1.Text) = "" Then
                    '----Во второе окно условие не введено
                    MySQLStr = MySQLStr & "AND ((UPPER(tbl_ForecastOrderR4_Main_RWH.Code) LIKE N'%" & UCase(MainForm.TextBox4.Text) & "%') "
                    MySQLStr = MySQLStr & " OR (UPPER(tbl_ForecastOrderR4_Main_RWH.Name) LIKE N'%" & UCase(MainForm.TextBox4.Text) & "%')) "
                Else
                    '----Условия введены в оба окна
                    MySQLStr = MySQLStr & "AND (((UPPER(tbl_ForecastOrderR4_Main_RWH.Code) LIKE N'%" & UCase(MainForm.TextBox4.Text) & "%') "
                    MySQLStr = MySQLStr & "AND (UPPER(tbl_ForecastOrderR4_Main_RWH.Code) LIKE N'%" & UCase(MainForm.TextBox1.Text) & "%')) "
                    MySQLStr = MySQLStr & " OR ((UPPER(tbl_ForecastOrderR4_Main_RWH.Name) LIKE N'%" & UCase(MainForm.TextBox4.Text) & "%') "
                    MySQLStr = MySQLStr & " AND (UPPER(tbl_ForecastOrderR4_Main_RWH.Name) LIKE N'%" & UCase(MainForm.TextBox1.Text) & "%'))) "
                End If
            End If
            MySQLStr = MySQLStr & "Order By tbl_ForecastOrderR4_Main_RWH.Code "
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
            DataGridView1.Columns(7).HeaderText = "авто МЖЗ"
            DataGridView1.Columns(7).Width = 60
            DataGridView1.Columns(8).HeaderText = "авто ROP"
            DataGridView1.Columns(8).Width = 60
            DataGridView1.Columns(9).HeaderText = "авто Страх уровень"
            DataGridView1.Columns(9).Width = 60
            DataGridView1.Columns(10).HeaderText = "ручн МЖЗ"
            DataGridView1.Columns(10).Width = 60
            DataGridView1.Columns(11).HeaderText = "ручн ROP"
            DataGridView1.Columns(11).Width = 60
            DataGridView1.Columns(12).HeaderText = "ручн Страх уровень"
            DataGridView1.Columns(12).Width = 60
            DataGridView1.Columns(13).HeaderText = "До даты"
            DataGridView1.Columns(13).Width = 100

            DataGridView1.SelectionMode = DataGridViewSelectionMode.FullRowSelect

            DataGridView1.Columns(7).DefaultCellStyle.Format = "### ##0.00"
            DataGridView1.Columns(8).DefaultCellStyle.Format = "### ##0.00"
            DataGridView1.Columns(9).DefaultCellStyle.Format = "### ##0.00"
            DataGridView1.Columns(10).DefaultCellStyle.Format = "### ##0.00"
            DataGridView1.Columns(10).DefaultCellStyle.BackColor = Color.LightBlue
            DataGridView1.Columns(11).DefaultCellStyle.Format = "### ##0.00"
            DataGridView1.Columns(11).DefaultCellStyle.BackColor = Color.LightBlue
            DataGridView1.Columns(12).DefaultCellStyle.Format = "### ##0.00"
            DataGridView1.Columns(12).DefaultCellStyle.BackColor = Color.LightBlue
            DataGridView1.Columns(13).DefaultCellStyle.Format = "d"

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
        If row.Cells(7).Value <> 0 Or row.Cells(8).Value <> 0 Or row.Cells(9).Value <> 0 Then
            row.DefaultCellStyle.BackColor = Color.Yellow
        End If
    End Sub
End Class