Public Class SupplierSelectList

    Private Sub Button5_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button5.Click
        '/////////////////////////////////////////////////////////////////////////////////////
        '//
        '// Выход без выбора поставщика
        '// 
        '/////////////////////////////////////////////////////////////////////////////////////

        Me.Close()
    End Sub

    Private Sub DataGridView1_CellFormatting(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellFormattingEventArgs) Handles DataGridView1.CellFormatting
        '////////////////////////////////////////////////////////////////////////////////
        '//
        '// подсветка информации по поставщикам
        '//
        '////////////////////////////////////////////////////////////////////////////////
        Dim row As DataGridViewRow

        row = DataGridView1.Rows(e.RowIndex)
        row.Cells(0).Style.Font = New Font(DataGridView1.DefaultCellStyle.Font, FontStyle.Bold)
        row.Cells(1).Style.Font = New Font(DataGridView1.DefaultCellStyle.Font, FontStyle.Bold)
        If Trim(row.Cells(3).Value.ToString) <> "0" Then
            row.Cells(3).Style.BackColor = Color.LightGreen
        End If
        If Trim(row.Cells(4).Value.ToString) <> "0" Then
            row.Cells(4).Style.BackColor = Color.LightGreen
        End If
        If Trim(row.Cells(5).Value.ToString) <> "0" Then
            row.Cells(5).Style.BackColor = Color.LightGreen
        End If
        If Trim(row.Cells(6).Value.ToString) <> "0" Then
            row.Cells(6).Style.BackColor = Color.LightGreen
        End If
    End Sub

    Private Sub DataGridView1_DoubleClick(ByVal sender As Object, ByVal e As System.EventArgs) Handles DataGridView1.DoubleClick
        '/////////////////////////////////////////////////////////////////////////////////////
        '//
        '// Выход с выбором события
        '// 
        '/////////////////////////////////////////////////////////////////////////////////////

        If DataGridView1.SelectedRows.Count <> 0 Then
            SupplierSelect()
        End If
    End Sub

    Private Sub Button4_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button4.Click
        '/////////////////////////////////////////////////////////////////////////////////////
        '//
        '// Выход с выбором события
        '// 
        '/////////////////////////////////////////////////////////////////////////////////////

        SupplierSelect()
    End Sub

    Private Sub SupplierSelect()
        '/////////////////////////////////////////////////////////////////////////////////////
        '//
        '// Выход с выбором события
        '// 
        '/////////////////////////////////////////////////////////////////////////////////////

        For i As Integer = 0 To MainForm.DataGridView1.Rows.Count - 1
            If Trim(MainForm.DataGridView1.Item(0, i).Value.ToString) = Trim(Me.DataGridView1.SelectedRows.Item(0).Cells(0).Value.ToString()) Then
                MainForm.DataGridView1.CurrentCell = MainForm.DataGridView1.Item(1, i)
                Me.Close()
                Exit Sub
            End If
        Next
        Me.Close()
    End Sub

    Private Sub SupplierSelectList_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles Me.KeyDown
        '/////////////////////////////////////////////////////////////////////////////////////
        '//
        '// выход без выбора по Escape
        '// 
        '/////////////////////////////////////////////////////////////////////////////////////

        If e.KeyData = Keys.Escape Then
            Me.Close()
        End If
    End Sub

    Private Sub SupplierSelectList_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        '/////////////////////////////////////////////////////////////////////////////////////
        '//
        '// При запуске загружаем список событий
        '// 
        '/////////////////////////////////////////////////////////////////////////////////////
        Dim MySQLStr As String
        Dim MyAdapter As SqlClient.SqlDataAdapter     'для списка поставщиков
        Dim MyDs As New DataSet                       '

        If MainForm.ComboBox2.Text = "Только активные поставщики" Then
            MySQLStr = "Exec spp_PurchaseWorkplace_SupplierListPrep 1, N'" & Trim(MainForm.ComboBox1.SelectedValue) & _
                "', N'" & Trim(MainForm.TextBox2.Text) & "', N'" & Trim(MainForm.TextBox3.Text) & "', N'" & _
                Trim(MainForm.ComboBox3.SelectedValue) & "' "
        Else
            MySQLStr = "Exec spp_PurchaseWorkplace_SupplierListPrep 0, N'" & Trim(MainForm.ComboBox1.SelectedValue) & _
                "', N'" & Trim(MainForm.TextBox2.Text) & "', N'" & Trim(MainForm.TextBox3.Text) & "', N'" & _
                Trim(MainForm.ComboBox3.SelectedValue) & "' "
        End If

        Try
            MyAdapter = New SqlClient.SqlDataAdapter(MySQLStr, Declarations.MyNETConnStr)
            MyAdapter.SelectCommand.CommandTimeout = 600
            MyAdapter.Fill(MyDs)
            DataGridView1.DataSource = MyDs.Tables(0)
        Catch ex As Exception
            MsgBox(ex.ToString)
        End Try

        '---заголовки
        DataGridView1.Columns(0).HeaderText = "Код постав щика"
        DataGridView1.Columns(0).Width = 100
        DataGridView1.Columns(1).HeaderText = "Поставщик"
        DataGridView1.Columns(1).Width = 224
        DataGridView1.Columns(2).HeaderText = "Адрес поставщика"
        DataGridView1.Columns(2).Width = 364
        DataGridView1.Columns(3).HeaderText = "Несгруппи рованных заказов"
        DataGridView1.Columns(3).Width = 110
        DataGridView1.Columns(4).HeaderText = "Неразме щенных заказов"
        DataGridView1.Columns(4).Width = 110
        DataGridView1.Columns(5).HeaderText = "Неподт вержденных заказов"
        DataGridView1.Columns(5).Width = 110
        DataGridView1.Columns(6).HeaderText = "Задол женных заказов"
        DataGridView1.Columns(6).Width = 110

        If DataGridView1.SelectedRows.Count = 0 Then
            Button4.Enabled = False
        Else
            Button4.Enabled = True
        End If
    End Sub
End Class