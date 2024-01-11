Public Class ItemSelectList2



    Private Sub ItemSelectList2_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Dim SQLStr As String
        Dim Adapter As SqlClient.SqlDataAdapter
        Dim DS As New DataSet

        SQLStr = "SELECT Code,SC01002 + SC01003 as ProductName ,WH,MinQty,MaxQty "
        SQLStr = SQLStr + " FROM [ScaDataDB].[dbo].[tbl_ConsStocks]  INNER JOIN ScaDataDB.dbo.SC010300 WITH(NOLOCK) ON Code = SC01001 "
        SQLStr = SQLStr + " WHERE WH = N'" & MainForm.ComboBox1.SelectedValue & "' "
        If Trim(MainForm.TextBox1.Text) = "" Then
            SQLStr = SQLStr & " AND UPPER(SC01001 + SC01002) LIKE N'%" & UCase(MainForm.TextBox4.Text) & "%'"
        Else
            If Trim(MainForm.TextBox4.Text) = "" Then
                SQLStr = SQLStr & " AND UPPER(SC01001 + SC01002) LIKE N'%" & UCase(MainForm.TextBox1.Text) & "%'"
            Else
                SQLStr = SQLStr & " AND ((UPPER(SC01001) LIKE N'%" & UCase(MainForm.TextBox1.Text) & "%' AND "
                SQLStr = SQLStr & " UPPER(SC01002) LIKE N'%" & UCase(MainForm.TextBox4.Text) & "%') OR "
                SQLStr = SQLStr & " (UPPER(SC01001) LIKE N'%" & UCase(MainForm.TextBox4.Text) & "%' AND "
                SQLStr = SQLStr & " UPPER(SC01002) LIKE N'%" & UCase(MainForm.TextBox1.Text) & "%'))"
            End If
        End If
        SQLStr = SQLStr & " ORDER BY SC01001"
        InitMyConn(False)

        Try
            Adapter = New SqlClient.SqlDataAdapter(SQLStr, Declarations.NETConnStr)
            Adapter.SelectCommand.CommandTimeout = 600
            Adapter.Fill(DS)
            DataGridView1.DataSource = DS.Tables(0)

            DataGridView1.Columns(0).HeaderText = "ID"
            DataGridView1.Columns(0).Width = 80
            DataGridView1.Columns(1).HeaderText = "Запас"
            DataGridView1.Columns(1).Width = 300
            DataGridView1.Columns(2).HeaderText = "Склад"
            DataGridView1.Columns(2).Width = 80
            DataGridView1.Columns(3).HeaderText = "Мин. кол-во"
            DataGridView1.Columns(3).Width = 80
            DataGridView1.Columns(4).HeaderText = "Макс. кол-во"
            DataGridView1.Columns(4).Width = 80

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

    Private Sub Button5_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button5.Click
        Me.Close()
    End Sub

    Private Sub Button4_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button4.Click
        For i As Integer = 0 To MainForm.DataGridView2.Rows.Count - 1
            If Trim(MainForm.DataGridView2.Item(0, i).Value.ToString) = Trim(Me.DataGridView1.SelectedRows.Item(0).Cells(0).Value.ToString()) Then
                MainForm.DataGridView2.CurrentCell = MainForm.DataGridView2.Item(0, i)
                Me.Close()
                Exit Sub
            End If
        Next
        Me.Close()
    End Sub
End Class