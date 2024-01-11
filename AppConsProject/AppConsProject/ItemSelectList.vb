Public Class ItemSelectList

    Private Sub Button5_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button5.Click

        Me.Close()

    End Sub

    Private Sub ItemSelectList_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        Dim SQLStr As String
        Dim Adapter As SqlClient.SqlDataAdapter
        Dim DS As New DataSet

        SQLStr = "SELECT SC01001, SC01002 + SC01003 AS SC01002  FROM ScaDataDB.dbo.SC010300 WITH(NOLOCK)"
        If Trim(MainForm.TextBox2.Text) = "" Then
            SQLStr = SQLStr & " WHERE UPPER(SC01001 + SC01002) LIKE N'%" & UCase(MainForm.TextBox3.Text) & "%'"
        Else
            If Trim(MainForm.TextBox3.Text) = "" Then
                SQLStr = SQLStr & " WHERE UPPER(SC01001 + SC01002) LIKE N'%" & UCase(MainForm.TextBox2.Text) & "%'"
            Else
                SQLStr = SQLStr & " WHERE (UPPER(SC01001) LIKE N'%" & UCase(MainForm.TextBox2.Text) & "%' AND "
                SQLStr = SQLStr & " UPPER(SC01002) LIKE N'%" & UCase(MainForm.TextBox3.Text) & "%') OR "
                SQLStr = SQLStr & " (UPPER(SC01001) LIKE N'%" & UCase(MainForm.TextBox3.Text) & "%' AND "
                SQLStr = SQLStr & " UPPER(SC01002) LIKE N'%" & UCase(MainForm.TextBox2.Text) & "%')"
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
            DataGridView1.Columns(1).Width = 600

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
    Private Sub Button4_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button4.Click

        For i As Integer = 0 To MainForm.DataGridView1.Rows.Count - 1
            If Trim(MainForm.DataGridView1.Item(0, i).Value.ToString) = Trim(Me.DataGridView1.SelectedRows.Item(0).Cells(0).Value.ToString()) Then
                MainForm.DataGridView1.CurrentCell = MainForm.DataGridView1.Item(0, i)
                Me.Close()
                Exit Sub
            End If
        Next
        Me.Close()
    End Sub
End Class