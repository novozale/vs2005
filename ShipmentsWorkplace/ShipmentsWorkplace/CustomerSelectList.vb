Public Class CustomerSelectList

    Private Sub Button5_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button5.Click
        '/////////////////////////////////////////////////////////////////////////////////////
        '//
        '// ����� ��� ������ ����������
        '// 
        '/////////////////////////////////////////////////////////////////////////////////////

        Me.Close()
    End Sub

    Private Sub Button4_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button4.Click
        '/////////////////////////////////////////////////////////////////////////////////////
        '//
        '// ����� � ������� ����������
        '// 
        '/////////////////////////////////////////////////////////////////////////////////////

        CustomerSelect()
    End Sub

    Private Sub DataGridView1_CellFormatting(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellFormattingEventArgs) Handles DataGridView1.CellFormatting
        '////////////////////////////////////////////////////////////////////////////////
        '//
        '// ��������� ���������� �� �����������
        '//
        '////////////////////////////////////////////////////////////////////////////////
        Dim row As DataGridViewRow

        row = DataGridView1.Rows(e.RowIndex)
        row.Cells(0).Style.Font = New Font(DataGridView1.DefaultCellStyle.Font, FontStyle.Bold)
        row.Cells(1).Style.Font = New Font(DataGridView1.DefaultCellStyle.Font, FontStyle.Bold)
        row.Cells(3).Style.Format = "n0"
        row.Cells(3).Style.Font = New Font(DataGridView1.DefaultCellStyle.Font, FontStyle.Bold)
        row.Cells(4).Style.Format = "n0"
        row.Cells(4).Style.Font = New Font(DataGridView1.DefaultCellStyle.Font, FontStyle.Bold)
        row.Cells(5).Style.Format = "n0"
        row.Cells(5).Style.Font = New Font(DataGridView1.DefaultCellStyle.Font, FontStyle.Bold)
        row.Cells(6).Style.Format = "n0"
        row.Cells(6).Style.Font = New Font(DataGridView1.DefaultCellStyle.Font, FontStyle.Bold)
        row.Cells(7).Style.Format = "n0"
        row.Cells(7).Style.Font = New Font(DataGridView1.DefaultCellStyle.Font, FontStyle.Bold)
        If Trim(row.Cells(3).Value.ToString) <> "" Then
            row.Cells(3).Style.BackColor = Color.LightGreen
        Else
            row.Cells(3).Style.BackColor = Color.Empty
        End If
        If Trim(row.Cells(4).Value.ToString) <> "" Then
            row.Cells(4).Style.BackColor = Color.Red
        Else
            row.Cells(4).Style.BackColor = Color.Empty
        End If
        If Trim(row.Cells(5).Value.ToString) <> "" Then
            row.Cells(5).Style.BackColor = Color.Orange
        Else
            row.Cells(5).Style.BackColor = Color.Empty
        End If
        If Trim(row.Cells(6).Value.ToString) <> "" Then
            row.Cells(6).Style.BackColor = Color.Yellow
        Else
            row.Cells(6).Style.BackColor = Color.Empty
        End If
        If Trim(row.Cells(7).Value.ToString) <> "" Then
            row.Cells(7).Style.BackColor = Color.Yellow
        Else
            row.Cells(7).Style.BackColor = Color.Empty
        End If
    End Sub

    Private Sub DataGridView1_DoubleClick(ByVal sender As Object, ByVal e As System.EventArgs) Handles DataGridView1.DoubleClick
        '/////////////////////////////////////////////////////////////////////////////////////
        '//
        '// ����� � ������� ����������
        '// 
        '/////////////////////////////////////////////////////////////////////////////////////

        If DataGridView1.SelectedRows.Count <> 0 Then
            CustomerSelect()
        End If
    End Sub

    Private Sub CustomerSelect()
        '/////////////////////////////////////////////////////////////////////////////////////
        '//
        '// ����� � ������� �������
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


    Private Sub CustomerSelectList_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles Me.KeyDown
        '/////////////////////////////////////////////////////////////////////////////////////
        '//
        '// ����� ��� ������ �� Escape
        '// 
        '/////////////////////////////////////////////////////////////////////////////////////

        If e.KeyData = Keys.Escape Then
            Me.Close()
        End If
    End Sub

    Private Sub CustomerSelectList_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        '/////////////////////////////////////////////////////////////////////////////////////
        '//
        '// �������� ������ ����������� (� ������������ � �����������)
        '// 
        '/////////////////////////////////////////////////////////////////////////////////////
        Dim MySQLStr As String
        Dim MyAdapter As SqlClient.SqlDataAdapter     '��� ������ �����������
        Dim MyDs As New DataSet                       '

        If MainForm.ComboBox2.Text = "������ �������� ����������" Then
            If MainForm.ComboBox3.Text = "�������������" Then
                MySQLStr = "Exec spp_Shipments_SalesmanWP_CommonInfo 1, N'" & Trim(MainForm.ComboBox1.SelectedValue) & "', N'" & Declarations.SalesmanCode & "', 0, N'" & Trim(MainForm.TextBox2.Text) & "', N'" & Trim(MainForm.TextBox3.Text) & "' "
            Else
                MySQLStr = "Exec spp_Shipments_SalesmanWP_CommonInfo 1, N'" & Trim(MainForm.ComboBox1.SelectedValue) & "', N'" & Declarations.SalesmanCode & "', 1, N'" & Trim(MainForm.TextBox2.Text) & "', N'" & Trim(MainForm.TextBox3.Text) & "' "
            End If
        Else
            If MainForm.ComboBox3.Text = "�������������" Then
                MySQLStr = "Exec spp_Shipments_SalesmanWP_CommonInfo 0, N'" & Trim(MainForm.ComboBox1.SelectedValue) & "', N'" & Declarations.SalesmanCode & "', 0, N'" & Trim(MainForm.TextBox2.Text) & "', N'" & Trim(MainForm.TextBox3.Text) & "' "
            Else
                MySQLStr = "Exec spp_Shipments_SalesmanWP_CommonInfo 0, N'" & Trim(MainForm.ComboBox1.SelectedValue) & "', N'" & Declarations.SalesmanCode & "', 1, N'" & Trim(MainForm.TextBox2.Text) & "', N'" & Trim(MainForm.TextBox3.Text) & "' "
            End If
        End If

        Try
            MyAdapter = New SqlClient.SqlDataAdapter(MySQLStr, Declarations.MyNETConnStr)
            MyAdapter.SelectCommand.CommandTimeout = 600
            MyAdapter.Fill(MyDs)
            DataGridView1.DataSource = MyDs.Tables(0)
        Catch ex As Exception
            MsgBox(ex.ToString)
        End Try

        '---���������
        DataGridView1.Columns(0).HeaderText = "��� ���� ������"
        DataGridView1.Columns(0).Width = 90
        DataGridView1.Columns(1).HeaderText = "����������"
        DataGridView1.Columns(1).Width = 210
        DataGridView1.Columns(2).HeaderText = "����� ����������"
        DataGridView1.Columns(2).Width = 361
        DataGridView1.Columns(3).HeaderText = "������� � ��������� � ���. 7 ����"
        DataGridView1.Columns(3).Width = 110
        DataGridView1.Columns(4).HeaderText = "������� � ������������ ���������"
        DataGridView1.Columns(4).Width = 110
        DataGridView1.Columns(5).HeaderText = "�������, � ������� ���� ������� ������ ���� ��������"
        DataGridView1.Columns(5).Width = 110
        DataGridView1.Columns(6).HeaderText = "�������, �� ���������� � ������� 7 ����"
        DataGridView1.Columns(6).Width = 110
        DataGridView1.Columns(7).HeaderText = "�������, �� ����������� � ������� 2 ����"
        DataGridView1.Columns(7).Width = 110

        If DataGridView1.SelectedRows.Count = 0 Then
            Button4.Enabled = False
        Else
            Button4.Enabled = True
        End If
    End Sub
End Class