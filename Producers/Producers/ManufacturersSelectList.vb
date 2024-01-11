Public Class ManufacturersSelectList

    Private Sub Button5_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button5.Click
        '/////////////////////////////////////////////////////////////////////////////////////
        '//
        '// ����� ��� ������ ����������
        '// 
        '/////////////////////////////////////////////////////////////////////////////////////

        Me.Close()
    End Sub

    Private Sub DataGridView1_DoubleClick(ByVal sender As Object, ByVal e As System.EventArgs) Handles DataGridView1.DoubleClick
        '/////////////////////////////////////////////////////////////////////////////////////
        '//
        '// ����� � ������� �������
        '// 
        '/////////////////////////////////////////////////////////////////////////////////////

        If DataGridView1.SelectedRows.Count <> 0 Then
            ManufacturersSelect()
        End If
    End Sub

    Private Sub Button4_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button4.Click
        '/////////////////////////////////////////////////////////////////////////////////////
        '//
        '// ����� � ������� �������
        '// 
        '/////////////////////////////////////////////////////////////////////////////////////

        ManufacturersSelect()
    End Sub

    Private Sub ManufacturersSelect()
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

    Private Sub ManufacturersSelectList_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles Me.KeyDown
        '/////////////////////////////////////////////////////////////////////////////////////
        '//
        '// ����� ��� ������ �� Escape
        '// 
        '/////////////////////////////////////////////////////////////////////////////////////

        If e.KeyData = Keys.Escape Then
            Me.Close()
        End If
    End Sub

    Private Sub ManufacturersSelectList_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        '/////////////////////////////////////////////////////////////////////////////////////
        '//
        '// �������� ������ ��������������
        '// 
        '/////////////////////////////////////////////////////////////////////////////////////
        Dim MySQLStr As String
        Dim MyAdapter As SqlClient.SqlDataAdapter     '��� ������ �����������
        Dim MyDs As New DataSet                       '

        If Trim(MainForm.TextBox2.Text) <> "" And Trim(MainForm.TextBox3.Text) <> "" Then
            MySQLStr = "SELECT * "
            MySQLStr = MySQLStr & "FROM tbl_Manufacturers "
            MySQLStr = MySQLStr & "WHERE ((UPPER(CONVERT(nvarchar(255), ID)) LIKE UPPER(N'%" & Trim(MainForm.TextBox2.Text) & "%')) AND "
            MySQLStr = MySQLStr & "(UPPER(CONVERT(nvarchar(255), ID)) LIKE UPPER(N'%" & Trim(MainForm.TextBox3.Text) & "%'))) OR "
            MySQLStr = MySQLStr & "((UPPER(Name) LIKE UPPER(N'%" & Trim(MainForm.TextBox2.Text) & "%')) AND "
            MySQLStr = MySQLStr & "(UPPER(Name) LIKE UPPER(N'%" & Trim(MainForm.TextBox3.Text) & "%'))) OR "
            MySQLStr = MySQLStr & "((UPPER(Address) LIKE UPPER(N'%" & Trim(MainForm.TextBox2.Text) & "%')) AND "
            MySQLStr = MySQLStr & "(UPPER(Address) LIKE UPPER(N'%" & Trim(MainForm.TextBox3.Text) & "%'))) OR "
            MySQLStr = MySQLStr & "((UPPER(ContactInfo) LIKE UPPER(N'%" & Trim(MainForm.TextBox2.Text) & "%')) AND "
            MySQLStr = MySQLStr & "(UPPER(ContactInfo) LIKE UPPER(N'%" & Trim(MainForm.TextBox3.Text) & "%'))) "
            MySQLStr = MySQLStr & "ORDER BY ID "
        ElseIf Trim(MainForm.TextBox2.Text) <> "" Then
            MySQLStr = "SELECT * "
            MySQLStr = MySQLStr & "FROM tbl_Manufacturers "
            MySQLStr = MySQLStr & "WHERE (UPPER(CONVERT(nvarchar(255), ID)) LIKE UPPER(N'%" & Trim(MainForm.TextBox2.Text) & "%')) OR "
            MySQLStr = MySQLStr & "(UPPER(Name) LIKE UPPER(N'%" & Trim(MainForm.TextBox2.Text) & "%'))  OR "
            MySQLStr = MySQLStr & "(UPPER(Address) LIKE UPPER(N'%" & Trim(MainForm.TextBox2.Text) & "%')) OR "
            MySQLStr = MySQLStr & "(UPPER(ContactInfo) LIKE UPPER(N'%" & Trim(MainForm.TextBox2.Text) & "%')) "
            MySQLStr = MySQLStr & "ORDER BY ID "
        Else
            MySQLStr = "SELECT * "
            MySQLStr = MySQLStr & "FROM tbl_Manufacturers "
            MySQLStr = MySQLStr & "WHERE (UPPER(CONVERT(nvarchar(255), ID)) LIKE UPPER(N'%" & Trim(MainForm.TextBox3.Text) & "%')) OR "
            MySQLStr = MySQLStr & "(UPPER(Name) LIKE UPPER(N'%" & Trim(MainForm.TextBox3.Text) & "%'))  OR "
            MySQLStr = MySQLStr & "(UPPER(Address) LIKE UPPER(N'%" & Trim(MainForm.TextBox3.Text) & "%')) OR "
            MySQLStr = MySQLStr & "(UPPER(ContactInfo) LIKE UPPER(N'%" & Trim(MainForm.TextBox3.Text) & "%')) "
            MySQLStr = MySQLStr & "ORDER BY ID "
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
        DataGridView1.Columns(0).HeaderText = "��� ����� ��������"
        DataGridView1.Columns(0).Width = 100
        DataGridView1.Columns(1).HeaderText = "������������"
        DataGridView1.Columns(1).Width = 300
        DataGridView1.Columns(2).HeaderText = "����� �������������"
        DataGridView1.Columns(2).Width = 410
        DataGridView1.Columns(3).HeaderText = "���������� ����������"
        DataGridView1.Columns(3).Width = 410
    End Sub
End Class