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

        For i As Integer = 0 To MainForm.DataGridView11.Rows.Count - 1
            If Trim(MainForm.DataGridView11.Item(0, i).Value.ToString) = Trim(Me.DataGridView1.SelectedRows.Item(0).Cells(0).Value.ToString()) Then
                MainForm.DataGridView11.CurrentCell = MainForm.DataGridView11.Item(0, i)
                Me.Close()
                Exit Sub
            End If
        Next
        Me.Close()
    End Sub

    Private Sub CustomerSelectList_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        '/////////////////////////////////////////////////////////////////////////////////////
        '//
        '// ��� ������� ��������� ������ �����������
        '// 
        '/////////////////////////////////////////////////////////////////////////////////////
        Dim MySQLStr As String                        '������� ������
        Dim MyAdapter As SqlClient.SqlDataAdapter     '
        Dim MyDs As New DataSet                       '

        If Trim(MainForm.TextBox1.Text) = "" Then
            '----� ������ ���� ������� �� ������� - �������, ��� �� ������ �������
            MySQLStr = "SELECT Code, Name, Address, Discount, Case WHEN WorkOverWEB = 1 THEN '��' ELSE '' END as WorkOverWEB, BasePrice "
            MySQLStr = MySQLStr & "FROM tbl_WEB_Clients "
            MySQLStr = MySQLStr & "WHERE (Upper(Code) LIKE N'%" & UCase(MainForm.TextBox2.Text) & "%') OR "
            MySQLStr = MySQLStr & "(Upper(Name) LIKE N'%" & UCase(MainForm.TextBox2.Text) & "%') OR "
            MySQLStr = MySQLStr & "(Upper(Address) LIKE N'%" & UCase(MainForm.TextBox2.Text) & "%') "
            MySQLStr = MySQLStr & "ORDER BY Code "
        Else
            '----� ������ ���� ������� �������
            If Trim(MainForm.TextBox2.Text) = "" Then
                '----�� ������ ���� ������� �� �������
                MySQLStr = "SELECT Code, Name, Address, Discount, Case WHEN WorkOverWEB = 1 THEN '��' ELSE '' END as WorkOverWEB, BasePrice "
                MySQLStr = MySQLStr & "FROM tbl_WEB_Clients "
                MySQLStr = MySQLStr & "WHERE (Upper(Code) LIKE N'%" & UCase(MainForm.TextBox1.Text) & "%') OR "
                MySQLStr = MySQLStr & "(Upper(Name) LIKE N'%" & UCase(MainForm.TextBox1.Text) & "%') OR "
                MySQLStr = MySQLStr & "(Upper(Address) LIKE N'%" & UCase(MainForm.TextBox1.Text) & "%') "
                MySQLStr = MySQLStr & "ORDER BY Code "
            Else
                '----������� ������� � ��� ����
                MySQLStr = "SELECT Code, Name, Address, Discount, Case WHEN WorkOverWEB = 1 THEN '��' ELSE '' END as WorkOverWEB, BasePrice "
                MySQLStr = MySQLStr & "FROM tbl_WEB_Clients "
                MySQLStr = MySQLStr & "WHERE ((Upper(Code) LIKE N'%" & UCase(MainForm.TextBox1.Text) & "%') AND "
                MySQLStr = MySQLStr & "(Upper(Code) LIKE N'%" & UCase(MainForm.TextBox2.Text) & "%')) OR "
                MySQLStr = MySQLStr & "((Upper(Name) LIKE N'%" & UCase(MainForm.TextBox1.Text) & "%') AND "
                MySQLStr = MySQLStr & "(Upper(Name) LIKE N'%" & UCase(MainForm.TextBox2.Text) & "%')) OR "
                MySQLStr = MySQLStr & "((Upper(Address) LIKE N'%" & UCase(MainForm.TextBox1.Text) & "%') AND "
                MySQLStr = MySQLStr & "(Upper(Address) LIKE N'%" & UCase(MainForm.TextBox2.Text) & "%')) "
                MySQLStr = MySQLStr & "ORDER BY Code "
            End If
        End If

        InitMyConn(False)
        Try
            MyAdapter = New SqlClient.SqlDataAdapter(MySQLStr, Declarations.MyNETConnStr)
            MyAdapter.SelectCommand.CommandTimeout = 600
            MyAdapter.Fill(MyDs)
            DataGridView1.DataSource = MyDs.Tables(0)
        Catch ex As Exception
            MsgBox(ex.ToString)
        End Try

        DataGridView1.Columns(0).HeaderText = "��� �������"
        DataGridView1.Columns(0).Width = 100
        DataGridView1.Columns(1).HeaderText = "�������� �������"
        DataGridView1.Columns(1).Width = 250
        DataGridView1.Columns(2).HeaderText = "�����"
        DataGridView1.Columns(2).Width = 520
        DataGridView1.Columns(3).HeaderText = "����� ������"
        DataGridView1.Columns(3).Width = 60
        DataGridView1.Columns(4).HeaderText = "�������� ����� WEB"
        DataGridView1.Columns(4).Width = 60
        DataGridView1.Columns(5).HeaderText = "������� �����"
        DataGridView1.Columns(5).Width = 60

        DataGridView1.SelectionMode = DataGridViewSelectionMode.FullRowSelect
    End Sub
End Class