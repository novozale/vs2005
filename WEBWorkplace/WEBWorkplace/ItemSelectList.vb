Public Class ItemSelectList
    Public MyBS As New BindingSource()

    Private Sub Button2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button2.Click
        '/////////////////////////////////////////////////////////////////////////////////////
        '//
        '// ����� ��� ������ ������
        '// 
        '/////////////////////////////////////////////////////////////////////////////////////

        Me.Close()
    End Sub

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        '/////////////////////////////////////////////////////////////////////////////////////
        '//
        '// ����� � ������� ������
        '// 
        '/////////////////////////////////////////////////////////////////////////////////////

        For i As Integer = 0 To MyItemList.DataGridView1.Rows.Count - 1
            If Trim(MyItemList.DataGridView1.Item(0, i).Value.ToString) = Trim(Me.DataGridView1.SelectedRows.Item(0).Cells(0).Value.ToString()) Then
                MyItemList.DataGridView1.CurrentCell = MyItemList.DataGridView1.Item(0, i)
                Me.Close()
                Exit Sub
            End If
        Next

        Me.Close()
    End Sub

    Private Sub ItemSelectList_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        '/////////////////////////////////////////////////////////////////////////////////////
        '//
        '// �������� ���������� � ����
        '// 
        '/////////////////////////////////////////////////////////////////////////////////////
        Dim MySQLStr As String                        '������� ������
        Dim MyAdapter As SqlClient.SqlDataAdapter     '
        Dim MyDs As New DataSet                       '

        If Trim(MyItemList.TextBox1.Text) = "" Then
            '----� ������ ���� ������� �� ������� - �������, ��� �� ������ �������
            MySQLStr = "SELECT tbl_WEB_Items.Code, tbl_WEB_Items.Name, tbl_WEB_Manufacturers.Name AS ManufacturerName, tbl_WEB_Items.ManufacturerItemCode, "
            MySQLStr = MySQLStr & "tbl_WEB_Items.Country,  CASE WHEN Ltrim(Rtrim(tbl_WEB_Items.SubGroupCode)) = '' THEN '' ELSE '��' END AS HasSubgroup "
            MySQLStr = MySQLStr & "FROM tbl_WEB_Items INNER JOIN "
            MySQLStr = MySQLStr & "tbl_WEB_Manufacturers ON tbl_WEB_Items.ManufacturerCode = tbl_WEB_Manufacturers.ID "
            MySQLStr = MySQLStr & "WHERE (UPPER(tbl_WEB_Items.Code) LIKE N'%" & UCase(MyItemList.TextBox2.Text) & "%') OR "
            MySQLStr = MySQLStr & "(UPPER(tbl_WEB_Items.Name) LIKE N'%" & UCase(MyItemList.TextBox2.Text) & "%') OR "
            MySQLStr = MySQLStr & "(UPPER(tbl_WEB_Items.ManufacturerItemCode) LIKE N'%" & UCase(MyItemList.TextBox2.Text) & "%')  "
            MySQLStr = MySQLStr & "ORDER BY tbl_WEB_Items.Code "
        Else
            '----� ������ ���� ������� �������
            If Trim(MyItemList.TextBox2.Text) = "" Then
                '----�� ������ ���� ������� �� �������
                MySQLStr = "SELECT tbl_WEB_Items.Code, tbl_WEB_Items.Name, tbl_WEB_Manufacturers.Name AS ManufacturerName, tbl_WEB_Items.ManufacturerItemCode, "
                MySQLStr = MySQLStr & "tbl_WEB_Items.Country,  CASE WHEN Ltrim(Rtrim(tbl_WEB_Items.SubGroupCode)) = '' THEN '' ELSE '��' END AS HasSubgroup "
                MySQLStr = MySQLStr & "FROM tbl_WEB_Items INNER JOIN "
                MySQLStr = MySQLStr & "tbl_WEB_Manufacturers ON tbl_WEB_Items.ManufacturerCode = tbl_WEB_Manufacturers.ID "
                MySQLStr = MySQLStr & "WHERE (UPPER(tbl_WEB_Items.Code) LIKE N'%" & UCase(MyItemList.TextBox1.Text) & "%') OR "
                MySQLStr = MySQLStr & "(UPPER(tbl_WEB_Items.Name) LIKE N'%" & UCase(MyItemList.TextBox1.Text) & "%') OR "
                MySQLStr = MySQLStr & "(UPPER(tbl_WEB_Items.ManufacturerItemCode) LIKE N'%" & UCase(MyItemList.TextBox1.Text) & "%')  "
                MySQLStr = MySQLStr & "ORDER BY tbl_WEB_Items.Code "
            Else
                '----������� ������� � ��� ����
                MySQLStr = "SELECT tbl_WEB_Items.Code, tbl_WEB_Items.Name, tbl_WEB_Manufacturers.Name AS ManufacturerName, tbl_WEB_Items.ManufacturerItemCode, "
                MySQLStr = MySQLStr & "tbl_WEB_Items.Country,  CASE WHEN Ltrim(Rtrim(tbl_WEB_Items.SubGroupCode)) = '' THEN '' ELSE '��' END AS HasSubgroup "
                MySQLStr = MySQLStr & "FROM tbl_WEB_Items INNER JOIN "
                MySQLStr = MySQLStr & "tbl_WEB_Manufacturers ON tbl_WEB_Items.ManufacturerCode = tbl_WEB_Manufacturers.ID "
                MySQLStr = MySQLStr & "WHERE ((UPPER(tbl_WEB_Items.Code) LIKE N'%" & UCase(MyItemList.TextBox1.Text) & "%') AND "
                MySQLStr = MySQLStr & "(UPPER(tbl_WEB_Items.Code) LIKE N'%" & UCase(MyItemList.TextBox2.Text) & "%')) OR "
                MySQLStr = MySQLStr & "((UPPER(tbl_WEB_Items.Name) LIKE N'%" & UCase(MyItemList.TextBox1.Text) & "%') AND "
                MySQLStr = MySQLStr & "(UPPER(tbl_WEB_Items.Name) LIKE N'%" & UCase(MyItemList.TextBox2.Text) & "%')) OR "
                MySQLStr = MySQLStr & "((UPPER(tbl_WEB_Items.ManufacturerItemCode) LIKE N'%" & UCase(MyItemList.TextBox1.Text) & "%') AND "
                MySQLStr = MySQLStr & "(UPPER(tbl_WEB_Items.ManufacturerItemCode) LIKE N'%" & UCase(MyItemList.TextBox2.Text) & "%'))  "
                MySQLStr = MySQLStr & "ORDER BY tbl_WEB_Items.Code "
            End If
        End If

        InitMyConn(False)
        Try
            MyAdapter = New SqlClient.SqlDataAdapter(MySQLStr, Declarations.MyNETConnStr)
            MyAdapter.SelectCommand.CommandTimeout = 600
            MyAdapter.Fill(MyDs)
            DataGridView1.DataSource = MyDs.Tables(0)
            '---������
            MyBS.DataSource = MyDs
            MyBS.DataMember = MyDs.Tables(0).TableName
            DataGridView1.DataSource = MyBS
            '---����� �������

        Catch ex As Exception
            MsgBox(ex.ToString)
        End Try

        DataGridView1.Columns(0).HeaderText = "��� ������ � Scala"
        DataGridView1.Columns(0).Width = 100
        DataGridView1.Columns(1).HeaderText = "��� ������ � Scala"
        DataGridView1.Columns(1).Width = 400
        DataGridView1.Columns(2).HeaderText = "�������������"
        DataGridView1.Columns(2).Width = 170
        DataGridView1.Columns(3).HeaderText = "��� ������ �������������"
        DataGridView1.Columns(3).Width = 170
        DataGridView1.Columns(4).HeaderText = "������"
        DataGridView1.Columns(4).Width = 100
        DataGridView1.Columns(5).HeaderText = "���� ���������"
        DataGridView1.Columns(5).Width = 80

        DataGridView1.SelectionMode = DataGridViewSelectionMode.FullRowSelect

        CheckButtonState()
    End Sub

    Private Sub CheckButtonState()
        '/////////////////////////////////////////////////////////////////////////////////////
        '//
        '// �������� � ����������� ��������� ������
        '// 
        '/////////////////////////////////////////////////////////////////////////////////////

        If DataGridView1.SelectedRows.Count = 0 Then
            Button1.Enabled = False
        Else
            If DataGridView1.SelectedRows.Item(0).Cells(5).Value = "" Then
                Button1.Enabled = False
            Else
                Button1.Enabled = True
            End If
        End If
    End Sub

    Private Sub DataGridView1_CellFormatting(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellFormattingEventArgs) Handles DataGridView1.CellFormatting
        '////////////////////////////////////////////////////////////////////////////////
        '//
        '// ��������� ����� ������� � ����������� �� �������
        '//
        '////////////////////////////////////////////////////////////////////////////////
        Dim row As DataGridViewRow

        row = DataGridView1.Rows(e.RowIndex)
        If row.Cells(5).Value = "" Then
            row.DefaultCellStyle.BackColor = Color.LightGray
        Else
            row.DefaultCellStyle.BackColor = Color.White
        End If
    End Sub

    Private Sub DataGridView1_CellMouseClick(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellMouseEventArgs) Handles DataGridView1.CellMouseClick
        '////////////////////////////////////////////////////////////////////////////////
        '//
        '// ����������� ���� ����������� �������
        '//
        '////////////////////////////////////////////////////////////////////////////////

        If e.Button = Windows.Forms.MouseButtons.Right Then
            Declarations.MyFilterColumn = e.ColumnIndex
            ContextMenuStrip1.Show(MousePosition.X, MousePosition.Y)
        End If
    End Sub

    Private Sub Button74_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button74.Click
        '////////////////////////////////////////////////////////////////////////////////
        '//
        '// ������ ������� 
        '//
        '////////////////////////////////////////////////////////////////////////////////

        MyBS.Filter = ""
        Label21.BackColor = Color.White
        For i As Integer = 0 To DataGridView1.Columns.Count - 1
            DataGridView1.Columns(i).HeaderCell.Style.ForeColor = Color.Black
        Next
    End Sub

    Private Sub ToolStripMenuItem1_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles ToolStripMenuItem1.Click
        '////////////////////////////////////////////////////////////////////////////////
        '//
        '// ����� ������������ ���� ��������� ������� ����
        '//
        '////////////////////////////////////////////////////////////////////////////////

        If Declarations.MyFilterColumn = 0 Then
            MyBS.Filter = "Code = '" & Trim(Me.DataGridView1.SelectedRows.Item(0).Cells(0).Value.ToString()) & "'"
        ElseIf Declarations.MyFilterColumn = 1 Then
            MyBS.Filter = "Name = '" & Trim(Me.DataGridView1.SelectedRows.Item(0).Cells(1).Value.ToString()) & "'"
        ElseIf Declarations.MyFilterColumn = 2 Then
            MyBS.Filter = "ManufacturerName = '" & Trim(Me.DataGridView1.SelectedRows.Item(0).Cells(2).Value.ToString()) & "'"
        ElseIf Declarations.MyFilterColumn = 3 Then
            MyBS.Filter = "ManufacturerItemCode = '" & Trim(Me.DataGridView1.SelectedRows.Item(0).Cells(3).Value.ToString()) & "'"
        ElseIf Declarations.MyFilterColumn = 4 Then
            MyBS.Filter = "Country = '" & Trim(Me.DataGridView1.SelectedRows.Item(0).Cells(4).Value.ToString()) & "'"
        ElseIf Declarations.MyFilterColumn = 5 Then
            MyBS.Filter = "HasSubgroup = '" & Trim(Me.DataGridView1.SelectedRows.Item(0).Cells(5).Value.ToString()) & "'"
        End If

        For i As Integer = 0 To DataGridView1.Columns.Count - 1
            If i = Declarations.MyFilterColumn Then
                DataGridView1.Columns(i).HeaderCell.Style.ForeColor = Color.Green
            Else
                DataGridView1.Columns(i).HeaderCell.Style.ForeColor = Color.Black
            End If
        Next
        Label21.BackColor = Color.Green
    End Sub

    Private Sub ToolStripMenuItem2_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles ToolStripMenuItem2.Click
        '////////////////////////////////////////////////////////////////////////////////
        '//
        '// ����� ������������ ���� ������ ������� 
        '//
        '////////////////////////////////////////////////////////////////////////////////

        MyBS.Filter = ""
        Label21.BackColor = Color.White
        For i As Integer = 0 To DataGridView1.Columns.Count - 1
            DataGridView1.Columns(i).HeaderCell.Style.ForeColor = Color.Black
        Next
    End Sub
End Class