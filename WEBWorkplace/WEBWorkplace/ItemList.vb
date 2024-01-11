Public Class ItemList
    Public StartParam As String
    Public MyBS As New BindingSource()

    Private Sub Button2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button2.Click
        '/////////////////////////////////////////////////////////////////////////////////////
        '//
        '// ����� ��� ������ ������
        '// 
        '/////////////////////////////////////////////////////////////////////////////////////

        Me.Close()
    End Sub

    Private Sub ItemList_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        '/////////////////////////////////////////////////////////////////////////////////////
        '//
        '// �������� ���������� � ����
        '// 
        '/////////////////////////////////////////////////////////////////////////////////////
        Dim MySQLStr As String                        '������� ������
        Dim MyAdapter As SqlClient.SqlDataAdapter     '
        Dim MyDs As New DataSet                       '

        MySQLStr = "SELECT tbl_WEB_Items.Code, tbl_WEB_Items.Name, tbl_WEB_Manufacturers.Name AS ManufacturerName, tbl_WEB_Items.ManufacturerItemCode, "
        MySQLStr = MySQLStr & "tbl_WEB_Items.Country,  CASE WHEN Ltrim(Rtrim(tbl_WEB_Items.SubGroupCode)) = '' THEN '' ELSE '��' END AS HasSubgroup "
        MySQLStr = MySQLStr & "FROM tbl_WEB_Items INNER JOIN "
        MySQLStr = MySQLStr & "tbl_WEB_Manufacturers ON tbl_WEB_Items.ManufacturerCode = tbl_WEB_Manufacturers.ID "
        MySQLStr = MySQLStr & "ORDER BY tbl_WEB_Items.Code "

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

    Private Sub DataGridView1_SelectionChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles DataGridView1.SelectionChanged
        '/////////////////////////////////////////////////////////////////////////////////////
        '//
        '// ������� ������
        '// 
        '/////////////////////////////////////////////////////////////////////////////////////

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
            Button54.Enabled = False
        Else
            If DataGridView1.SelectedRows.Item(0).Cells(5).Value = "" Then
                Button1.Enabled = False
            Else
                Button1.Enabled = True
            End If
            Button54.Enabled = True
        End If
    End Sub

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        '/////////////////////////////////////////////////////////////////////////////////////
        '//
        '// ����� ������ �� ������
        '// 
        '/////////////////////////////////////////////////////////////////////////////////////

        If StartParam = "DiscountItem" Then
            MyDiscountItem.TextBox1.Text = DataGridView1.SelectedRows.Item(0).Cells(0).Value
            MyDiscountItem.Label3.Text = DataGridView1.SelectedRows.Item(0).Cells(1).Value
        ElseIf StartParam = "AgreedRange" Then
            MyAgreedRange.TextBox1.Text = DataGridView1.SelectedRows.Item(0).Cells(0).Value
            MyAgreedRange.Label3.Text = DataGridView1.SelectedRows.Item(0).Cells(1).Value
        End If

        Me.Close()
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

    Private Sub Button54_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button54.Click
        '////////////////////////////////////////////////////////////////////////////////
        '//
        '// ����� ���� ���������� �� �������� ����������� � ��������� ����
        '//
        '////////////////////////////////////////////////////////////////////////////////

        If Trim(TextBox1.Text) = "" And Trim(TextBox2.Text) = "" Then
            MsgBox("���������� ������ �������� ������", MsgBoxStyle.OkOnly, "��������!")
            TextBox1.Select()
        Else
            MyItemSelectList = New ItemSelectList
            MyItemSelectList.ShowDialog()
        End If
    End Sub
End Class