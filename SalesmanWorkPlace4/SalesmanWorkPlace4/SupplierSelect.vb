Public Class SupplierSelect

    Public MySrcWin As String                         '����, �� �������� �������

    Private Sub Button5_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button5.Click
        '////////////////////////////////////////////////////////////////////////////////
        '//
        '// ����� �� ���� ��� ������ ����������
        '//
        '////////////////////////////////////////////////////////////////////////////////

        Me.Close()
    End Sub

    Private Sub CustomerSelect_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        '/////////////////////////////////////////////////////////////////////////////////////
        '//
        '// ��� ������� ��������� ������ �����������
        '// 
        '/////////////////////////////////////////////////////////////////////////////////////

        Dim MySQLStr As String                        '������� ������
        Dim MyAdapter As SqlClient.SqlDataAdapter     '
        Dim MyDs As New DataSet                       '

        MySQLStr = "SELECT PL01001, PL01002, PL01003 + PL01004 + PL01005 AS PL01003, PL01025 "
        MySQLStr = MySQLStr & "FROM PL01" & Declarations.CompanyID & "00 WITH(NOLOCK) "
        MySQLStr = MySQLStr & "ORDER BY PL01002 "

        Try
            MyAdapter = New SqlClient.SqlDataAdapter(MySQLStr, Declarations.MyNETConnStr)
            MyAdapter.SelectCommand.CommandTimeout = 600
            MyAdapter.Fill(MyDs)
            DataGridView1.DataSource = MyDs.Tables(0)
        Catch ex As Exception
            MsgBox(ex.ToString)
        End Try

        DataGridView1.Columns(0).HeaderText = "��� ����������"
        DataGridView1.Columns(0).Width = 90
        DataGridView1.Columns(1).HeaderText = "��� ����������"
        DataGridView1.Columns(1).Width = 140
        DataGridView1.Columns(2).HeaderText = "����� ����������"
        DataGridView1.Columns(3).HeaderText = "��� ����������"
        DataGridView1.Columns(3).Width = 130

        CheckButtons()
    End Sub

    Private Sub CheckButtons()
        '/////////////////////////////////////////////////////////////////////////////////////
        '//
        '// ����������� ��������� ������
        '// 
        '/////////////////////////////////////////////////////////////////////////////////////

        If DataGridView1.SelectedRows.Count = 0 Then
            Button4.Enabled = False
        Else
            Button4.Enabled = True
        End If
    End Sub

    Private Sub DataGridView1_ColumnHeaderMouseClick(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellMouseEventArgs) Handles DataGridView1.ColumnHeaderMouseClick
        '/////////////////////////////////////////////////////////////////////////////////////
        '//
        '// ������ �� ��������� ������� 
        '// 
        '/////////////////////////////////////////////////////////////////////////////////////

        Button6.Text = "���������� ���"
    End Sub

    Private Sub DataGridView1_DoubleClick(ByVal sender As Object, ByVal e As System.EventArgs) Handles DataGridView1.DoubleClick
        '////////////////////////////////////////////////////////////////////////////////
        '//
        '// ����� �� ���� � ������� ����������
        '//
        '////////////////////////////////////////////////////////////////////////////////

        SupplierSelect()
    End Sub

    Private Sub DataGridView1_SelectionChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles DataGridView1.SelectionChanged
        '/////////////////////////////////////////////////////////////////////////////////////
        '//
        '// �������� ��������� ������ ��� ��������� ���������  
        '// 
        '/////////////////////////////////////////////////////////////////////////////////////

        CheckButtons()
    End Sub

    Private Sub Button4_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button4.Click
        '////////////////////////////////////////////////////////////////////////////////
        '//
        '// ����� �� ���� � ������� ����������
        '//
        '////////////////////////////////////////////////////////////////////////////////

        SupplierSelect()
    End Sub

    Private Sub SupplierSelect()
        '////////////////////////////////////////////////////////////////////////////////
        '//
        '// ����� �� ���� � ������� ����������
        '//
        '////////////////////////////////////////////////////////////////////////////////
        Dim MySQLStr As String
        Dim MyRez As Double

        If MySrcWin = "OrderLines" Then
            MyOrderLines.TextBox1.Text = Trim(Me.DataGridView1.SelectedRows.Item(0).Cells(0).Value.ToString())
        ElseIf MySrcWin = "ItemSelect" Then
            MyItemSelect.TextBox1.Text = Trim(Me.DataGridView1.SelectedRows.Item(0).Cells(0).Value.ToString())
        ElseIf MySrcWin = "AddToOrder" Then
            MyAddToOrder.TextBox13.Text = Trim(Me.DataGridView1.SelectedRows.Item(0).Cells(0).Value.ToString())
        ElseIf MySrcWin = "EditInOrder" Then
            MyEditInOrder.TextBox15.Text = Trim(Me.DataGridView1.SelectedRows.Item(0).Cells(0).Value.ToString())
        End If

        MySQLStr = "SELECT COUNT(*) AS CC "
        MySQLStr = MySQLStr & "FROM PL010300 WITH (NOLOCK) "
        MySQLStr = MySQLStr & "WHERE (PL01001 = N'" & Trim(Me.DataGridView1.SelectedRows.Item(0).Cells(0).Value.ToString()) & "')"
        InitMyConn(False)
        InitMyRec(False, MySQLStr)
        MyRez = Declarations.MyRec.Fields("CC").Value
        trycloseMyRec()
        If MyRez = 1 Then
            MySQLStr = "SELECT PL01002, PL01003 + ' ' + PL01004 + ' ' + PL01005 AS PL01003 "
            MySQLStr = MySQLStr & "FROM PL010300 WITH (NOLOCK) "
            MySQLStr = MySQLStr & "WHERE (PL01001 = N'" & Trim(Me.DataGridView1.SelectedRows.Item(0).Cells(0).Value.ToString()) & "') "
            InitMyConn(False)
            InitMyRec(False, MySQLStr)
            If MySrcWin = "OrderLines" Then
                MyOrderLines.Label3.Text = Declarations.MyRec.Fields("PL01002").Value & " " & Declarations.MyRec.Fields("PL01003").Value
            ElseIf MySrcWin = "ItemSelect" Then
                MyItemSelect.Label3.Text = Declarations.MyRec.Fields("PL01002").Value & " " & Declarations.MyRec.Fields("PL01003").Value
            ElseIf MySrcWin = "AddToOrder" Then
                MyAddToOrder.TextBox14.Text = Declarations.MyRec.Fields("PL01002").Value
                MyAddToOrder.TextBox14.Enabled = False
                MyAddToOrder.TextBox14.BackColor = Color.FromName("ButtonFace")
            ElseIf MySrcWin = "EditInOrder" Then
                MyEditInOrder.TextBox16.Text = Declarations.MyRec.Fields("PL01002").Value
                MyEditInOrder.TextBox16.Enabled = False
                MyEditInOrder.TextBox16.BackColor = Color.FromName("ButtonFace")
            End If
            trycloseMyRec()
        Else
            If MySrcWin = "OrderLines" Then
                MyOrderLines.Label3.Text = ""
            ElseIf MySrcWin = "ItemSelect" Then
                MyItemSelect.Label3.Text = ""
            ElseIf MySrcWin = "AddToOrder" Then
                MyAddToOrder.TextBox14.Text = ""
                MyAddToOrder.TextBox14.Enabled = True
                MyAddToOrder.TextBox14.BackColor = Color.FromName("Window")
            ElseIf MySrcWin = "AddToOrder" Then
                MyEditInOrder.TextBox16.Text = ""
                MyEditInOrder.TextBox16.Enabled = True
                MyEditInOrder.TextBox16.BackColor = Color.FromName("Window")
            End If
        End If
        If MySrcWin = "OrderLines" Then
            MyOrderLines.RefreshProductList()
        ElseIf MySrcWin = "ItemSelect" Then
            MyItemSelect.RefreshProductList()
        End If
        Me.Close()

    End Sub

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        '////////////////////////////////////////////////////////////////////////////////
        '//
        '// ����� ������� ����������� �� �������� ����������
        '//
        '////////////////////////////////////////////////////////////////////////////////

        If Trim(TextBox1.Text) = "" And Trim(TextBox2.Text) = "" Then
        Else
            For i As Integer = 0 To DataGridView1.Rows.Count - 1
                If InStr(UCase(Trim(DataGridView1.Item(0, i).Value.ToString)), UCase(Trim(TextBox1.Text))) <> 0 And InStr(UCase(Trim(DataGridView1.Item(0, i).Value.ToString)), UCase(Trim(TextBox2.Text))) <> 0 _
                    Or InStr(UCase(Trim(DataGridView1.Item(1, i).Value.ToString)), UCase(Trim(TextBox1.Text))) <> 0 And InStr(UCase(Trim(DataGridView1.Item(1, i).Value.ToString)), UCase(Trim(TextBox2.Text))) <> 0 _
                    Or InStr(UCase(Trim(DataGridView1.Item(2, i).Value.ToString)), UCase(Trim(TextBox1.Text))) <> 0 And InStr(UCase(Trim(DataGridView1.Item(2, i).Value.ToString)), UCase(Trim(TextBox2.Text))) <> 0 _
                    Or InStr(UCase(Trim(DataGridView1.Item(3, i).Value.ToString)), UCase(Trim(TextBox1.Text))) <> 0 And InStr(UCase(Trim(DataGridView1.Item(3, i).Value.ToString)), UCase(Trim(TextBox2.Text))) <> 0 Then
                    DataGridView1.CurrentCell = DataGridView1.Item(0, i)
                    Exit Sub
                End If
            Next
        End If
    End Sub

    Private Sub Button2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button2.Click
        '////////////////////////////////////////////////////////////////////////////////
        '//
        '// ����� ���������� ����������� �� �������� ����������
        '//
        '////////////////////////////////////////////////////////////////////////////////
        Dim MyRez As Object

        If Trim(TextBox1.Text) = "" And Trim(TextBox2.Text) = "" Then
        Else
            For i As Integer = DataGridView1.CurrentCellAddress.Y + 1 To DataGridView1.Rows.Count
                If i = DataGridView1.Rows.Count Then
                    MyRez = MsgBox("����� ����� �� ����� ������. ������ �������?", MsgBoxStyle.YesNo, "��������!")
                    If MyRez = 6 Then
                        i = 0
                    Else
                        Exit Sub
                    End If
                End If
                If DataGridView1.Rows.Count = 0 Then
                Else
                    If InStr(UCase(Trim(DataGridView1.Item(0, i).Value.ToString)), UCase(Trim(TextBox1.Text))) <> 0 And InStr(UCase(Trim(DataGridView1.Item(0, i).Value.ToString)), UCase(Trim(TextBox2.Text))) <> 0 _
                        Or InStr(UCase(Trim(DataGridView1.Item(1, i).Value.ToString)), UCase(Trim(TextBox1.Text))) <> 0 And InStr(UCase(Trim(DataGridView1.Item(1, i).Value.ToString)), UCase(Trim(TextBox2.Text))) <> 0 _
                        Or InStr(UCase(Trim(DataGridView1.Item(2, i).Value.ToString)), UCase(Trim(TextBox1.Text))) <> 0 And InStr(UCase(Trim(DataGridView1.Item(2, i).Value.ToString)), UCase(Trim(TextBox2.Text))) <> 0 _
                        Or InStr(UCase(Trim(DataGridView1.Item(3, i).Value.ToString)), UCase(Trim(TextBox1.Text))) <> 0 And InStr(UCase(Trim(DataGridView1.Item(3, i).Value.ToString)), UCase(Trim(TextBox2.Text))) <> 0 Then
                        DataGridView1.CurrentCell = DataGridView1.Item(0, i)
                        Exit Sub
                    End If
                End If
            Next i
        End If
    End Sub

    Private Sub Button6_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button6.Click
        '////////////////////////////////////////////////////////////////////////////////
        '//
        '// ������������� ���� ���������� �� �������� �����������
        '//
        '////////////////////////////////////////////////////////////////////////////////

        If Button6.Text = "���������� ���" Then
            For i As Integer = 0 To DataGridView1.Rows.Count - 1
                If InStr(UCase(Trim(DataGridView1.Item(0, i).Value.ToString)), UCase(Trim(TextBox1.Text))) <> 0 And InStr(UCase(Trim(DataGridView1.Item(0, i).Value.ToString)), UCase(Trim(TextBox2.Text))) <> 0 _
                    Or InStr(UCase(Trim(DataGridView1.Item(1, i).Value.ToString)), UCase(Trim(TextBox1.Text))) <> 0 And InStr(UCase(Trim(DataGridView1.Item(1, i).Value.ToString)), UCase(Trim(TextBox2.Text))) <> 0 _
                    Or InStr(UCase(Trim(DataGridView1.Item(2, i).Value.ToString)), UCase(Trim(TextBox1.Text))) <> 0 And InStr(UCase(Trim(DataGridView1.Item(2, i).Value.ToString)), UCase(Trim(TextBox2.Text))) <> 0 _
                    Or InStr(UCase(Trim(DataGridView1.Item(3, i).Value.ToString)), UCase(Trim(TextBox1.Text))) <> 0 And InStr(UCase(Trim(DataGridView1.Item(3, i).Value.ToString)), UCase(Trim(TextBox2.Text))) <> 0 Then
                    DataGridView1.Rows(i).DefaultCellStyle.BackColor = Color.Yellow
                Else
                    DataGridView1.Rows(i).DefaultCellStyle.BackColor = Color.Empty
                End If
            Next
            Button6.Text = "����� ���������"
        Else
            For i As Integer = 0 To DataGridView1.Rows.Count - 1
                DataGridView1.Rows(i).DefaultCellStyle.BackColor = Color.Empty
            Next
            Button6.Text = "���������� ���"
        End If
    End Sub

    Private Sub Button3_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button3.Click
        '////////////////////////////////////////////////////////////////////////////////
        '//
        '// ����� ���� ���������� �� �������� ����������� � ��������� ����
        '//
        '////////////////////////////////////////////////////////////////////////////////

        If Trim(TextBox1.Text) = "" And Trim(TextBox2.Text) = "" Then
            MsgBox("���������� ������ �������� ������", MsgBoxStyle.OkOnly, "��������!")
            TextBox1.Select()
        Else
            MySupplierSelectList = New SupplierSelectList
            MySupplierSelectList.ShowDialog()
        End If
    End Sub
End Class
