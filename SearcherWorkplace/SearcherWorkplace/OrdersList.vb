Public Class OrdersList
    Public LoadFlag As Integer

    Private Sub OrdersList_FormClosed(ByVal sender As Object, ByVal e As System.Windows.Forms.FormClosedEventArgs) Handles Me.FormClosed
        '////////////////////////////////////////////////////////////////////////////////
        '//
        '// ��������� ���������� ���������� ��� ������ �� �����
        '//
        '////////////////////////////////////////////////////////////////////////////////

        MyOrdersList = Nothing
    End Sub

    Private Sub OrdersList_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles Me.KeyDown
        '////////////////////////////////////////////////////////////////////////////////
        '//
        '// ������ �������� ���� �� ALT+F4
        '//
        '////////////////////////////////////////////////////////////////////////////////

        If e.KeyData = Keys.Alt + Keys.F4 Then
            e.Handled = True
        End If
    End Sub

    Private Sub Button13_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button13.Click
        '////////////////////////////////////////////////////////////////////////////////
        '//
        '// �����
        '//
        '////////////////////////////////////////////////////////////////////////////////

        Me.Close()
    End Sub

    Private Sub OrdersList_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        '////////////////////////////////////////////////////////////////////////////////
        '//
        '// �������� ����, ������ � ����...
        '//
        '////////////////////////////////////////////////////////////////////////////////

        LoadFlag = 1

        '---���������� ���������� - ��� ��������
        'If Declarations.IsManager = 1 Then
        '    ComboBoxAct.Items.Clear()
        '    ComboBoxAct.Items.Add(Declarations.PurchName)
        '    ComboBoxAct.Items.Add("��� ����������")
        '    ComboBoxAct.SelectedIndex = 0
        'Else
        '    ComboBoxAct.Items.Clear()
        '    ComboBoxAct.Items.Add(Declarations.PurchName)
        '    ComboBoxAct.SelectedIndex = 0
        'End If

        ComboBoxAct.Items.Clear()
        ComboBoxAct.Items.Add(Declarations.PurchName)
        ComboBoxAct.Items.Add("��� ����������")
        ComboBoxAct.SelectedIndex = 1

        LoadFlag = 0
        Application.DoEvents()
        Windows.Forms.Cursor.Current = Cursors.WaitCursor
        LoadOrders()
        LoadItems()
        LoadSearchRez()
        CheckMarkButtons()
        CheckAttachButtons()
        Windows.Forms.Cursor.Current = Cursors.Default
    End Sub

    Private Sub LoadOrders()
        '////////////////////////////////////////////////////////////////////////////////
        '//
        '// �������� ������ �� �������
        '//
        '////////////////////////////////////////////////////////////////////////////////
        Dim MySQLStr As String
        Dim MyAdapter As SqlClient.SqlDataAdapter     '��� ������ �����������
        Dim MyDs As New DataSet

        If LoadFlag = 0 Then
            MySQLStr = "exec spp_SupplSearch_Orders0TypeInfo N'" + Trim(ComboBoxAct.Text) + "' "

            Try
                MyAdapter = New SqlClient.SqlDataAdapter(MySQLStr, Declarations.MyNETConnStr)
                MyAdapter.SelectCommand.CommandTimeout = 600
                MyAdapter.Fill(MyDs)
                DataGridView1.DataSource = MyDs.Tables(0)
            Catch ex As Exception
                MsgBox(ex.ToString)
            End Try

            '---���������
            DataGridView1.Columns(0).HeaderText = "N ������"
            DataGridView1.Columns(0).Width = 100
            DataGridView1.Columns(1).HeaderText = "��� ������ ����"
            DataGridView1.Columns(1).Width = 50
            DataGridView1.Columns(2).HeaderText = "���������"
            DataGridView1.Columns(2).Width = 250
            DataGridView1.Columns(3).HeaderText = "���� ������"
            DataGridView1.Columns(3).Width = 100
            DataGridView1.Columns(3).DefaultCellStyle.Format = "dd/MM/yyyy"
            DataGridView1.Columns(4).HeaderText = "����� ������"
            DataGridView1.Columns(4).Width = 200
            DataGridView1.Columns(4).DefaultCellStyle.Format = "n2"
            DataGridView1.Columns(5).HeaderText = "������ ������"
            DataGridView1.Columns(5).Width = 50
            DataGridView1.Columns(6).HeaderText = "��� ����� ����"
            DataGridView1.Columns(6).Width = 50
            DataGridView1.Columns(7).HeaderText = "��������"
            DataGridView1.Columns(7).Width = 200
            DataGridView1.Columns(8).HeaderText = "N ������ �� �������"
            DataGridView1.Columns(8).Width = 100
            DataGridView1.Columns(9).HeaderText = "��� ��� �����"
            DataGridView1.Columns(9).Width = 50
            DataGridView1.Columns(10).HeaderText = "��������"
            DataGridView1.Columns(10).Width = 200

        End If
    End Sub

    Private Sub LoadItems()
        '////////////////////////////////////////////////////////////////////////////////
        '//
        '// �������� ������ �� ������� � ������
        '//
        '////////////////////////////////////////////////////////////////////////////////
        Dim MySQLStr As String
        Dim MyAdapter As SqlClient.SqlDataAdapter     '��� ������ �����������
        Dim MyDs As New DataSet

        If LoadFlag = 0 Then
            If DataGridView1.SelectedRows.Count <> 0 Then
                MySQLStr = "Exec spp_SupplSearch_ItemsInOrder0TypeInfo N'" & Trim(DataGridView1.SelectedRows.Item(0).Cells(0).Value.ToString()) & "' "
            Else
                MySQLStr = "Exec spp_SupplSearch_ItemsInOrder0TypeInfo N'' "
            End If

            Try
                MyAdapter = New SqlClient.SqlDataAdapter(MySQLStr, Declarations.MyNETConnStr)
                MyAdapter.SelectCommand.CommandTimeout = 600
                MyAdapter.Fill(MyDs)
                DataGridView2.DataSource = MyDs.Tables(0)
            Catch ex As Exception
                MsgBox(ex.ToString)
            End Try

            '---���������
            DataGridView2.Columns(0).HeaderText = "N ������"
            DataGridView2.Columns(0).Width = 80
            DataGridView2.Columns(1).HeaderText = "��� ������"
            DataGridView2.Columns(1).Width = 150
            DataGridView2.Columns(2).HeaderText = "�������� ������"
            DataGridView2.Columns(2).Width = 250
            DataGridView2.Columns(3).HeaderText = "����������"
            DataGridView2.Columns(3).Width = 100
            DataGridView2.Columns(3).DefaultCellStyle.Format = "n3"
            DataGridView2.Columns(4).HeaderText = "����"
            DataGridView2.Columns(4).Width = 100
            DataGridView2.Columns(4).DefaultCellStyle.Format = "n2"
            DataGridView2.Columns(5).HeaderText = "������"
            DataGridView2.Columns(5).Width = 50

        End If
    End Sub

    Private Sub LoadSearchRez()
        '////////////////////////////////////////////////////////////////////////////////
        '//
        '// �������� ������ �� ����������� ������ ���������� ������ �� ���
        '//
        '////////////////////////////////////////////////////////////////////////////////
        Dim MySQLStr As String
        Dim MyAdapter As SqlClient.SqlDataAdapter     '��� ������ �����������
        Dim MyDs As New DataSet

        If LoadFlag = 0 Then
            If DataGridView2.SelectedRows.Count <> 0 Then
                MySQLStr = "Exec spp_SupplSearch_ItemsInSearchInfo N'" & Trim(DataGridView2.SelectedRows.Item(0).Cells(1).Value.ToString()) & "' "
            Else
                MySQLStr = "Exec spp_SupplSearch_ItemsInSearchInfo N'999999999999999999999999999999' "
            End If

            Try
                MyAdapter = New SqlClient.SqlDataAdapter(MySQLStr, Declarations.MyNETConnStr)
                MyAdapter.SelectCommand.CommandTimeout = 600
                MyAdapter.Fill(MyDs)
                DataGridView3.DataSource = MyDs.Tables(0)
            Catch ex As Exception
                MsgBox(ex.ToString)
            End Try

            '---���������
            DataGridView3.Columns(0).HeaderText = "ID"
            DataGridView3.Columns(0).Width = 50
            DataGridView3.Columns(1).HeaderText = "���� ������"
            DataGridView3.Columns(1).Width = 100
            DataGridView3.Columns(1).DefaultCellStyle.Format = "dd/MM/yyyy"
            DataGridView3.Columns(2).HeaderText = "��� ������ ����"
            DataGridView3.Columns(2).Width = 50
            DataGridView3.Columns(3).HeaderText = "���������"
            DataGridView3.Columns(3).Width = 250
            DataGridView3.Columns(4).HeaderText = "����"
            DataGridView3.Columns(4).Width = 100
            DataGridView3.Columns(4).DefaultCellStyle.Format = "n2"
            DataGridView3.Columns(5).HeaderText = "����������� ��������"
            DataGridView3.Columns(5).Width = 150
            DataGridView3.Columns(6).HeaderText = "���������"
            DataGridView3.Columns(6).Width = 150
            DataGridView3.Columns(7).HeaderText = "��������� ��������"
            DataGridView3.Columns(7).Width = 100
            DataGridView3.Columns(8).HeaderText = "��������� ����������"
            DataGridView3.Columns(8).Width = 100
            DataGridView3.Columns(9).HeaderText = "��� �������"
            DataGridView3.Columns(9).Width = 70
            DataGridView3.Columns(10).HeaderText = "�������� �������"
            DataGridView3.Columns(10).Width = 150
        End If
    End Sub

    Private Sub CheckAttachButtons()
        '////////////////////////////////////////////////////////////////////////////////
        '//
        '// �������� ��������� ������ �������� �����������
        '//
        '////////////////////////////////////////////////////////////////////////////////

        If DataGridView3.SelectedRows.Count = 0 Then
            Button12.Enabled = False
            Button14.Enabled = False
        Else
            '-----��������� ��������
            If DataGridView3.SelectedRows.Item(0).Cells(7).Value = "" Then
                Button12.Enabled = False
            Else
                Button12.Enabled = True
            End If

            '-----��������� ����������
            If DataGridView3.SelectedRows.Item(0).Cells(8).Value = "" Then
                Button14.Enabled = False
            Else
                Button14.Enabled = True
            End If
        End If
    End Sub

    Private Sub CheckMarkButtons()
        '////////////////////////////////////////////////////////////////////////////////
        '//
        '// �������� ��������� ������ ������� � ��������
        '//
        '////////////////////////////////////////////////////////////////////////////////

        If DataGridView3.SelectedRows.Count = 0 Then
            Button1.Enabled = False
        Else
            Button1.Enabled = True
        End If
    End Sub

    Private Sub ComboBoxAct_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles ComboBoxAct.SelectedIndexChanged
        '////////////////////////////////////////////////////////////////////////////////
        '//
        '// ����� ������ ����������
        '//
        '////////////////////////////////////////////////////////////////////////////////

        Windows.Forms.Cursor.Current = Cursors.WaitCursor
        Application.DoEvents()
        LoadOrders()
        LoadItems()
        LoadSearchRez()
        CheckAttachButtons()
        CheckMarkButtons()
        Windows.Forms.Cursor.Current = Cursors.Default
    End Sub

    Private Sub Button16_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button16.Click
        '////////////////////////////////////////////////////////////////////////////////
        '//
        '// ���������� ����������
        '//
        '////////////////////////////////////////////////////////////////////////////////

        Windows.Forms.Cursor.Current = Cursors.WaitCursor
        Application.DoEvents()
        LoadOrders()
        LoadItems()
        LoadSearchRez()
        CheckAttachButtons()
        CheckMarkButtons()
        Windows.Forms.Cursor.Current = Cursors.Default
    End Sub

    Private Sub DataGridView1_SelectionChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles DataGridView1.SelectionChanged
        '////////////////////////////////////////////////////////////////////////////////
        '//
        '// ����� ������ ������ �� �������
        '//
        '////////////////////////////////////////////////////////////////////////////////

        If LoadFlag = 0 Then
            Windows.Forms.Cursor.Current = Cursors.WaitCursor
            LoadItems()
            LoadSearchRez()
            CheckAttachButtons()
            CheckMarkButtons()
            Windows.Forms.Cursor.Current = Cursors.Default
        End If
    End Sub

    Private Sub DataGridView2_SelectionChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles DataGridView2.SelectionChanged
        '////////////////////////////////////////////////////////////////////////////////
        '//
        '// ����� ������ ������ � ������
        '//
        '////////////////////////////////////////////////////////////////////////////////

        If LoadFlag = 0 Then
            Windows.Forms.Cursor.Current = Cursors.WaitCursor
            LoadSearchRez()
            CheckAttachButtons()
            CheckMarkButtons()
            Windows.Forms.Cursor.Current = Cursors.Default
        End If
    End Sub

    Private Sub Button12_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button12.Click
        '////////////////////////////////////////////////////////////////////////////////
        '//
        '// �������� ���� � ������������ ��������
        '//
        '////////////////////////////////////////////////////////////////////////////////
        Declarations.MyRequestNum = DataGridView3.SelectedRows.Item(0).Cells(0).Value
        MyAttachmentsList = New AttachmentsList
        MyAttachmentsList.AttType = "Sales"
        MyAttachmentsList.WhoStart = "Search"
        MyAttachmentsList.MyPlace = "OrderList"
        MyAttachmentsList.ShowDialog()
    End Sub

    Private Sub Button14_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button14.Click
        '////////////////////////////////////////////////////////////////////////////////
        '//
        '// �������� ���� � ������������ ����������
        '//
        '////////////////////////////////////////////////////////////////////////////////
        Declarations.MyRequestNum = DataGridView3.SelectedRows.Item(0).Cells(0).Value
        MyAttachmentsList = New AttachmentsList
        MyAttachmentsList.AttType = "Search"
        MyAttachmentsList.WhoStart = "Sales"
        MyAttachmentsList.MyPlace = "OrderList"
        MyAttachmentsList.ShowDialog()
    End Sub

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        '////////////////////////////////////////////////////////////////////////////////
        '//
        '// ������� ������ �� ������� ��� "�����������"
        '//
        '////////////////////////////////////////////////////////////////////////////////
        Dim MySQLStr As String
        Dim MyOrderNum As String

        MyOrderNum = Trim(DataGridView1.SelectedRows.Item(0).Cells(0).Value)
        '------------------������ ������ � "�����������"
        MySQLStr = "DELETE FROM tbl_SupplSearch_OrdersMarkAsChecked "
        MySQLStr = MySQLStr & "WHERE (PurchOrderNum = N'" & MyOrderNum & "') "
        InitMyConn(False)
        Declarations.MyConn.Execute(MySQLStr)

        MySQLStr = "INSERT INTO tbl_SupplSearch_OrdersMarkAsChecked (PurchOrderNum) "
        MySQLStr = MySQLStr & "VALUES (N'" & MyOrderNum & "') "
        InitMyConn(False)
        Declarations.MyConn.Execute(MySQLStr)

        '------------------���������� ������
        Windows.Forms.Cursor.Current = Cursors.WaitCursor
        Application.DoEvents()
        LoadOrders()
        LoadItems()
        LoadSearchRez()
        CheckMarkButtons()
        CheckAttachButtons()
        Windows.Forms.Cursor.Current = Cursors.Default
    End Sub
End Class