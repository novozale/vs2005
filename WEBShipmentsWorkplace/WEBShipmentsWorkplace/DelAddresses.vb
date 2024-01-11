Public Class DelAddresses

    Private Sub Button5_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button5.Click
        '////////////////////////////////////////////////////////////////////////////////
        '//
        '// ����� ��� ������
        '//
        '////////////////////////////////////////////////////////////////////////////////

        Me.Close()
    End Sub

    Private Sub DelAddresses_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles Me.KeyDown
        '/////////////////////////////////////////////////////////////////////////////////////
        '//
        '// ����� ��� ������ �� Escape
        '// 
        '/////////////////////////////////////////////////////////////////////////////////////

        If e.KeyData = Keys.Escape Then
            Me.Close()
        End If
    End Sub

    Private Sub DelAddresses_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        '/////////////////////////////////////////////////////////////////////////////////////
        '//
        '// ��� ������� ��������� ������ �������
        '// 
        '/////////////////////////////////////////////////////////////////////////////////////

        Label2.Text = Trim(MyShipment.LblCustomerCode.Text) + " " + Trim(MyShipment.LblCustomerName.Text)
        LoadData()
        CheckButtons()
    End Sub

    Private Sub LoadData()
        '/////////////////////////////////////////////////////////////////////////////////////
        '//
        '// �������� ������ �������
        '// 
        '/////////////////////////////////////////////////////////////////////////////////////

        Dim MySQLStr As String                        '������� ������
        Dim MyAdapter As SqlClient.SqlDataAdapter     '
        Dim MyDs As New DataSet                       '

        MySQLStr = "SELECT     SL14002 AS ID, LTRIM(RTRIM(LTRIM(RTRIM(SL14003)) + ' ' + LTRIM(RTRIM(SL14004)) + ' ' + LTRIM(RTRIM(SL14005)) + ' ' + LTRIM(RTRIM(SL14006)))) AS Address "
        MySQLStr = MySQLStr & "FROM SL140300 "
        MySQLStr = MySQLStr & "WHERE (SL14001 = N'" & Trim(MyShipment.LblCustomerCode.Text) & "') "
        MySQLStr = MySQLStr & "ORDER BY ID "

        Try
            MyAdapter = New SqlClient.SqlDataAdapter(MySQLStr, Declarations.MyNETConnStr)
            MyAdapter.SelectCommand.CommandTimeout = 600
            MyAdapter.Fill(MyDs)
            DataGridView1.DataSource = MyDs.Tables(0)
        Catch ex As Exception
            MsgBox(ex.ToString)
        End Try

        DataGridView1.Columns(0).HeaderText = "ID"
        DataGridView1.Columns(0).Width = 40
        DataGridView1.Columns(1).HeaderText = "�����"
        DataGridView1.Columns(1).Width = 700
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

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        '/////////////////////////////////////////////////////////////////////////////////////
        '//
        '// ���������� ������ �������
        '// 
        '/////////////////////////////////////////////////////////////////////////////////////

        Label2.Text = Trim(MyShipment.LblCustomerCode.Text) + " " + Trim(MyShipment.LblCustomerName.Text)
        LoadData()
        CheckButtons()
    End Sub

    Private Sub DataGridView1_CellDoubleClick(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs) Handles DataGridView1.CellDoubleClick
        '////////////////////////////////////////////////////////////////////////////////
        '//
        '// ����� �� ���� � ������� ������
        '//
        '////////////////////////////////////////////////////////////////////////////////

        If DataGridView1.SelectedRows.Count <> 0 Then
            AddressSelect()
        End If
    End Sub

    Private Sub Button4_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button4.Click
        '////////////////////////////////////////////////////////////////////////////////
        '//
        '// ����� �� ���� � ������� ������
        '//
        '////////////////////////////////////////////////////////////////////////////////

        AddressSelect()
    End Sub

    Private Sub AddressSelect()
        '////////////////////////////////////////////////////////////////////////////////
        '//
        '// ����� �� ���� � ������� ������
        '//
        '////////////////////////////////////////////////////////////////////////////////

        MyShipment.TextBox2.Text = Trim(Me.DataGridView1.SelectedRows.Item(0).Cells(1).Value.ToString())
        Me.Close()
    End Sub
End Class