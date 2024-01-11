Public Class TransferNamesDescrToDB

    Private Sub Button3_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button3.Click
        '/////////////////////////////////////////////////////////////////////////////////////
        '//
        '// ����� ��� ����������
        '// 
        '/////////////////////////////////////////////////////////////////////////////////////

        Me.Close()
    End Sub

    Private Sub TransferNamesDescrToDB_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles Me.KeyDown
        '/////////////////////////////////////////////////////////////////////////////////////
        '//
        '// ������ ������ �� alt - F4
        '// 
        '/////////////////////////////////////////////////////////////////////////////////////

        If e.KeyData = Keys.Alt + Keys.F4 Then
            e.Handled = True
        End If
    End Sub

    Private Sub Button2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button2.Click
        '/////////////////////////////////////////////////////////////////////////////////////
        '//
        '// ������� ������ WEB � ��
        '// 
        '/////////////////////////////////////////////////////////////////////////////////////
        Dim MySelection As Integer

        If CheckBox1.Checked = True Then
            MySelection = 1
        Else
            MySelection = 0
        End If

        Me.Cursor = Cursors.WaitCursor
        MoveNamesDescrToDB(ComboBox1.SelectedValue, ComboBox2.SelectedItem, MySelection)
        Me.Cursor = Cursors.Default
    End Sub

    Private Sub TransferNamesDescrToDB_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        '/////////////////////////////////////////////////////////////////////////////////////
        '//
        '// �������� ������ � ����
        '// 
        '/////////////////////////////////////////////////////////////////////////////////////
        Dim MySQLStr As String
        Dim MyAdapter As SqlClient.SqlDataAdapter     '��� ����� �������
        Dim MyDs As New DataSet

        InitMyConn(False)
        '---���������� �������� � �������
        MySQLStr = "SELECT DISTINCT ID, Convert(nvarchar(10),ID) + '  ' + Ltrim(Rtrim(CompanyName)) AS CompanyName "
        MySQLStr = MySQLStr & "FROM tbl_WEBSearchScrapping_Companies "
        MySQLStr = MySQLStr & "ORDER BY ID "

        Try
            MyAdapter = New SqlClient.SqlDataAdapter(MySQLStr, Declarations.MyNETConnStr)
            MyAdapter.SelectCommand.CommandTimeout = 600
            MyAdapter.Fill(MyDs)
            ComboBox1.DisplayMember = "CompanyName" '��� �� ��� ����� ������������
            ComboBox1.ValueMember = "ID"   '��� �� ��� ����� ���������
            ComboBox1.DataSource = MyDs.Tables(0).DefaultView
        Catch ex As Exception
            MsgBox(ex.ToString)
        End Try

        ComboBox2.SelectedItem = "��������"
    End Sub

    Private Sub MoveNamesDescrToDB(ByVal MyCompany As String, ByVal MyItemPart As String, ByVal MySelection As Integer)
        '/////////////////////////////////////////////////////////////////////////////////////
        '//
        '// ��������� �������� ������ WEB � ��
        '// 
        '/////////////////////////////////////////////////////////////////////////////////////

        MyCheckUpdateNamesDescr = New CheckUpdateNamesDescr
        MyCheckUpdateNamesDescr.MyItemPart = MyItemPart
        MyCheckUpdateNamesDescr.MySelection = MySelection
        MyCheckUpdateNamesDescr.MyCompanyCode = MyCompany
        MyCheckUpdateNamesDescr.ShowDialog()

    End Sub
End Class