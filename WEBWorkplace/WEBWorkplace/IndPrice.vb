Public Class IndPrice

    Private Sub Button2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button2.Click
        '/////////////////////////////////////////////////////////////////////////////////////
        '//
        '// ����� ��� ����������
        '// 
        '/////////////////////////////////////////////////////////////////////////////////////

        Me.Close()
    End Sub

    Private Sub IndPrice_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        '/////////////////////////////////////////////////////////////////////////////////////
        '//
        '// �������� ������
        '// 
        '/////////////////////////////////////////////////////////////////////////////////////
        Dim MySQLStr As String                        '������� ������
        Dim MyAdapter As SqlClient.SqlDataAdapter     '��� ������ ��������
        Dim MyDs As New DataSet

        '---------------������ ��������
        MySQLStr = "SELECT Code, LTRIM(RTRIM(LTRIM(RTRIM(Code)) + ' ' + LTRIM(RTRIM(Name)))) AS Name "
        MySQLStr = MySQLStr & "FROM tbl_WEB_Clients "
        MySQLStr = MySQLStr & "WHERE (WorkOverWEB = 1) "
        MySQLStr = MySQLStr & "ORDER BY Name "

        InitMyConn(False)
        Try
            MyAdapter = New SqlClient.SqlDataAdapter(MySQLStr, Declarations.MyNETConnStr)
            MyAdapter.SelectCommand.CommandTimeout = 600
            MyAdapter.Fill(MyDs)
            ComboBox1.DisplayMember = "Name" '��� �� ��� ����� ������������
            ComboBox1.ValueMember = "Code"   '��� �� ��� ����� ���������
            ComboBox1.DataSource = MyDs.Tables(0).DefaultView
        Catch ex As Exception
            MsgBox(ex.ToString)
        End Try
    End Sub

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        '////////////////////////////////////////////////////////////////////////////////
        '//
        '// �������� � Excel ��������������� ����� �����
        '//
        '////////////////////////////////////////////////////////////////////////////////

        Me.Cursor = Cursors.WaitCursor
        If My.Settings.UseOffice = "LibreOffice" Then
            UploadIndividualPriceToLO(ComboBox1.SelectedValue, ComboBox1.Text, CheckBox1.Checked)
        Else
            UploadIndividualPriceToExcel(ComboBox1.SelectedValue, ComboBox1.Text, CheckBox1.Checked)
        End If

        Me.Cursor = Cursors.Default
    End Sub
End Class