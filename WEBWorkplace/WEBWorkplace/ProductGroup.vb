Public Class ProductGroup

    Private Sub Button2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button2.Click
        '/////////////////////////////////////////////////////////////////////////////////////
        '//
        '// ����� ��� ����������
        '// 
        '/////////////////////////////////////////////////////////////////////////////////////

        Me.Close()
    End Sub

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        '/////////////////////////////////////////////////////////////////////////////////////
        '//
        '// ���������� ������
        '// 
        '/////////////////////////////////////////////////////////////////////////////////////
        Dim MySQLStr As String

        MySQLStr = "UPDATE tbl_WEB_ItemGroup "
        MySQLStr = MySQLStr & "SET WEBName = N'" & Trim(TextBox3.Text) & "', "
        MySQLStr = MySQLStr & "RMStatus = CASE WHEN RMStatus = 1 THEN 1 ELSE 3 END, "
        MySQLStr = MySQLStr & "WEBStatus = CASE WHEN WEBStatus = 1 THEN 1 ELSE 3 END "
        MySQLStr = MySQLStr & "WHERE (Code = N'" & CStr(Declarations.MyProductGroupID) & "') "
        InitMyConn(False)
        Declarations.MyConn.Execute(MySQLStr)
        Me.Close()
    End Sub

    Private Sub ProductGroup_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        '/////////////////////////////////////////////////////////////////////////////////////
        '//
        '// �������� ������
        '// 
        '/////////////////////////////////////////////////////////////////////////////////////
        Dim MySQLStr As String                        '������� ������

        MySQLStr = "SELECT Code, Name, WEBName "
        MySQLStr = MySQLStr & "FROM  tbl_WEB_ItemGroup WITH(NOLOCK) "
        MySQLStr = MySQLStr & "WHERE (Code = " & UCase(Trim(Declarations.MyProductGroupID)) & ") "
        InitMyConn(False)
        InitMyRec(False, MySQLStr)
        If Declarations.MyRec.EOF = True And Declarations.MyRec.BOF = True Then
            MsgBox("���������� ������ ��������� �� �������, �������� ������� ������ �������������. �������� � �������� �� ����� ������� ����� ���������.", MsgBoxStyle.Critical, "��������!")
            trycloseMyRec()
            Me.Close()
        Else
            TextBox1.Text = Declarations.MyRec.Fields("Code").Value.ToString
            TextBox2.Text = Declarations.MyRec.Fields("Name").Value
            TextBox3.Text = Declarations.MyRec.Fields("WEBName").Value.ToString
            trycloseMyRec()
        End If
        TextBox1.Enabled = False
        TextBox2.Enabled = False
    End Sub
End Class