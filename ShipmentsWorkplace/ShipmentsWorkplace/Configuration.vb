Public Class Configuration

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        '////////////////////////////////////////////////////////////////////////////////
        '//
        '// �����
        '//
        '////////////////////////////////////////////////////////////////////////////////

        Me.Close()
    End Sub

    Private Sub Configuration_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        '////////////////////////////////////////////////////////////////////////////////
        '//
        '// �������� ������
        '//
        '////////////////////////////////////////////////////////////////////////////////
        Dim MySQLStr As String

        '-----EMail ��� �����������
        MySQLStr = "SELECT Value "
        MySQLStr = MySQLStr & "FROM tbl_Shipments_SalesmanWP_Config "
        MySQLStr = MySQLStr & "WHERE (UserId = " & Declarations.UserID & ") "
        MySQLStr = MySQLStr & "AND (Parameter = N'EMail ��� �����������') "
        InitMyConn(False)
        InitMyRec(False, MySQLStr)
        If Declarations.MyRec.EOF = True And Declarations.MyRec.BOF = True Then
            ComboBoxEMail.Text = "�������"
            trycloseMyRec()
        Else
            If Declarations.MyRec.Fields("Value").Value = "0" Then
                ComboBoxEMail.Text = "�������"
            Else
                ComboBoxEMail.Text = "��������� �� CRM �������"
            End If
            trycloseMyRec()
        End If

        '-----���������� ���� �������
        MySQLStr = "SELECT Value "
        MySQLStr = MySQLStr & "FROM tbl_Shipments_SalesmanWP_Config "
        MySQLStr = MySQLStr & "WHERE (UserId = " & Declarations.UserID & ") "
        MySQLStr = MySQLStr & "AND (Parameter = N'���������� ���� �������') "
        InitMyConn(False)
        InitMyRec(False, MySQLStr)
        If Declarations.MyRec.EOF = True And Declarations.MyRec.BOF = True Then
            ComboBoxContact.Text = "�������"
            trycloseMyRec()
        Else
            If Declarations.MyRec.Fields("Value").Value = "0" Then
                ComboBoxContact.Text = "�������"
            Else
                ComboBoxContact.Text = "�� Scala"
            End If
            trycloseMyRec()
        End If
    End Sub

    Private Sub Button2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button2.Click
        '////////////////////////////////////////////////////////////////////////////////
        '//
        '// ����� � ����������� ������
        '//
        '////////////////////////////////////////////////////////////////////////////////

        SaveInfo()
        Me.Close()
    End Sub

    Private Sub SaveInfo()
        '////////////////////////////////////////////////////////////////////////////////
        '//
        '// ���������� ������ �� ������������
        '//
        '////////////////////////////////////////////////////////////////////////////////
        Dim MySQLStr As String

        '-----EMail ��� �����������
        MySQLStr = "DELETE FROM tbl_Shipments_SalesmanWP_Config "
        MySQLStr = MySQLStr & "WHERE (UserId = " & Declarations.UserID & ") "
        MySQLStr = MySQLStr & "AND (Parameter = N'EMail ��� �����������') "
        InitMyConn(False)
        Declarations.MyConn.Execute(MySQLStr)

        MySQLStr = "INSERT INTO tbl_Shipments_SalesmanWP_Config "
        MySQLStr = MySQLStr & "(UserId, Parameter, Value) "
        MySQLStr = MySQLStr & "VALUES (" & Declarations.UserID & ", "
        MySQLStr = MySQLStr & "N'EMail ��� �����������', "
        If ComboBoxEMail.Text = "�������" Then
            MySQLStr = MySQLStr & "N'0') "
            MyEmail = 0
        Else
            MySQLStr = MySQLStr & "N'1') "
            MyEmail = 1
        End If
        InitMyConn(False)
        Declarations.MyConn.Execute(MySQLStr)

        '-----���������� ���� �������
        MySQLStr = "DELETE FROM tbl_Shipments_SalesmanWP_Config "
        MySQLStr = MySQLStr & "WHERE (UserId = " & Declarations.UserID & ") "
        MySQLStr = MySQLStr & "AND (Parameter = N'���������� ���� �������') "
        InitMyConn(False)
        Declarations.MyConn.Execute(MySQLStr)

        MySQLStr = "INSERT INTO tbl_Shipments_SalesmanWP_Config "
        MySQLStr = MySQLStr & "(UserId, Parameter, Value) "
        MySQLStr = MySQLStr & "VALUES (" & Declarations.UserID & ", "
        MySQLStr = MySQLStr & "N'���������� ���� �������', "
        If ComboBoxContact.Text = "�������" Then
            MySQLStr = MySQLStr & "N'0') "
            MyContact = 0
        Else
            MySQLStr = MySQLStr & "N'1') "
            MyContact = 1
        End If
        InitMyConn(False)
        Declarations.MyConn.Execute(MySQLStr)
    End Sub
End Class