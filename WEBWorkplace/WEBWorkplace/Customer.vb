Public Class Customer

    Private Sub Button2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button2.Click
        '/////////////////////////////////////////////////////////////////////////////////////
        '//
        '// ����� ��� ����������
        '// 
        '/////////////////////////////////////////////////////////////////////////////////////

        Me.Close()
    End Sub

    Private Sub Customer_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        '/////////////////////////////////////////////////////////////////////////////////////
        '//
        '// �������� ������
        '// 
        '/////////////////////////////////////////////////////////////////////////////////////
        Dim MySQLStr As String                        '������� ������
        Dim MyAdapter As SqlClient.SqlDataAdapter     '��� ������ ��������
        Dim MyDs As New DataSet

        '---------------������ ������� ����� ������
        MySQLStr = "SELECT DISTINCT SY240300.SY24002 AS Code, SY240300.SY24002 + ' ' + SY240300.SY24003 AS Name "
        MySQLStr = MySQLStr & "FROM SY240300 INNER JOIN "
        MySQLStr = MySQLStr & "SC390300 ON SY240300.SY24002 = SC390300.SC39002 "
        MySQLStr = MySQLStr & "WHERE (SY240300.SY24001 = N'IL') "
        MySQLStr = MySQLStr & "ORDER BY SY240300.SY24002 "
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

        '-------------�������� ��������
        MySQLStr = "SELECT Code, Name, Address, Discount, WorkOverWEB, BasePrice "
        MySQLStr = MySQLStr & "FROM tbl_WEB_Clients "
        MySQLStr = MySQLStr & "WHERE (Code = N'" & Declarations.MyCustomerID & "') "
        InitMyConn(False)
        InitMyRec(False, MySQLStr)
        If Declarations.MyRec.EOF = True And Declarations.MyRec.BOF = True Then
            MsgBox("���������� ������ �� ������, �������� ������ ������ �������������. �������� � �������� �� ����� ������� ��������.", MsgBoxStyle.Critical, "��������!")
            trycloseMyRec()
            Me.Close()
        Else
            TextBox1.Text = Declarations.MyRec.Fields("Code").Value
            TextBox2.Text = Declarations.MyRec.Fields("Name").Value
            TextBox3.Text = Declarations.MyRec.Fields("Address").Value
            TextBox4.Text = Declarations.MyRec.Fields("Discount").Value
            CheckBox1.Checked = Declarations.MyRec.Fields("WorkOverWEB").Value
            ComboBox1.SelectedValue = Declarations.MyRec.Fields("BasePrice").Value
            trycloseMyRec()
        End If
    End Sub

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        '/////////////////////////////////////////////////////////////////////////////////////
        '//
        '// ���������� ������
        '// 
        '/////////////////////////////////////////////////////////////////////////////////////
        Dim MySQLStr As String

        MySQLStr = "UPDATE tbl_WEB_Clients "
        If CheckBox1.Checked = True Then
            MySQLStr = MySQLStr & "SET WorkOverWEB = 1, "
        Else
            MySQLStr = MySQLStr & "SET WorkOverWEB = 0, "
        End If
        MySQLStr = MySQLStr & "BasePrice = N'" & ComboBox1.SelectedValue & "' "
        MySQLStr = MySQLStr & "WHERE (Code = N'" & Declarations.MyCustomerID & "') "
        InitMyConn(False)
        Declarations.MyConn.Execute(MySQLStr)

        If CheckBox1.Checked = False Then
            '---------����� ����� ������� ��� ������ � ������������� �����������, ���� ������ ������ �� �������� ����� WEB
            '---������ �� �������
            MySQLStr = "DELETE FROM tbl_WEB_DiscountGroup "
            MySQLStr = MySQLStr & "WHERE (ClientCode = N'" & Declarations.MyCustomerID & "') "
            InitMyConn(False)
            Declarations.MyConn.Execute(MySQLStr)

            '---������ �� ����������
            MySQLStr = "DELETE FROM tbl_WEB_DiscountSubgroup "
            MySQLStr = MySQLStr & "WHERE (ClientCode = N'" & Declarations.MyCustomerID & "') "
            InitMyConn(False)
            Declarations.MyConn.Execute(MySQLStr)

            '---������ �� �������
            MySQLStr = "DELETE FROM tbl_WEB_DiscountItem "
            MySQLStr = MySQLStr & "WHERE (ClientCode = N'" & Declarations.MyCustomerID & "') "
            InitMyConn(False)
            Declarations.MyConn.Execute(MySQLStr)

            '---������������� �����������
            MySQLStr = "DELETE FROM tbl_WEB_AgreedRange "
            MySQLStr = MySQLStr & "WHERE (ClientCode = N'" & Declarations.MyCustomerID & "') "
            InitMyConn(False)
            Declarations.MyConn.Execute(MySQLStr)

        End If
        Me.Close()
    End Sub

    Private Sub CheckBox1_Validating(ByVal sender As Object, ByVal e As System.ComponentModel.CancelEventArgs) Handles CheckBox1.Validating
        '/////////////////////////////////////////////////////////////////////////////////////
        '//
        '// �������� ��������� ��������� � ��������� �� ����
        '// 
        '/////////////////////////////////////////////////////////////////////////////////////
        Dim MyRez As Object

        If CheckBox1.Checked = False Then
            MyRez = MsgBox("�� �������� �������, ��� ������ �������� ����� WEB. ��� ���� ����� ������� ��� ������, ����������� ������� ��� ������ ����� WEB � ������������� �����������. �� �������, ��� ������ ����� ��� �������?", MsgBoxStyle.YesNo, "��������!")
            If MyRez = 6 Then '--��
            Else
                e.Cancel = True
            End If
        End If
    End Sub

    Private Sub Label11_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Label11.Click

    End Sub

    Private Sub ComboBox1_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ComboBox1.SelectedIndexChanged

    End Sub
End Class