Public Class DiscountGroup
    Public StartParam As String

    Private Sub Button2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button2.Click
        '/////////////////////////////////////////////////////////////////////////////////////
        '//
        '// ����� ��� ����������
        '// 
        '/////////////////////////////////////////////////////////////////////////////////////

        Me.Close()
    End Sub

    Private Sub DiscountGroup_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        '/////////////////////////////////////////////////////////////////////////////////////
        '//
        '// �������� ������
        '// 
        '/////////////////////////////////////////////////////////////////////////////////////
        Dim MySQLStr As String                        '������� ������
        Dim MyAdapter As SqlClient.SqlDataAdapter     '��� ������ ��������
        Dim MyDs As New DataSet

        '--------------������ �����
        If StartParam = "Edit" Then
            MySQLStr = "SELECT Ltrim(Rtrim(Code)) as Code, LTRIM(RTRIM(LTRIM(RTRIM(Code)) + ' ' + LTRIM(RTRIM(Name)))) AS Name "
            MySQLStr = MySQLStr & "FROM tbl_WEB_ItemGroup "
            MySQLStr = MySQLStr & "ORDER BY Code "
        Else
            MySQLStr = "SELECT Ltrim(Rtrim(Code)) as Code, LTRIM(RTRIM(LTRIM(RTRIM(Code)) + ' ' + LTRIM(RTRIM(Name)))) AS Name "
            MySQLStr = MySQLStr & "FROM tbl_WEB_ItemGroup "
            MySQLStr = MySQLStr & "WHERE (Code NOT IN "
            MySQLStr = MySQLStr & "(SELECT GroupCode "
            MySQLStr = MySQLStr & "FROM tbl_WEB_DiscountGroup "
            MySQLStr = MySQLStr & "WHERE (ClientCode = N'" & Declarations.MyCustomerID & "'))) "
            MySQLStr = MySQLStr & "ORDER BY Code "
        End If
        
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

        '-----------�������� ������
        If StartParam = "Edit" Then
            MySQLStr = "SELECT Discount "
            MySQLStr = MySQLStr & "FROM tbl_WEB_DiscountGroup "
            MySQLStr = MySQLStr & "WHERE (ClientCode = N'" & Declarations.MyCustomerID & "') "
            MySQLStr = MySQLStr & "AND (GroupCode = N'" & Declarations.MyProductGroupID & "') "
            InitMyConn(False)
            InitMyRec(False, MySQLStr)
            If Declarations.MyRec.EOF = True And Declarations.MyRec.BOF = True Then
                MsgBox("���������� ������ �� ������ ������� �� �������, �������� ������� ������ �������������. �������� � �������� �� ����� ������� ������ �� ������ �������.", MsgBoxStyle.Critical, "��������!")
                trycloseMyRec()
                Me.Close()
            Else
                ComboBox1.SelectedValue = Declarations.MyProductGroupID
                TextBox3.Text = Declarations.MyRec.Fields("Discount").Value
                trycloseMyRec()
            End If
            ComboBox1.Enabled = False
        Else
            ComboBox1.Enabled = True
        End If
    End Sub

    Private Sub TextBox3_Validating(ByVal sender As Object, ByVal e As System.ComponentModel.CancelEventArgs) Handles TextBox3.Validating
        '////////////////////////////////////////////////////////////////////////////////
        '//
        '// �������� - �������� �� �������� �������
        '//
        '////////////////////////////////////////////////////////////////////////////////
        Dim MyRez As Double
        Dim aa As New System.Globalization.NumberFormatInfo

        If Trim(TextBox3.Text) <> "" Then
            If InStr(TextBox3.Text, aa.CurrentInfo.NumberGroupSeparator) Then
                MsgBox("� ���� ""������ (%)"" ������ ���� ������� �����", MsgBoxStyle.Critical, "��������!")
                e.Cancel = True
                Exit Sub
            Else
                Try
                    MyRez = TextBox3.Text
                Catch ex As Exception
                    MsgBox("� ���� ""������ (%)"" ������ ���� ������� �����", MsgBoxStyle.Critical, "��������!")
                    e.Cancel = True
                    Exit Sub
                End Try

                If MyRez <= 0 Then
                    MsgBox("������ ������ ���� ������ ����.", MsgBoxStyle.Critical, "��������!")
                    e.Cancel = True
                    Exit Sub
                End If
            End If
        End If
        e.Cancel = False
    End Sub

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        '/////////////////////////////////////////////////////////////////////////////////////
        '//
        '// ���������� ������
        '// 
        '/////////////////////////////////////////////////////////////////////////////////////
        Dim MySQLStr As String

        If CheckData() = True Then
            If StartParam = "Edit" Then
                MySQLStr = "UPDATE tbl_WEB_DiscountGroup "
                MySQLStr = MySQLStr & "SET Discount = " & Replace(Trim(TextBox3.Text), ",", ".") & " "
                MySQLStr = MySQLStr & "WHERE (ClientCode = N'" & Declarations.MyCustomerID & "') "
                MySQLStr = MySQLStr & "AND (GroupCode = N'" & Declarations.MyProductGroupID & "') "

                InitMyConn(False)
                Declarations.MyConn.Execute(MySQLStr)
            Else
                Declarations.MyProductGroupID = ComboBox1.SelectedValue

                MySQLStr = "INSERT INTO tbl_WEB_DiscountGroup "
                MySQLStr = MySQLStr & "(ID, ClientCode, GroupCode, Discount) "
                MySQLStr = MySQLStr & "VALUES (NEWID(), "
                MySQLStr = MySQLStr & "N'" & Declarations.MyCustomerID & "', "
                MySQLStr = MySQLStr & "N'" & Declarations.MyProductGroupID & "', "
                MySQLStr = MySQLStr & Replace(Trim(TextBox3.Text), ",", ".") & ") "

                InitMyConn(False)
                Declarations.MyConn.Execute(MySQLStr)
            End If

            Me.Close()
        End If
    End Sub

    Private Function CheckData() As Boolean
        '/////////////////////////////////////////////////////////////////////////////////////
        '//
        '// �������� ���������� �����
        '//
        '/////////////////////////////////////////////////////////////////////////////////////

        If Trim(TextBox3.Text) = "" Then
            MsgBox("���� ""������ (%)"" ������ ���� ���������.")
            CheckData = False
            TextBox3.Select()
            Exit Function
        End If

        If ComboBox1.SelectedValue = Nothing Then
            MsgBox("������ ��������� ��� ������ ������ ���� �������.")
            CheckData = False
            ComboBox1.Select()
            Exit Function
        End If

        CheckData = True
    End Function
End Class