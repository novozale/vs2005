Public Class AgreedRange
    Public StartParam As String

    Private Sub Button2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button2.Click
        '/////////////////////////////////////////////////////////////////////////////////////
        '//
        '// ����� ��� ����������
        '// 
        '/////////////////////////////////////////////////////////////////////////////////////

        Me.Close()
    End Sub

    Private Sub TextBox1_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles TextBox1.KeyDown
        '////////////////////////////////////////////////////////////////////////////////
        '//
        '// ������� �� ��������� ���� �� Enter
        '//
        '////////////////////////////////////////////////////////////////////////////////

        If e.KeyData = Keys.Enter Then
            Me.SelectNextControl(sender, True, True, True, False)
        End If
    End Sub

    Private Sub TextBox1_Validating(ByVal sender As Object, ByVal e As System.ComponentModel.CancelEventArgs) Handles TextBox1.Validating
        '/////////////////////////////////////////////////////////////////////////////////////
        '//
        '// �������� ������������ ����� ���� ������
        '// 
        '/////////////////////////////////////////////////////////////////////////////////////
        Dim MySQLStr As String                        '������� ������

        If Trim(TextBox1.Text) <> "" Then
            MySQLStr = "SELECT Code, Name, LTRIM(RTRIM(SubGroupCode)) AS SubGroupCode "
            MySQLStr = MySQLStr & "FROM tbl_WEB_Items "
            MySQLStr = MySQLStr & "WHERE (Code = N'" & Trim(TextBox1.Text) & "') "
            InitMyConn(False)
            InitMyRec(False, MySQLStr)
            If Declarations.MyRec.EOF = True And Declarations.MyRec.BOF = True Then
                MsgBox("��� ������ �� ������ � ��.", MsgBoxStyle.Critical, "��������!")
                Label3.Text = ""
                e.Cancel = True
                trycloseMyRec()
                Exit Sub
            Else
                If Declarations.MyRec.Fields("SubGroupCode").Value = "" Then
                    MsgBox("��� ������� ���� ������ �� ��������� ���������. ������������� ����������� �� ����� ������ �������� �� �����", MsgBoxStyle.Critical, "��������!")
                    Label3.Text = ""
                    e.Cancel = True
                    trycloseMyRec()
                    Exit Sub
                Else
                    Label3.Text = Declarations.MyRec.Fields("Name").Value
                    Declarations.MyProductID = Trim(TextBox1.Text)
                    trycloseMyRec()
                End If
            End If
            e.Cancel = False
        Else
            Label3.Text = ""
            e.Cancel = False
        End If
    End Sub

    Private Sub TextBox3_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles TextBox3.KeyDown
        '////////////////////////////////////////////////////////////////////////////////
        '//
        '// ������� �� ��������� ���� �� Enter
        '//
        '////////////////////////////////////////////////////////////////////////////////

        If e.KeyData = Keys.Enter Then
            Me.SelectNextControl(sender, True, True, True, False)
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
                MsgBox("� ���� ""������������� ���� (��� ���)"" ������ ���� ������� �����", MsgBoxStyle.Critical, "��������!")
                e.Cancel = True
                Exit Sub
            Else
                Try
                    MyRez = TextBox3.Text
                Catch ex As Exception
                    MsgBox("� ���� ""������������� ���� (��� ���)"" ������ ���� ������� �����", MsgBoxStyle.Critical, "��������!")
                    e.Cancel = True
                    Exit Sub
                End Try

                If MyRez <= 0 Then
                    MsgBox("������������� ���� (��� ���) ������ ���� ������ ����.", MsgBoxStyle.Critical, "��������!")
                    e.Cancel = True
                    Exit Sub
                End If
            End If
        End If
        e.Cancel = False
    End Sub

    Private Function CheckData() As Boolean
        '/////////////////////////////////////////////////////////////////////////////////////
        '//
        '// �������� ���������� �����
        '//
        '/////////////////////////////////////////////////////////////////////////////////////

        If Trim(TextBox3.Text) = "" Then
            MsgBox("���� ""������������� ���� (��� ���)"" ������ ���� ���������.")
            CheckData = False
            TextBox3.Select()
            Exit Function
        End If

        If Trim(TextBox1.Text) = "" Then
            MsgBox("���� ""��� ������"" ������ ���� ���������.")
            CheckData = False
            TextBox1.Select()
            Exit Function
        End If

        CheckData = True
    End Function

    Private Sub Button3_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button3.Click
        '/////////////////////////////////////////////////////////////////////////////////////
        '//
        '// �������� ���� �� ������� �������
        '// 
        '/////////////////////////////////////////////////////////////////////////////////////

        MyItemList = New ItemList
        MyItemList.StartParam = "AgreedRange"
        MyItemList.ShowDialog()
    End Sub

    Private Sub AgreedRange_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        '/////////////////////////////////////////////////////////////////////////////////////
        '//
        '// �������� ������
        '// 
        '/////////////////////////////////////////////////////////////////////////////////////
        Dim MySQLStr As String                        '������� ������
        Dim MyAdapter As SqlClient.SqlDataAdapter     '��� ������ �����
        Dim MyDs As New DataSet

        '--------------������ �����
        MySQLStr = "SELECT SYCD001 AS Code, CONVERT(nvarchar(5), SYCD001) + ' ' + SYCD009 AS Name "
        MySQLStr = MySQLStr & "FROM SYCD0100 "
        MySQLStr = MySQLStr & "WHERE (SYCD001 IN (0, 1, 4, 12)) "
        MySQLStr = MySQLStr & "ORDER BY Code "

        InitMyConn(False)
        Try
            MyAdapter = New SqlClient.SqlDataAdapter(MySQLStr, Declarations.MyNETConnStr)
            MyAdapter.SelectCommand.CommandTimeout = 600
            MyAdapter.Fill(MyDs)
            ComboBox2.DisplayMember = "Name" '��� �� ��� ����� ������������
            ComboBox2.ValueMember = "Code"   '��� �� ��� ����� ���������
            ComboBox2.DataSource = MyDs.Tables(0).DefaultView
        Catch ex As Exception
            MsgBox(ex.ToString)
        End Try

        '-----------�������� ������
        If StartParam = "Edit" Then
            MySQLStr = "SELECT tbl_WEB_AgreedRange.ItemCode, tbl_WEB_Items.Name, tbl_WEB_AgreedRange.AgreedPrice, tbl_WEB_AgreedRange.CurrCode "
            MySQLStr = MySQLStr & "FROM tbl_WEB_AgreedRange LEFT OUTER JOIN "
            MySQLStr = MySQLStr & "tbl_WEB_Items ON tbl_WEB_AgreedRange.ItemCode = tbl_WEB_Items.Code "
            MySQLStr = MySQLStr & "WHERE (tbl_WEB_AgreedRange.ClientCode = N'" & Declarations.MyCustomerID & "') "
            MySQLStr = MySQLStr & "AND (tbl_WEB_AgreedRange.ItemCode = N'" & Declarations.MyProductID & "') "

            InitMyConn(False)
            InitMyRec(False, MySQLStr)
            If Declarations.MyRec.EOF = True And Declarations.MyRec.BOF = True Then
                MsgBox("���������� ������ � ������������� ������������ �� �������, �������� ������� ������ �������������. �������� � �������� �� ����� ������� �������������� ������������.", MsgBoxStyle.Critical, "��������!")
                trycloseMyRec()
                Me.Close()
            Else
                TextBox1.Text = Declarations.MyProductID
                Label3.Text = Declarations.MyRec.Fields("Name").Value
                TextBox3.Text = Declarations.MyRec.Fields("AgreedPrice").Value
                ComboBox2.SelectedValue = Declarations.MyRec.Fields("CurrCode").Value
                trycloseMyRec()
            End If
            TextBox1.Enabled = False
            Button3.Enabled = False
        Else
            TextBox1.Text = ""
            Label3.Text = ""
            TextBox3.Text = ""
            TextBox1.Enabled = True
            Button3.Enabled = True
        End If

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
                MySQLStr = "UPDATE tbl_WEB_AgreedRange "
                MySQLStr = MySQLStr & "SET AgreedPrice = " & Replace(Trim(TextBox3.Text), ",", ".") & ", "
                MySQLStr = MySQLStr & "CurrCode = " & ComboBox2.SelectedValue & " "
                MySQLStr = MySQLStr & "WHERE (ClientCode = N'" & Declarations.MyCustomerID & "') "
                MySQLStr = MySQLStr & "AND (ItemCode = N'" & Declarations.MyProductID & "') "

                InitMyConn(False)
                Declarations.MyConn.Execute(MySQLStr)
            Else
                Declarations.MyProductID = Trim(TextBox1.Text)
                '----�������� - ����� ����, ������ �� ���� ����� ����� ������� ��� ����
                MySQLStr = "SELECT COUNT(*) AS CC "
                MySQLStr = MySQLStr & "FROM tbl_WEB_AgreedRange "
                MySQLStr = MySQLStr & "WHERE (ItemCode = N'" & Declarations.MyProductID & "') "
                MySQLStr = MySQLStr & "AND (ClientCode = N'" & Declarations.MyCustomerID & "') "
                InitMyConn(False)
                InitMyRec(False, MySQLStr)
                If Declarations.MyRec.EOF = True And Declarations.MyRec.BOF = True Then
                    MsgBox("���������� ��������� �������� ������� ���� ������. ���������� � ��������������.", MsgBoxStyle.Critical, "��������!")
                    trycloseMyRec()
                    Exit Sub
                Else
                    If Declarations.MyRec.Fields("CC").Value <> 0 Then
                        MsgBox("�� ������ ����� ������� ���������� ��� ������� ������������� �����������. �������������� �������� �������������� �������� ���� �������������� ������������.", MsgBoxStyle.Critical, "��������!")
                        trycloseMyRec()
                        Exit Sub
                    Else
                        trycloseMyRec()
                        MySQLStr = "INSERT INTO tbl_WEB_AgreedRange "
                        MySQLStr = MySQLStr & "(ID, ItemCode, ClientCode, AgreedPrice, CurrCode) "
                        MySQLStr = MySQLStr & "VALUES (NEWID(), "
                        MySQLStr = MySQLStr & "N'" & Declarations.MyProductID & "', "
                        MySQLStr = MySQLStr & "N'" & Declarations.MyCustomerID & "', "
                        MySQLStr = MySQLStr & Replace(Trim(TextBox3.Text), ",", ".") & ", "
                        MySQLStr = MySQLStr & ComboBox2.SelectedValue & ") "

                        InitMyConn(False)
                        Declarations.MyConn.Execute(MySQLStr)
                    End If
                End If
            End If

            Me.Close()
        End If
    End Sub
End Class