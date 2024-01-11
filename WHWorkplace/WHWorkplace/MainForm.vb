Public Class MainForm
    Public LoadFlag As Integer

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        '////////////////////////////////////////////////////////////////////////////////
        '//
        '// �����
        '//
        '////////////////////////////////////////////////////////////////////////////////

        Application.Exit()
    End Sub

    Private Sub MainForm_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles Me.KeyDown
        '////////////////////////////////////////////////////////////////////////////////
        '//
        '// ������ �������� ���� �� ALT+F4
        '//
        '////////////////////////////////////////////////////////////////////////////////

        If e.KeyData = Keys.Alt + Keys.F4 Then
            e.Handled = True
        End If
    End Sub

    Private Sub MainForm_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        '/////////////////////////////////////////////////////////////////////////////////////
        '//
        '// ��� ������� ���������� ��������� - ���, ��������, ������������ � �.�.
        '//  
        '/////////////////////////////////////////////////////////////////////////////////////
        '/////////////////////////////////////////////////////////////////////////////////////
        '//
        '// ��� ������� ���������� ��������� - ���, ��������, ������������ � �.�.
        '// ����� ���� ������� ������ ����������� 
        '/////////////////////////////////////////////////////////////////////////////////////
        Dim MySQLStr As String                        '������� ������
        Dim MyAdapter As SqlClient.SqlDataAdapter     '��� ������ �������
        Dim MyDs As New DataSet                       '
        Dim MyAdapter1 As SqlClient.SqlDataAdapter    '��� ������ ����������� �� 1 ��������
        Dim MyDs1 As New DataSet                      '

        LoadFlag = 1
        '---��������� �������
        Try
            Dim Scala As New SfwIII.Application

            Declarations.CompanyID = Scala.ActiveProcess.CommonVars.CompanyCode
            Declarations.Year = Mid(Scala.ActiveProcess.CommonVars.FiscalYear, 3)
            Declarations.UserCode = Scala.ActiveProcess.CommonVars.UserCode
            'Declarations.UserCode = "Novozhilov"
        Catch
            MsgBox("��������� ������ ����������� ������ �� ���� Scala", MsgBoxStyle.Critical, "��������!")
            Application.Exit()
        End Try

        '---ID ������������
        MySQLStr = "SELECT UserID, FullName, UserName "
        MySQLStr = MySQLStr & "FROM  ScalaSystemDB.dbo.ScaUsers WITH(NOLOCK) "
        MySQLStr = MySQLStr & "WHERE (Upper(UserName) = N'" & UCase(Trim(Declarations.UserCode)) & "') "
        InitMyConn(False)
        InitMyRec(False, MySQLStr)
        If Declarations.MyRec.EOF = True And Declarations.MyRec.BOF = True Then
            MsgBox("�� ������ ID ����������, ��������������� ������ �� ���� � Scala. ���������� � ��������������.", MsgBoxStyle.Critical, "��������!")
            trycloseMyRec()
            Application.Exit()
        Else
            Declarations.UserID = Declarations.MyRec.Fields("UserID").Value
            Declarations.FullName = Declarations.MyRec.Fields("FullName").Value
            Declarations.UserName = Declarations.MyRec.Fields("UserName").Value
            trycloseMyRec()
        End If

        '---������ �������
        MySQLStr = "SELECT SC23001 AS WHCode, SC23001 + ' ' + SC23002 AS WHName "
        MySQLStr = MySQLStr & "FROM SC230300 WITH(NOLOCK) "
        MySQLStr = MySQLStr & "WHERE (SC23006 = N'1') AND (SC23001 IN('01','03')) "
        MySQLStr = MySQLStr & "ORDER BY WHCode "
        Try
            MyAdapter = New SqlClient.SqlDataAdapter(MySQLStr, Declarations.MyNETConnStr)
            MyAdapter.SelectCommand.CommandTimeout = 600
            MyAdapter.Fill(MyDs)
            ComboBox1.DisplayMember = "WHName" '��� �� ��� ����� ������������
            ComboBox1.ValueMember = "WHCode"   '��� �� ��� ����� ���������
            ComboBox1.DataSource = MyDs.Tables(0).DefaultView
        Catch ex As Exception
            MsgBox(ex.ToString)
        End Try

        '---�������� ������ �� ����������� ����������
        MySQLStr = "SELECT Value "
        MySQLStr = MySQLStr & "FROM tbl_WHWorkplace_Config "
        MySQLStr = MySQLStr & "WHERE (UserLogin = N'" & Declarations.UserName & "') "
        MySQLStr = MySQLStr & "AND (Parameter = N'������ �� ������� �����') "
        InitMyConn(False)
        InitMyRec(False, MySQLStr)
        If Declarations.MyRec.EOF = True And Declarations.MyRec.BOF = True Then
            trycloseMyRec()
        Else
            ComboBox1.SelectedValue = Declarations.MyRec.Fields("Value").Value
            trycloseMyRec()
        End If
        Declarations.MyWH = Trim(ComboBox1.SelectedValue)

        '---������ �����������
        MySQLStr = "SELECT '---' AS UserName, ' ���' AS FullName "
        MySQLStr = MySQLStr & "UNION ALL "
        MySQLStr = MySQLStr & "SELECT ScalaSystemDB.dbo.ScaUsers.UserName, ScalaSystemDB.dbo.ScaUsers.FullName "
        MySQLStr = MySQLStr & "FROM tbl_WHWorkplace_Employees INNER JOIN "
        MySQLStr = MySQLStr & "ScalaSystemDB.dbo.ScaUsers ON tbl_WHWorkplace_Employees.WHEmployee = ScalaSystemDB.dbo.ScaUsers.UserName "
        MySQLStr = MySQLStr & "WHERE (tbl_WHWorkplace_Employees.WHCode = N'" & Trim(ComboBox1.SelectedValue) & "') "
        MySQLStr = MySQLStr & "ORDER BY FullName "
        Try
            MyAdapter1 = New SqlClient.SqlDataAdapter(MySQLStr, Declarations.MyNETConnStr)
            MyAdapter1.SelectCommand.CommandTimeout = 600
            MyAdapter1.Fill(MyDs1)
            ComboBox2.DisplayMember = "FullName" '��� �� ��� ����� ������������
            ComboBox2.ValueMember = "UserName"   '��� �� ��� ����� ���������
            ComboBox2.DataSource = MyDs1.Tables(0).DefaultView
        Catch ex As Exception
            MsgBox(ex.ToString)
        End Try
        Declarations.MyEmployee = "---"

        '---�������� ���������� �������� �� ������������ ���������
        MySQLStr = "SELECT Value "
        MySQLStr = MySQLStr & "FROM tbl_WHWorkplace_Config "
        MySQLStr = MySQLStr & "WHERE (UserLogin = N'" & Declarations.UserName & "') "
        MySQLStr = MySQLStr & "AND (Parameter = N'���������� ��������') "
        InitMyConn(False)
        InitMyRec(False, MySQLStr)
        If Declarations.MyRec.EOF = True And Declarations.MyRec.BOF = True Then
            ComboBox3.Text = "���"
            trycloseMyRec()
            Declarations.MyActivity = "0"
        Else
            If Declarations.MyRec.Fields("Value").Value = 0 Then '---���
                ComboBox3.Text = "���"
                Declarations.MyActivity = "0"
            Else                '---������ ��������
                ComboBox3.Text = "������ ��������"
                Declarations.MyActivity = "1"
            End If
            trycloseMyRec()
        End If


        Application.DoEvents()
        Windows.Forms.Cursor.Current = Cursors.WaitCursor
        '---�������� ������
        ShDataLoading()
        Application.DoEvents()
        Windows.Forms.Cursor.Current = Cursors.WaitCursor
        AvlForShDataLoading()
        Application.DoEvents()
        Windows.Forms.Cursor.Current = Cursors.WaitCursor
        '---�������� ��������� ������
        Check1LeftButtonsState()
        Application.DoEvents()
        Windows.Forms.Cursor.Current = Cursors.WaitCursor
        Check1RightButtonsState()
        Application.DoEvents()
        Windows.Forms.Cursor.Current = Cursors.Default
        LoadFlag = 0
    End Sub

    Public Sub ShDataLoading()
        '/////////////////////////////////////////////////////////////////////////////////////
        '//
        '// �������� ������� �� �������� ������� �� ��������
        '// 
        '/////////////////////////////////////////////////////////////////////////////////////
        Dim MySQLStr As String
        Dim MyAdapter As SqlClient.SqlDataAdapter
        Dim MyDs As New DataSet

        MySQLStr = "Exec spp_WHWorkplace_CollectionInfo N'" & Trim(Declarations.MyWH) & "', N'" & Declarations.MyEmployee & "', " & Declarations.MyActivity
        Try
            MyAdapter = New SqlClient.SqlDataAdapter(MySQLStr, Declarations.MyNETConnStr)
            MyAdapter.SelectCommand.CommandTimeout = 600
            MyAdapter.Fill(MyDs)
            DataGridView1.DataSource = MyDs.Tables(0)
        Catch ex As Exception
            MsgBox(ex.ToString)
        End Try

        '---���������
        DataGridView1.Columns(0).HeaderText = "ID ������ ��������"
        DataGridView1.Columns(0).Width = 80
        DataGridView1.Columns(0).Visible = False
        DataGridView1.Columns(1).HeaderText = "���������"
        DataGridView1.Columns(1).Width = 100
        DataGridView1.Columns(2).HeaderText = "N ������"
        DataGridView1.Columns(2).Width = 80
        DataGridView1.Columns(3).HeaderText = "��������"
        DataGridView1.Columns(3).Width = 80
        DataGridView1.Columns(4).HeaderText = "����������� ����"
        DataGridView1.Columns(4).Width = 80
        DataGridView1.Columns(5).HeaderText = "������"
        DataGridView1.Columns(5).Width = 60
        DataGridView1.Columns(6).HeaderText = "������"
        DataGridView1.Columns(6).Width = 180
        DataGridView1.Columns(7).HeaderText = "����� ��������"
        DataGridView1.Columns(7).Width = 200
        DataGridView1.Columns(8).HeaderText = "��������"
        DataGridView1.Columns(8).Width = 100

        FormatDataGridView1()
    End Sub

    Private Sub FormatDataGridView1()
        '////////////////////////////////////////////////////////////////////////////////
        '//
        '// ��������� ���������� �� ���������
        '//
        '////////////////////////////////////////////////////////////////////////////////
        Dim i As Integer

        For i = 0 To DataGridView1.Rows.Count - 1
            If DataGridView1.Rows(i).Cells(5).Value = "" Then
                DataGridView1.Rows(i).DefaultCellStyle.BackColor = Color.White
            Else
                DataGridView1.Rows(i).DefaultCellStyle.BackColor = Color.LightGray
            End If
        Next
    End Sub

    Public Sub AvlForShDataLoading()
        '/////////////////////////////////////////////////////////////////////////////////////
        '//
        '// �������� ������� �� ��������, ��������� ��� ��������
        '// 
        '/////////////////////////////////////////////////////////////////////////////////////
        Dim MySQLStr As String
        Dim MyAdapter As SqlClient.SqlDataAdapter
        Dim MyDs As New DataSet

        MySQLStr = "Exec spp_WHWorkplace_AvlCollectionInfo N'" & Trim(Declarations.MyWH) & "'"
        Try
            MyAdapter = New SqlClient.SqlDataAdapter(MySQLStr, Declarations.MyNETConnStr)
            MyAdapter.SelectCommand.CommandTimeout = 600
            MyAdapter.Fill(MyDs)
            DataGridView2.DataSource = MyDs.Tables(0)
        Catch ex As Exception
            MsgBox(ex.ToString)
        End Try

        '---���������
        DataGridView2.Columns(0).HeaderText = "ID ������ ��������"
        DataGridView2.Columns(0).Width = 80
        DataGridView2.Columns(0).Visible = False
        DataGridView2.Columns(1).HeaderText = "N ������"
        DataGridView2.Columns(1).Width = 80
        DataGridView2.Columns(2).HeaderText = "��������"
        DataGridView2.Columns(2).Width = 80
        DataGridView2.Columns(3).HeaderText = "����������� ����"
        DataGridView2.Columns(3).Width = 80
        DataGridView2.Columns(4).HeaderText = "��� �������"
        DataGridView2.Columns(4).Width = 80
        DataGridView2.Columns(5).HeaderText = "������"
        DataGridView2.Columns(5).Width = 200
        DataGridView2.Columns(6).HeaderText = "����� ��������"
        DataGridView2.Columns(6).Width = 200
        DataGridView2.Columns(7).HeaderText = "��������"
        DataGridView2.Columns(7).Width = 100

    End Sub

    Public Sub Check1LeftButtonsState()
        '/////////////////////////////////////////////////////////////////////////////////////
        '//
        '// �������� ��������� ������ ������� �� �������� ������� �� ��������
        '// 
        '/////////////////////////////////////////////////////////////////////////////////////

        If DataGridView1.SelectedRows.Count = 0 Then         '----���  �������
            Button3.Enabled = False
            Button4.Enabled = False
        Else
            If DataGridView1.SelectedRows.Item(0).Cells(5).Value = "" Then
                Button3.Enabled = True
                Button4.Enabled = False
            Else
                Button3.Enabled = False
                Button4.Enabled = True
            End If
        End If
    End Sub

    Public Sub Check1RightButtonsState()
        '/////////////////////////////////////////////////////////////////////////////////////
        '//
        '// �������� ��������� ������ ������� �� ��������, ��������� ��� ��������
        '// 
        '/////////////////////////////////////////////////////////////////////////////////////

        If DataGridView1.SelectedRows.Count = 0 Then         '----���  �������
            Button11.Enabled = False
        Else
            If DataGridView1.SelectedRows.Item(0).Cells(5).Value = "" Then
                Button11.Enabled = True
            Else
                Button11.Enabled = False
            End If
        End If

        If DataGridView2.Rows.Count = 0 Then         '----���  �������
            Button9.Enabled = False
        Else
            If ComboBox2.SelectedValue = "---" Then
                Button9.Enabled = False
            Else
                Button9.Enabled = True
            End If
        End If
    End Sub

    Private Sub ComboBox1_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles ComboBox1.SelectedIndexChanged
        '////////////////////////////////////////////////////////////////////////////////
        '//
        '// �������� ������ � ������������ � ��������� �������
        '//
        '////////////////////////////////////////////////////////////////////////////////
        Dim MySQLStr As String                        '������� ������
        Dim MyAdapter1 As SqlClient.SqlDataAdapter    '��� ������ ����������� �� 1 ��������
        Dim MyDs1 As New DataSet                      '

        If LoadFlag = 0 Then
            Declarations.MyWH = Trim(ComboBox1.SelectedValue)
            SaveWHChoice()
            Windows.Forms.Cursor.Current = Cursors.WaitCursor

            '---������ �����������
            MySQLStr = "SELECT '---' AS UserName, ' ���' AS FullName "
            MySQLStr = MySQLStr & "UNION ALL "
            MySQLStr = MySQLStr & "SELECT ScalaSystemDB.dbo.ScaUsers.UserName, ScalaSystemDB.dbo.ScaUsers.FullName "
            MySQLStr = MySQLStr & "FROM tbl_WHWorkplace_Employees INNER JOIN "
            MySQLStr = MySQLStr & "ScalaSystemDB.dbo.ScaUsers ON tbl_WHWorkplace_Employees.WHEmployee = ScalaSystemDB.dbo.ScaUsers.UserName "
            MySQLStr = MySQLStr & "WHERE (tbl_WHWorkplace_Employees.WHCode = N'" & Trim(ComboBox1.SelectedValue) & "') "
            MySQLStr = MySQLStr & "ORDER BY FullName "
            Try
                MyAdapter1 = New SqlClient.SqlDataAdapter(MySQLStr, Declarations.MyNETConnStr)
                MyAdapter1.SelectCommand.CommandTimeout = 600
                MyAdapter1.Fill(MyDs1)
                ComboBox2.DisplayMember = "FullName" '��� �� ��� ����� ������������
                ComboBox2.ValueMember = "UserName"   '��� �� ��� ����� ���������
                ComboBox2.DataSource = MyDs1.Tables(0).DefaultView
            Catch ex As Exception
                MsgBox(ex.ToString)
            End Try
            '---�������� ������
            ShDataLoading()
            AvlForShDataLoading()
            '---�������� ��������� ������
            Check1LeftButtonsState()
            Check1RightButtonsState()
            Windows.Forms.Cursor.Current = Cursors.Default
        End If
    End Sub

    Private Sub ComboBox2_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles ComboBox2.SelectedIndexChanged
        '////////////////////////////////////////////////////////////////////////////////
        '//
        '// �������� ������ � ������������ � ��������� �����������
        '//
        '////////////////////////////////////////////////////////////////////////////////

        If LoadFlag = 0 Then
            Declarations.MyEmployee = Trim(ComboBox2.SelectedValue)
            Windows.Forms.Cursor.Current = Cursors.WaitCursor
            '---�������� ������
            ShDataLoading()
            AvlForShDataLoading()
            '---�������� ��������� ������
            Check1LeftButtonsState()
            Check1RightButtonsState()
            Windows.Forms.Cursor.Current = Cursors.Default
        End If
    End Sub

    Private Sub Button5_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button5.Click
        '////////////////////////////////////////////////////////////////////////////////
        '//
        '// ���������� ������
        '//
        '////////////////////////////////////////////////////////////////////////////////

        Windows.Forms.Cursor.Current = Cursors.WaitCursor
        '---�������� ������
        ShDataLoading()
        AvlForShDataLoading()
        '---�������� ��������� ������
        Check1LeftButtonsState()
        Check1RightButtonsState()
        Windows.Forms.Cursor.Current = Cursors.Default
    End Sub

    Public Sub SaveWHChoice()
        '////////////////////////////////////////////////////////////////////////////////
        '//
        '// ���������� ��������� ����� - ����� ������
        '//
        '////////////////////////////////////////////////////////////////////////////////
        Dim MySQLStr As String

        MySQLStr = "DELETE FROM tbl_WHWorkplace_Config "
        MySQLStr = MySQLStr & "WHERE (UserLogin = N'" & Declarations.UserName & "') "
        MySQLStr = MySQLStr & "AND (Parameter = N'������ �� ������� �����') "
        InitMyConn(False)
        Declarations.MyConn.Execute(MySQLStr)

        MySQLStr = "INSERT INTO tbl_WHWorkplace_Config "
        MySQLStr = MySQLStr & "(UserLogin, Parameter, Value) "
        MySQLStr = MySQLStr & "VALUES (N'" & Declarations.UserName & "', "
        MySQLStr = MySQLStr & "N'������ �� ������� �����', "
        MySQLStr = MySQLStr & "N'" & Declarations.MyWH & "') "

        InitMyConn(False)
        Declarations.MyConn.Execute(MySQLStr)
    End Sub

    Private Sub ComboBox3_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles ComboBox3.SelectedIndexChanged
        '////////////////////////////////////////////////////////////////////////////////
        '//
        '// �������� ������ � ������������ � ��������� ����������� ��������
        '//
        '////////////////////////////////////////////////////////////////////////////////

        If LoadFlag = 0 Then
            If ComboBox3.Text = "���" Then
                Declarations.MyActivity = "0"
            Else
                Declarations.MyActivity = "1"
            End If
            SaveActivityChoice()
            Windows.Forms.Cursor.Current = Cursors.WaitCursor
            '---�������� ������
            ShDataLoading()
            AvlForShDataLoading()
            '---�������� ��������� ������
            Check1LeftButtonsState()
            Check1RightButtonsState()
            Windows.Forms.Cursor.Current = Cursors.Default
        End If
    End Sub

    Public Sub SaveActivityChoice()
        '////////////////////////////////////////////////////////////////////////////////
        '//
        '// ���������� ��������� ����� - �������� ��� �������� ��� ������ ��������
        '//
        '////////////////////////////////////////////////////////////////////////////////
        Dim MySQLStr As String

        MySQLStr = "DELETE FROM tbl_WHWorkplace_Config "
        MySQLStr = MySQLStr & "WHERE (UserLogin = N'" & Declarations.UserName & "') "
        MySQLStr = MySQLStr & "AND (Parameter = N'���������� ��������') "
        InitMyConn(False)
        Declarations.MyConn.Execute(MySQLStr)

        MySQLStr = "INSERT INTO tbl_WHWorkplace_Config "
        MySQLStr = MySQLStr & "(UserLogin, Parameter, Value) "
        MySQLStr = MySQLStr & "VALUES (N'" & Declarations.UserName & "', "
        MySQLStr = MySQLStr & "N'���������� ��������', "
        MySQLStr = MySQLStr & "N'" & Declarations.MyActivity & "') "

        InitMyConn(False)
        Declarations.MyConn.Execute(MySQLStr)
    End Sub

    Private Sub DataGridView1_ColumnHeaderMouseClick(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellMouseEventArgs) Handles DataGridView1.ColumnHeaderMouseClick
        '/////////////////////////////////////////////////////////////////////////////////////
        '//
        '// ������ �� ��������� ������� 
        '// 
        '/////////////////////////////////////////////////////////////////////////////////////

        FormatDataGridView1()
    End Sub

    Private Sub DataGridView1_SelectionChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles DataGridView1.SelectionChanged
        '/////////////////////////////////////////////////////////////////////////////////////
        '//
        '// ��������� ������ �������� 
        '// 
        '/////////////////////////////////////////////////////////////////////////////////////

        If LoadFlag = 0 Then
            Windows.Forms.Cursor.Current = Cursors.WaitCursor
            '---�������� ��������� ������
            Check1LeftButtonsState()
            Check1RightButtonsState()
            Windows.Forms.Cursor.Current = Cursors.Default
        End If
    End Sub

    Private Sub Button9_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button9.Click
        '/////////////////////////////////////////////////////////////////////////////////////
        '//
        '// ���������� ������� �� �������� ���������� ������
        '// 
        '/////////////////////////////////////////////////////////////////////////////////////
        Dim i As Integer

        For i = 0 To DataGridView2.SelectedRows.Count - 1
            AddOrderToCollection(DataGridView2.SelectedRows.Item(i).Cells(0).Value, ComboBox2.SelectedValue)
        Next
        Windows.Forms.Cursor.Current = Cursors.WaitCursor
        '---�������� ������
        ShDataLoading()
        AvlForShDataLoading()
        '---�������� ��������� ������
        Check1LeftButtonsState()
        Check1RightButtonsState()
        Windows.Forms.Cursor.Current = Cursors.Default
    End Sub

    Private Sub AddOrderToCollection(ByVal MyCode As Integer, ByVal UserCode As String)
        '/////////////////////////////////////////////////////////////////////////////////////
        '//
        '// ���������� ������ �� �������� ���������� ������
        '// 
        '/////////////////////////////////////////////////////////////////////////////////////
        Dim MySQLStr As String

        MySQLStr = "INSERT INTO tbl_WHWorkplace_ShipmentsJobs "
        MySQLStr = MySQLStr & "(OrderShipmentID, WHEmployee, IsClosed) "
        MySQLStr = MySQLStr & "VALUES (" & MyCode.ToString & ", "
        MySQLStr = MySQLStr & "N'" & UserCode & "', "
        MySQLStr = MySQLStr & "0)"
        InitMyConn(False)
        Declarations.MyConn.Execute(MySQLStr)
    End Sub

    Private Sub Button11_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button11.Click
        '/////////////////////////////////////////////////////////////////////////////////////
        '//
        '// ������ ���������� ������� �� �������� ���������� ������
        '// 
        '/////////////////////////////////////////////////////////////////////////////////////
        Dim i As Integer

        For i = 0 To DataGridView1.SelectedRows.Count - 1
            RemoveOrderFromCollection(DataGridView1.SelectedRows.Item(i).Cells(0).Value)
        Next
        Windows.Forms.Cursor.Current = Cursors.WaitCursor
        '---�������� ������
        ShDataLoading()
        AvlForShDataLoading()
        '---�������� ��������� ������
        Check1LeftButtonsState()
        Check1RightButtonsState()
        Windows.Forms.Cursor.Current = Cursors.Default
    End Sub

    Private Sub RemoveOrderFromCollection(ByVal MyCode As Integer)
        '/////////////////////////////////////////////////////////////////////////////////////
        '//
        '// ������ ���������� ������ �� �������� ���������� ������
        '// 
        '/////////////////////////////////////////////////////////////////////////////////////
        Dim MySQLStr As String

        MySQLStr = "DELETE FROM tbl_WHWorkplace_ShipmentsJobs "
        MySQLStr = MySQLStr & "WHERE (ID = " & MyCode & ") "

        InitMyConn(False)
        Declarations.MyConn.Execute(MySQLStr)
    End Sub

    Private Sub Button3_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button3.Click
        '/////////////////////////////////////////////////////////////////////////////////////
        '//
        '// �������� ��������
        '// 
        '/////////////////////////////////////////////////////////////////////////////////////
        Dim MySQLStr As String

        MySQLStr = "UPDATE tbl_WHWorkplace_ShipmentsJobs "
        MySQLStr = MySQLStr & "SET IsClosed = 1 "
        MySQLStr = MySQLStr & "WHERE (ID = " & DataGridView1.SelectedRows.Item(0).Cells(0).Value & ") "

        InitMyConn(False)
        Declarations.MyConn.Execute(MySQLStr)

        DataGridView1.SelectedRows.Item(0).Cells(5).Value = "+"
        FormatDataGridView1()
        '---�������� ��������� ������
        Check1LeftButtonsState()
        Check1RightButtonsState()
    End Sub

    Private Sub Button4_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button4.Click
        '/////////////////////////////////////////////////////////////////////////////////////
        '//
        '// �������� ��������
        '// 
        '/////////////////////////////////////////////////////////////////////////////////////
        Dim MySQLStr As String

        MySQLStr = "UPDATE tbl_WHWorkplace_ShipmentsJobs "
        MySQLStr = MySQLStr & "SET IsClosed = 0 "
        MySQLStr = MySQLStr & "WHERE (ID = " & DataGridView1.SelectedRows.Item(0).Cells(0).Value & ") "

        InitMyConn(False)
        Declarations.MyConn.Execute(MySQLStr)

        DataGridView1.SelectedRows.Item(0).Cells(5).Value = ""
        FormatDataGridView1()
        '---�������� ��������� ������
        Check1LeftButtonsState()
        Check1RightButtonsState()
    End Sub
End Class
