Public Class MainForm
    Public MyLoadFlag = 1

    Private Sub MainForm_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        '/////////////////////////////////////////////////////////////////////////////////////
        '//
        '// ��� ������� ���������� ��������� - ���, ��������, ������������ � �.�.
        '// ����� ���� ������� ������ ���������� 
        '/////////////////////////////////////////////////////////////////////////////////////
        Dim MySQLStr As String                        '������� ������
        Dim MyAdapter As SqlClient.SqlDataAdapter
        Dim MyDs As DataSet

        '---��������� �������
        Try
            Dim Scala As New SfwIII.Application

            Declarations.CompanyID = Scala.ActiveProcess.CommonVars.CompanyCode
            Declarations.Year = Mid(Scala.ActiveProcess.CommonVars.FiscalYear, 3)
            Declarations.UserCode = Scala.ActiveProcess.CommonVars.UserCode

        Catch
            MsgBox("��������� ������ ����������� ������ �� ���� Scala", MsgBoxStyle.Critical, "��������!")
            Application.Exit()
        End Try

        '---����----------------------------
        MySQLStr = "SELECT CASE WHEN RIGHT(name, 2) < '50' THEN '20' + RIGHT(name, 2) ELSE '19' + RIGHT(name, 2) END AS MyYear "
        MySQLStr = MySQLStr & "FROM sys.sysobjects  WITH(NOLOCK) "
        MySQLStr = MySQLStr & "WHERE (xtype = 'U') AND (name LIKE N'GL0603%') "
        MySQLStr = MySQLStr & "AND (CASE WHEN RIGHT(name, 2) < '50' THEN '20' + RIGHT(name, 2) ELSE '19' + RIGHT(name, 2) END > 2006) "
        MySQLStr = MySQLStr & "ORDER BY MyYear "
        InitMyConn(False)
        Try
            MyAdapter = New SqlClient.SqlDataAdapter(MySQLStr, Declarations.MyNETConnStr)
            MyAdapter.SelectCommand.CommandTimeout = 600
            MyDs = New DataSet
            MyAdapter.Fill(MyDs)
            ComboBox1.DisplayMember = "MyYear" '��� �� ��� ����� ������������
            ComboBox1.ValueMember = "MyYear"   '��� �� ��� ����� ���������
            ComboBox1.DataSource = MyDs.Tables(0).DefaultView
        Catch ex As Exception
            MsgBox(ex.ToString)
        End Try

        '---��������-------------------------
        MySQLStr = "SELECT ST01001 AS Code, ST01001 + ' ' + ST01002 AS Name "
        MySQLStr = MySQLStr & "FROM ST010300  WITH(NOLOCK)"
        MySQLStr = MySQLStr & "ORDER BY ST01002 "
        Try
            MyAdapter = New SqlClient.SqlDataAdapter(MySQLStr, Declarations.MyNETConnStr)
            MyAdapter.SelectCommand.CommandTimeout = 600
            MyDs = New DataSet
            MyAdapter.Fill(MyDs)
            ComboBox2.DisplayMember = "Name" '��� �� ��� ����� ������������
            ComboBox2.ValueMember = "Code"   '��� �� ��� ����� ���������
            ComboBox2.DataSource = MyDs.Tables(0).DefaultView
        Catch ex As Exception
            MsgBox(ex.ToString)
        End Try
        MyLoadFlag = 0
    End Sub

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        '/////////////////////////////////////////////////////////////////////////////////////
        '//
        '// �������� ������ � Excel
        '// 
        '/////////////////////////////////////////////////////////////////////////////////////

        If My.Settings.UseOffice = "LibreOffice" Then
            UploadToLO()
        Else
            UploadToExcel()
        End If

    End Sub

    Private Sub RadioButton1_CheckedChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles RadioButton1.CheckedChanged
        '/////////////////////////////////////////////////////////////////////////////////////
        '//
        '// ����� ���� ��������������
        '// 
        '/////////////////////////////////////////////////////////////////////////////////////

        ChangeCombobox2Data()
    End Sub

    Private Sub ChangeCombobox2Data()
        '/////////////////////////////////////////////////////////////////////////////////////
        '//
        '// ����� ���� �������������� - ����� ������
        '// 
        '/////////////////////////////////////////////////////////////////////////////////////
        Dim MySQLStr As String                        '������� ������
        Dim MyAdapter As SqlClient.SqlDataAdapter
        Dim MyDs As DataSet

        If MyLoadFlag = 0 Then
            If RadioButton1.Checked = True Then     '---�������������� �� ���������
                Label1.Text = "��������"
                MySQLStr = "SELECT ST01001 AS Code, ST01001 + ' ' + ST01002 AS Name "
                MySQLStr = MySQLStr & "FROM ST010300  WITH(NOLOCK)"
                MySQLStr = MySQLStr & "ORDER BY ST01002 "
                Try
                    MyAdapter = New SqlClient.SqlDataAdapter(MySQLStr, Declarations.MyNETConnStr)
                    MyAdapter.SelectCommand.CommandTimeout = 600
                    MyDs = New DataSet
                    MyAdapter.Fill(MyDs)
                    ComboBox2.DisplayMember = "Name" '��� �� ��� ����� ������������
                    ComboBox2.ValueMember = "Code"   '��� �� ��� ����� ���������
                    ComboBox2.DataSource = MyDs.Tables(0).DefaultView
                Catch ex As Exception
                    MsgBox(ex.ToString)
                End Try
            Else                                    '---�������������� �� ���� �������
                Label1.Text = "���� �����"
                MySQLStr = "SELECT DISTINCT View_5.GL03002 AS Code, View_5.GL03002 + ' ' + View_5.GL03003 AS Name "
                MySQLStr = MySQLStr & "FROM ST010300 INNER JOIN "
                MySQLStr = MySQLStr & "(SELECT GL03002, GL03003 "
                MySQLStr = MySQLStr & "FROM GL0303" & Microsoft.VisualBasic.Right(ComboBox1.SelectedValue.ToString, 2) & " "
                MySQLStr = MySQLStr & "WHERE (GL03001 = N'B')) AS View_5 ON SUBSTRING(ST010300.ST01021, 7, 3) = View_5.GL03002 "
                MySQLStr = MySQLStr & "Order By View_5.GL03002 "
                Try
                    MyAdapter = New SqlClient.SqlDataAdapter(MySQLStr, Declarations.MyNETConnStr)
                    MyAdapter.SelectCommand.CommandTimeout = 600
                    MyDs = New DataSet
                    MyAdapter.Fill(MyDs)
                    ComboBox2.DisplayMember = "Name" '��� �� ��� ����� ������������
                    ComboBox2.ValueMember = "Code"   '��� �� ��� ����� ���������
                    ComboBox2.DataSource = MyDs.Tables(0).DefaultView
                Catch ex As Exception
                    MsgBox(ex.ToString)
                End Try
            End If
        End If
    End Sub

    Private Sub ComboBox1_Validated(ByVal sender As Object, ByVal e As System.EventArgs) Handles ComboBox1.Validated
        '/////////////////////////////////////////////////////////////////////////////////////
        '//
        '// ����� ���� �������������� - ����� ������
        '// 
        '/////////////////////////////////////////////////////////////////////////////////////
        ChangeCombobox2Data()
    End Sub
End Class
