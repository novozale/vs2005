Public Class MainForm

    Private Sub MainForm_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        '/////////////////////////////////////////////////////////////////////////////////////
        '//
        '// ��� ������� ���������� ��������� - ���, ��������, ������������ � �.�.
        '// ����� ���� ������� ������ ����������� ������� ������������
        '/////////////////////////////////////////////////////////////////////////////////////
        Dim MySQLStr As String                        '������� ������
        Dim MyAdapter As SqlClient.SqlDataAdapter     '
        Dim MyDs As New DataSet                       '

        '---��������� �������
        Try
            Dim Scala As New SfwIII.Application

            Declarations.CompanyID = Scala.ActiveProcess.CommonVars.CompanyCode
            Declarations.Year = Mid(Scala.ActiveProcess.CommonVars.FiscalYear, 3)
            Declarations.UserCode = Scala.ActiveProcess.CommonVars.UserCode
            'Declarations.UserCode = "pupinina"

            MySQLStr = "SELECT ST010300.ST01001 AS SC, ST010300.ST01002 AS FullName "
            MySQLStr = MySQLStr & "FROM ScalaSystemDB.dbo.ScaUsers WITH (NOLOCK) INNER JOIN "
            MySQLStr = MySQLStr & "ST010300 ON ScalaSystemDB.dbo.ScaUsers.FullName = ST010300.ST01002 "
            MySQLStr = MySQLStr & "WHERE (UPPER(ScalaSystemDB.dbo.ScaUsers.UserName) = UPPER(N'" & Declarations.UserCode & "')) "
            InitMyConn(False)
            InitMyRec(False, MySQLStr)
            If Declarations.MyRec.EOF = True And Declarations.MyRec.BOF = True Then
                MsgBox("�� ������ ��� ��������, ��������������� ������ �� ���� � Scala. ���������� � ��������������.", MsgBoxStyle.Critical, "��������!")
                trycloseMyRec()
                Application.Exit()
            Else
                Declarations.SalesmanCode = Declarations.MyRec.Fields("SC").Value
                Declarations.SalesmanName = Declarations.MyRec.Fields("FullName").Value
                trycloseMyRec()
                'Label1.Text = "������ ������� ��� ���������� �� �������� �������� " & Declarations.SalesmanCode & " " & Declarations.SalesmanName
            End If
        Catch ex As Exception
            MsgBox("��������� ������ ����������� ������ �� ���� Scala", MsgBoxStyle.Critical, "��������!")
            Application.Exit()
        End Try


        '---����� ������ � ����
        MySQLStr = "SELECT OR010300.OR01001 AS OrderN, OR010300.OR01015 AS OrderDate, OR010300.OR01002 AS OrderType, "
        MySQLStr = MySQLStr & "OR010300.OR01019 AS SalesmanCode, "
        MySQLStr = MySQLStr & "ISNULL(ST010300.ST01002, N'����������') AS SalesmanName, OR010300.OR01003 AS CustomerCode, "
        MySQLStr = MySQLStr & "ISNULL(SL010300.SL01002, N'') AS CustomerName, "
        MySQLStr = MySQLStr & "LTRIM(RTRIM(ISNULL(SL010300.SL01003, N' ') + ' ' + ISNULL(SL010300.SL01004, N''))) AS CustomerAddress "
        MySQLStr = MySQLStr & "FROM OR010300 WITH (NOLOCK) LEFT OUTER JOIN "
        MySQLStr = MySQLStr & "ST010300 ON OR010300.OR01019 = ST010300.ST01001 LEFT OUTER JOIN "
        MySQLStr = MySQLStr & "SL010300 ON OR010300.OR01003 = SL010300.SL01001 "
        MySQLStr = MySQLStr & "WHERE (OR010300.OR01002 <> 0) AND (OR010300.OR01008 <> 3) "
        MySQLStr = MySQLStr & "ORDER BY OrderN "

        InitMyConn(False)

        Try
            MyAdapter = New SqlClient.SqlDataAdapter(MySQLStr, Declarations.MyNETConnStr)
            MyAdapter.SelectCommand.CommandTimeout = 600
            MyAdapter.Fill(MyDs)
            DataGridView1.DataSource = MyDs.Tables(0)
        Catch ex As Exception
            MsgBox(ex.ToString)
        End Try

        DataGridView1.Columns(0).HeaderText = "����� ������"
        DataGridView1.Columns(0).Width = 100
        DataGridView1.Columns(1).HeaderText = "���� ��������"
        DataGridView1.Columns(1).Width = 100
        DataGridView1.Columns(2).HeaderText = "��� ������"
        DataGridView1.Columns(2).Width = 80
        DataGridView1.Columns(3).HeaderText = "��� ��������"
        DataGridView1.Columns(3).Width = 80
        DataGridView1.Columns(4).HeaderText = "��������"
        DataGridView1.Columns(4).Width = 120
        DataGridView1.Columns(5).HeaderText = "��� ����������"
        DataGridView1.Columns(5).Width = 120
        DataGridView1.Columns(6).HeaderText = "��� ����������"
        DataGridView1.Columns(6).Width = 200
        DataGridView1.Columns(7).HeaderText = "����� ����������"

        CheckButtons()

        DataGridView1.SelectionMode = DataGridViewSelectionMode.FullRowSelect

    End Sub

    Private Sub CheckButtons()
        '/////////////////////////////////////////////////////////////////////////////////////
        '//
        '// ����������� ��������� ������
        '// 
        '/////////////////////////////////////////////////////////////////////////////////////

        If DataGridView1.SelectedRows.Count = 0 Then
            Button2.Enabled = False
        Else
            Button2.Enabled = True
        End If
    End Sub

    Private Sub DataGridView1_SelectionChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles DataGridView1.SelectionChanged
        '////////////////////////////////////////////////////////////////////////////////
        '//
        '// �������� ��������� ������
        '//
        '////////////////////////////////////////////////////////////////////////////////

        CheckButtons()
    End Sub

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        '////////////////////////////////////////////////////////////////////////////////
        '//
        '// �����
        '//
        '////////////////////////////////////////////////////////////////////////////////

        Application.Exit()
    End Sub

    Private Sub ReloadData()
        '/////////////////////////////////////////////////////////////////////////////////////
        '//
        '// �������� ������ 
        '// 
        '/////////////////////////////////////////////////////////////////////////////////////
        Dim MySQLStr As String                        '������� ������
        Dim MyAdapter As SqlClient.SqlDataAdapter     '
        Dim MyDs As New DataSet                       '

        MySQLStr = "SELECT OR010300.OR01001 AS OrderN, OR010300.OR01015 AS OrderDate, OR010300.OR01002 AS OrderType, "
        MySQLStr = MySQLStr & "OR010300.OR01019 AS SalesmanCode, "
        MySQLStr = MySQLStr & "ISNULL(ST010300.ST01002, N'����������') AS SalesmanName, OR010300.OR01003 AS CustomerCode, "
        MySQLStr = MySQLStr & "ISNULL(SL010300.SL01002, N'') AS CustomerName, "
        MySQLStr = MySQLStr & "LTRIM(RTRIM(ISNULL(SL010300.SL01003, N' ') + ' ' + ISNULL(SL010300.SL01004, N''))) AS CustomerAddress "
        MySQLStr = MySQLStr & "FROM OR010300 WITH (NOLOCK) LEFT OUTER JOIN "
        MySQLStr = MySQLStr & "ST010300 ON OR010300.OR01019 = ST010300.ST01001 LEFT OUTER JOIN "
        MySQLStr = MySQLStr & "SL010300 ON OR010300.OR01003 = SL010300.SL01001 "
        MySQLStr = MySQLStr & "WHERE (OR010300.OR01002 <> 0) AND (OR010300.OR01008 <> 3) "
        MySQLStr = MySQLStr & "ORDER BY OrderN "
        InitMyConn(False)

        Try
            MyAdapter = New SqlClient.SqlDataAdapter(MySQLStr, Declarations.MyNETConnStr)
            MyAdapter.SelectCommand.CommandTimeout = 600
            MyAdapter.Fill(MyDs)
            DataGridView1.DataSource = MyDs.Tables(0)
        Catch ex As Exception
            MsgBox(ex.ToString)
        End Try

    End Sub

    Private Sub Button2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button2.Click
        '/////////////////////////////////////////////////////////////////////////////////////
        '//
        '// ������ ���������� �� �������� 
        '// 
        '/////////////////////////////////////////////////////////////////////////////////////

        Declarations.OrderID = DataGridView1.SelectedRows.Item(0).Cells(0).Value.ToString
        ExecShippingAllovance(Declarations.OrderID)
        ReloadData()
        CheckButtons()
    End Sub

    Private Sub Button3_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button3.Click
        '/////////////////////////////////////////////////////////////////////////////////////
        '//
        '// ����� ���������� ������ � ������ 
        '// 
        '/////////////////////////////////////////////////////////////////////////////////////
        Dim MyOrderNum As String

        MyOrderNum = Microsoft.VisualBasic.Right("0000000000" + Trim(TextBox1.Text), 10)
        If Trim(TextBox1.Text) = "" Then
        Else
            Windows.Forms.Cursor.Current = Cursors.WaitCursor
            For i As Integer = 0 To DataGridView1.Rows.Count - 1
                If InStr(UCase(Trim(DataGridView1.Item(0, i).Value.ToString)), UCase(Trim(MyOrderNum))) <> 0 Then
                    DataGridView1.CurrentCell = DataGridView1.Item(0, i)
                    Windows.Forms.Cursor.Current = Cursors.Default
                    Exit Sub
                End If
            Next
            Windows.Forms.Cursor.Current = Cursors.Default
        End If
    End Sub
End Class
