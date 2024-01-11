Public Class ItemSelectList
    Public MySrcWin As String                         '����, �� �������� �������

    Private Sub Button5_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button5.Click
        '/////////////////////////////////////////////////////////////////////////////////////
        '//
        '// ����� ��� ������ ������
        '// 
        '/////////////////////////////////////////////////////////////////////////////////////

        Me.Close()
    End Sub

    Private Sub ItemSelectList_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles Me.KeyDown
        '/////////////////////////////////////////////////////////////////////////////////////
        '//
        '// ����� ��� ���������� �� Escape
        '// 
        '/////////////////////////////////////////////////////////////////////////////////////

        If e.KeyData = Keys.Escape Then
            Me.Close()
        End If
    End Sub

    Private Sub Button4_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button4.Click
        '/////////////////////////////////////////////////////////////////////////////////////
        '//
        '// ����� � ������� ������
        '// 
        '/////////////////////////////////////////////////////////////////////////////////////

        ItemSelect()
    End Sub

    Private Sub DataGridView1_DoubleClick(ByVal sender As Object, ByVal e As System.EventArgs) Handles DataGridView1.DoubleClick
        '/////////////////////////////////////////////////////////////////////////////////////
        '//
        '// ����� � ������� ������
        '// 
        '/////////////////////////////////////////////////////////////////////////////////////

        ItemSelect()
    End Sub

    Private Sub ItemSelect()
        '/////////////////////////////////////////////////////////////////////////////////////
        '//
        '// ����� � ������� ������
        '// 
        '/////////////////////////////////////////////////////////////////////////////////////

        If MyItemSelectList.MySrcWin = "SELECT" Then
            For i As Integer = 0 To MyItemSelect.DataGridView1.Rows.Count - 1
                If Trim(MyItemSelect.DataGridView1.Item(0, i).Value.ToString) = Trim(Me.DataGridView1.SelectedRows.Item(0).Cells(0).Value.ToString()) Then
                    MyItemSelect.DataGridView1.CurrentCell = MyItemSelect.DataGridView1.Item(2, i)
                    Me.Close()
                    Exit Sub
                End If
            Next
        ElseIf MyItemSelectList.MySrcWin = "" Then

        End If
        Me.Close()
    End Sub

    Private Sub ItemSelectList_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        '/////////////////////////////////////////////////////////////////////////////////////
        '//
        '// ��� ������� ��������� ������ �������
        '// 
        '/////////////////////////////////////////////////////////////////////////////////////
        Dim MySQLStr As String                        '������� ������
        Dim MyAdapter As SqlClient.SqlDataAdapter     '
        Dim MyDs As New DataSet                       '

        If MyItemSelectList.MySrcWin = "SELECT" Then
            MySQLStr = "SELECT SC010300.SC01001 AS ID, "
            MySQLStr = MySQLStr & "SC010300.SC01002 + ' ' + SC010300.SC01003 AS Name, "
            MySQLStr = MySQLStr & "ROUND(ISNULL(t2.SC39005, 0), 2) AS Price, "
            MySQLStr = MySQLStr & "CASE WHEN SC010300.SC01042 <= 0 THEN 0 ELSE SC010300.SC01052 END AS PriCost, "
            MySQLStr = MySQLStr & "SC010300.SC01060 AS SuppID, "
            MySQLStr = MySQLStr & "View_1.txt AS UnitName, "
            MySQLStr = MySQLStr & "SC010300.SC01042 AS TotalQty, "
            MySQLStr = MySQLStr & "SC010300.SC01058 AS SuppCode, "
            MySQLStr = MySQLStr & "ISNULL(PL010300.PL01002, N'') + ' ' + ISNULL(PL010300.PL01003, N'') AS SuppName "
            MySQLStr = MySQLStr & "FROM SC010300 WITH(NOLOCK) INNER JOIN "
            MySQLStr = MySQLStr & "(SELECT 0 AS Expr1, SC09002 AS txt "
            MySQLStr = MySQLStr & "FROM SC090300 WITH(NOLOCK) "
            MySQLStr = MySQLStr & "WHERE (SC09001 = 'RUS') "
            MySQLStr = MySQLStr & "UNION "
            MySQLStr = MySQLStr & "SELECT 1 AS Expr1, SC09003 "
            MySQLStr = MySQLStr & "FROM SC090300 AS SC090300_40 WITH(NOLOCK) "
            MySQLStr = MySQLStr & "WHERE (SC09001 = 'RUS') "
            MySQLStr = MySQLStr & "UNION "
            MySQLStr = MySQLStr & "SELECT 2 AS Expr1, SC09004 "
            MySQLStr = MySQLStr & "FROM SC090300 AS SC090300_39 WITH(NOLOCK) "
            MySQLStr = MySQLStr & "WHERE (SC09001 = 'RUS') "
            MySQLStr = MySQLStr & "UNION "
            MySQLStr = MySQLStr & "SELECT 3 AS Expr1, SC09005 "
            MySQLStr = MySQLStr & "FROM SC090300 AS SC090300_38 WITH(NOLOCK) "
            MySQLStr = MySQLStr & "WHERE (SC09001 = 'RUS') "
            MySQLStr = MySQLStr & "UNION "
            MySQLStr = MySQLStr & "SELECT 4 AS Expr1, SC09006 "
            MySQLStr = MySQLStr & "FROM SC090300 AS SC090300_37 WITH(NOLOCK) "
            MySQLStr = MySQLStr & "WHERE (SC09001 = 'RUS') "
            MySQLStr = MySQLStr & "UNION "
            MySQLStr = MySQLStr & "SELECT 5 AS Expr1, SC09007 "
            MySQLStr = MySQLStr & "FROM SC090300 AS SC090300_36 WITH(NOLOCK) "
            MySQLStr = MySQLStr & "WHERE (SC09001 = 'RUS') "
            MySQLStr = MySQLStr & "UNION "
            MySQLStr = MySQLStr & "SELECT 6 AS Expr1, SC09008 "
            MySQLStr = MySQLStr & "FROM SC090300 AS SC090300_35 WITH(NOLOCK) "
            MySQLStr = MySQLStr & "WHERE (SC09001 = 'RUS') "
            MySQLStr = MySQLStr & "UNION "
            MySQLStr = MySQLStr & "SELECT 7 AS Expr1, SC09009 "
            MySQLStr = MySQLStr & "FROM SC090300 AS SC090300_34 WITH(NOLOCK) "
            MySQLStr = MySQLStr & "WHERE (SC09001 = 'RUS') "
            MySQLStr = MySQLStr & "UNION "
            MySQLStr = MySQLStr & "SELECT 8 AS Expr1, SC09010 "
            MySQLStr = MySQLStr & "FROM SC090300 AS SC090300_33 WITH(NOLOCK) "
            MySQLStr = MySQLStr & "WHERE (SC09001 = 'RUS') "
            MySQLStr = MySQLStr & "UNION "
            MySQLStr = MySQLStr & "SELECT 9 AS Expr1, SC09011 "
            MySQLStr = MySQLStr & "FROM SC090300 AS SC090300_32 WITH(NOLOCK) "
            MySQLStr = MySQLStr & "WHERE (SC09001 = 'RUS') "
            MySQLStr = MySQLStr & "UNION "
            MySQLStr = MySQLStr & "SELECT 10 AS Expr1, SC09012 "
            MySQLStr = MySQLStr & "FROM SC090300 AS SC090300_31 WITH(NOLOCK) "
            MySQLStr = MySQLStr & "WHERE (SC09001 = 'RUS') "
            MySQLStr = MySQLStr & "UNION "
            MySQLStr = MySQLStr & "SELECT 11 AS Expr1, SC09013 "
            MySQLStr = MySQLStr & "FROM SC090300 AS SC090300_30 WITH(NOLOCK) "
            MySQLStr = MySQLStr & "WHERE (SC09001 = 'RUS') "
            MySQLStr = MySQLStr & "UNION "
            MySQLStr = MySQLStr & "SELECT 12 AS Expr1, SC09014 "
            MySQLStr = MySQLStr & "FROM SC090300 AS SC090300_29 WITH(NOLOCK)"
            MySQLStr = MySQLStr & "WHERE (SC09001 = 'RUS') "
            MySQLStr = MySQLStr & "UNION "
            MySQLStr = MySQLStr & "SELECT 13 AS Expr1, SC09015 "
            MySQLStr = MySQLStr & "FROM SC090300 AS SC090300_28 WITH(NOLOCK) "
            MySQLStr = MySQLStr & "WHERE (SC09001 = 'RUS') "
            MySQLStr = MySQLStr & "UNION "
            MySQLStr = MySQLStr & "SELECT 14 AS Expr1, SC09016 "
            MySQLStr = MySQLStr & "FROM SC090300 AS SC090300_27 WITH(NOLOCK) "
            MySQLStr = MySQLStr & "WHERE (SC09001 = 'RUS') "
            MySQLStr = MySQLStr & "UNION "
            MySQLStr = MySQLStr & "SELECT 15 AS Expr1, SC09017 "
            MySQLStr = MySQLStr & "FROM SC090300 AS SC090300_26 WITH(NOLOCK) "
            MySQLStr = MySQLStr & "WHERE (SC09001 = 'RUS') "
            MySQLStr = MySQLStr & "UNION "
            MySQLStr = MySQLStr & "SELECT 16 AS Expr1, SC09018 "
            MySQLStr = MySQLStr & "FROM SC090300 AS SC090300_25 WITH(NOLOCK) "
            MySQLStr = MySQLStr & "WHERE (SC09001 = 'RUS') "
            MySQLStr = MySQLStr & "UNION "
            MySQLStr = MySQLStr & "SELECT 17 AS Expr1, SC09019 "
            MySQLStr = MySQLStr & "FROM SC090300 AS SC090300_24 WITH(NOLOCK) "
            MySQLStr = MySQLStr & "WHERE (SC09001 = 'RUS') "
            MySQLStr = MySQLStr & "UNION "
            MySQLStr = MySQLStr & "SELECT 18 AS Expr1, SC09020 "
            MySQLStr = MySQLStr & "FROM SC090300 AS SC090300_23 WITH(NOLOCK) "
            MySQLStr = MySQLStr & "WHERE (SC09001 = 'RUS') "
            MySQLStr = MySQLStr & "UNION "
            MySQLStr = MySQLStr & "SELECT 19 AS Expr1, SC09021 "
            MySQLStr = MySQLStr & "FROM SC090300 AS SC090300_22 WITH(NOLOCK) "
            MySQLStr = MySQLStr & "WHERE (SC09001 = 'RUS') "
            MySQLStr = MySQLStr & "UNION "
            MySQLStr = MySQLStr & "SELECT 20 AS Expr1, SC09022 "
            MySQLStr = MySQLStr & "FROM SC090300 AS SC090300_21 WITH(NOLOCK) "
            MySQLStr = MySQLStr & "WHERE (SC09001 = 'RUS') "
            MySQLStr = MySQLStr & "UNION "
            MySQLStr = MySQLStr & "SELECT 21 AS Expr1, SC09023 "
            MySQLStr = MySQLStr & "FROM SC090300 AS SC090300_20 WITH(NOLOCK) "
            MySQLStr = MySQLStr & "WHERE (SC09001 = 'RUS') "
            MySQLStr = MySQLStr & "UNION "
            MySQLStr = MySQLStr & "SELECT 22 AS Expr1, SC09024 "
            MySQLStr = MySQLStr & "FROM SC090300 AS SC090300_19 WITH(NOLOCK) "
            MySQLStr = MySQLStr & "WHERE (SC09001 = 'RUS') "
            MySQLStr = MySQLStr & "UNION "
            MySQLStr = MySQLStr & "SELECT 23 AS Expr1, SC09025 "
            MySQLStr = MySQLStr & "FROM SC090300 AS SC090300_18 WITH(NOLOCK) "
            MySQLStr = MySQLStr & "WHERE (SC09001 = 'RUS') "
            MySQLStr = MySQLStr & "UNION "
            MySQLStr = MySQLStr & "SELECT 24 AS Expr1, SC09026 "
            MySQLStr = MySQLStr & "FROM SC090300 AS SC090300_17 WITH(NOLOCK) "
            MySQLStr = MySQLStr & "WHERE (SC09001 = 'RUS') "
            MySQLStr = MySQLStr & "UNION "
            MySQLStr = MySQLStr & "SELECT 25 AS Expr1, SC09027 "
            MySQLStr = MySQLStr & "FROM SC090300 AS SC090300_16 WITH(NOLOCK) "
            MySQLStr = MySQLStr & "WHERE (SC09001 = 'RUS') "
            MySQLStr = MySQLStr & "UNION "
            MySQLStr = MySQLStr & "SELECT 26 AS Expr1, SC09028 "
            MySQLStr = MySQLStr & "FROM SC090300 AS SC090300_15 WITH(NOLOCK) "
            MySQLStr = MySQLStr & "WHERE (SC09001 = 'RUS') "
            MySQLStr = MySQLStr & "UNION "
            MySQLStr = MySQLStr & "SELECT 27 AS Expr1, SC09029 "
            MySQLStr = MySQLStr & "FROM SC090300 AS SC090300_14 WITH(NOLOCK) "
            MySQLStr = MySQLStr & "WHERE (SC09001 = 'RUS') "
            MySQLStr = MySQLStr & "UNION "
            MySQLStr = MySQLStr & "SELECT 28 AS Expr1, SC09030 "
            MySQLStr = MySQLStr & "FROM SC090300 AS SC090300_13 WITH(NOLOCK) "
            MySQLStr = MySQLStr & "WHERE (SC09001 = 'RUS') "
            MySQLStr = MySQLStr & "UNION "
            MySQLStr = MySQLStr & "SELECT 29 AS Expr1, SC09031 "
            MySQLStr = MySQLStr & "FROM SC090300 AS SC090300_12 WITH(NOLOCK) "
            MySQLStr = MySQLStr & "WHERE (SC09001 = 'RUS') "
            MySQLStr = MySQLStr & "UNION "
            MySQLStr = MySQLStr & "SELECT 30 AS Expr1, SC09032 "
            MySQLStr = MySQLStr & "FROM SC090300 AS SC090300_11 WITH(NOLOCK) "
            MySQLStr = MySQLStr & "WHERE (SC09001 = 'RUS') "
            MySQLStr = MySQLStr & "UNION "
            MySQLStr = MySQLStr & "SELECT 31 AS Expr1, SC09033 "
            MySQLStr = MySQLStr & "FROM SC090300 AS SC090300_10 WITH(NOLOCK) "
            MySQLStr = MySQLStr & "WHERE (SC09001 = 'RUS') "
            MySQLStr = MySQLStr & "UNION "
            MySQLStr = MySQLStr & "SELECT 32 AS Expr1, SC09034 "
            MySQLStr = MySQLStr & "FROM SC090300 AS SC090300_9 WITH(NOLOCK) "
            MySQLStr = MySQLStr & "WHERE (SC09001 = 'RUS') "
            MySQLStr = MySQLStr & "UNION "
            MySQLStr = MySQLStr & "SELECT 33 AS Expr1, SC09035 "
            MySQLStr = MySQLStr & "FROM SC090300 AS SC090300_8 WITH(NOLOCK) "
            MySQLStr = MySQLStr & "WHERE (SC09001 = 'RUS') "
            MySQLStr = MySQLStr & "UNION "
            MySQLStr = MySQLStr & "SELECT 34 AS Expr1, SC09036 "
            MySQLStr = MySQLStr & "FROM SC090300 AS SC090300_7 WITH(NOLOCK) "
            MySQLStr = MySQLStr & "WHERE (SC09001 = 'RUS') "
            MySQLStr = MySQLStr & "UNION "
            MySQLStr = MySQLStr & "SELECT 35 AS Expr1, SC09037 "
            MySQLStr = MySQLStr & "FROM SC090300 AS SC090300_6 WITH(NOLOCK) "
            MySQLStr = MySQLStr & "WHERE (SC09001 = 'RUS') "
            MySQLStr = MySQLStr & "UNION "
            MySQLStr = MySQLStr & "SELECT 36 AS Expr1, SC09038 "
            MySQLStr = MySQLStr & "FROM SC090300 AS SC090300_5 WITH(NOLOCK) "
            MySQLStr = MySQLStr & "WHERE (SC09001 = 'RUS') "
            MySQLStr = MySQLStr & "UNION "
            MySQLStr = MySQLStr & "SELECT 37 AS Expr1, SC09039 "
            MySQLStr = MySQLStr & "FROM SC090300 AS SC090300_4 WITH(NOLOCK) "
            MySQLStr = MySQLStr & "WHERE (SC09001 = 'RUS') "
            MySQLStr = MySQLStr & "UNION "
            MySQLStr = MySQLStr & "SELECT 38 AS Expr1, SC09040 "
            MySQLStr = MySQLStr & "FROM SC090300 AS SC090300_3 WITH(NOLOCK) "
            MySQLStr = MySQLStr & "WHERE (SC09001 = 'RUS') "
            MySQLStr = MySQLStr & "UNION "
            MySQLStr = MySQLStr & "SELECT 39 AS Expr1, SC09041 "
            MySQLStr = MySQLStr & "FROM SC090300 AS SC090300_2 WITH(NOLOCK) "
            MySQLStr = MySQLStr & "WHERE (SC09001 = 'RUS') "
            MySQLStr = MySQLStr & "UNION "
            MySQLStr = MySQLStr & "SELECT 40 AS Expr1, SC09042 "
            MySQLStr = MySQLStr & "FROM SC090300 AS SC090300_1 WITH(NOLOCK)) AS View_1 ON "
            MySQLStr = MySQLStr & "SC010300.SC01135 = View_1.Expr1 LEFT OUTER JOIN "
            MySQLStr = MySQLStr & "PL010300 ON SC010300.SC01058 = PL010300.PL01001 LEFT OUTER JOIN "
            MySQLStr = MySQLStr & "(SELECT SC39001, SC39005 "
            MySQLStr = MySQLStr & "FROM SC390300 WITH(NOLOCK) "
            MySQLStr = MySQLStr & "WHERE (SC39002 = N'00')) AS t2 ON "
            MySQLStr = MySQLStr & "SC010300.SC01001 = t2.SC39001 "
            MySQLStr = MySQLStr & "WHERE (SC010300.SC01001 <> N'00000000') AND "
            MySQLStr = MySQLStr & "(LTRIM(RTRIM(SC010300.SC01066)) <> N'8') "
            If Trim(MyItemSelect.TextBox1.Text) = "" Then
            Else
                MySQLStr = MySQLStr & "AND (SC010300.SC01058 = N'" & Trim(MyItemSelect.TextBox1.Text) & "') "
            End If

            If Trim(MyItemSelect.TextBox2.Text) = "" Then
                '----� ������ ���� ������� �� ������� - �������, ��� �� ������ �������
                MySQLStr = MySQLStr & "AND ((UPPER(SC010300.SC01001) LIKE N'%" & UCase(MyItemSelect.TextBox3.Text) & "%') "
                MySQLStr = MySQLStr & "OR (UPPER(SC010300.SC01002 + ' ' + SC010300.SC01003) LIKE N'%" & UCase(MyItemSelect.TextBox3.Text) & "%') "
                MySQLStr = MySQLStr & "OR (UPPER(SC010300.SC01060) LIKE N'%" & UCase(MyItemSelect.TextBox3.Text) & "%')) "
            Else
                '----� ������ ���� ������� �������
                If Trim(MyItemSelect.TextBox3.Text) = "" Then
                    '----�� ������ ���� ������� �� �������
                    MySQLStr = MySQLStr & "AND ((UPPER(SC010300.SC01001) LIKE N'%" & UCase(MyItemSelect.TextBox2.Text) & "%') "
                    MySQLStr = MySQLStr & "OR (UPPER(SC010300.SC01002 + ' ' + SC010300.SC01003) LIKE N'%" & UCase(MyItemSelect.TextBox2.Text) & "%') "
                    MySQLStr = MySQLStr & "OR (UPPER(SC010300.SC01060) LIKE N'%" & UCase(MyItemSelect.TextBox2.Text) & "%')) "
                Else
                    '----������� ������� � ��� ����
                    MySQLStr = MySQLStr & "AND (((UPPER(SC010300.SC01001) LIKE N'%" & UCase(MyItemSelect.TextBox2.Text) & "%') "
                    MySQLStr = MySQLStr & "AND (UPPER(SC010300.SC01001) LIKE N'%" & UCase(MyItemSelect.TextBox3.Text) & "%')) "
                    MySQLStr = MySQLStr & "OR ((UPPER(SC010300.SC01002 + ' ' + SC010300.SC01003) LIKE N'%" & UCase(MyItemSelect.TextBox2.Text) & "%') "
                    MySQLStr = MySQLStr & "AND (UPPER(SC010300.SC01002 + ' ' + SC010300.SC01003) LIKE N'%" & UCase(MyItemSelect.TextBox3.Text) & "%')) "
                    MySQLStr = MySQLStr & "OR ((UPPER(SC010300.SC01060) LIKE N'%" & UCase(MyItemSelect.TextBox2.Text) & "%') "
                    MySQLStr = MySQLStr & "AND (UPPER(SC010300.SC01060) LIKE N'%" & UCase(MyItemSelect.TextBox3.Text) & "%'))) "
                End If
            End If
            MySQLStr = MySQLStr & "ORDER BY dbo.SC010300.SC01001  "

        ElseIf MyItemSelectList.MySrcWin = "" Then
        End If

        Try
            MyAdapter = New SqlClient.SqlDataAdapter(MySQLStr, Declarations.MyNETConnStr)
            MyAdapter.SelectCommand.CommandTimeout = 600
            MyAdapter.Fill(MyDs)
            DataGridView1.DataSource = MyDs.Tables(0)
        Catch ex As Exception
            MsgBox(ex.ToString)
        End Try

        DataGridView1.Columns(0).HeaderText = "��� Scala"
        DataGridView1.Columns(0).Width = 110
        DataGridView1.Columns(1).HeaderText = "��� ��������"
        DataGridView1.Columns(1).Width = 300
        DataGridView1.Columns(2).HeaderText = "�����"
        DataGridView1.Columns(2).Width = 80
        DataGridView1.Columns(3).HeaderText = "������� ������"
        DataGridView1.Columns(3).Width = 80
        DataGridView1.Columns(4).HeaderText = "��� ������"
        DataGridView1.Columns(4).Width = 110
        DataGridView1.Columns(5).HeaderText = "�� ���"
        DataGridView1.Columns(5).Width = 40
        If MyItemSelectList.MySrcWin = "SELECT" Then
            DataGridView1.Columns(6).HeaderText = "����� �� �������"
        ElseIf MyItemSelectList.MySrcWin = "" Then
            DataGridView1.Columns(6).HeaderText = "�������� ��� ������ �� ���� �������"
        End If
        DataGridView1.Columns(6).Width = 115
        DataGridView1.Columns(7).HeaderText = "��������� ID"
        DataGridView1.Columns(7).Width = 70
        DataGridView1.Columns(8).HeaderText = "���������"
        DataGridView1.Columns(8).Width = 300
    End Sub
End Class