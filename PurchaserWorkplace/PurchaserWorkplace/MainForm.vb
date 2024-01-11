Public Class MainForm

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

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        '////////////////////////////////////////////////////////////////////////////////
        '//
        '// �����
        '//
        '////////////////////////////////////////////////////////////////////////////////

        Application.Exit()
    End Sub

    Private Sub MainForm_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        '/////////////////////////////////////////////////////////////////////////////////////
        '//
        '// ��� ������� ���������� ��������� - ���, ��������, ������������ � �.�.
        '// ����� ���� ������� ������ ����������� 
        '/////////////////////////////////////////////////////////////////////////////////////
        Dim MySQLStr As String                        '������� ������
        Dim MyAdapter As SqlClient.SqlDataAdapter     '��� ������ �������
        Dim MyDs As New DataSet                       '
        Dim MyAdapter1 As SqlClient.SqlDataAdapter     '
        Dim MyDs1 As New DataSet                       '

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
        MySQLStr = MySQLStr & "WHERE (SC23006 = N'1') OR (LEFT(SC23006, 2) = N'TR') "
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

        '---������ ����������
        MySQLStr = "SELECT '---' AS PurchCode, ' ���' AS PurchName "
        MySQLStr = MySQLStr & "UNION "
        MySQLStr = MySQLStr & "SELECT DISTINCT Ltrim(Rtrim(View_17.SYPD001)) AS PurchCode, Ltrim(Rtrim(View_17.SYPD001)) + ' ' + Ltrim(Rtrim(View_17.SYPD003)) AS PurchName "
        MySQLStr = MySQLStr & "FROM (SELECT SYPD001, SYPD003 "
        MySQLStr = MySQLStr & "FROM SYPD0300 "
        MySQLStr = MySQLStr & "WHERE (SYPD002 = N'ENG')) AS View_17 INNER JOIN "
        MySQLStr = MySQLStr & "tbl_SupplierCard0300 ON View_17.SYPD001 = tbl_SupplierCard0300.Purchaser "
        
        Try
            MyAdapter1 = New SqlClient.SqlDataAdapter(MySQLStr, Declarations.MyNETConnStr)
            MyAdapter1.SelectCommand.CommandTimeout = 600
            MyAdapter1.Fill(MyDs1)
            ComboBox3.DisplayMember = "PurchName" '��� �� ��� ����� ������������
            ComboBox3.ValueMember = "PurchCode"   '��� �� ��� ����� ���������
            ComboBox3.DataSource = MyDs1.Tables(0).DefaultView
        Catch ex As Exception
            MsgBox(ex.ToString)
        End Try

        '---��������� ������� - ���������
        ComboBox2.Text = "������ �������� ����������"

        '---�������� ������
        DataLoading()
        '---�������� ��������� ������
        CheckButtonsState()

    End Sub

    Public Function DataLoading()
        '/////////////////////////////////////////////////////////////////////////////////////
        '//
        '// �������� ������ ����������� (� ������������ � �����������)
        '// 
        '/////////////////////////////////////////////////////////////////////////////////////
        Dim MySQLStr As String
        Dim MyAdapter As SqlClient.SqlDataAdapter     '��� ������ �����������
        Dim MyDs As New DataSet                       '

        If ComboBox2.Text = "������ �������� ����������" Then
            MySQLStr = "Exec spp_PurchaseWorkplace_SupplierListPrep 1, N'" & Trim(ComboBox1.SelectedValue) & "', N'', N'', N'" & Trim(ComboBox3.SelectedValue) & "' "
        Else
            MySQLStr = "Exec spp_PurchaseWorkplace_SupplierListPrep 0, N'" & Trim(ComboBox1.SelectedValue) & "', N'', N'', N'" & Trim(ComboBox3.SelectedValue) & "' "
        End If

        Try
            MyAdapter = New SqlClient.SqlDataAdapter(MySQLStr, Declarations.MyNETConnStr)
            MyAdapter.SelectCommand.CommandTimeout = 600
            MyAdapter.Fill(MyDs)
            DataGridView1.DataSource = MyDs.Tables(0)
        Catch ex As Exception
            MsgBox(ex.ToString)
        End Try

        '---���������
        DataGridView1.Columns(0).HeaderText = "��� ������ ����"
        DataGridView1.Columns(0).Width = 100
        DataGridView1.Columns(1).HeaderText = "���������"
        DataGridView1.Columns(1).Width = 204
        DataGridView1.Columns(2).HeaderText = "����� ����������"
        DataGridView1.Columns(2).Width = 364
        DataGridView1.Columns(3).HeaderText = "��������� �������� �������"
        DataGridView1.Columns(3).Width = 90
        DataGridView1.Columns(4).HeaderText = "������� ������ �������"
        DataGridView1.Columns(4).Width = 90
        DataGridView1.Columns(5).HeaderText = "������ ���������� �������"
        DataGridView1.Columns(5).Width = 90
        DataGridView1.Columns(6).HeaderText = "����� ������ �������"
        DataGridView1.Columns(6).Width = 90
        DataGridView1.Columns(7).HeaderText = "������� � ���������� ������"
        DataGridView1.Columns(7).Width = 90
        DataGridView1.Columns(8).HeaderText = "������� � ����������� ������"
        DataGridView1.Columns(8).Width = 90
    End Function

    Public Function CheckButtonsState()
        '////////////////////////////////////////////////////////////////////////////////
        '//
        '// �������� � ����������� ��������� ������
        '//
        '////////////////////////////////////////////////////////////////////////////////

        If DataGridView1.SelectedRows.Count = 0 Then
            Button2.Enabled = False
        Else
            Button2.Enabled = True
        End If
    End Function

    Private Sub Button3_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button3.Click
        '////////////////////////////////////////////////////////////////////////////////
        '//
        '// ���������� ������
        '//
        '////////////////////////////////////////////////////////////////////////////////

        '---�������� ������
        DataLoading()
        '---�������� ��������� ������
        CheckButtonsState()
    End Sub

    Private Sub ComboBox1_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles ComboBox1.SelectedIndexChanged
        '////////////////////////////////////////////////////////////////////////////////
        '//
        '// �������� ������ � ������������ � ��������� �������
        '//
        '////////////////////////////////////////////////////////////////////////////////

        '---�������� ������
        DataLoading()
        '---�������� ��������� ������
        CheckButtonsState()
    End Sub

    Private Sub ComboBox2_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles ComboBox2.SelectedIndexChanged
        '////////////////////////////////////////////////////////////////////////////////
        '//
        '// �������� ������ � ������������ � ��������� ������ - ���������� ���� ����������� ��� ������ ��������
        '//
        '////////////////////////////////////////////////////////////////////////////////

        '---�������� ������
        DataLoading()
        '---�������� ��������� ������
        CheckButtonsState()
    End Sub

    Private Sub DataGridView1_CellFormatting(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellFormattingEventArgs) Handles DataGridView1.CellFormatting
        '////////////////////////////////////////////////////////////////////////////////
        '//
        '// ��������� ���������� �� �����������
        '//
        '////////////////////////////////////////////////////////////////////////////////
        Dim row As DataGridViewRow

        row = DataGridView1.Rows(e.RowIndex)
        row.Cells(0).Style.Font = New Font(DataGridView1.DefaultCellStyle.Font, FontStyle.Bold)
        row.Cells(1).Style.Font = New Font(DataGridView1.DefaultCellStyle.Font, FontStyle.Bold)
        If Trim(row.Cells(3).Value.ToString) <> "0" Then
            row.Cells(3).Style.BackColor = Color.LightGreen
        Else
            row.Cells(3).Style.BackColor = Color.Empty
        End If
        If Trim(row.Cells(4).Value.ToString) <> "0" Then
            row.Cells(4).Style.BackColor = Color.LightGreen
        Else
            row.Cells(4).Style.BackColor = Color.Empty
        End If
        If Trim(row.Cells(5).Value.ToString) <> "0" Then
            row.Cells(5).Style.BackColor = Color.LightGreen
        Else
            row.Cells(5).Style.BackColor = Color.Empty
        End If
        If Trim(row.Cells(6).Value.ToString) <> "0" Then
            row.Cells(6).Style.BackColor = Color.LightCoral
        Else
            row.Cells(6).Style.BackColor = Color.Empty
        End If
        If Trim(row.Cells(7).Value.ToString) <> "0" Then
            row.Cells(7).Style.BackColor = Color.Yellow
        Else
            row.Cells(7).Style.BackColor = Color.Empty
        End If
        If Trim(row.Cells(8).Value.ToString) <> "0" Then
            row.Cells(8).Style.BackColor = Color.Yellow
        Else
            row.Cells(8).Style.BackColor = Color.Empty
        End If
    End Sub

    Private Sub Button8_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button8.Click
        '////////////////////////////////////////////////////////////////////////////////
        '//
        '// ����� ������� ����������� ����������
        '//
        '////////////////////////////////////////////////////////////////////////////////

        If Trim(TextBox2.Text) = "" And Trim(TextBox3.Text) = "" Then
        Else
            Windows.Forms.Cursor.Current = Cursors.WaitCursor
            For i As Integer = 0 To DataGridView1.Rows.Count - 1
                If InStr(UCase(Trim(DataGridView1.Item(0, i).Value.ToString)), UCase(Trim(TextBox2.Text))) <> 0 And InStr(UCase(Trim(DataGridView1.Item(0, i).Value.ToString)), UCase(Trim(TextBox3.Text))) <> 0 _
                    Or InStr(UCase(Trim(DataGridView1.Item(1, i).Value.ToString)), UCase(Trim(TextBox2.Text))) <> 0 And InStr(UCase(Trim(DataGridView1.Item(1, i).Value.ToString)), UCase(Trim(TextBox3.Text))) <> 0 _
                    Or InStr(UCase(Trim(DataGridView1.Item(2, i).Value.ToString)), UCase(Trim(TextBox2.Text))) <> 0 And InStr(UCase(Trim(DataGridView1.Item(2, i).Value.ToString)), UCase(Trim(TextBox3.Text))) <> 0 Then
                    DataGridView1.CurrentCell = DataGridView1.Item(0, i)
                    Windows.Forms.Cursor.Current = Cursors.Default
                    Exit Sub
                End If
            Next
            Windows.Forms.Cursor.Current = Cursors.Default
        End If
    End Sub

    Private Sub Button7_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button7.Click
        '////////////////////////////////////////////////////////////////////////////////
        '//
        '// ����� ���������� ����������� �������
        '//
        '////////////////////////////////////////////////////////////////////////////////
        Dim MyRez As Object

        If Trim(TextBox2.Text) = "" And Trim(TextBox3.Text) = "" Then
        Else
            Windows.Forms.Cursor.Current = Cursors.WaitCursor
            For i As Integer = DataGridView1.CurrentCellAddress.Y + 1 To DataGridView1.Rows.Count
                If i = DataGridView1.Rows.Count Then
                    MyRez = MsgBox("����� ����� �� ����� ������. ������ �������?", MsgBoxStyle.YesNo, "��������!")
                    If MyRez = 6 Then
                        i = 0
                    Else
                        Windows.Forms.Cursor.Current = Cursors.Default
                        Exit Sub
                    End If
                End If
                If DataGridView1.Rows.Count = 0 Then
                Else
                    If InStr(UCase(Trim(DataGridView1.Item(0, i).Value.ToString)), UCase(Trim(TextBox2.Text))) <> 0 And InStr(UCase(Trim(DataGridView1.Item(0, i).Value.ToString)), UCase(Trim(TextBox3.Text))) <> 0 _
                        Or InStr(UCase(Trim(DataGridView1.Item(1, i).Value.ToString)), UCase(Trim(TextBox2.Text))) <> 0 And InStr(UCase(Trim(DataGridView1.Item(1, i).Value.ToString)), UCase(Trim(TextBox3.Text))) <> 0 _
                        Or InStr(UCase(Trim(DataGridView1.Item(2, i).Value.ToString)), UCase(Trim(TextBox2.Text))) <> 0 And InStr(UCase(Trim(DataGridView1.Item(2, i).Value.ToString)), UCase(Trim(TextBox3.Text))) <> 0 Then
                        DataGridView1.CurrentCell = DataGridView1.Item(0, i)
                        Windows.Forms.Cursor.Current = Cursors.Default
                        Exit Sub
                    End If
                End If
            Next i
            Windows.Forms.Cursor.Current = Cursors.Default
        End If
    End Sub

    Private Sub Button6_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button6.Click
        '////////////////////////////////////////////////////////////////////////////////
        '//
        '// ������������� ���� ���������� �� �������� �������
        '//
        '////////////////////////////////////////////////////////////////////////////////

        If Button6.Text = "���������� ���" Then
            Windows.Forms.Cursor.Current = Cursors.WaitCursor
            For i As Integer = 0 To DataGridView1.Rows.Count - 1
                If InStr(UCase(Trim(DataGridView1.Item(0, i).Value.ToString)), UCase(Trim(TextBox2.Text))) <> 0 And InStr(UCase(Trim(DataGridView1.Item(0, i).Value.ToString)), UCase(Trim(TextBox3.Text))) <> 0 _
                        Or InStr(UCase(Trim(DataGridView1.Item(1, i).Value.ToString)), UCase(Trim(TextBox2.Text))) <> 0 And InStr(UCase(Trim(DataGridView1.Item(1, i).Value.ToString)), UCase(Trim(TextBox3.Text))) <> 0 _
                        Or InStr(UCase(Trim(DataGridView1.Item(2, i).Value.ToString)), UCase(Trim(TextBox2.Text))) <> 0 And InStr(UCase(Trim(DataGridView1.Item(2, i).Value.ToString)), UCase(Trim(TextBox3.Text))) <> 0 Then
                    DataGridView1.Rows(i).Cells(0).Style.BackColor = Color.Yellow
                    DataGridView1.Rows(i).Cells(1).Style.BackColor = Color.Yellow
                    DataGridView1.Rows(i).Cells(2).Style.BackColor = Color.Yellow
                Else
                    DataGridView1.Rows(i).Cells(0).Style.BackColor = Color.Empty
                    DataGridView1.Rows(i).Cells(1).Style.BackColor = Color.Empty
                    DataGridView1.Rows(i).Cells(2).Style.BackColor = Color.Empty
                End If
            Next
            Windows.Forms.Cursor.Current = Cursors.Default
            Button6.Text = "����� �����."
        Else
            Windows.Forms.Cursor.Current = Cursors.WaitCursor
            For i As Integer = 0 To DataGridView1.Rows.Count - 1
                DataGridView1.Rows(i).Cells(0).Style.BackColor = Color.Empty
                DataGridView1.Rows(i).Cells(1).Style.BackColor = Color.Empty
                DataGridView1.Rows(i).Cells(2).Style.BackColor = Color.Empty
            Next
            Windows.Forms.Cursor.Current = Cursors.Default
            Button6.Text = "���������� ���"
        End If
    End Sub

    Private Sub DataGridView1_ColumnHeaderMouseClick(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellMouseEventArgs) Handles DataGridView1.ColumnHeaderMouseClick
        '/////////////////////////////////////////////////////////////////////////////////////
        '//
        '// ������ �� ��������� ������� 
        '// 
        '/////////////////////////////////////////////////////////////////////////////////////

        Button6.Text = "���������� ���"
    End Sub

    Private Sub Button9_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button9.Click
        '////////////////////////////////////////////////////////////////////////////////
        '//
        '// ����� ���� ���������� �� �������� ����������� � ��������� ����
        '//
        '////////////////////////////////////////////////////////////////////////////////

        If Trim(TextBox2.Text) = "" And Trim(TextBox3.Text) = "" Then
            MsgBox("���������� ������ �������� ������", MsgBoxStyle.OkOnly, "��������!")
            TextBox2.Select()
        Else
            MySupplierSelectList = New SupplierSelectList
            MySupplierSelectList.ShowDialog()
        End If
    End Sub

    Private Sub Button2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button2.Click
        '////////////////////////////////////////////////////////////////////////////////
        '//
        '// ��������� ���� ������������ ����������������� �������
        '//
        '////////////////////////////////////////////////////////////////////////////////

        OpenConsolidationWindow()
    End Sub

    Private Sub DataGridView1_DoubleClick(ByVal sender As Object, ByVal e As System.EventArgs) Handles DataGridView1.DoubleClick
        '////////////////////////////////////////////////////////////////////////////////
        '//
        '// ��������� ���� ������������ ����������������� �������
        '//
        '////////////////////////////////////////////////////////////////////////////////

        OpenConsolidationWindow()
    End Sub

    Private Sub OpenConsolidationWindow()
        '////////////////////////////////////////////////////////////////////////////////
        '//
        '// ��������� ���� ������������ ����������������� �������
        '//
        '////////////////////////////////////////////////////////////////////////////////

        Declarations.MySupplierCode = Trim(Me.DataGridView1.SelectedRows.Item(0).Cells(0).Value.ToString())
        Declarations.MyWH = Trim(Me.ComboBox1.SelectedValue)
        MyConsolidatedOrders = New ConsolidatedOrders
        MyConsolidatedOrders.ShowDialog()
        DataLoading()
        '---������� ������� ������� ���������������
        For i As Integer = 0 To DataGridView1.Rows.Count - 1
            If Trim(DataGridView1.Item(0, i).Value.ToString) = Declarations.MySupplierCode Then
                DataGridView1.CurrentCell = DataGridView1.Item(0, i)
            End If
        Next
        CheckButtonsState()
    End Sub

    Private Sub ComboBox3_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles ComboBox3.SelectedIndexChanged
        '////////////////////////////////////////////////////////////////////////////////
        '//
        '// �������� ������ � ������������ � ��������� ����������
        '//
        '////////////////////////////////////////////////////////////////////////////////

        '---�������� ������
        DataLoading()
        '---�������� ��������� ������
        CheckButtonsState()
    End Sub
End Class
