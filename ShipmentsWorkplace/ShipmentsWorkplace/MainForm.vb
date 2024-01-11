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
        '// ����� ���� ������� ������ ����������� 
        '/////////////////////////////////////////////////////////////////////////////////////
        Dim MySQLStr As String                        '������� ������
        Dim MyAdapter As SqlClient.SqlDataAdapter     '��� ������ �������
        Dim MyDs As New DataSet                       '

        LoadFlag = 1
        '---��������� �������
        Try
            Dim Scala As New SfwIII.Application

            Declarations.CompanyID = Scala.ActiveProcess.CommonVars.CompanyCode
            Declarations.Year = Mid(Scala.ActiveProcess.CommonVars.FiscalYear, 3)
            Declarations.UserCode = Scala.ActiveProcess.CommonVars.UserCode
            'Declarations.UserCode = "galkina"
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
            Declarations.SalesmanName = Declarations.FullName
            trycloseMyRec()
        End If

        '---��� ��������
        MySQLStr = "Select ST01001 "
        MySQLStr = MySQLStr & "FROM ST010300 "
        MySQLStr = MySQLStr & "WHERE (ST01002 = N'" & Trim(Declarations.FullName) & "')"
        InitMyConn(False)
        InitMyRec(False, MySQLStr)
        If Declarations.MyRec.EOF = True And Declarations.MyRec.BOF = True Then
            MsgBox("�� ������ ��� ��������, ��������������� ������ �� ���� � Scala. ���������� � ��������������.", MsgBoxStyle.Critical, "��������!")
            trycloseMyRec()
            Application.Exit()
        Else
            Declarations.SalesmanCode = Declarations.MyRec.Fields("ST01001").Value
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

        '---��������� ������� - ���������
        ComboBox2.Text = "������ �������� ����������"

        '---��������� ������ � ������ ��� �������������
        MySQLStr = "SELECT Value "
        MySQLStr = MySQLStr & "FROM tbl_Shipments_SalesmanWP_Config "
        MySQLStr = MySQLStr & "WHERE (UserId = " & Declarations.UserID & ") "
        MySQLStr = MySQLStr & "AND (Parameter = N'������ � ������ ��� �������������') "
        InitMyConn(False)
        InitMyRec(False, MySQLStr)
        If Declarations.MyRec.EOF = True And Declarations.MyRec.BOF = True Then
            ComboBox3.Text = "�������������"
            trycloseMyRec()
        Else
            If Declarations.MyRec.Fields("Value").Value = "0" Then
                ComboBox3.Text = "�������������"
            Else
                ComboBox3.Text = "� ������"
            End If
            trycloseMyRec()
        End If

        '--------------������ ������������ � ���������� ����������--------------------------
        '-----EMail ��� �����������
        MySQLStr = "SELECT Value "
        MySQLStr = MySQLStr & "FROM tbl_Shipments_SalesmanWP_Config "
        MySQLStr = MySQLStr & "WHERE (UserId = " & Declarations.UserID & ") "
        MySQLStr = MySQLStr & "AND (Parameter = N'EMail ��� �����������') "
        InitMyConn(False)
        InitMyRec(False, MySQLStr)
        If Declarations.MyRec.EOF = True And Declarations.MyRec.BOF = True Then
            MyEmail = 0
            trycloseMyRec()
        Else
            MyEmail = Declarations.MyRec.Fields("Value").Value
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
            MyContact = 0
            trycloseMyRec()
        Else
            MyContact = Declarations.MyRec.Fields("Value").Value
            trycloseMyRec()
        End If


        LoadFlag = 0
        Application.DoEvents()
        Windows.Forms.Cursor.Current = Cursors.WaitCursor
        '---�������� ������
        DataLoading()
        Application.DoEvents()
        Windows.Forms.Cursor.Current = Cursors.WaitCursor
        '---�������� ��������� ������
        CheckButtonsState()
        Application.DoEvents()
        Windows.Forms.Cursor.Current = Cursors.Default
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

        If LoadFlag = 0 Then
            If ComboBox2.Text = "������ �������� ����������" Then
                If ComboBox3.Text = "�������������" Then
                    MySQLStr = "Exec spp_Shipments_SalesmanWP_CommonInfo 1, N'" & Trim(ComboBox1.SelectedValue) & "', N'" & Declarations.SalesmanCode & "', 0, N'', N'' "
                Else
                    MySQLStr = "Exec spp_Shipments_SalesmanWP_CommonInfo 1, N'" & Trim(ComboBox1.SelectedValue) & "', N'" & Declarations.SalesmanCode & "', 1, N'', N'' "
                End If
            Else
                If ComboBox3.Text = "�������������" Then
                    MySQLStr = "Exec spp_Shipments_SalesmanWP_CommonInfo 0, N'" & Trim(ComboBox1.SelectedValue) & "', N'" & Declarations.SalesmanCode & "', 0, N'', N'' "
                Else
                    MySQLStr = "Exec spp_Shipments_SalesmanWP_CommonInfo 0, N'" & Trim(ComboBox1.SelectedValue) & "', N'" & Declarations.SalesmanCode & "', 1, N'', N'' "
                End If
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
            DataGridView1.Columns(0).HeaderText = "��� ���� ������"
            DataGridView1.Columns(0).Width = 90
            DataGridView1.Columns(1).HeaderText = "����������"
            DataGridView1.Columns(1).Width = 210
            DataGridView1.Columns(2).HeaderText = "����� ����������"
            DataGridView1.Columns(2).Width = 361
            DataGridView1.Columns(3).HeaderText = "������� � ��������� � ���. 7 ����"
            DataGridView1.Columns(3).Width = 110
            DataGridView1.Columns(4).HeaderText = "������� � ������������ ���������"
            DataGridView1.Columns(4).Width = 110
            DataGridView1.Columns(5).HeaderText = "�������, � ������� ���� ������� ������ ���� ��������"
            DataGridView1.Columns(5).Width = 110
            DataGridView1.Columns(6).HeaderText = "�������, �� ���������� � ������� 7 ����"
            DataGridView1.Columns(6).Width = 110
            DataGridView1.Columns(7).HeaderText = "�������, �� ����������� � ������� 2 ����"
            DataGridView1.Columns(7).Width = 110
        End If
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

        Windows.Forms.Cursor.Current = Cursors.WaitCursor
        '---�������� ������
        DataLoading()
        Application.DoEvents()
        '---�������� ��������� ������
        CheckButtonsState()
        Application.DoEvents()
        Windows.Forms.Cursor.Current = Cursors.Default
    End Sub

    Private Sub ComboBox1_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles ComboBox1.SelectedIndexChanged
        '////////////////////////////////////////////////////////////////////////////////
        '//
        '// �������� ������ � ������������ � ��������� �������
        '//
        '////////////////////////////////////////////////////////////////////////////////

        Windows.Forms.Cursor.Current = Cursors.WaitCursor
        '---�������� ������
        DataLoading()
        Application.DoEvents()
        Windows.Forms.Cursor.Current = Cursors.WaitCursor
        '---�������� ��������� ������
        CheckButtonsState()
        Application.DoEvents()
        Windows.Forms.Cursor.Current = Cursors.Default
    End Sub

    Private Sub ComboBox2_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles ComboBox2.SelectedIndexChanged
        '////////////////////////////////////////////////////////////////////////////////
        '//
        '// �������� ������ � ������������ � ��������� ������ - ���������� ���� ����������� ��� ������ ��������
        '//
        '////////////////////////////////////////////////////////////////////////////////

        Windows.Forms.Cursor.Current = Cursors.WaitCursor
        '---�������� ������
        DataLoading()
        Application.DoEvents()
        Windows.Forms.Cursor.Current = Cursors.WaitCursor
        '---�������� ��������� ������
        CheckButtonsState()
        Application.DoEvents()
        Windows.Forms.Cursor.Current = Cursors.Default
    End Sub

    Private Sub Button2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button2.Click
        '////////////////////////////////////////////////////////////////////////////////
        '//
        '// ��������� ���� ������������ ��������
        '//
        '////////////////////////////////////////////////////////////////////////////////

        OpenShipmentWindow()
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
        row.Cells(3).Style.Format = "n0"
        row.Cells(3).Style.Font = New Font(DataGridView1.DefaultCellStyle.Font, FontStyle.Bold)
        row.Cells(4).Style.Format = "n0"
        row.Cells(4).Style.Font = New Font(DataGridView1.DefaultCellStyle.Font, FontStyle.Bold)
        row.Cells(5).Style.Format = "n0"
        row.Cells(5).Style.Font = New Font(DataGridView1.DefaultCellStyle.Font, FontStyle.Bold)
        row.Cells(6).Style.Format = "n0"
        row.Cells(6).Style.Font = New Font(DataGridView1.DefaultCellStyle.Font, FontStyle.Bold)
        row.Cells(7).Style.Format = "n0"
        row.Cells(7).Style.Font = New Font(DataGridView1.DefaultCellStyle.Font, FontStyle.Bold)
        If Trim(row.Cells(3).Value.ToString) <> "" Then
            row.Cells(3).Style.BackColor = Color.LightGreen
        Else
            row.Cells(3).Style.BackColor = Color.Empty
        End If
        If Trim(row.Cells(4).Value.ToString) <> "" Then
            row.Cells(4).Style.BackColor = Color.Red
        Else
            row.Cells(4).Style.BackColor = Color.Empty
        End If
        If Trim(row.Cells(5).Value.ToString) <> "" Then
            row.Cells(5).Style.BackColor = Color.Orange
        Else
            row.Cells(5).Style.BackColor = Color.Empty
        End If
        If Trim(row.Cells(6).Value.ToString) <> "" Then
            row.Cells(6).Style.BackColor = Color.Yellow
        Else
            row.Cells(6).Style.BackColor = Color.Empty
        End If
        If Trim(row.Cells(7).Value.ToString) <> "" Then
            row.Cells(7).Style.BackColor = Color.Yellow
        Else
            row.Cells(7).Style.BackColor = Color.Empty
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


    Private Sub DataGridView1_DoubleClick(ByVal sender As Object, ByVal e As System.EventArgs) Handles DataGridView1.DoubleClick
        '////////////////////////////////////////////////////////////////////////////////
        '//
        '// ��������� ���� ������������ ����������������� �������
        '//
        '////////////////////////////////////////////////////////////////////////////////

        OpenShipmentWindow()
    End Sub

    Private Sub OpenShipmentWindow()
        '////////////////////////////////////////////////////////////////////////////////
        '//
        '// ��������� ���� ������������ ��������
        '//
        '////////////////////////////////////////////////////////////////////////////////

        Declarations.MyCustomerCode = Trim(Me.DataGridView1.SelectedRows.Item(0).Cells(0).Value.ToString())
        Declarations.MyWH = Trim(Me.ComboBox1.SelectedValue)
        If ComboBox3.Text = "�������������" Then
            Declarations.MyGroupOrIndividualFlag = 0
        Else
            Declarations.MyGroupOrIndividualFlag = 1
        End If
        MyShipmentsList = New ShipmentsList
        MyShipmentsList.ShowDialog()
        Application.DoEvents()
        Windows.Forms.Cursor.Current = Cursors.WaitCursor
        DataLoading()
        '---������� ������� ������� ���������������
        For i As Integer = 0 To DataGridView1.Rows.Count - 1
            If Trim(DataGridView1.Item(0, i).Value.ToString) = Declarations.MyCustomerCode Then
                DataGridView1.CurrentCell = DataGridView1.Item(0, i)
            End If
        Next
        CheckButtonsState()
    End Sub

    Private Sub ComboBox3_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles ComboBox3.SelectedIndexChanged
        '////////////////////////////////////////////////////////////////////////////////
        '//
        '// �������� ������ � ������������ � ��������� ������ - ������ � ������ ��� �������������
        '//
        '////////////////////////////////////////////////////////////////////////////////

        Windows.Forms.Cursor.Current = Cursors.WaitCursor
        '---���������� ������
        SaveJobTypeChoice()
        Application.DoEvents()
        Windows.Forms.Cursor.Current = Cursors.WaitCursor
        '---�������� ������
        DataLoading()
        Application.DoEvents()
        Windows.Forms.Cursor.Current = Cursors.WaitCursor
        '---�������� ��������� ������
        CheckButtonsState()
        Application.DoEvents()
        Windows.Forms.Cursor.Current = Cursors.Default
    End Sub

    Public Sub SaveJobTypeChoice()
        '////////////////////////////////////////////////////////////////////////////////
        '//
        '// ���������� ��������� ����� - ������ � ������ ��� �������������
        '//
        '////////////////////////////////////////////////////////////////////////////////
        Dim MySQLStr As String

        MySQLStr = "DELETE FROM tbl_Shipments_SalesmanWP_Config "
        MySQLStr = MySQLStr & "WHERE (UserId = " & Declarations.UserID & ") "
        MySQLStr = MySQLStr & "AND (Parameter = N'������ � ������ ��� �������������') "
        InitMyConn(False)
        Declarations.MyConn.Execute(MySQLStr)

        MySQLStr = "INSERT INTO tbl_Shipments_SalesmanWP_Config "
        MySQLStr = MySQLStr & "(UserId, Parameter, Value) "
        MySQLStr = MySQLStr & "VALUES (" & Declarations.UserID & ", "
        MySQLStr = MySQLStr & "N'������ � ������ ��� �������������', "
        If ComboBox3.Text = "�������������" Then
            MySQLStr = MySQLStr & "N'0') "
        Else
            MySQLStr = MySQLStr & "N'1') "
        End If
        InitMyConn(False)
        Declarations.MyConn.Execute(MySQLStr)
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
                If (InStr(UCase(Trim(DataGridView1.Item(0, i).Value.ToString)), UCase(Trim(TextBox2.Text))) <> 0 _
                    Or InStr(UCase(Trim(DataGridView1.Item(1, i).Value.ToString)), UCase(Trim(TextBox2.Text))) <> 0 _
                    Or InStr(UCase(Trim(DataGridView1.Item(2, i).Value.ToString)), UCase(Trim(TextBox2.Text))) <> 0) _
                    And (InStr(UCase(Trim(DataGridView1.Item(0, i).Value.ToString)), UCase(Trim(TextBox3.Text))) <> 0 _
                    Or InStr(UCase(Trim(DataGridView1.Item(1, i).Value.ToString)), UCase(Trim(TextBox3.Text))) <> 0 _
                    Or InStr(UCase(Trim(DataGridView1.Item(2, i).Value.ToString)), UCase(Trim(TextBox3.Text))) <> 0) Then
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
                    If (InStr(UCase(Trim(DataGridView1.Item(0, i).Value.ToString)), UCase(Trim(TextBox2.Text))) <> 0 _
                    Or InStr(UCase(Trim(DataGridView1.Item(1, i).Value.ToString)), UCase(Trim(TextBox2.Text))) <> 0 _
                    Or InStr(UCase(Trim(DataGridView1.Item(2, i).Value.ToString)), UCase(Trim(TextBox2.Text))) <> 0) _
                    And (InStr(UCase(Trim(DataGridView1.Item(0, i).Value.ToString)), UCase(Trim(TextBox3.Text))) <> 0 _
                    Or InStr(UCase(Trim(DataGridView1.Item(1, i).Value.ToString)), UCase(Trim(TextBox3.Text))) <> 0 _
                    Or InStr(UCase(Trim(DataGridView1.Item(2, i).Value.ToString)), UCase(Trim(TextBox3.Text))) <> 0) Then
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
                If (InStr(UCase(Trim(DataGridView1.Item(0, i).Value.ToString)), UCase(Trim(TextBox2.Text))) <> 0 _
                    Or InStr(UCase(Trim(DataGridView1.Item(1, i).Value.ToString)), UCase(Trim(TextBox2.Text))) <> 0 _
                    Or InStr(UCase(Trim(DataGridView1.Item(2, i).Value.ToString)), UCase(Trim(TextBox2.Text))) <> 0) _
                    And (InStr(UCase(Trim(DataGridView1.Item(0, i).Value.ToString)), UCase(Trim(TextBox3.Text))) <> 0 _
                    Or InStr(UCase(Trim(DataGridView1.Item(1, i).Value.ToString)), UCase(Trim(TextBox3.Text))) <> 0 _
                    Or InStr(UCase(Trim(DataGridView1.Item(2, i).Value.ToString)), UCase(Trim(TextBox3.Text))) <> 0) Then
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
            MyCustomerSelectList = New CustomerSelectList
            MyCustomerSelectList.ShowDialog()
        End If
    End Sub

    Private Sub Button4_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button4.Click
        '////////////////////////////////////////////////////////////////////////////////
        '//
        '// �������� ���� ��������
        '//
        '////////////////////////////////////////////////////////////////////////////////

        MyConfiguration = New Configuration
        MyConfiguration.ShowDialog()
    End Sub
End Class