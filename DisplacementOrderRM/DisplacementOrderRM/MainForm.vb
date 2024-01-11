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
        '// ����� ���� ������� ������ ������� ��������� 
        '/////////////////////////////////////////////////////////////////////////////////////
        Dim MySQLStr As String                        '������� ������
        Dim MyAdapter As SqlClient.SqlDataAdapter     '��� ������ �������
        Dim MyDs As New DataSet

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
        MySQLStr = "SELECT UserID, FullName "
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
            trycloseMyRec()
        End If

        '---------������������� - ��������� ������� ����� ��������
        CheckRights(Declarations.UserCode, "��������� ���������")


        '---������ �������
        MySQLStr = "SELECT SC23001 AS WHCode, SC23001 + ' ' + SC23002 AS WHName "
        MySQLStr = MySQLStr & "FROM SC230300 WITH(NOLOCK) "
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
        ComboBox2.Text = "������ �������� ������ �������"

        '---�������� ������
        DataLoading()
        '---�������� ��������� ������
        CheckButtonsState()
    End Sub

    Public Function DataLoading()
        '/////////////////////////////////////////////////////////////////////////////////////
        '//
        '// �������� ������ ������� ������� � �����������
        '// 
        '/////////////////////////////////////////////////////////////////////////////////////
        Dim MySQLStr As String
        Dim MyAdapter As SqlClient.SqlDataAdapter     '��� ������ �����������
        Dim MyDs As New DataSet                       '

        If ComboBox2.Text = "������ �������� ������ �������" Then
            MySQLStr = "Exec spp_DisplacementWorkplace_WHToListPrep 1, N'" & Trim(ComboBox1.SelectedValue) & "' "
        Else
            MySQLStr = "Exec spp_DisplacementWorkplace_WHToListPrep 0, N'" & Trim(ComboBox1.SelectedValue) & "' "
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
        DataGridView1.Columns(0).HeaderText = "��� ������"
        DataGridView1.Columns(0).Width = 100
        DataGridView1.Columns(1).HeaderText = "�����"
        DataGridView1.Columns(1).Width = 224
        DataGridView1.Columns(2).HeaderText = "��������� �������� �������"
        DataGridView1.Columns(2).Width = 110
        DataGridView1.Columns(3).HeaderText = "����� �������� � ������"
        DataGridView1.Columns(3).Width = 110
        DataGridView1.Columns(4).HeaderText = "������������� ��������"
        DataGridView1.Columns(4).Width = 110
        DataGridView1.Columns(5).HeaderText = "���������� ��������"
        DataGridView1.Columns(5).Width = 110
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
        '// �������� ������ � ������������ � ��������� ������ - ���������� ��� ������ ��� ������ ��������
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
        '// ��������� ���������� �� �������
        '//
        '////////////////////////////////////////////////////////////////////////////////
        Dim row As DataGridViewRow

        row = DataGridView1.Rows(e.RowIndex)
        row.Cells(0).Style.Font = New Font(DataGridView1.DefaultCellStyle.Font, FontStyle.Bold)
        row.Cells(1).Style.Font = New Font(DataGridView1.DefaultCellStyle.Font, FontStyle.Bold)
        If Trim(row.Cells(2).Value.ToString) <> "0" Then
            row.Cells(2).Style.BackColor = Color.LightCoral
        Else
            row.Cells(2).Style.BackColor = Color.Empty
        End If
        If Trim(row.Cells(3).Value.ToString) <> "0" Then
            row.Cells(3).Style.BackColor = Color.LightGreen
        Else
            row.Cells(3).Style.BackColor = Color.Empty
        End If
        If Trim(row.Cells(4).Value.ToString) <> "0" Then
            row.Cells(4).Style.BackColor = Color.LightCoral
        Else
            row.Cells(4).Style.BackColor = Color.Empty
        End If
        If Trim(row.Cells(5).Value.ToString) <> "0" Then
            row.Cells(5).Style.BackColor = Color.LightCoral
        Else
            row.Cells(5).Style.BackColor = Color.Empty
        End If
    End Sub

    Private Sub Button2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button2.Click
        '////////////////////////////////////////////////////////////////////////////////
        '//
        '// ��������� ���� ������������ ��������
        '//
        '////////////////////////////////////////////////////////////////////////////////

        OpenConsolidationWindow()
    End Sub

    Private Sub DataGridView1_DoubleClick(ByVal sender As Object, ByVal e As System.EventArgs) Handles DataGridView1.DoubleClick
        '////////////////////////////////////////////////////////////////////////////////
        '//
        '// ��������� ���� ������������ ��������
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

        Declarations.WHToCode = Trim(Me.DataGridView1.SelectedRows.Item(0).Cells(0).Value.ToString())
        Declarations.WHFromCode = Trim(Me.ComboBox1.SelectedValue)
        MyConsolidatedOrders = New ConsolidatedOrders
        MyConsolidatedOrders.ShowDialog()
        DataLoading()
        '---������� ������� ������� ���������������
        For i As Integer = 0 To DataGridView1.Rows.Count - 1
            If Trim(DataGridView1.Item(0, i).Value.ToString) = Declarations.WHToCode Then
                DataGridView1.CurrentCell = DataGridView1.Item(0, i)
            End If
        Next
        CheckButtonsState()
    End Sub
End Class
