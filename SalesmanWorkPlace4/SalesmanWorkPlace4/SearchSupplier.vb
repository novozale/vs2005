

Public Class SearchSupplier
    Public LoadFlag As Integer
    Public FullInfoFlag As Integer

    Private Sub Button13_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button13.Click
        '////////////////////////////////////////////////////////////////////////////////
        '//
        '// �����
        '//
        '////////////////////////////////////////////////////////////////////////////////

        Me.Close()
    End Sub

    Private Sub SearchSupplier_FormClosed(ByVal sender As Object, ByVal e As System.Windows.Forms.FormClosedEventArgs) Handles Me.FormClosed
        '////////////////////////////////////////////////////////////////////////////////
        '//
        '// ��������� ���������� ���������� ��� ������ �� �����
        '//
        '////////////////////////////////////////////////////////////////////////////////

        MySearchSupplier = Nothing
    End Sub

    Private Sub SearchSupplier_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles Me.KeyDown
        '////////////////////////////////////////////////////////////////////////////////
        '//
        '// ������ �������� ���� �� ALT+F4
        '//
        '////////////////////////////////////////////////////////////////////////////////

        If e.KeyData = Keys.Alt + Keys.F4 Then
            e.Handled = True
        End If
    End Sub

    Private Sub SearchSupplier_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        '////////////////////////////////////////////////////////////////////////////////
        '//
        '// ��������� ���� ������ �� ����� ����������
        '//
        '////////////////////////////////////////////////////////////////////////////////

        LoadFlag = 1
        FullInfoFlag = 0
        '----------------------------------
        Label9.Text = Declarations.SalesmanCode & " " & Declarations.SalesmanName
        ComboBoxAct.SelectedIndex = 0
        ComboBox1.SelectedIndex = 2
        ComboBox3.Text = "�������������"

        '----------------------------------
        LoadFlag = 0
        Application.DoEvents()
        System.Windows.Forms.Cursor.Current = Cursors.WaitCursor
        LoadRequests()
        LoadItems()
        LoadProposal()
        CheckRequestButtons()
        CheckItemButtons()
        CheckProposalButtons()
        System.Windows.Forms.Cursor.Current = Cursors.Default
    End Sub

    Private Sub LoadRequests()
        '////////////////////////////////////////////////////////////////////////////////
        '//
        '// �������� ������ �� �������� �� �����
        '//
        '////////////////////////////////////////////////////////////////////////////////
        Dim MySQLStr As String
        Dim MyAdapter As SqlClient.SqlDataAdapter     '��� ������ �����������
        Dim MyDs As New DataSet

        If LoadFlag = 0 Then
            If ComboBoxAct.Text = "������ �������� ��� ��������" Then
                MySQLStr = "exec spp_SupplSearch_SearchRequestInfo " + Declarations.SalesmanCode + ", " + "-1"
            ElseIf ComboBoxAct.Text = "������ ��������" Then
                MySQLStr = "exec spp_SupplSearch_SearchRequestInfo " + Declarations.SalesmanCode + ", " + "0"
            ElseIf ComboBoxAct.Text = "����������������" Then
                MySQLStr = "exec spp_SupplSearch_SearchRequestInfo " + Declarations.SalesmanCode + ", " + "4"
            Else
                MySQLStr = "exec spp_SupplSearch_SearchRequestInfo " + Declarations.SalesmanCode + ", " + "1"
            End If

            If ComboBox3.Text = "�������������" Then
                MySQLStr = MySQLStr + ", 0 "
            Else
                MySQLStr = MySQLStr + ", 1 "
            End If

            MySQLStr = MySQLStr + ", N'" + Trim(TextBox1.Text) + "' "

            Try
                MyAdapter = New SqlClient.SqlDataAdapter(MySQLStr, Declarations.MyNETConnStr)
                MyAdapter.SelectCommand.CommandTimeout = 600
                MyAdapter.Fill(MyDs)
                DataGridView1.DataSource = MyDs.Tables(0)
            Catch ex As Exception
                MsgBox(ex.ToString)
            End Try

            '---���������
            DataGridView1.Columns("ID").HeaderText = "N �� �����"
            DataGridView1.Columns("ID").Width = 50
            DataGridView1.Columns("StartDate").HeaderText = "���� �� �����"
            DataGridView1.Columns("StartDate").Width = 100
            DataGridView1.Columns("StartDate").DefaultCellStyle.Format = "dd/MM/yyyy HH:mm"
            DataGridView1.Columns("Customer").HeaderText = "������"
            DataGridView1.Columns("Customer").Width = 150
            DataGridView1.Columns("CustomerContactName").HeaderText = "���������� ����"
            DataGridView1.Columns("CustomerContactName").Width = 150
            DataGridView1.Columns("CustomerPhone").HeaderText = "�������"
            DataGridView1.Columns("CustomerPhone").Width = 150
            DataGridView1.Columns("CustomerEmail").HeaderText = "Email"
            DataGridView1.Columns("CustomerEmail").Width = 150
            DataGridView1.Columns("RequestDate").HeaderText = "���� ������������� ��"
            DataGridView1.Columns("RequestDate").Width = 100
            DataGridView1.Columns("RequestDate").DefaultCellStyle.Format = "dd/MM/yyyy"
            DataGridView1.Columns("Salesman").HeaderText = "��������"
            DataGridView1.Columns("Salesman").Width = 150
            DataGridView1.Columns("SalesStatusID").HeaderText = "ID ������ ��������"
            DataGridView1.Columns("SalesStatusID").Visible = False
            DataGridView1.Columns("SalesStatus").HeaderText = "������ ��������"
            DataGridView1.Columns("SalesStatus").Width = 150
            DataGridView1.Columns("Comments").HeaderText = "����������� ��������"
            DataGridView1.Columns("Comments").Width = 250
            DataGridView1.Columns("Comments").DefaultCellStyle.Font = New Font(DataGridView1.DefaultCellStyle.Font, FontStyle.Bold)
            DataGridView1.Columns("Searcher").HeaderText = "���������"
            DataGridView1.Columns("Searcher").Width = 150
            DataGridView1.Columns("SearchStatusID").HeaderText = "ID ������ ����������"
            DataGridView1.Columns("SearchStatusID").Visible = False
            DataGridView1.Columns("SearchStatus").HeaderText = "������ ����������"
            DataGridView1.Columns("SearchStatus").Width = 150
            DataGridView1.Columns("SearcherComments").HeaderText = "����������� ����������"
            DataGridView1.Columns("SearcherComments").Width = 250
            DataGridView1.Columns("SearcherComments").DefaultCellStyle.Font = New Font(DataGridView1.DefaultCellStyle.Font, FontStyle.Bold)
            DataGridView1.Columns("CCSal").HeaderText = "��������� ��������"
            DataGridView1.Columns("CCSal").Width = 80
            DataGridView1.Columns("CCSearch").HeaderText = "��������� ����������"
            DataGridView1.Columns("CCSearch").Width = 80
            DataGridView1.Columns("CustomerRequestNum").HeaderText = "N ������� �� �������"
            DataGridView1.Columns("CustomerRequestNum").Width = 200
            DataGridView1.Columns("CPNum").HeaderText = "N ������������� �����������"
            DataGridView1.Columns("CPNum").Width = 150
            DataGridView1.Columns("CancelReason").HeaderText = "������� ������"
            DataGridView1.Columns("CancelReason").Width = 200
            DataGridView1.Columns("PaymentTerms").HeaderText = "������� ������ ��������"
            DataGridView1.Columns("PaymentTerms").Width = 200

            FormatDataGridView1()
            ChangeColumnsVisibility()
        End If
    End Sub

    Private Sub FormatDataGridView1()
        '////////////////////////////////////////////////////////////////////////////////
        '//
        '// ��������� ���������� �� �������� �� �����
        '//
        '////////////////////////////////////////////////////////////////////////////////
        Dim i As Integer

        For i = 0 To DataGridView1.Rows.Count - 1
            If DataGridView1.Rows(i).Cells("SalesStatusID").Value = -1 Then
                DataGridView1.Rows(i).DefaultCellStyle.BackColor = Color.LightSkyBlue
                '-----���� ����������� 
                DataGridView1.Rows(i).Cells("Comments").Style.BackColor = Color.FromArgb(236, 244, 250)
                DataGridView1.Rows(i).Cells("SearcherComments").Style.BackColor = Color.FromArgb(236, 244, 250)
            ElseIf DataGridView1.Rows(i).Cells("SalesStatusID").Value = 0 Then
                '-----������� ���������
                If DataGridView1.Rows(i).Cells("SearchStatusID").Value = 0 Then
                    '-----��� �� � ������ �����������
                    DataGridView1.Rows(i).DefaultCellStyle.BackColor = Color.White
                    '-----���� ����������� 
                    DataGridView1.Rows(i).Cells("Comments").Style.BackColor = Color.White
                    DataGridView1.Rows(i).Cells("SearcherComments").Style.BackColor = Color.White
                ElseIf DataGridView1.Rows(i).Cells("SearchStatusID").Value = 1 Then
                    '-----� ������ �����������
                    DataGridView1.Rows(i).DefaultCellStyle.BackColor = Color.FromArgb(246, 255, 140)
                    '-----���� ����������� 
                    DataGridView1.Rows(i).Cells("Comments").Style.BackColor = Color.FromArgb(252, 255, 213)
                    DataGridView1.Rows(i).Cells("SearcherComments").Style.BackColor = Color.FromArgb(252, 255, 213)
                ElseIf DataGridView1.Rows(i).Cells("SearchStatusID").Value = 2 Then
                    '-----��������� ����������� �����������
                    DataGridView1.Rows(i).DefaultCellStyle.BackColor = Color.FromArgb(170, 255, 143)
                    '-----���� ����������� 
                    DataGridView1.Rows(i).Cells("Comments").Style.BackColor = Color.FromArgb(217, 255, 205)
                    DataGridView1.Rows(i).Cells("SearcherComments").Style.BackColor = Color.FromArgb(217, 255, 205)
                ElseIf DataGridView1.Rows(i).Cells("SearchStatusID").Value = 4 Then
                    '-----��������� �������
                    DataGridView1.Rows(i).DefaultCellStyle.BackColor = Color.Red
                    '-----���� ����������� 
                    DataGridView1.Rows(i).Cells("Comments").Style.BackColor = Color.FromArgb(255, 179, 179)
                    DataGridView1.Rows(i).Cells("SearcherComments").Style.BackColor = Color.FromArgb(255, 179, 179)
                End If
                If CDate(DataGridView1.Rows(i).Cells("RequestDate").Value).AddDays(1) < Now() Then
                    '-----����������
                    DataGridView1.Rows(i).Cells("RequestDate").Style.BackColor = Color.Red
                End If
            ElseIf DataGridView1.Rows(i).Cells("SalesStatusID").Value = 1 Then
                '-----�������� ���������� �����������
                If DataGridView1.Rows(i).Cells("SearchStatusID").Value = 2 Then
                    '-----��������� ����������� �����������
                    DataGridView1.Rows(i).DefaultCellStyle.BackColor = Color.Yellow
                    '-----���� ����������� 
                    DataGridView1.Rows(i).Cells("Comments").Style.BackColor = Color.FromArgb(255, 255, 185)
                    DataGridView1.Rows(i).Cells("SearcherComments").Style.BackColor = Color.FromArgb(255, 255, 185)
                ElseIf DataGridView1.Rows(i).Cells("SearchStatusID").Value = 3 Then
                    '-----��������� ������ �����
                    DataGridView1.Rows(i).DefaultCellStyle.BackColor = Color.FromArgb(65, 255, 5)
                    '-----���� ����������� 
                    DataGridView1.Rows(i).Cells("Comments").Style.BackColor = Color.FromArgb(170, 255, 143)
                    DataGridView1.Rows(i).Cells("SearcherComments").Style.BackColor = Color.FromArgb(170, 255, 143)
                ElseIf DataGridView1.Rows(i).Cells("SearchStatusID").Value = 4 Then
                    '-----��������� �������
                    DataGridView1.Rows(i).DefaultCellStyle.BackColor = Color.Red
                    '-----���� ����������� 
                    DataGridView1.Rows(i).Cells("Comments").Style.BackColor = Color.FromArgb(255, 179, 179)
                    DataGridView1.Rows(i).Cells("SearcherComments").Style.BackColor = Color.FromArgb(255, 179, 179)
                End If
                If CDate(DataGridView1.Rows(i).Cells("RequestDate").Value).AddDays(1) < Now() Then
                    '-----����������
                    DataGridView1.Rows(i).Cells("RequestDate").Style.BackColor = Color.Red
                End If
            ElseIf DataGridView1.Rows(i).Cells("SalesStatusID").Value = 2 Then
                '-----�������� �� ���������� �����������
                If DataGridView1.Rows(i).Cells("SearchStatusID").Value = 1 Then
                    '-----� ������ �����������
                    DataGridView1.Rows(i).DefaultCellStyle.BackColor = Color.Orange
                    '-----���� ����������� 
                    DataGridView1.Rows(i).Cells("Comments").Style.BackColor = Color.FromArgb(255, 255, 117)
                    DataGridView1.Rows(i).Cells("SearcherComments").Style.BackColor = Color.FromArgb(255, 255, 117)
                ElseIf DataGridView1.Rows(i).Cells("SearchStatusID").Value = 2 Then
                    '-----��������� ����������� �����������
                    DataGridView1.Rows(i).DefaultCellStyle.BackColor = Color.LimeGreen
                    '-----���� ����������� 
                    DataGridView1.Rows(i).Cells("Comments").Style.BackColor = Color.FromArgb(163, 255, 163)
                    DataGridView1.Rows(i).Cells("SearcherComments").Style.BackColor = Color.FromArgb(163, 255, 163)
                ElseIf DataGridView1.Rows(i).Cells("SearchStatusID").Value = 3 Then
                    '-----��������� ������ �����
                    DataGridView1.Rows(i).DefaultCellStyle.BackColor = Color.Green
                    '-----���� ����������� 
                    DataGridView1.Rows(i).Cells("Comments").Style.BackColor = Color.FromArgb(204, 233, 173)
                    DataGridView1.Rows(i).Cells("SearcherComments").Style.BackColor = Color.FromArgb(204, 233, 173)
                ElseIf DataGridView1.Rows(i).Cells("SearchStatusID").Value = 4 Then
                    '-----��������� �������
                    DataGridView1.Rows(i).DefaultCellStyle.BackColor = Color.Red
                    '-----���� ����������� 
                    DataGridView1.Rows(i).Cells("Comments").Style.BackColor = Color.FromArgb(255, 179, 179)
                    DataGridView1.Rows(i).Cells("SearcherComments").Style.BackColor = Color.FromArgb(255, 179, 179)
                End If
                If CDate(DataGridView1.Rows(i).Cells("RequestDate").Value).AddDays(1) < Now() Then
                    '-----����������
                    DataGridView1.Rows(i).Cells("RequestDate").Style.BackColor = Color.Red
                End If
            ElseIf DataGridView1.Rows(i).Cells("SalesStatusID").Value = 4 Then
                '-----�������� ������������ ������ (�������� �� �����) (4)
                DataGridView1.Rows(i).DefaultCellStyle.BackColor = Color.FromArgb(179, 129, 217)
                '-----���� ����������� 
                DataGridView1.Rows(i).Cells("Comments").Style.BackColor = Color.FromArgb(216, 190, 236)
                DataGridView1.Rows(i).Cells(14).Style.BackColor = Color.FromArgb(216, 190, 236)
            Else
                '-----�������� ��������� ������ ������ (3)
                DataGridView1.Rows(i).DefaultCellStyle.BackColor = Color.LightGray
            End If
            If Not IsDBNull(DataGridView1.Rows(i).Cells("SearcherComments").Value) Then
                DataGridView1.Rows(i).Cells("ID").ToolTipText = DataGridView1.Rows(i).Cells("SearcherComments").Value
                DataGridView1.Rows(i).Cells("StartDate").ToolTipText = DataGridView1.Rows(i).Cells("SearcherComments").Value
                DataGridView1.Rows(i).Cells("Customer").ToolTipText = DataGridView1.Rows(i).Cells("SearcherComments").Value
                DataGridView1.Rows(i).Cells("CustomerContactName").ToolTipText = DataGridView1.Rows(i).Cells("SearcherComments").Value
                DataGridView1.Rows(i).Cells("CustomerPhone").ToolTipText = DataGridView1.Rows(i).Cells("SearcherComments").Value
                DataGridView1.Rows(i).Cells("CustomerEmail").ToolTipText = DataGridView1.Rows(i).Cells("SearcherComments").Value
                DataGridView1.Rows(i).Cells("RequestDate").ToolTipText = DataGridView1.Rows(i).Cells("SearcherComments").Value
                DataGridView1.Rows(i).Cells("Salesman").ToolTipText = DataGridView1.Rows(i).Cells("SearcherComments").Value
                DataGridView1.Rows(i).Cells("SalesStatusID").ToolTipText = DataGridView1.Rows(i).Cells("SearcherComments").Value
                DataGridView1.Rows(i).Cells("SalesStatus").ToolTipText = DataGridView1.Rows(i).Cells("SearcherComments").Value
                DataGridView1.Rows(i).Cells("Searcher").ToolTipText = DataGridView1.Rows(i).Cells("SearcherComments").Value
                DataGridView1.Rows(i).Cells("SearchStatusID").ToolTipText = DataGridView1.Rows(i).Cells("SearcherComments").Value
                DataGridView1.Rows(i).Cells("SearchStatus").ToolTipText = DataGridView1.Rows(i).Cells("SearcherComments").Value
                DataGridView1.Rows(i).Cells("CCSal").ToolTipText = DataGridView1.Rows(i).Cells("SearcherComments").Value
                DataGridView1.Rows(i).Cells("CCSearch").ToolTipText = DataGridView1.Rows(i).Cells("SearcherComments").Value
                DataGridView1.Rows(i).Cells("CustomerRequestNum").ToolTipText = DataGridView1.Rows(i).Cells("SearcherComments").Value
            End If

        Next
    End Sub

    Private Sub CheckRequestButtons()
        '////////////////////////////////////////////////////////////////////////////////
        '//
        '// ��������� ������ ����������� � �������
        '//
        '////////////////////////////////////////////////////////////////////////////////

        If DataGridView1.SelectedRows.Count = 0 Then
            Button10.Enabled = False
            Button7.Enabled = False
            Button4.Enabled = False
            Button6.Enabled = False
            Button17.Enabled = False
            Button18.Enabled = False
            Button12.Enabled = False
            Button14.Enabled = False
            Button19.Enabled = False
            ButtonPause.Enabled = False
            ButtonContinue.Enabled = False
            ButtonPause.BackColor = SystemColors.Control
            ButtonContinue.BackColor = SystemColors.Control
        Else
            Button10.Enabled = True
            Button12.Enabled = True
            Button14.Enabled = True

            If DataGridView1.SelectedRows.Item(0).Cells("SalesStatusID").Value = -1 Then
                Button7.Enabled = True
                Button6.Enabled = True
                Button17.Enabled = False
                Button18.Enabled = False
                Button19.Enabled = True
                Button4.Enabled = False
                ButtonPause.Enabled = False
                ButtonContinue.Enabled = False
                ButtonPause.BackColor = SystemColors.Control
                ButtonContinue.BackColor = SystemColors.Control
            ElseIf DataGridView1.SelectedRows.Item(0).Cells("SalesStatusID").Value = 0 Then
                '-----������� ���������
                ButtonPause.Enabled = True
                ButtonContinue.Enabled = False
                ButtonPause.BackColor = Color.Pink
                ButtonContinue.BackColor = SystemColors.Control
                If DataGridView1.SelectedRows.Item(0).Cells("SearchStatusID").Value = 0 Then
                    '-----��� �� � ������ �����������
                    Button7.Enabled = True
                    Button6.Enabled = True
                    Button17.Enabled = False
                    Button18.Enabled = False
                    Button19.Enabled = False
                    ButtonPause.Enabled = False
                    ButtonContinue.Enabled = False
                    ButtonPause.BackColor = SystemColors.Control
                    ButtonContinue.BackColor = SystemColors.Control
                ElseIf DataGridView1.SelectedRows.Item(0).Cells("SearchStatusID").Value = -1 Then
                    '-----��������� ��������, �� ��� �� ��������� � ������
                    Button7.Enabled = False
                    Button6.Enabled = False
                    Button17.Enabled = False
                    Button18.Enabled = False
                    Button19.Enabled = False
                    ButtonPause.Enabled = False
                    ButtonContinue.Enabled = False
                    ButtonPause.BackColor = SystemColors.Control
                    ButtonContinue.BackColor = SystemColors.Control
                ElseIf DataGridView1.SelectedRows.Item(0).Cells("SearchStatusID").Value = 1 Then
                    '-----� ������ �����������
                    Button7.Enabled = False
                    Button6.Enabled = False
                    Button17.Enabled = False
                    Button18.Enabled = False
                    Button19.Enabled = False
                    ButtonPause.Enabled = False
                    ButtonContinue.Enabled = False
                    ButtonPause.BackColor = SystemColors.Control
                    ButtonContinue.BackColor = SystemColors.Control
                ElseIf DataGridView1.SelectedRows.Item(0).Cells("SearchStatusID").Value = 2 Then
                    '-----��������� ����������� �����������
                    Button7.Enabled = False
                    Button6.Enabled = False
                    Button17.Enabled = True
                    Button18.Enabled = True
                    Button19.Enabled = False
                ElseIf DataGridView1.SelectedRows.Item(0).Cells("SearchStatusID").Value = 4 Then
                    '-----��������� �������
                    Button7.Enabled = False
                    Button6.Enabled = False
                    Button17.Enabled = False
                    Button18.Enabled = False
                    Button19.Enabled = False
                    ButtonPause.Enabled = False
                    ButtonContinue.Enabled = False
                    ButtonPause.BackColor = SystemColors.Control
                    ButtonContinue.BackColor = SystemColors.Control
                End If
                Button4.Enabled = True

            ElseIf DataGridView1.SelectedRows.Item(0).Cells("SalesStatusID").Value = 1 Then
                '-----�������� ���������� �����������
                ButtonPause.Enabled = True
                ButtonContinue.Enabled = False
                ButtonPause.BackColor = Color.Pink
                ButtonContinue.BackColor = SystemColors.Control
                If DataGridView1.SelectedRows.Item(0).Cells("SearchStatusID").Value = 2 Then
                    '-----��������� ����������� �����������
                    Button7.Enabled = False
                    Button6.Enabled = False
                    Button17.Enabled = False
                    Button18.Enabled = False
                    Button19.Enabled = False
                    ButtonPause.Enabled = False
                    ButtonContinue.Enabled = False
                    ButtonPause.BackColor = SystemColors.Control
                    ButtonContinue.BackColor = SystemColors.Control
                ElseIf DataGridView1.SelectedRows.Item(0).Cells("SearchStatusID").Value = 3 Then
                    '-----��������� ������ �����
                    Button7.Enabled = False
                    Button6.Enabled = False
                    Button17.Enabled = False
                    Button18.Enabled = False
                    Button19.Enabled = False
                    ButtonPause.Enabled = False
                    ButtonContinue.Enabled = False
                    ButtonPause.BackColor = SystemColors.Control
                    ButtonContinue.BackColor = SystemColors.Control
                ElseIf DataGridView1.SelectedRows.Item(0).Cells("SearchStatusID").Value = 4 Then
                    '-----��������� �������
                    Button7.Enabled = False
                    Button6.Enabled = False
                    Button17.Enabled = False
                    Button18.Enabled = False
                    Button19.Enabled = False
                    ButtonPause.Enabled = False
                    ButtonContinue.Enabled = False
                    ButtonPause.BackColor = SystemColors.Control
                    ButtonContinue.BackColor = SystemColors.Control
                End If
                Button4.Enabled = True

            ElseIf DataGridView1.SelectedRows.Item(0).Cells("SalesStatusID").Value = 2 Then
                '-----�������� �� ���������� �����������
                ButtonPause.Enabled = True
                ButtonContinue.Enabled = False
                ButtonPause.BackColor = Color.Pink
                ButtonContinue.BackColor = SystemColors.Control
                If DataGridView1.SelectedRows.Item(0).Cells("SearchStatusID").Value = 1 Then
                    '-----� ������ �����������
                    Button7.Enabled = False
                    Button6.Enabled = False
                    Button17.Enabled = False
                    Button18.Enabled = False
                    Button19.Enabled = False
                    ButtonPause.Enabled = False
                    ButtonContinue.Enabled = False
                    ButtonPause.BackColor = SystemColors.Control
                    ButtonContinue.BackColor = SystemColors.Control
                ElseIf DataGridView1.SelectedRows.Item(0).Cells("SearchStatusID").Value = 2 Then
                    '-----��������� ����������� �����������
                    Button7.Enabled = False
                    Button6.Enabled = False
                    Button17.Enabled = True
                    Button18.Enabled = True
                    Button19.Enabled = False
                ElseIf DataGridView1.SelectedRows.Item(0).Cells("SearchStatusID").Value = 4 Then
                    '-----��������� �������
                    Button7.Enabled = False
                    Button6.Enabled = False
                    Button17.Enabled = False
                    Button18.Enabled = False
                    Button19.Enabled = False
                    ButtonPause.Enabled = False
                    ButtonContinue.Enabled = False
                    ButtonPause.BackColor = SystemColors.Control
                    ButtonContinue.BackColor = SystemColors.Control
                End If

                Button4.Enabled = True

            ElseIf DataGridView1.SelectedRows.Item(0).Cells("SalesStatusID").Value = 4 Then
                '-----�������� ������������ ������ (�������� �� �����) (4)
                ButtonPause.Enabled = False
                ButtonContinue.Enabled = True
                ButtonPause.BackColor = SystemColors.Control
                ButtonContinue.BackColor = Color.Cyan

            Else
                '-----�������� ��������� ������ ������ (3)
                Button7.Enabled = False
                Button4.Enabled = False
                Button6.Enabled = False
                Button17.Enabled = False
                Button18.Enabled = False
                Button19.Enabled = False
                ButtonPause.Enabled = False
                ButtonContinue.Enabled = False
            End If

        End If
    End Sub

    Private Sub LoadItems()
        '////////////////////////////////////////////////////////////////////////////////
        '//
        '// �������� ����������� �� �������
        '//
        '////////////////////////////////////////////////////////////////////////////////
        Dim MySQLStr As String
        Dim MyAdapter As SqlClient.SqlDataAdapter     '��� ������ �����������
        Dim MyDs As New DataSet                       '

        If LoadFlag = 0 Then
            If DataGridView1.SelectedRows.Count <> 0 Then
                MySQLStr = "Exec spp_SupplSearch_SearchItemsInfo " & Trim(DataGridView1.SelectedRows.Item(0).Cells("ID").Value.ToString())
            Else
                MySQLStr = "Exec spp_SupplSearch_SearchItemsInfo 0 "
            End If
            Try
                MyAdapter = New SqlClient.SqlDataAdapter(MySQLStr, Declarations.MyNETConnStr)
                MyAdapter.SelectCommand.CommandTimeout = 600
                MyAdapter.Fill(MyDs)
                DataGridView2.DataSource = MyDs.Tables(0)
            Catch ex As Exception
                MsgBox(ex.ToString)
            End Try

            '---���������
            DataGridView2.Columns(0).HeaderText = "ID"
            DataGridView2.Columns(0).Width = 50
            DataGridView2.Columns(0).HeaderCell.Style.BackColor = Color.LightBlue
            DataGridView2.Columns(1).HeaderText = "��� ������ � �����"
            DataGridView2.Columns(1).Width = 100
            DataGridView2.Columns(1).Visible = False
            DataGridView2.Columns(2).HeaderText = "��� ������ �������������"
            DataGridView2.Columns(2).Width = 130
            DataGridView2.Columns(2).HeaderCell.Style.BackColor = Color.LightBlue
            DataGridView2.Columns(3).HeaderText = "�������� ������"
            DataGridView2.Columns(3).Width = 250
            DataGridView2.Columns(3).HeaderCell.Style.BackColor = Color.LightBlue
            DataGridView2.Columns(4).HeaderText = "�� ���� �����"
            DataGridView2.Columns(4).Width = 50
            DataGridView2.Columns(4).HeaderCell.Style.BackColor = Color.LightBlue
            DataGridView2.Columns(5).HeaderText = "���-��"
            DataGridView2.Columns(5).Width = 100
            DataGridView2.Columns(5).HeaderCell.Style.BackColor = Color.LightBlue
            DataGridView2.Columns(5).DefaultCellStyle.Format = "n3"
            DataGridView2.Columns(6).HeaderText = "���� �������� (���)"
            DataGridView2.Columns(6).Width = 100
            DataGridView2.Columns(6).Visible = False
            DataGridView2.Columns(6).DefaultCellStyle.Format = "n2"
            DataGridView2.Columns(7).HeaderText = "�����������"
            DataGridView2.Columns(7).Width = 250
            '
            FormatDataGridView2()
        End If
    End Sub

    Private Sub FormatDataGridView2()
        '////////////////////////////////////////////////////////////////////////////////
        '//
        '// ��������� ���������� �� ������� � ������
        '//
        '////////////////////////////////////////////////////////////////////////////////
        Dim i As Integer

        For i = 0 To DataGridView2.Rows.Count - 1

            If Not IsDBNull(DataGridView2.Rows(i).Cells(7).Value) Then
                DataGridView2.Rows(i).Cells(0).ToolTipText = DataGridView2.Rows(i).Cells(7).Value
                DataGridView2.Rows(i).Cells(1).ToolTipText = DataGridView2.Rows(i).Cells(7).Value
                DataGridView2.Rows(i).Cells(2).ToolTipText = DataGridView2.Rows(i).Cells(7).Value
                DataGridView2.Rows(i).Cells(3).ToolTipText = DataGridView2.Rows(i).Cells(7).Value
                DataGridView2.Rows(i).Cells(4).ToolTipText = DataGridView2.Rows(i).Cells(7).Value
                DataGridView2.Rows(i).Cells(5).ToolTipText = DataGridView2.Rows(i).Cells(7).Value
                DataGridView2.Rows(i).Cells(6).ToolTipText = DataGridView2.Rows(i).Cells(7).Value
            End If

        Next
    End Sub

    Private Sub CheckItemButtons()
        '////////////////////////////////////////////////////////////////////////////////
        '//
        '// ��������� ������ ����������� � ������
        '//
        '////////////////////////////////////////////////////////////////////////////////

        If DataGridView2.SelectedRows.Count = 0 Then
            If DataGridView1.SelectedRows.Count = 0 Then
                Button5.Enabled = False
                Button1.Enabled = False
                Button3.Enabled = False
                Button9.Enabled = False
                Button11.Enabled = False
            Else
                If DataGridView1.SelectedRows.Item(0).Cells("SalesStatusID").Value = -1 Then
                    Button5.Enabled = False
                    Button1.Enabled = True
                    Button3.Enabled = True
                    Button9.Enabled = False
                    Button11.Enabled = False
                Else
                    Button5.Enabled = False
                    Button1.Enabled = False
                    Button3.Enabled = False
                    Button9.Enabled = False
                    Button11.Enabled = False
                End If
            End If
        Else
            If DataGridView1.SelectedRows.Count = 0 Then
                Button5.Enabled = False
                Button1.Enabled = False
                Button3.Enabled = False
                Button9.Enabled = False
                Button11.Enabled = False
            Else
                If DataGridView1.SelectedRows.Item(0).Cells("SalesStatusID").Value = -1 Then
                    Button5.Enabled = True
                    Button1.Enabled = True
                    Button3.Enabled = True
                    Button9.Enabled = True
                    Button11.Enabled = True
                Else
                    Button5.Enabled = True
                    Button1.Enabled = False
                    Button3.Enabled = False
                    Button9.Enabled = False
                    Button11.Enabled = False
                End If
            End If
        End If
    End Sub

    Private Sub LoadProposal()
        '////////////////////////////////////////////////////////////////////////////////
        '//
        '// �������� ����������� �����������
        '//
        '////////////////////////////////////////////////////////////////////////////////
        Dim MySQLStr As String
        Dim MyAdapter As SqlClient.SqlDataAdapter     '��� ������ �����������
        Dim MyDs As New DataSet                       '

        If LoadFlag = 0 Then
            If DataGridView1.SelectedRows.Count <> 0 Then
                If ComboBox1.Text = "������������ ������������" Then
                    MySQLStr = "Exec spp_SupplSearch_GetProposalInfo " & Trim(DataGridView1.SelectedRows.Item(0).Cells("ID").Value.ToString()) & ", 1 "
                ElseIf ComboBox1.Text = "��� ��������" Then
                    MySQLStr = "Exec spp_SupplSearch_GetProposalInfo " & Trim(DataGridView1.SelectedRows.Item(0).Cells("ID").Value.ToString()) & ", 0 "
                Else
                    MySQLStr = "Exec spp_SupplSearch_GetProposalInfo " & Trim(DataGridView1.SelectedRows.Item(0).Cells("ID").Value.ToString()) & ", 2 "
                End If
            Else
                MySQLStr = "Exec spp_SupplSearch_GetProposalInfo 0, 0 "
            End If
            Try
                MyAdapter = New SqlClient.SqlDataAdapter(MySQLStr, Declarations.MyNETConnStr)
                MyAdapter.SelectCommand.CommandTimeout = 600
                MyAdapter.Fill(MyDs)
                DataGridView3.DataSource = MyDs.Tables(0)
            Catch ex As Exception
                MsgBox(ex.ToString)
            End Try

            '---���������
            DataGridView3.Columns("ID").HeaderText = "ID"
            DataGridView3.Columns("ID").Width = 40
            DataGridView3.Columns("ID").HeaderCell.Style.BackColor = Color.LightBlue
            DataGridView3.Columns("ID").SortMode = DataGridViewColumnSortMode.Programmatic
            DataGridView3.Columns("IsSelected").HeaderText = "��� ��� ���"
            DataGridView3.Columns("IsSelected").Width = 30
            DataGridView3.Columns("IsSelected").SortMode = DataGridViewColumnSortMode.Programmatic
            DataGridView3.Columns("SelectedBySalesman").HeaderText = "��� ���"
            DataGridView3.Columns("SelectedBySalesman").Width = 30
            DataGridView3.Columns("SelectedBySalesman").SortMode = DataGridViewColumnSortMode.Programmatic
            DataGridView3.Columns("ItemID").HeaderText = "ItemID"
            DataGridView3.Columns("ItemID").Width = 0
            DataGridView3.Columns("ItemID").SortMode = DataGridViewColumnSortMode.Programmatic
            DataGridView3.Columns("ItemID").Visible = False
            DataGridView3.Columns("ItemCode").HeaderText = "��� ������ � �����"
            DataGridView3.Columns("ItemCode").Width = 100
            DataGridView3.Columns("ItemCode").HeaderCell.Style.BackColor = Color.LightBlue
            DataGridView3.Columns("ItemCode").SortMode = DataGridViewColumnSortMode.Programmatic
            DataGridView3.Columns("ItemSuppCode").HeaderText = "��� ������ ����������"
            DataGridView3.Columns("ItemSuppCode").Width = 100
            DataGridView3.Columns("ItemSuppCode").HeaderCell.Style.BackColor = Color.LightBlue
            DataGridView3.Columns("ItemSuppCode").SortMode = DataGridViewColumnSortMode.Programmatic
            DataGridView3.Columns("ItemName").HeaderText = "�������� ������"
            DataGridView3.Columns("ItemName").Width = 180
            DataGridView3.Columns("ItemName").HeaderCell.Style.BackColor = Color.LightBlue
            DataGridView3.Columns("ItemName").SortMode = DataGridViewColumnSortMode.Programmatic
            DataGridView3.Columns("UOM").HeaderText = "�� ���� �����"
            DataGridView3.Columns("UOM").Width = 40
            DataGridView3.Columns("UOM").HeaderCell.Style.BackColor = Color.LightBlue
            DataGridView3.Columns("UOM").SortMode = DataGridViewColumnSortMode.Programmatic
            DataGridView3.Columns("QTY").HeaderText = "���-��"
            DataGridView3.Columns("QTY").Width = 70
            DataGridView3.Columns("QTY").HeaderCell.Style.BackColor = Color.LightBlue
            DataGridView3.Columns("QTY").DefaultCellStyle.Format = "n3"
            DataGridView3.Columns("QTY").SortMode = DataGridViewColumnSortMode.Programmatic
            DataGridView3.Columns("Price").HeaderText = "����� ���� ��� ���"
            DataGridView3.Columns("Price").Width = 70
            DataGridView3.Columns("Price").DefaultCellStyle.Format = "n2"
            DataGridView3.Columns("Price").SortMode = DataGridViewColumnSortMode.Programmatic
            DataGridView3.Columns("PriCost").HeaderText = "��������� �.����� ��� ������� � ��� ��������"
            DataGridView3.Columns("PriCost").Width = 100
            DataGridView3.Columns("PriCost").DefaultCellStyle.Format = "n2"
            DataGridView3.Columns("PriCost").SortMode = DataGridViewColumnSortMode.Programmatic
            DataGridView3.Columns("Curr").HeaderText = "������"
            DataGridView3.Columns("Curr").Width = 50
            DataGridView3.Columns("Curr").SortMode = DataGridViewColumnSortMode.Programmatic
            DataGridView3.Columns("LeadTimeWeek").HeaderText = "���� ������ �� (���)"
            DataGridView3.Columns("LeadTimeWeek").Width = 50
            DataGridView3.Columns("LeadTimeWeek").DefaultCellStyle.Format = "n2"
            DataGridView3.Columns("LeadTimeWeek").SortMode = DataGridViewColumnSortMode.Programmatic
            DataGridView3.Columns("SupplierID").HeaderText = "SupplierID"
            DataGridView3.Columns("SupplierID").Width = 0
            DataGridView3.Columns("SupplierID").Visible = False
            DataGridView3.Columns("SupplierID").SortMode = DataGridViewColumnSortMode.Programmatic
            DataGridView3.Columns("SupplierCode").HeaderText = "��� ������ ���� � �����"
            DataGridView3.Columns("SupplierCode").Width = 80
            DataGridView3.Columns("SupplierCode").SortMode = DataGridViewColumnSortMode.Programmatic
            DataGridView3.Columns("SupplierName").HeaderText = "���������"
            DataGridView3.Columns("SupplierName").Width = 120
            DataGridView3.Columns("SupplierName").SortMode = DataGridViewColumnSortMode.Programmatic
            DataGridView3.Columns("ShippingCost").HeaderText = "����� ����������"
            DataGridView3.Columns("ShippingCost").Width = 80
            DataGridView3.Columns("ShippingCost").DefaultCellStyle.Format = "n2"
            DataGridView3.Columns("ShippingCost").SortMode = DataGridViewColumnSortMode.Programmatic
            DataGridView3.Columns("Comments").HeaderText = "����������� ����������"
            DataGridView3.Columns("Comments").Width = 150
            DataGridView3.Columns("Comments").SortMode = DataGridViewColumnSortMode.Programmatic
            DataGridView3.Columns("SalesmanComments").HeaderText = "����������� ��������"
            DataGridView3.Columns("SalesmanComments").Width = 150
            DataGridView3.Columns("SalesmanComments").SortMode = DataGridViewColumnSortMode.Programmatic
            DataGridView3.Columns("AlternateTo").HeaderText = "������������ ������"
            DataGridView3.Columns("AlternateTo").Width = 150
            DataGridView3.Columns("AlternateTo").SortMode = DataGridViewColumnSortMode.Programmatic
            DataGridView3.Columns("DueDate").HeaderText = "������. ��"
            DataGridView3.Columns("DueDate").Width = 100
            DataGridView3.Columns("DueDate").SortMode = DataGridViewColumnSortMode.Programmatic

            FormatDataGridView3()
        End If
    End Sub

    Private Sub FormatDataGridView3()
        '////////////////////////////////////////////////////////////////////////////////
        '//
        '// ��������� ���������� �� ������� � �����������
        '//
        '////////////////////////////////////////////////////////////////////////////////
        Dim i As Integer

        For i = 0 To DataGridView3.Rows.Count - 1
            If DataGridView3.Rows(i).Cells("AlternateTo").Value = "" Then
                DataGridView3.Rows(i).DefaultCellStyle.BackColor = Color.White
            Else
                DataGridView3.Rows(i).DefaultCellStyle.BackColor = Color.FromArgb(255, 255, 189)
            End If

            If DataGridView3.Rows(i).Cells("IsSelected").Value = False And DataGridView3.Rows(i).Cells("SelectedBySalesman").Value = True Then
                DataGridView3.Rows(i).Cells("IsSelected").Style.BackColor = Color.FromArgb(255, 179, 179)
                'DataGridView3.Rows(i).DefaultCellStyle.BackColor = Color.FromArgb(255, 179, 179)
            Else
                DataGridView3.Rows(i).Cells("IsSelected").Style.BackColor = Color.White
                'DataGridView3.Rows(i).DefaultCellStyle.BackColor = Color.White
            End If

            If Not IsDBNull(DataGridView3.Rows(i).Cells("SupplierName").Value) Then
                DataGridView3.Rows(i).Cells("ID").ToolTipText = DataGridView3.Rows(i).Cells("SupplierName").Value
                DataGridView3.Rows(i).Cells("IsSelected").ToolTipText = DataGridView3.Rows(i).Cells("SupplierName").Value
                DataGridView3.Rows(i).Cells("SelectedBySalesman").ToolTipText = DataGridView3.Rows(i).Cells("SupplierName").Value
                DataGridView3.Rows(i).Cells("ItemID").ToolTipText = DataGridView3.Rows(i).Cells("SupplierName").Value
                DataGridView3.Rows(i).Cells("ItemCode").ToolTipText = DataGridView3.Rows(i).Cells("SupplierName").Value
                DataGridView3.Rows(i).Cells("ItemSuppCode").ToolTipText = DataGridView3.Rows(i).Cells("SupplierName").Value
                DataGridView3.Rows(i).Cells("ItemName").ToolTipText = DataGridView3.Rows(i).Cells("SupplierName").Value
                DataGridView3.Rows(i).Cells("UOM").ToolTipText = DataGridView3.Rows(i).Cells("SupplierName").Value
                DataGridView3.Rows(i).Cells("QTY").ToolTipText = DataGridView3.Rows(i).Cells("SupplierName").Value
                DataGridView3.Rows(i).Cells("Price").ToolTipText = DataGridView3.Rows(i).Cells("SupplierName").Value
                DataGridView3.Rows(i).Cells("PriCost").ToolTipText = DataGridView3.Rows(i).Cells("SupplierName").Value
                DataGridView3.Rows(i).Cells("Curr").ToolTipText = DataGridView3.Rows(i).Cells("SupplierName").Value
                DataGridView3.Rows(i).Cells("LeadTimeWeek").ToolTipText = DataGridView3.Rows(i).Cells("SupplierName").Value
                DataGridView3.Rows(i).Cells("SupplierID").ToolTipText = DataGridView3.Rows(i).Cells("SupplierName").Value
                DataGridView3.Rows(i).Cells("SupplierCode").ToolTipText = DataGridView3.Rows(i).Cells("SupplierName").Value
            End If

        Next
    End Sub

    Private Sub CheckProposalButtons()
        '////////////////////////////////////////////////////////////////////////////////
        '//
        '// �������� ������ ����������� �����������
        '//
        '////////////////////////////////////////////////////////////////////////////////

        If DataGridView3.SelectedRows.Count = 0 Then
            Button15.Enabled = False
            Button25.Enabled = False
        Else
            Button15.Enabled = True
            Button25.Enabled = True
        End If

        If DataGridView1.SelectedRows.Count = 0 Then
            Button2.Enabled = False
            Button21.Enabled = False
        Else
            If DataGridView3.SelectedRows.Count = 0 Then
                Button2.Enabled = False
                Button21.Enabled = False
            Else
                If DataGridView1.SelectedRows.Item(0).Cells("SalesStatusID").Value = -1 Then
                    Button2.Enabled = False
                    Button21.Enabled = False
                ElseIf DataGridView1.SelectedRows.Item(0).Cells("SalesStatusID").Value = 0 Then
                    '-----������� ���������
                    If DataGridView1.SelectedRows.Item(0).Cells("SearchStatusID").Value = 0 Then
                        '-----��� �� � ������ �����������
                        Button2.Enabled = False
                        Button21.Enabled = False
                    ElseIf DataGridView1.SelectedRows.Item(0).Cells("SearchStatusID").Value = 1 Then
                        '-----� ������ �����������
                        Button2.Enabled = False
                        Button21.Enabled = False
                    ElseIf DataGridView1.SelectedRows.Item(0).Cells("SearchStatusID").Value = 2 Then
                        '-----��������� ����������� �����������
                        Button2.Enabled = True
                        Button21.Enabled = True
                    ElseIf DataGridView1.SelectedRows.Item(0).Cells("SearchStatusID").Value = 4 Then
                        '-----��������� �������
                        Button2.Enabled = False
                        Button21.Enabled = False
                    End If

                ElseIf DataGridView1.SelectedRows.Item(0).Cells("SalesStatusID").Value = 1 Then
                    '-----�������� ���������� �����������
                    If DataGridView1.SelectedRows.Item(0).Cells("SearchStatusID").Value = 2 Then
                        '-----��������� ����������� �����������
                        Button2.Enabled = True
                        Button21.Enabled = False
                    ElseIf DataGridView1.SelectedRows.Item(0).Cells("SearchStatusID").Value = 3 Then
                        '-----��������� ������ �����
                        Button2.Enabled = True
                        Button21.Enabled = False
                    ElseIf DataGridView1.SelectedRows.Item(0).Cells("SearchStatusID").Value = 4 Then
                        '-----��������� �������
                        Button2.Enabled = False
                        Button21.Enabled = False
                    End If

                ElseIf DataGridView1.SelectedRows.Item(0).Cells("SalesStatusID").Value = 2 Then
                    '-----�������� �� ���������� �����������
                    If DataGridView1.SelectedRows.Item(0).Cells("SearchStatusID").Value = 1 Then
                        '-----� ������ �����������
                        Button2.Enabled = False
                        Button21.Enabled = False
                    ElseIf DataGridView1.SelectedRows.Item(0).Cells("SearchStatusID").Value = 2 Then
                        '-----��������� ����������� �����������
                        Button2.Enabled = True
                        Button21.Enabled = True
                    ElseIf DataGridView1.SelectedRows.Item(0).Cells("SearchStatusID").Value = 4 Then
                        '-----��������� �������
                        Button2.Enabled = False
                        Button21.Enabled = False
                    End If

                Else
                    '-----�������� ��������� ������ ������ (3)
                    Button2.Enabled = False
                    Button21.Enabled = False
                End If
            End If
        End If

    End Sub


    Private Sub ComboBoxAct_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles ComboBoxAct.SelectedIndexChanged
        '////////////////////////////////////////////////////////////////////////////////
        '//
        '// ����� ������ ���������� ��������
        '//
        '////////////////////////////////////////////////////////////////////////////////

        LoadRequests()
        LoadItems()
        LoadProposal()
        CheckRequestButtons()
        CheckItemButtons()
        CheckProposalButtons()
    End Sub

    Private Sub DataGridView1_CellMouseClick(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellMouseEventArgs) Handles DataGridView1.CellMouseClick
        '////////////////////////////////////////////////////////////////////////////////
        '//
        '// ����������� ���� �����������
        '//
        '////////////////////////////////////////////////////////////////////////////////

        If e.Button = System.Windows.Forms.MouseButtons.Right Then
            Declarations.MyRowIndex = e.RowIndex
            Declarations.MyRequestNum = DataGridView1.Rows(Declarations.MyRowIndex).Cells("ID").Value
            If DataGridView1.Rows(Declarations.MyRowIndex).Cells("SalesStatusID").Value = 0 Or _
                DataGridView1.Rows(Declarations.MyRowIndex).Cells("SalesStatusID").Value = 1 Or _
                DataGridView1.Rows(Declarations.MyRowIndex).Cells("SalesStatusID").Value = 2 Then
                ContextMenuStrip1.Show(MousePosition.X, MousePosition.Y)
            End If
        End If
    End Sub

    Private Sub DataGridView1_SelectionChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles DataGridView1.SelectionChanged
        '////////////////////////////////////////////////////////////////////////////////
        '//
        '// ����� ������ �������� / ��������
        '//
        '////////////////////////////////////////////////////////////////////////////////

        If LoadFlag = 0 Then
            System.Windows.Forms.Cursor.Current = Cursors.WaitCursor
            LoadItems()
            LoadProposal()
            CheckRequestButtons()
            CheckItemButtons()
            CheckProposalButtons()
            System.Windows.Forms.Cursor.Current = Cursors.Default
        End If
    End Sub

    Private Sub DataGridView1_Sorted(ByVal sender As Object, ByVal e As System.EventArgs) Handles DataGridView1.Sorted
        '////////////////////////////////////////////////////////////////////////////////
        '//
        '// ����� ����������
        '//
        '////////////////////////////////////////////////////////////////////////////////

        FormatDataGridView1()
    End Sub

    Private Sub Button16_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button16.Click
        '////////////////////////////////////////////////////////////////////////////////
        '//
        '// ���������� ���������� � ����
        '//
        '////////////////////////////////////////////////////////////////////////////////

        LoadRequests()
        LoadItems()
        LoadProposal()
        CheckRequestButtons()
        CheckItemButtons()
        CheckProposalButtons()
    End Sub

    Private Sub Button8_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button8.Click
        '////////////////////////////////////////////////////////////////////////////////
        '//
        '// �������� ������� �� �����
        '//
        '////////////////////////////////////////////////////////////////////////////////

        MyEditRequest = New EditRequest
        MyEditRequest.StartParam = "Create"
        MyEditRequest.WindowFrom = "SearchSupplier"
        MyEditRequest.ShowDialog()
        LoadRequests()
        For i As Integer = 0 To DataGridView1.Rows.Count - 1
            If Trim(DataGridView1.Item("ID", i).Value.ToString) = Trim(Declarations.MyRequestNum) Then
                DataGridView1.CurrentCell = DataGridView1.Item("ID", i)
                Exit For
            End If
        Next
        CheckRequestButtons()
    End Sub

    Private Sub Button7_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button7.Click
        '////////////////////////////////////////////////////////////////////////////////
        '//
        '// �������������� ������� �� �����
        '//
        '////////////////////////////////////////////////////////////////////////////////

        MyEditRequest = New EditRequest
        MyEditRequest.StartParam = "Edit"
        MyEditRequest.WindowFrom = "SearchSupplier"
        MyEditRequest.ShowDialog()
        LoadRequests()
        For i As Integer = 0 To DataGridView1.Rows.Count - 1
            If Trim(DataGridView1.Item("ID", i).Value.ToString) = Trim(Declarations.MyRequestNum) Then
                DataGridView1.CurrentCell = DataGridView1.Item("ID", i)
                Exit For
            End If
        Next
        CheckRequestButtons()
    End Sub

    Private Sub Button10_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button10.Click
        '////////////////////////////////////////////////////////////////////////////////
        '//
        '// ����������� ������� �� �����
        '//
        '////////////////////////////////////////////////////////////////////////////////
        Dim MySQLStr As String

        Try
            MySQLStr = "exec spp_SupplSearch_SearchRequestCopy " & Trim(DataGridView1.SelectedRows.Item(0).Cells("ID").Value.ToString())
            InitMyConn(False)
            InitMyRec(False, MySQLStr)
            If Declarations.MyRec.BOF = True And Declarations.MyRec.EOF = True Then
                MsgBox("������ ����������� ������")
                trycloseMyRec()
            Else
                Declarations.MyRequestNum = Declarations.MyRec.Fields("MyNewID").Value
                trycloseMyRec()
                '---� �������� �������� ������ ��������
                MySQLStr = "UPDATE tbl_SupplSearch "
                MySQLStr = MySQLStr + "SET SalesmanID = N'" + Trim(Declarations.SalesmanCode) + "', "
                MySQLStr = MySQLStr + "SalesmanName = N'" + Trim(Declarations.SalesmanName) + "' "
                MySQLStr = MySQLStr + "WHERE (ID = " + CStr(Declarations.MyRequestNum) + ")"
                InitMyConn(False)
                Declarations.MyConn.Execute(MySQLStr)
                LoadRequests()
                For i As Integer = 0 To DataGridView1.Rows.Count - 1
                    If Trim(DataGridView1.Item("ID", i).Value.ToString) = Trim(Declarations.MyRequestNum) Then
                        DataGridView1.CurrentCell = DataGridView1.Item("ID", i)
                        Exit For
                    End If
                Next
                CheckRequestButtons()
                MsgBox("������ �����������")
            End If
        Catch ex As Exception
            MsgBox(ex.ToString)
        End Try
    End Sub

    Private Sub Button6_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button6.Click
        '////////////////////////////////////////////////////////////////////////////////
        '//
        '// �������� ������� �� �����
        '//
        '////////////////////////////////////////////////////////////////////////////////
        Dim MySQLStr As String

        MySQLStr = "DELETE FROM tbl_SupplSearchItems "
        MySQLStr = MySQLStr + "WHERE (SupplSearchID = " + Trim(DataGridView1.SelectedRows.Item(0).Cells("ID").Value.ToString()) + ") "
        InitMyConn(False)
        Declarations.MyConn.Execute(MySQLStr)

        MySQLStr = "DELETE FROM tbl_SupplSearch "
        MySQLStr = MySQLStr + "WHERE (ID = " + Trim(DataGridView1.SelectedRows.Item(0).Cells("ID").Value.ToString()) + ")"
        InitMyConn(False)
        Declarations.MyConn.Execute(MySQLStr)

        LoadRequests()
        CheckRequestButtons()
        MsgBox("������ �������")
    End Sub

    Private Sub Button12_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button12.Click
        '////////////////////////////////////////////////////////////////////////////////
        '//
        '// �������� ���� � ������������ ��������
        '//
        '////////////////////////////////////////////////////////////////////////////////

        Declarations.MyRequestNum = DataGridView1.SelectedRows.Item(0).Cells("ID").Value
        MyAttachmentsList = New AttachmentsList
        MyAttachmentsList.AttType = "Sales"
        MyAttachmentsList.WhoStart = "Sales"
        MyAttachmentsList.MyPlace = "List"
        MyAttachmentsList.ShowDialog()
    End Sub

    Private Sub Button14_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button14.Click
        '////////////////////////////////////////////////////////////////////////////////
        '//
        '// �������� ���� � ������������ ����������
        '//
        '////////////////////////////////////////////////////////////////////////////////

        Declarations.MyRequestNum = DataGridView1.SelectedRows.Item(0).Cells("ID").Value
        MyAttachmentsList = New AttachmentsList
        MyAttachmentsList.AttType = "Search"
        MyAttachmentsList.WhoStart = "Sales"
        MyAttachmentsList.MyPlace = "List"
        MyAttachmentsList.ShowDialog()
    End Sub

    Private Sub Button19_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button19.Click
        '////////////////////////////////////////////////////////////////////////////////
        '//
        '// �������� ������� "� ������"
        '//
        '////////////////////////////////////////////////////////////////////////////////
        Dim MySQLStr As String
        Dim MyID As Integer
        Dim EmailStr As String
        Dim RequestStatus As String

        If DataGridView2.Rows.Count = 0 Then
            MsgBox("���������� ������ ���� �� ���� ����� ��� ������", MsgBoxStyle.Information, "��������!")
        Else
            MyID = DataGridView1.SelectedRows.Item(0).Cells("ID").Value
            MySQLStr = "UPDATE tbl_SupplSearch "
            MySQLStr = MySQLStr + "SET SalesStatus = 0 "
            MySQLStr = MySQLStr + "WHERE (ID = " & MyID & ") "
            InitMyConn(False)
            Declarations.MyConn.Execute(MySQLStr)
            '--------------�������� �����
            'EmailStr = GetEmailFromDB(Trim(Strings.Left(DataGridView1.SelectedRows.Item(0).Cells(11).Value.ToString, 4)))
            EmailStr = GetSrchManagerEmailFromDB()
            If EmailStr = "" Then
                MsgBox("��� ������������ ����������� � �� �� �������� ����� ��� �� �� ������ � ������� tbl_SupplSearch_Searchers. ���������� � ��������������.", MsgBoxStyle.Critical, "��������!")
            Else
                RequestStatus = "��������� ������"
                SendInfoByEmail(DataGridView1.SelectedRows.Item(0).Cells("ID").Value.ToString(), DataGridView1.SelectedRows.Item(0).Cells("StartDate").Value.ToString(), _
                   EmailStr, DataGridView1.SelectedRows.Item(0).Cells("Customer").Value.ToString(), DataGridView1.SelectedRows.Item(0).Cells("Salesman").Value.ToString(), _
                   RequestStatus)
            End If
            '---------------------------
            LoadRequests()
            For i As Integer = 0 To DataGridView1.Rows.Count - 1
                If DataGridView1.Item("ID", i).Value = MyID Then
                    DataGridView1.CurrentCell = DataGridView1.Item("ID", i)
                    Exit For
                End If
            Next
            CheckRequestButtons()
        End If
    End Sub

    Private Sub Button4_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button4.Click
        '////////////////////////////////////////////////////////////////////////////////
        '//
        '// �������� �������
        '//
        '////////////////////////////////////////////////////////////////////////////////
        Dim MySQLStr As String
        Dim MyID As Integer
        Dim EmailStr As String
        Dim RequestStatus As String

        MyID = DataGridView1.SelectedRows.Item(0).Cells("ID").Value
        'myValue = ""
        'myValue = InputBox("������� ����������� (���� ����������)", "�����������", "")
        MyCommentAndCancelReason = New CommentAndCancelReason
        MyCommentAndCancelReason.MyID = MyID
        MyCommentAndCancelReason.ShowDialog()
        MySQLStr = "UPDATE tbl_SupplSearch "
        MySQLStr = MySQLStr + "SET SalesStatus = 3 "
        'MySQLStr = MySQLStr + "Comments = ISNULL(Comments, '') + " + Chr(10) + Chr(13) + " + '--" + Format(Now, "dd/MM/yyyy HH:mm") + "-->' + N'" & myValue & "' "
        MySQLStr = MySQLStr + "WHERE (ID = " & MyID & ") "
        InitMyConn(False)
        Declarations.MyConn.Execute(MySQLStr)
        '--------------�������� �����
        EmailStr = GetEmailFromDB(Trim(Strings.Left(DataGridView1.SelectedRows.Item(0).Cells("Searcher").Value.ToString, 4)))
        If EmailStr = "" Then
            MsgBox("��� ������������ " & DataGridView1.SelectedRows.Item(0).Cells("Searcher").Value.ToString & " � �� �� �������� �����. ���������� � ��������������.", MsgBoxStyle.Critical, "��������!")
        Else
            RequestStatus = "������ ������"
            SendInfoByEmail(DataGridView1.SelectedRows.Item(0).Cells("ID").Value.ToString(), DataGridView1.SelectedRows.Item(0).Cells("StartDate").Value.ToString(), _
                EmailStr, DataGridView1.SelectedRows.Item(0).Cells("Customer").Value.ToString(), DataGridView1.SelectedRows.Item(0).Cells("Salesman").Value.ToString(), _
                RequestStatus)
        End If
        '---------------------------
        LoadRequests()
        For i As Integer = 0 To DataGridView1.Rows.Count - 1
            If DataGridView1.Item("ID", i).Value = MyID Then
                DataGridView1.CurrentCell = DataGridView1.Item("ID", i)
                Exit For
            End If
        Next
        CheckRequestButtons()
    End Sub

    Private Sub Button17_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button17.Click
        '////////////////////////////////////////////////////////////////////////////////
        '//
        '// ������������� �������� ������������� �����������
        '//
        '////////////////////////////////////////////////////////////////////////////////
        Dim MySQLStr As String
        Dim myValue As String
        Dim MyID As Integer
        Dim EmailStr As String
        Dim RequestStatus As String

        MyID = DataGridView1.SelectedRows.Item(0).Cells("ID").Value
        myValue = ""
        myValue = InputBox("������� ����������� (���� ����������)", "�����������", "")
        MySQLStr = "UPDATE tbl_SupplSearch "
        MySQLStr = MySQLStr + "SET SalesStatus = 1, "
        MySQLStr = MySQLStr + "Comments = ISNULL(Comments, '') + '" + CStr(Chr(10) + Chr(13)) + "' + '--" + Format(Now, "dd/MM/yyyy HH:mm") + "-->' + N'" & myValue & "', "
        MySQLStr = MySQLStr + "CancelReason = N'' "
        MySQLStr = MySQLStr + "WHERE (ID = " & MyID & ") "
        InitMyConn(False)
        Declarations.MyConn.Execute(MySQLStr)
        '--------------�������� �����
        EmailStr = GetEmailFromDB(Trim(Strings.Left(DataGridView1.SelectedRows.Item(0).Cells("Searcher").Value.ToString, 4)))
        If EmailStr = "" Then
            MsgBox("��� ������������ " & DataGridView1.SelectedRows.Item(0).Cells("Searcher").Value.ToString & " � �� �� �������� �����. ���������� � ��������������.", MsgBoxStyle.Critical, "��������!")
        Else
            RequestStatus = "������� �����������"
            SendInfoByEmail(DataGridView1.SelectedRows.Item(0).Cells("ID").Value.ToString(), DataGridView1.SelectedRows.Item(0).Cells("StartDate").Value.ToString(), _
               EmailStr, DataGridView1.SelectedRows.Item(0).Cells("Customer").Value.ToString(), DataGridView1.SelectedRows.Item(0).Cells("Salesman").Value.ToString(), _
               RequestStatus)
        End If
        '---------------------------
        LoadRequests()
        For i As Integer = 0 To DataGridView1.Rows.Count - 1
            If DataGridView1.Item("ID", i).Value = MyID Then
                DataGridView1.CurrentCell = DataGridView1.Item("ID", i)
                Exit For
            End If
        Next
        CheckRequestButtons()
    End Sub

    Private Sub Button18_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button18.Click
        '////////////////////////////////////////////////////////////////////////////////
        '//
        '// ���������� � ���������, ������������ �����������
        '//
        '////////////////////////////////////////////////////////////////////////////////
        Dim MySQLStr As String
        Dim MyDate As DateTime
        Dim MyID As Integer
        Dim EmailStr As String
        Dim RequestStatus As String

        MyID = DataGridView1.SelectedRows.Item(0).Cells("ID").Value
        MyDate = DataGridView1.SelectedRows.Item(0).Cells("RequestDate").Value
        '---����� ������������� ���� �������������� ��
        MyCorrectRequestDate = New CorrectRequestDate
        MyCorrectRequestDate.MyID = MyID
        MyCorrectRequestDate.MyDate = MyDate
        MyCorrectRequestDate.ShowDialog()

        '---����������� � ������� ������
        MyCommentAndCancelReason = New CommentAndCancelReason
        MyCommentAndCancelReason.MyID = MyID
        MyCommentAndCancelReason.ShowDialog()
        MySQLStr = "UPDATE tbl_SupplSearch "
        MySQLStr = MySQLStr + "SET SalesStatus = 2, "
        MySQLStr = MySQLStr + "SearchStatus = 1 "
        'MySQLStr = MySQLStr + "Comments = ISNULL(Comments, '') + " + Chr(10) + Chr(13) + " + '--" + Format(Now, "dd/MM/yyyy HH:mm") + "-->' + N'" & myValue & "' "
        MySQLStr = MySQLStr + "WHERE (ID = " & MyID & ") "
        InitMyConn(False)
        Declarations.MyConn.Execute(MySQLStr)
        '--------------�������� �����
        EmailStr = GetEmailFromDB(Trim(Strings.Left(DataGridView1.SelectedRows.Item(0).Cells("Searcher").Value.ToString, 4)))
        If EmailStr = "" Then
            MsgBox("��� ������������ " & DataGridView1.SelectedRows.Item(0).Cells("Searcher").Value.ToString & " � �� �� �������� �����. ���������� � ��������������.", MsgBoxStyle.Critical, "��������!")
        Else
            RequestStatus = "������� �� ���������� / ���������"
            SendInfoByEmail(DataGridView1.SelectedRows.Item(0).Cells("ID").Value.ToString(), DataGridView1.SelectedRows.Item(0).Cells("StartDate").Value.ToString(), _
               EmailStr, DataGridView1.SelectedRows.Item(0).Cells("Customer").Value.ToString(), DataGridView1.SelectedRows.Item(0).Cells("Salesman").Value.ToString(), _
               RequestStatus)
        End If
        '---------------------------
        LoadRequests()
        For i As Integer = 0 To DataGridView1.Rows.Count - 1
            If DataGridView1.Item("ID", i).Value = MyID Then
                DataGridView1.CurrentCell = DataGridView1.Item("ID", i)
                Exit For
            End If
        Next
        CheckRequestButtons()
    End Sub

    Private Sub Button3_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button3.Click
        '////////////////////////////////////////////////////////////////////////////////
        '//
        '// �������� ������ � ������ � �������
        '//
        '////////////////////////////////////////////////////////////////////////////////

        MyAddItem = New AddItem
        MyAddItem.StartParam = "Create"
        MyAddItem.ShowDialog()
        LoadItems()

        'DataGridView2.CurrentCell = DataGridView2.Item(0, DataGridView2.RowCount - 1)
        CheckItemButtons()
    End Sub

    Private Sub Button9_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button9.Click
        '////////////////////////////////////////////////////////////////////////////////
        '//
        '// �������������� ������ � ������ � �������
        '//
        '////////////////////////////////////////////////////////////////////////////////

        Declarations.MyItemSrchID = DataGridView2.SelectedRows.Item(0).Cells(0).Value
        MyAddItem = New AddItem
        MyAddItem.StartParam = "Edit"
        MyAddItem.ShowDialog()
        LoadItems()

        For i As Integer = 0 To DataGridView2.Rows.Count - 1
            If DataGridView2.Item(0, i).Value = Declarations.MyItemSrchID Then
                DataGridView2.CurrentCell = DataGridView2.Item(0, i)
                Exit For
            End If
        Next
        CheckItemButtons()
    End Sub

    Private Sub Button11_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button11.Click
        '////////////////////////////////////////////////////////////////////////////////
        '//
        '// �������� ������ � ������ � �������
        '//
        '////////////////////////////////////////////////////////////////////////////////
        Dim MySQLStr As String

        MySQLStr = "DELETE FROM tbl_SupplSearchItems "
        MySQLStr = MySQLStr & "WHERE (ID = " & DataGridView2.SelectedRows.Item(0).Cells(0).Value & ") "
        InitMyConn(False)
        Declarations.MyConn.Execute(MySQLStr)
        LoadItems()
        CheckItemButtons()
    End Sub

    Private Sub Button5_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button5.Click
        '////////////////////////////////////////////////////////////////////////////////
        '//
        '// �������� ����������� ������� � ������������
        '//
        '////////////////////////////////////////////////////////////////////////////////

        If My.Settings.UseOffice = "LibreOffice" Then
            If DataGridView2.SortedColumn Is Nothing Then
                ExportToSpecToLO(0, 1)
            Else
                ExportToSpecToLO(DataGridView2.SortedColumn.Index, DataGridView2.SortOrder)
            End If

        Else
            If DataGridView2.SortedColumn Is Nothing Then
                ExportToSpecToExcel(0, 1)
            Else
                ExportToSpecToExcel(DataGridView2.SortedColumn.Index, DataGridView2.SortOrder)
            End If
        End If
    End Sub

    Private Sub ExportToSpecToExcel(ByVal MyCol As Integer, ByVal MySort As Integer)
        '////////////////////////////////////////////////////////////////////////////////
        '//
        '// �������� ����������� ������� � ������������ � Excel
        '//
        '////////////////////////////////////////////////////////////////////////////////
        Dim MySQLStr As String
        Dim SpecVersion As String               '--������ ������������
        Dim MyObj As Object
        Dim MyWRKBook As Object
        Dim i As Integer

        MyObj = CreateObject("Excel.Application")
        Dim oldCI As System.Globalization.CultureInfo = System.Threading.Thread.CurrentThread.CurrentCulture
        System.Threading.Thread.CurrentThread.CurrentCulture = New System.Globalization.CultureInfo("en-US")

        MyObj.SheetsInNewWorkbook = 1
        MyObj.ReferenceStyle = 1
        MyWRKBook = MyObj.Workbooks.Add

        '---���������
        '---������ ������������
        MySQLStr = "SELECT Version "
        MySQLStr = MySQLStr & "FROM tbl_VersionImportItemsFromExcel WITH (NOLOCK) "
        MySQLStr = MySQLStr & "WHERE (Name = N'������������ �����������') "
        InitMyConn(False)
        InitMyRec(False, MySQLStr)
        If Declarations.MyRec.BOF = True And Declarations.MyRec.EOF = True Then
            MsgBox("� Scala �� ����������� ������� ������ ����� Excel. ���������� � ��������������.", vbCritical, "��������!")
            trycloseMyRec()
            Exit Sub
        Else
            SpecVersion = Trim(Declarations.MyRec.Fields("Version").Value)
            trycloseMyRec()
        End If
        MyWRKBook.ActiveSheet.Range("A1") = SpecVersion
        With MyWRKBook.ActiveSheet.Range("A1").Font
            .Name = "Calibri"
            .Size = 9
            '.Color = -16776961
            .ColorIndex = 3
        End With

        MyWRKBook.ActiveSheet.Range("B2") = "Skandika"
        With MyWRKBook.ActiveSheet.Range("B2").Font
            .Name = "Calibri"
            .Size = 16
            .Color = 0
            .Bold = True
        End With

        MyWRKBook.ActiveSheet.Range("A4") = "OOO ""��������"""
        With MyWRKBook.ActiveSheet.Range("A4").Font
            .Name = "Tahoma"
            .Size = 9
            .Color = 0
            .Bold = True
        End With

        MyWRKBook.ActiveSheet.Range("A5") = "�����:"
        With MyWRKBook.ActiveSheet.Range("A5").Font
            .Name = "Tahoma"
            .Size = 9
            .Color = 0
            .Bold = True
        End With

        MyWRKBook.ActiveSheet.Range("B5:H6").MergeCells = True
        MyWRKBook.ActiveSheet.Range("B5") = "������,195027, ����� ���������, ������� ��., �.4 Tel: +7 (812) 325 2040, Fax: +7 (812) 325 2039, www.skandikagroup.ru"
        With MyWRKBook.ActiveSheet.Range("B5").Font
            .Name = "Tahoma"
            .Size = 9
            .Color = 0
            .Bold = True
        End With
        MyWRKBook.ActiveSheet.Range("B5:H6").WrapText = True
        MyWRKBook.ActiveSheet.Range("B5:H6").VerticalAlignment = -4160

        MyWRKBook.ActiveSheet.Range("D8") = "������������ ��������"
        With MyWRKBook.ActiveSheet.Range("D8").Font
            .Name = "Tahoma"
            .Size = 11.5
            .Color = 0
            .Bold = True
        End With

        '--��������� �������
        MyWRKBook.ActiveSheet.Range("A10") = "N �/�"
        MyWRKBook.ActiveSheet.Range("B10") = "��� ������ Scala"
        MyWRKBook.ActiveSheet.Range("C10") = "��� ������ ����������"
        MyWRKBook.ActiveSheet.Range("D10") = "������������ ������"
        MyWRKBook.ActiveSheet.Range("E10") = "�� ���������"
        MyWRKBook.ActiveSheet.Range("F10") = "���-��"
        MyWRKBook.ActiveSheet.Range("G10") = "���� ��� ���"
        MyWRKBook.ActiveSheet.Range("H10") = "����� ��� ���"
        MyWRKBook.ActiveSheet.Range("I10") = "���� �������� (���)"
        MyWRKBook.ActiveSheet.Range("A10:I10").Select()
        MyWRKBook.ActiveSheet.Range("A10:I10").Borders(5).LineStyle = -4142
        MyWRKBook.ActiveSheet.Range("A10:I10").Borders(6).LineStyle = -4142
        With MyWRKBook.ActiveSheet.Range("A10:I10").Borders(7)
            .LineStyle = 1
            .Weight = 2
            .ColorIndex = -4105
        End With
        With MyWRKBook.ActiveSheet.Range("A10:I10").Borders(8)
            .LineStyle = 1
            .Weight = 2
            .ColorIndex = -4105
        End With
        With MyWRKBook.ActiveSheet.Range("A10:I10").Borders(9)
            .LineStyle = 1
            .Weight = 2
            .ColorIndex = -4105
        End With
        With MyWRKBook.ActiveSheet.Range("A10:I10").Borders(10)
            .LineStyle = 1
            .Weight = 2
            .ColorIndex = -4105
        End With
        With MyWRKBook.ActiveSheet.Range("A10:I10").Borders(11)
            .LineStyle = 1
            .Weight = 2
            .ColorIndex = -4105
        End With
        MyWRKBook.ActiveSheet.Range("A10:I10").WrapText = True
        MyWRKBook.ActiveSheet.Range("A10:I10").VerticalAlignment = -4108
        MyWRKBook.ActiveSheet.Range("A10:I10").HorizontalAlignment = -4108
        With MyWRKBook.ActiveSheet.Range("A10:I10").Font
            .Name = "Calibri"
            .Size = 9
            .Color = 0
            .Bold = False
        End With

        MyWRKBook.ActiveSheet.Columns("A:A").ColumnWidth = 4
        MyWRKBook.ActiveSheet.Columns("B:B").ColumnWidth = 10
        MyWRKBook.ActiveSheet.Columns("C:C").ColumnWidth = 10
        MyWRKBook.ActiveSheet.Columns("D:D").ColumnWidth = 30
        MyWRKBook.ActiveSheet.Columns("E:E").ColumnWidth = 5
        MyWRKBook.ActiveSheet.Columns("F:F").ColumnWidth = 8
        MyWRKBook.ActiveSheet.Columns("G:G").ColumnWidth = 10
        MyWRKBook.ActiveSheet.Columns("H:H").ColumnWidth = 12
        MyWRKBook.ActiveSheet.Columns("I:I").ColumnWidth = 12

        '--����� ����� ������������
        MySQLStr = "SELECT ISNULL(tbl_SupplSearchItems.ItemID, N'') AS ItemCode, ISNULL(tbl_SupplSearchItems.ItemSuppID, N'') AS SuppItemCode, "
        MySQLStr = MySQLStr & "CASE WHEN LTRIM(RTRIM(LTRIM(RTRIM(ISNULL(SC010300.SC01002, ''))) + LTRIM(RTRIM(ISNULL(SC010300.SC01003, ''))))) "
        MySQLStr = MySQLStr & "= '' THEN ISNULL(tbl_SupplSearchItems.ItemName, '') ELSE LTRIM(RTRIM(LTRIM(RTRIM(ISNULL(SC010300.SC01002, ''))) "
        MySQLStr = MySQLStr & "+ LTRIM(RTRIM(ISNULL(SC010300.SC01003, ''))))) END AS ItemName, CASE WHEN ISNULL(View_1.txt, N'') "
        MySQLStr = MySQLStr & "= '��' THEN 'pcs(��.)' ELSE CASE WHEN ISNULL(View_1.txt, N'') = '�' THEN 'm (�)' ELSE CASE WHEN ISNULL(View_1.txt, N'') "
        MySQLStr = MySQLStr & "= '��' THEN 'kg (��)' ELSE CASE WHEN ISNULL(View_1.txt, N'') = '��' THEN 'km (��)' ELSE CASE WHEN ISNULL(View_1.txt, N'') "
        MySQLStr = MySQLStr & "= '�' THEN 'litre (����)' ELSE CASE WHEN ISNULL(View_1.txt, N'') = '����' THEN 'pack (����.)' ELSE CASE WHEN ISNULL(View_1.txt, N'') "
        MySQLStr = MySQLStr & "= '�����' THEN 'set (�����.)' ELSE CASE WHEN ISNULL(View_1.txt, N'') = '���' THEN 'pair (����)' END END END END END END END END AS UOM, "
        MySQLStr = MySQLStr & "tbl_SupplSearchItems.QTY, '' AS Price, ISNULL(tbl_SupplSearchItems.LeadTimeWeek, 1) AS WeekQTY "
        MySQLStr = MySQLStr & "FROM tbl_SupplSearchItems LEFT OUTER JOIN "
        MySQLStr = MySQLStr & "SC010300 ON tbl_SupplSearchItems.ItemID = SC010300.SC01001 LEFT OUTER JOIN "
        MySQLStr = MySQLStr & "(SELECT     0 AS Expr1, SC09002 AS txt "
        MySQLStr = MySQLStr & "FROM          SC090300 WITH (NOLOCK) "
        MySQLStr = MySQLStr & "WHERE      (SC09001 = 'RUS') "
        MySQLStr = MySQLStr & "UNION "
        MySQLStr = MySQLStr & "SELECT     1 AS Expr1, SC09003 "
        MySQLStr = MySQLStr & "FROM         SC090300 AS SC090300_40 WITH (NOLOCK) "
        MySQLStr = MySQLStr & "WHERE     (SC09001 = 'RUS') "
        MySQLStr = MySQLStr & "UNION "
        MySQLStr = MySQLStr & "SELECT     2 AS Expr1, SC09004 "
        MySQLStr = MySQLStr & "FROM         SC090300 AS SC090300_39 WITH (NOLOCK) "
        MySQLStr = MySQLStr & "WHERE     (SC09001 = 'RUS') "
        MySQLStr = MySQLStr & "UNION "
        MySQLStr = MySQLStr & "SELECT     3 AS Expr1, SC09005 "
        MySQLStr = MySQLStr & "FROM         SC090300 AS SC090300_38 WITH (NOLOCK) "
        MySQLStr = MySQLStr & "WHERE     (SC09001 = 'RUS') "
        MySQLStr = MySQLStr & "UNION "
        MySQLStr = MySQLStr & "SELECT     4 AS Expr1, SC09006 "
        MySQLStr = MySQLStr & "FROM         SC090300 AS SC090300_37 WITH (NOLOCK) "
        MySQLStr = MySQLStr & "WHERE     (SC09001 = 'RUS') "
        MySQLStr = MySQLStr & "UNION "
        MySQLStr = MySQLStr & "SELECT     5 AS Expr1, SC09007 "
        MySQLStr = MySQLStr & "FROM         SC090300 AS SC090300_36 WITH (NOLOCK) "
        MySQLStr = MySQLStr & "WHERE     (SC09001 = 'RUS') "
        MySQLStr = MySQLStr & "UNION "
        MySQLStr = MySQLStr & "SELECT     6 AS Expr1, SC09008 "
        MySQLStr = MySQLStr & "FROM         SC090300 AS SC090300_35 WITH (NOLOCK) "
        MySQLStr = MySQLStr & "WHERE     (SC09001 = 'RUS') "
        MySQLStr = MySQLStr & "UNION "
        MySQLStr = MySQLStr & "SELECT     7 AS Expr1, SC09009 "
        MySQLStr = MySQLStr & "FROM         SC090300 AS SC090300_34 WITH (NOLOCK) "
        MySQLStr = MySQLStr & "WHERE     (SC09001 = 'RUS') "
        MySQLStr = MySQLStr & "UNION "
        MySQLStr = MySQLStr & "SELECT     8 AS Expr1, SC09010 "
        MySQLStr = MySQLStr & "FROM         SC090300 AS SC090300_33 WITH (NOLOCK) "
        MySQLStr = MySQLStr & "WHERE     (SC09001 = 'RUS') "
        MySQLStr = MySQLStr & "UNION "
        MySQLStr = MySQLStr & "SELECT     9 AS Expr1, SC09011 "
        MySQLStr = MySQLStr & "FROM         SC090300 AS SC090300_32 WITH (NOLOCK) "
        MySQLStr = MySQLStr & "WHERE     (SC09001 = 'RUS') "
        MySQLStr = MySQLStr & "UNION "
        MySQLStr = MySQLStr & "SELECT     10 AS Expr1, SC09012 "
        MySQLStr = MySQLStr & "FROM         SC090300 AS SC090300_31 WITH (NOLOCK) "
        MySQLStr = MySQLStr & "WHERE     (SC09001 = 'RUS') "
        MySQLStr = MySQLStr & "UNION "
        MySQLStr = MySQLStr & "SELECT     11 AS Expr1, SC09013 "
        MySQLStr = MySQLStr & "FROM         SC090300 AS SC090300_30 WITH (NOLOCK) "
        MySQLStr = MySQLStr & "WHERE     (SC09001 = 'RUS') "
        MySQLStr = MySQLStr & "UNION "
        MySQLStr = MySQLStr & "SELECT     12 AS Expr1, SC09014 "
        MySQLStr = MySQLStr & "FROM         SC090300 AS SC090300_29 WITH (NOLOCK) "
        MySQLStr = MySQLStr & "WHERE     (SC09001 = 'RUS') "
        MySQLStr = MySQLStr & "UNION "
        MySQLStr = MySQLStr & "SELECT     13 AS Expr1, SC09015 "
        MySQLStr = MySQLStr & "FROM         SC090300 AS SC090300_28 WITH (NOLOCK) "
        MySQLStr = MySQLStr & "WHERE     (SC09001 = 'RUS') "
        MySQLStr = MySQLStr & "UNION "
        MySQLStr = MySQLStr & "SELECT     14 AS Expr1, SC09016 "
        MySQLStr = MySQLStr & "FROM         SC090300 AS SC090300_27 WITH (NOLOCK) "
        MySQLStr = MySQLStr & "WHERE     (SC09001 = 'RUS') "
        MySQLStr = MySQLStr & "UNION "
        MySQLStr = MySQLStr & "SELECT     15 AS Expr1, SC09017 "
        MySQLStr = MySQLStr & "FROM         SC090300 AS SC090300_26 WITH (NOLOCK) "
        MySQLStr = MySQLStr & "WHERE     (SC09001 = 'RUS') "
        MySQLStr = MySQLStr & "UNION "
        MySQLStr = MySQLStr & "SELECT     16 AS Expr1, SC09018 "
        MySQLStr = MySQLStr & "FROM         SC090300 AS SC090300_25 WITH (NOLOCK) "
        MySQLStr = MySQLStr & "WHERE     (SC09001 = 'RUS') "
        MySQLStr = MySQLStr & "UNION "
        MySQLStr = MySQLStr & "SELECT     17 AS Expr1, SC09019 "
        MySQLStr = MySQLStr & "FROM         SC090300 AS SC090300_24 WITH (NOLOCK) "
        MySQLStr = MySQLStr & "WHERE     (SC09001 = 'RUS') "
        MySQLStr = MySQLStr & "UNION "
        MySQLStr = MySQLStr & "SELECT     18 AS Expr1, SC09020 "
        MySQLStr = MySQLStr & "FROM         SC090300 AS SC090300_23 WITH (NOLOCK) "
        MySQLStr = MySQLStr & "WHERE     (SC09001 = 'RUS') "
        MySQLStr = MySQLStr & "UNION "
        MySQLStr = MySQLStr & "SELECT     19 AS Expr1, SC09021 "
        MySQLStr = MySQLStr & "FROM         SC090300 AS SC090300_22 WITH (NOLOCK) "
        MySQLStr = MySQLStr & "WHERE     (SC09001 = 'RUS') "
        MySQLStr = MySQLStr & "UNION "
        MySQLStr = MySQLStr & "SELECT     20 AS Expr1, SC09022 "
        MySQLStr = MySQLStr & "FROM         SC090300 AS SC090300_21 WITH (NOLOCK) "
        MySQLStr = MySQLStr & "WHERE     (SC09001 = 'RUS') "
        MySQLStr = MySQLStr & "UNION "
        MySQLStr = MySQLStr & "SELECT     21 AS Expr1, SC09023 "
        MySQLStr = MySQLStr & "FROM         SC090300 AS SC090300_20 WITH (NOLOCK) "
        MySQLStr = MySQLStr & "WHERE     (SC09001 = 'RUS') "
        MySQLStr = MySQLStr & "UNION "
        MySQLStr = MySQLStr & "SELECT     22 AS Expr1, SC09024 "
        MySQLStr = MySQLStr & "FROM         SC090300 AS SC090300_19 WITH (NOLOCK) "
        MySQLStr = MySQLStr & "WHERE     (SC09001 = 'RUS') "
        MySQLStr = MySQLStr & "UNION "
        MySQLStr = MySQLStr & "SELECT     23 AS Expr1, SC09025 "
        MySQLStr = MySQLStr & "FROM         SC090300 AS SC090300_18 WITH (NOLOCK) "
        MySQLStr = MySQLStr & "WHERE     (SC09001 = 'RUS') "
        MySQLStr = MySQLStr & "UNION "
        MySQLStr = MySQLStr & "SELECT     24 AS Expr1, SC09026 "
        MySQLStr = MySQLStr & "FROM         SC090300 AS SC090300_17 WITH (NOLOCK) "
        MySQLStr = MySQLStr & "WHERE     (SC09001 = 'RUS') "
        MySQLStr = MySQLStr & "UNION "
        MySQLStr = MySQLStr & "SELECT     25 AS Expr1, SC09027 "
        MySQLStr = MySQLStr & "FROM         SC090300 AS SC090300_16 WITH (NOLOCK) "
        MySQLStr = MySQLStr & "WHERE     (SC09001 = 'RUS') "
        MySQLStr = MySQLStr & "UNION "
        MySQLStr = MySQLStr & "SELECT     26 AS Expr1, SC09028 "
        MySQLStr = MySQLStr & "FROM         SC090300 AS SC090300_15 WITH (NOLOCK) "
        MySQLStr = MySQLStr & "WHERE     (SC09001 = 'RUS') "
        MySQLStr = MySQLStr & "UNION "
        MySQLStr = MySQLStr & "SELECT     27 AS Expr1, SC09029 "
        MySQLStr = MySQLStr & "FROM         SC090300 AS SC090300_14 WITH (NOLOCK) "
        MySQLStr = MySQLStr & "WHERE     (SC09001 = 'RUS') "
        MySQLStr = MySQLStr & "UNION "
        MySQLStr = MySQLStr & "SELECT     28 AS Expr1, SC09030 "
        MySQLStr = MySQLStr & "FROM         SC090300 AS SC090300_13 WITH (NOLOCK) "
        MySQLStr = MySQLStr & "WHERE     (SC09001 = 'RUS') "
        MySQLStr = MySQLStr & "UNION "
        MySQLStr = MySQLStr & "SELECT     29 AS Expr1, SC09031 "
        MySQLStr = MySQLStr & "FROM         SC090300 AS SC090300_12 WITH (NOLOCK) "
        MySQLStr = MySQLStr & "WHERE     (SC09001 = 'RUS') "
        MySQLStr = MySQLStr & "UNION "
        MySQLStr = MySQLStr & "SELECT     30 AS Expr1, SC09032 "
        MySQLStr = MySQLStr & "FROM         SC090300 AS SC090300_11 WITH (NOLOCK) "
        MySQLStr = MySQLStr & "WHERE     (SC09001 = 'RUS') "
        MySQLStr = MySQLStr & "UNION "
        MySQLStr = MySQLStr & "SELECT     31 AS Expr1, SC09033 "
        MySQLStr = MySQLStr & "FROM         SC090300 AS SC090300_10 WITH (NOLOCK) "
        MySQLStr = MySQLStr & "WHERE     (SC09001 = 'RUS') "
        MySQLStr = MySQLStr & "UNION "
        MySQLStr = MySQLStr & "SELECT     32 AS Expr1, SC09034 "
        MySQLStr = MySQLStr & "FROM         SC090300 AS SC090300_9 WITH (NOLOCK) "
        MySQLStr = MySQLStr & "WHERE     (SC09001 = 'RUS') "
        MySQLStr = MySQLStr & "UNION "
        MySQLStr = MySQLStr & "SELECT     33 AS Expr1, SC09035 "
        MySQLStr = MySQLStr & "FROM         SC090300 AS SC090300_8 WITH (NOLOCK) "
        MySQLStr = MySQLStr & "WHERE     (SC09001 = 'RUS') "
        MySQLStr = MySQLStr & "UNION "
        MySQLStr = MySQLStr & "SELECT     34 AS Expr1, SC09036 "
        MySQLStr = MySQLStr & "FROM         SC090300 AS SC090300_7 WITH (NOLOCK) "
        MySQLStr = MySQLStr & "WHERE     (SC09001 = 'RUS') "
        MySQLStr = MySQLStr & "UNION "
        MySQLStr = MySQLStr & "SELECT     35 AS Expr1, SC09037 "
        MySQLStr = MySQLStr & "FROM         SC090300 AS SC090300_6 WITH (NOLOCK) "
        MySQLStr = MySQLStr & "WHERE     (SC09001 = 'RUS') "
        MySQLStr = MySQLStr & "UNION "
        MySQLStr = MySQLStr & "SELECT     36 AS Expr1, SC09038 "
        MySQLStr = MySQLStr & "FROM         SC090300 AS SC090300_5 WITH (NOLOCK) "
        MySQLStr = MySQLStr & "WHERE     (SC09001 = 'RUS') "
        MySQLStr = MySQLStr & "UNION "
        MySQLStr = MySQLStr & "SELECT     37 AS Expr1, SC09039 "
        MySQLStr = MySQLStr & "FROM         SC090300 AS SC090300_4 WITH (NOLOCK) "
        MySQLStr = MySQLStr & "WHERE     (SC09001 = 'RUS') "
        MySQLStr = MySQLStr & "UNION "
        MySQLStr = MySQLStr & "SELECT     38 AS Expr1, SC09040 "
        MySQLStr = MySQLStr & "FROM         SC090300 AS SC090300_3 WITH (NOLOCK) "
        MySQLStr = MySQLStr & "WHERE     (SC09001 = 'RUS') "
        MySQLStr = MySQLStr & "UNION "
        MySQLStr = MySQLStr & "SELECT     39 AS Expr1, SC09041 "
        MySQLStr = MySQLStr & "FROM         SC090300 AS SC090300_2 WITH (NOLOCK) "
        MySQLStr = MySQLStr & "WHERE     (SC09001 = 'RUS') "
        MySQLStr = MySQLStr & "UNION "
        MySQLStr = MySQLStr & "SELECT     40 AS Expr1, SC09042 "
        MySQLStr = MySQLStr & "FROM         SC090300 AS SC090300_1 WITH (NOLOCK)) AS View_1 ON tbl_SupplSearchItems.UOM = View_1.Expr1 "
        MySQLStr = MySQLStr & "WHERE (tbl_SupplSearchItems.SupplSearchID = " & DataGridView1.SelectedRows.Item(0).Cells("ID").Value & ") "
        Select Case MyCol
            Case 0
                MySQLStr = MySQLStr & "ORDER BY tbl_SupplSearchItems.ID "
            Case 2
                MySQLStr = MySQLStr & "ORDER BY SuppItemCode "
            Case 3
                MySQLStr = MySQLStr & "ORDER BY ItemName "
            Case 4
                MySQLStr = MySQLStr & "ORDER BY UOM "
            Case 5
                MySQLStr = MySQLStr & "ORDER BY tbl_SupplSearchItems.QTY "
            Case Else
                MySQLStr = MySQLStr & "ORDER BY tbl_SupplSearchItems.ID "
        End Select
        If MySort = 2 Then
            MySQLStr = MySQLStr & "Desc "
        End If

        InitMyConn(False)
        InitMyRec(False, MySQLStr)
        If Declarations.MyRec.BOF = True And Declarations.MyRec.EOF = True Then
            trycloseMyRec()
        Else
            Declarations.MyRec.MoveFirst()
            i = 0
            While Declarations.MyRec.EOF = False
                MyWRKBook.ActiveSheet.Range("A" & CStr(i + 11)).NumberFormat = "@"
                MyWRKBook.ActiveSheet.Range("A" & CStr(i + 11)) = i + 1
                MyWRKBook.ActiveSheet.Range("B" & CStr(i + 11)).NumberFormat = "@"
                MyWRKBook.ActiveSheet.Range("B" & CStr(i + 11)) = Declarations.MyRec.Fields("ItemCode").Value
                MyWRKBook.ActiveSheet.Range("C" & CStr(i + 11)).NumberFormat = "@"
                MyWRKBook.ActiveSheet.Range("C" & CStr(i + 11)) = Declarations.MyRec.Fields("SuppItemCode").Value
                MyWRKBook.ActiveSheet.Range("D" & CStr(i + 11)) = Declarations.MyRec.Fields("ItemName").Value
                MyWRKBook.ActiveSheet.Range("E" & CStr(i + 11)) = Declarations.MyRec.Fields("UOM").Value
                MyWRKBook.ActiveSheet.Range("F" & CStr(i + 11)) = Declarations.MyRec.Fields("QTY").Value
                MyWRKBook.ActiveSheet.Range("G" & CStr(i + 11)) = Declarations.MyRec.Fields("Price").Value
                MyWRKBook.ActiveSheet.Range("I" & CStr(i + 11)) = Declarations.MyRec.Fields("WeekQTY").Value

                i = i + 1
                Declarations.MyRec.MoveNext()
            End While
            trycloseMyRec()
        End If
        '---

        MyWRKBook.ActiveSheet.Range("A11:I11").Select()
        MyWRKBook.ActiveSheet.Range("A11:I11").Borders(5).LineStyle = -4142
        MyWRKBook.ActiveSheet.Range("A11:I11").Borders(6).LineStyle = -4142
        With MyWRKBook.ActiveSheet.Range("A11:I1011").Borders(7)
            .LineStyle = 1
            .Weight = 2
            .ColorIndex = -4105
        End With
        With MyWRKBook.ActiveSheet.Range("A11:I1011").Borders(8)
            .LineStyle = 1
            .Weight = 2
            .ColorIndex = -4105
        End With
        With MyWRKBook.ActiveSheet.Range("A11:I1011").Borders(9)
            .LineStyle = 1
            .Weight = 2
            .ColorIndex = -4105
        End With
        With MyWRKBook.ActiveSheet.Range("A11:I1011").Borders(10)
            .LineStyle = 1
            .Weight = 2
            .ColorIndex = -4105
        End With
        With MyWRKBook.ActiveSheet.Range("A11:I1011").Borders(11)
            .LineStyle = 1
            .Weight = 2
            .ColorIndex = -4105
        End With
        With MyWRKBook.ActiveSheet.Range("A11:I1011").Borders(12)
            .LineStyle = 1
            .Weight = 2
            .ColorIndex = -4105
        End With
        With MyWRKBook.ActiveSheet.Range("A11:I1011").Font
            .Name = "Calibri"
            .Size = 9
            .Color = 0
            .Bold = False
        End With
        With MyWRKBook.ActiveSheet.Range("H11:H1011")
            '.FormulaR1C1 = "=����(RC[-2]*RC[-1] = 0; """"; RC[-2]*RC[-1])"
            .FormulaR1C1 = "=IF(RC[-2]*RC[-1] = 0, """", RC[-2]*RC[-1])"
        End With
        System.Threading.Thread.CurrentThread.CurrentCulture = oldCI

        '---�������� ���������� ����������
        MyWRKBook.ActiveSheet.Range("N10") = "������� ���������"
        MyWRKBook.ActiveSheet.Range("N11") = 0
        MyWRKBook.ActiveSheet.Range("O11") = "pcs(��.)"
        MyWRKBook.ActiveSheet.Range("N12") = 1
        MyWRKBook.ActiveSheet.Range("O12") = "m (�)"
        MyWRKBook.ActiveSheet.Range("N13") = 2
        MyWRKBook.ActiveSheet.Range("O13") = "kg (��)"
        MyWRKBook.ActiveSheet.Range("N14") = 3
        MyWRKBook.ActiveSheet.Range("O14") = "km (��)"
        MyWRKBook.ActiveSheet.Range("N15") = 4
        MyWRKBook.ActiveSheet.Range("O15") = "litre (����)"
        MyWRKBook.ActiveSheet.Range("N16") = 5
        MyWRKBook.ActiveSheet.Range("O16") = "pack (����.)"
        MyWRKBook.ActiveSheet.Range("N17") = 6
        MyWRKBook.ActiveSheet.Range("O17") = "set (�����.)"
        MyWRKBook.ActiveSheet.Range("N18") = 7
        MyWRKBook.ActiveSheet.Range("O18") = "pair (����)"

        MyWRKBook.ActiveSheet.Range("N10:O18").Font.Color = 16777215
        'MyWRKBook.ActiveSheet.Range("N10:O18").Font.TintAndShade = 0
        MyWRKBook.ActiveSheet.Range("E11:E1011").Validation.Add(Type:=3, AlertStyle:=1, Operator:=1, Formula1:="=$O$11:$O$18")
        'MyWRKBook.ActiveSheet.Range("E11:E1011").Validation.Add(Type:=3, AlertStyle:=1, Operator:=1, Formula1:="=R11C15:R18C15")

        MyWRKBook.ActiveSheet.Cells.Locked = True
        MyWRKBook.ActiveSheet.Cells.FormulaHidden = True

        MyWRKBook.ActiveSheet.Range("A11:G1011").Locked = False
        MyWRKBook.ActiveSheet.Range("A11:G1011").FormulaHidden = False
        MyWRKBook.ActiveSheet.Range("I11:I1011").Locked = False
        MyWRKBook.ActiveSheet.Range("I11:I1011").FormulaHidden = False

        MyWRKBook.ActiveSheet.Protect(Password:="!pass2009", DrawingObjects:=True, Contents:=True, Scenarios:=True)

        'MyWRKBook.Application.PrintCommunication = True
        'MyWRKBook.ActiveSheet.PageSetup.PrintArea = "$A$1:$I$1011"
        'MyWRKBook.Application.PrintCommunication = False
        'MyWRKBook.ActiveSheet.PageSetup.FitToPagesWide = 1
        'MyWRKBook.ActiveSheet.PageSetup.FitToPagesTall = 0
        'MyWRKBook.Application.PrintCommunication = True

        MyWRKBook.ActiveSheet.Range("A11").Select()
        MyObj.Application.Visible = True
        MyWRKBook = Nothing
        MyObj = Nothing
        oldCI = Nothing
    End Sub

    Private Sub ExportToSpecToLO(ByVal MyCol As Integer, ByVal MySort As Integer)
        '////////////////////////////////////////////////////////////////////////////////
        '//
        '// �������� ����������� ������� � ������������ � LibreOffice
        '//
        '////////////////////////////////////////////////////////////////////////////////
        Dim oServiceManager As Object
        Dim oDispatcher As Object
        Dim oDesktop As Object
        Dim oWorkBook As Object
        Dim oSheet As Object
        Dim oFrame As Object
        Dim MySQLStr As String
        Dim SpecVersion As String               '--������ ������������
        Dim i As Integer

        LOSetNotation(0)
        oServiceManager = CreateObject("com.sun.star.ServiceManager")
        oDesktop = oServiceManager.createInstance("com.sun.star.frame.Desktop")
        oDispatcher = oServiceManager.createInstance("com.sun.star.frame.DispatchHelper")
        Dim arg(1)
        arg(0) = mAkePropertyValue("Hidden", True, oServiceManager)
        arg(1) = mAkePropertyValue("MacroExecutionMode", 4, oServiceManager)
        oWorkBook = oDesktop.loadComponentFromURL("private:factory/scalc", "_blank", 0, arg)
        oSheet = oWorkBook.getSheets().getByIndex(0)
        oFrame = oWorkBook.getCurrentController.getFrame

        '---������ �������
        oSheet.getColumns().getByName("A").Width = 1390
        oSheet.getColumns().getByName("B").Width = 2280
        oSheet.getColumns().getByName("C").Width = 2570
        oSheet.getColumns().getByName("D").Width = 5590
        oSheet.getColumns().getByName("E").Width = 1150
        oSheet.getColumns().getByName("F").Width = 1770
        oSheet.getColumns().getByName("G").Width = 2190
        oSheet.getColumns().getByName("H").Width = 2260
        oSheet.getColumns().getByName("I").Width = 2260
        '---������ �����
        LOSetCellProtection(oServiceManager, oDispatcher, oFrame, "", True)
        '---���������
        '---������ ������������
        MySQLStr = "SELECT Version "
        MySQLStr = MySQLStr & "FROM tbl_VersionImportItemsFromExcel WITH (NOLOCK) "
        MySQLStr = MySQLStr & "WHERE (Name = N'������������ �����������') "
        InitMyConn(False)
        InitMyRec(False, MySQLStr)
        If Declarations.MyRec.BOF = True And Declarations.MyRec.EOF = True Then
            MsgBox("� Scala �� ����������� ������� ������ ����� Excel. ���������� � ��������������.", vbCritical, "��������!")
            trycloseMyRec()
            Exit Sub
        Else
            SpecVersion = Trim(Declarations.MyRec.Fields("Version").Value)
            trycloseMyRec()
        End If
        oSheet.getCellRangeByName("A1").String = SpecVersion
        LOFontSetFamilyName(oServiceManager, oDispatcher, oFrame, "A1", "Calibri")
        LOFontSetSize(oServiceManager, oDispatcher, oFrame, "A1", 11)
        oSheet.getCellRangeByName("A1").CharColor = RGB(61, 65, 239) '---��� ����� �������� ������� - �� ����� ���� ��� �� RGB � BGR!!!! ������ ���� �����

        oSheet.getCellRangeByName("B2").String = "Skandika"
        LOFontSetFamilyName(oServiceManager, oDispatcher, oFrame, "B2", "Calibri")
        LOFontSetBold(oServiceManager, oDispatcher, oFrame, "B2")
        LOFontSetSize(oServiceManager, oDispatcher, oFrame, "B2", 16)

        oSheet.getCellRangeByName("A4").String = "OOO ""��������"""
        LOFontSetFamilyName(oServiceManager, oDispatcher, oFrame, "A4", "Tahoma")
        LOFontSetBold(oServiceManager, oDispatcher, oFrame, "A4")
        LOFontSetSize(oServiceManager, oDispatcher, oFrame, "A4", 11)

        oSheet.getCellRangeByName("A5").String = "�����:"
        LOFontSetFamilyName(oServiceManager, oDispatcher, oFrame, "A5", "Tahoma")
        LOFontSetBold(oServiceManager, oDispatcher, oFrame, "A5")
        LOFontSetSize(oServiceManager, oDispatcher, oFrame, "A5", 9)

        LOMergeCells(oServiceManager, oDispatcher, oFrame, "B5:H6")
        oSheet.getCellRangeByName("B5").String = "������,195027, ����� ���������, ������� ��., �.4 Tel: +7 (812) 325 2040, Fax: +7 (812) 325 2039, www.skandikagroup.ru"
        LOFontSetFamilyName(oServiceManager, oDispatcher, oFrame, "B5", "Tahoma")
        LOFontSetBold(oServiceManager, oDispatcher, oFrame, "B5")
        LOFontSetSize(oServiceManager, oDispatcher, oFrame, "B5", 9)
        LOWrapText(oServiceManager, oDispatcher, oFrame, "B5:H6")

        oSheet.getCellRangeByName("D8").String = "������������ ��������"
        LOFontSetFamilyName(oServiceManager, oDispatcher, oFrame, "D8", "Tahoma")
        LOFontSetBold(oServiceManager, oDispatcher, oFrame, "D8")
        LOFontSetSize(oServiceManager, oDispatcher, oFrame, "D8", 11.5)

        '--��������� �������
        oSheet.getCellRangeByName("A10").String = "N �/�"
        oSheet.getCellRangeByName("B10").String = "��� ������ Scala"
        oSheet.getCellRangeByName("C10").String = "��� ������ ����������"
        oSheet.getCellRangeByName("D10").String = "������������ ������"
        oSheet.getCellRangeByName("E10").String = "�� ���������"
        oSheet.getCellRangeByName("F10").String = "���-��"
        oSheet.getCellRangeByName("G10").String = "���� ��� ���"
        oSheet.getCellRangeByName("H10").String = "����� ��� ���"
        oSheet.getCellRangeByName("I10").String = "���� �������� (���)"
        LOFontSetFamilyName(oServiceManager, oDispatcher, oFrame, "A10:I10", "Calibri")
        LOFontSetSize(oServiceManager, oDispatcher, oFrame, "A10:I10", 9)
        LOWrapText(oServiceManager, oDispatcher, oFrame, "A10:I10")
        Dim LineFormat As Object
        LineFormat = oServiceManager.Bridge_GetStruct("com.sun.star.table.BorderLine2")
        LineFormat.LineStyle = 0
        LineFormat.LineWidth = 40
        oSheet.getCellRangeByName("A10:I10").TopBorder = LineFormat
        oSheet.getCellRangeByName("A10:I10").RightBorder = LineFormat
        oSheet.getCellRangeByName("A10:I10").LeftBorder = LineFormat
        oSheet.getCellRangeByName("A10:I10").BottomBorder = LineFormat
        oSheet.getCellRangeByName("A10:I10").VertJustify = 2
        oSheet.getCellRangeByName("A10:I10").HoriJustify = 2

        '--����� ����� ������������
        MySQLStr = "SELECT ISNULL(tbl_SupplSearchItems.ItemID, N'') AS ItemCode, ISNULL(tbl_SupplSearchItems.ItemSuppID, N'') AS SuppItemCode, "
        MySQLStr = MySQLStr & "CASE WHEN LTRIM(RTRIM(LTRIM(RTRIM(ISNULL(SC010300.SC01002, ''))) + LTRIM(RTRIM(ISNULL(SC010300.SC01003, ''))))) "
        MySQLStr = MySQLStr & "= '' THEN ISNULL(tbl_SupplSearchItems.ItemName, '') ELSE LTRIM(RTRIM(LTRIM(RTRIM(ISNULL(SC010300.SC01002, ''))) "
        MySQLStr = MySQLStr & "+ LTRIM(RTRIM(ISNULL(SC010300.SC01003, ''))))) END AS ItemName, CASE WHEN ISNULL(View_1.txt, N'') "
        MySQLStr = MySQLStr & "= '��' THEN 'pcs(��.)' ELSE CASE WHEN ISNULL(View_1.txt, N'') = '�' THEN 'm (�)' ELSE CASE WHEN ISNULL(View_1.txt, N'') "
        MySQLStr = MySQLStr & "= '��' THEN 'kg (��)' ELSE CASE WHEN ISNULL(View_1.txt, N'') = '��' THEN 'km (��)' ELSE CASE WHEN ISNULL(View_1.txt, N'') "
        MySQLStr = MySQLStr & "= '�' THEN 'litre (����)' ELSE CASE WHEN ISNULL(View_1.txt, N'') = '����' THEN 'pack (����.)' ELSE CASE WHEN ISNULL(View_1.txt, N'') "
        MySQLStr = MySQLStr & "= '�����' THEN 'set (�����.)' ELSE CASE WHEN ISNULL(View_1.txt, N'') = '���' THEN 'pair (����)' END END END END END END END END AS UOM, "
        MySQLStr = MySQLStr & "tbl_SupplSearchItems.QTY, '' AS Price, ISNULL(tbl_SupplSearchItems.LeadTimeWeek, 1) AS WeekQTY "
        MySQLStr = MySQLStr & "FROM tbl_SupplSearchItems LEFT OUTER JOIN "
        MySQLStr = MySQLStr & "SC010300 ON tbl_SupplSearchItems.ItemID = SC010300.SC01001 LEFT OUTER JOIN "
        MySQLStr = MySQLStr & "(SELECT     0 AS Expr1, SC09002 AS txt "
        MySQLStr = MySQLStr & "FROM          SC090300 WITH (NOLOCK) "
        MySQLStr = MySQLStr & "WHERE      (SC09001 = 'RUS') "
        MySQLStr = MySQLStr & "UNION "
        MySQLStr = MySQLStr & "SELECT     1 AS Expr1, SC09003 "
        MySQLStr = MySQLStr & "FROM         SC090300 AS SC090300_40 WITH (NOLOCK) "
        MySQLStr = MySQLStr & "WHERE     (SC09001 = 'RUS') "
        MySQLStr = MySQLStr & "UNION "
        MySQLStr = MySQLStr & "SELECT     2 AS Expr1, SC09004 "
        MySQLStr = MySQLStr & "FROM         SC090300 AS SC090300_39 WITH (NOLOCK) "
        MySQLStr = MySQLStr & "WHERE     (SC09001 = 'RUS') "
        MySQLStr = MySQLStr & "UNION "
        MySQLStr = MySQLStr & "SELECT     3 AS Expr1, SC09005 "
        MySQLStr = MySQLStr & "FROM         SC090300 AS SC090300_38 WITH (NOLOCK) "
        MySQLStr = MySQLStr & "WHERE     (SC09001 = 'RUS') "
        MySQLStr = MySQLStr & "UNION "
        MySQLStr = MySQLStr & "SELECT     4 AS Expr1, SC09006 "
        MySQLStr = MySQLStr & "FROM         SC090300 AS SC090300_37 WITH (NOLOCK) "
        MySQLStr = MySQLStr & "WHERE     (SC09001 = 'RUS') "
        MySQLStr = MySQLStr & "UNION "
        MySQLStr = MySQLStr & "SELECT     5 AS Expr1, SC09007 "
        MySQLStr = MySQLStr & "FROM         SC090300 AS SC090300_36 WITH (NOLOCK) "
        MySQLStr = MySQLStr & "WHERE     (SC09001 = 'RUS') "
        MySQLStr = MySQLStr & "UNION "
        MySQLStr = MySQLStr & "SELECT     6 AS Expr1, SC09008 "
        MySQLStr = MySQLStr & "FROM         SC090300 AS SC090300_35 WITH (NOLOCK) "
        MySQLStr = MySQLStr & "WHERE     (SC09001 = 'RUS') "
        MySQLStr = MySQLStr & "UNION "
        MySQLStr = MySQLStr & "SELECT     7 AS Expr1, SC09009 "
        MySQLStr = MySQLStr & "FROM         SC090300 AS SC090300_34 WITH (NOLOCK) "
        MySQLStr = MySQLStr & "WHERE     (SC09001 = 'RUS') "
        MySQLStr = MySQLStr & "UNION "
        MySQLStr = MySQLStr & "SELECT     8 AS Expr1, SC09010 "
        MySQLStr = MySQLStr & "FROM         SC090300 AS SC090300_33 WITH (NOLOCK) "
        MySQLStr = MySQLStr & "WHERE     (SC09001 = 'RUS') "
        MySQLStr = MySQLStr & "UNION "
        MySQLStr = MySQLStr & "SELECT     9 AS Expr1, SC09011 "
        MySQLStr = MySQLStr & "FROM         SC090300 AS SC090300_32 WITH (NOLOCK) "
        MySQLStr = MySQLStr & "WHERE     (SC09001 = 'RUS') "
        MySQLStr = MySQLStr & "UNION "
        MySQLStr = MySQLStr & "SELECT     10 AS Expr1, SC09012 "
        MySQLStr = MySQLStr & "FROM         SC090300 AS SC090300_31 WITH (NOLOCK) "
        MySQLStr = MySQLStr & "WHERE     (SC09001 = 'RUS') "
        MySQLStr = MySQLStr & "UNION "
        MySQLStr = MySQLStr & "SELECT     11 AS Expr1, SC09013 "
        MySQLStr = MySQLStr & "FROM         SC090300 AS SC090300_30 WITH (NOLOCK) "
        MySQLStr = MySQLStr & "WHERE     (SC09001 = 'RUS') "
        MySQLStr = MySQLStr & "UNION "
        MySQLStr = MySQLStr & "SELECT     12 AS Expr1, SC09014 "
        MySQLStr = MySQLStr & "FROM         SC090300 AS SC090300_29 WITH (NOLOCK) "
        MySQLStr = MySQLStr & "WHERE     (SC09001 = 'RUS') "
        MySQLStr = MySQLStr & "UNION "
        MySQLStr = MySQLStr & "SELECT     13 AS Expr1, SC09015 "
        MySQLStr = MySQLStr & "FROM         SC090300 AS SC090300_28 WITH (NOLOCK) "
        MySQLStr = MySQLStr & "WHERE     (SC09001 = 'RUS') "
        MySQLStr = MySQLStr & "UNION "
        MySQLStr = MySQLStr & "SELECT     14 AS Expr1, SC09016 "
        MySQLStr = MySQLStr & "FROM         SC090300 AS SC090300_27 WITH (NOLOCK) "
        MySQLStr = MySQLStr & "WHERE     (SC09001 = 'RUS') "
        MySQLStr = MySQLStr & "UNION "
        MySQLStr = MySQLStr & "SELECT     15 AS Expr1, SC09017 "
        MySQLStr = MySQLStr & "FROM         SC090300 AS SC090300_26 WITH (NOLOCK) "
        MySQLStr = MySQLStr & "WHERE     (SC09001 = 'RUS') "
        MySQLStr = MySQLStr & "UNION "
        MySQLStr = MySQLStr & "SELECT     16 AS Expr1, SC09018 "
        MySQLStr = MySQLStr & "FROM         SC090300 AS SC090300_25 WITH (NOLOCK) "
        MySQLStr = MySQLStr & "WHERE     (SC09001 = 'RUS') "
        MySQLStr = MySQLStr & "UNION "
        MySQLStr = MySQLStr & "SELECT     17 AS Expr1, SC09019 "
        MySQLStr = MySQLStr & "FROM         SC090300 AS SC090300_24 WITH (NOLOCK) "
        MySQLStr = MySQLStr & "WHERE     (SC09001 = 'RUS') "
        MySQLStr = MySQLStr & "UNION "
        MySQLStr = MySQLStr & "SELECT     18 AS Expr1, SC09020 "
        MySQLStr = MySQLStr & "FROM         SC090300 AS SC090300_23 WITH (NOLOCK) "
        MySQLStr = MySQLStr & "WHERE     (SC09001 = 'RUS') "
        MySQLStr = MySQLStr & "UNION "
        MySQLStr = MySQLStr & "SELECT     19 AS Expr1, SC09021 "
        MySQLStr = MySQLStr & "FROM         SC090300 AS SC090300_22 WITH (NOLOCK) "
        MySQLStr = MySQLStr & "WHERE     (SC09001 = 'RUS') "
        MySQLStr = MySQLStr & "UNION "
        MySQLStr = MySQLStr & "SELECT     20 AS Expr1, SC09022 "
        MySQLStr = MySQLStr & "FROM         SC090300 AS SC090300_21 WITH (NOLOCK) "
        MySQLStr = MySQLStr & "WHERE     (SC09001 = 'RUS') "
        MySQLStr = MySQLStr & "UNION "
        MySQLStr = MySQLStr & "SELECT     21 AS Expr1, SC09023 "
        MySQLStr = MySQLStr & "FROM         SC090300 AS SC090300_20 WITH (NOLOCK) "
        MySQLStr = MySQLStr & "WHERE     (SC09001 = 'RUS') "
        MySQLStr = MySQLStr & "UNION "
        MySQLStr = MySQLStr & "SELECT     22 AS Expr1, SC09024 "
        MySQLStr = MySQLStr & "FROM         SC090300 AS SC090300_19 WITH (NOLOCK) "
        MySQLStr = MySQLStr & "WHERE     (SC09001 = 'RUS') "
        MySQLStr = MySQLStr & "UNION "
        MySQLStr = MySQLStr & "SELECT     23 AS Expr1, SC09025 "
        MySQLStr = MySQLStr & "FROM         SC090300 AS SC090300_18 WITH (NOLOCK) "
        MySQLStr = MySQLStr & "WHERE     (SC09001 = 'RUS') "
        MySQLStr = MySQLStr & "UNION "
        MySQLStr = MySQLStr & "SELECT     24 AS Expr1, SC09026 "
        MySQLStr = MySQLStr & "FROM         SC090300 AS SC090300_17 WITH (NOLOCK) "
        MySQLStr = MySQLStr & "WHERE     (SC09001 = 'RUS') "
        MySQLStr = MySQLStr & "UNION "
        MySQLStr = MySQLStr & "SELECT     25 AS Expr1, SC09027 "
        MySQLStr = MySQLStr & "FROM         SC090300 AS SC090300_16 WITH (NOLOCK) "
        MySQLStr = MySQLStr & "WHERE     (SC09001 = 'RUS') "
        MySQLStr = MySQLStr & "UNION "
        MySQLStr = MySQLStr & "SELECT     26 AS Expr1, SC09028 "
        MySQLStr = MySQLStr & "FROM         SC090300 AS SC090300_15 WITH (NOLOCK) "
        MySQLStr = MySQLStr & "WHERE     (SC09001 = 'RUS') "
        MySQLStr = MySQLStr & "UNION "
        MySQLStr = MySQLStr & "SELECT     27 AS Expr1, SC09029 "
        MySQLStr = MySQLStr & "FROM         SC090300 AS SC090300_14 WITH (NOLOCK) "
        MySQLStr = MySQLStr & "WHERE     (SC09001 = 'RUS') "
        MySQLStr = MySQLStr & "UNION "
        MySQLStr = MySQLStr & "SELECT     28 AS Expr1, SC09030 "
        MySQLStr = MySQLStr & "FROM         SC090300 AS SC090300_13 WITH (NOLOCK) "
        MySQLStr = MySQLStr & "WHERE     (SC09001 = 'RUS') "
        MySQLStr = MySQLStr & "UNION "
        MySQLStr = MySQLStr & "SELECT     29 AS Expr1, SC09031 "
        MySQLStr = MySQLStr & "FROM         SC090300 AS SC090300_12 WITH (NOLOCK) "
        MySQLStr = MySQLStr & "WHERE     (SC09001 = 'RUS') "
        MySQLStr = MySQLStr & "UNION "
        MySQLStr = MySQLStr & "SELECT     30 AS Expr1, SC09032 "
        MySQLStr = MySQLStr & "FROM         SC090300 AS SC090300_11 WITH (NOLOCK) "
        MySQLStr = MySQLStr & "WHERE     (SC09001 = 'RUS') "
        MySQLStr = MySQLStr & "UNION "
        MySQLStr = MySQLStr & "SELECT     31 AS Expr1, SC09033 "
        MySQLStr = MySQLStr & "FROM         SC090300 AS SC090300_10 WITH (NOLOCK) "
        MySQLStr = MySQLStr & "WHERE     (SC09001 = 'RUS') "
        MySQLStr = MySQLStr & "UNION "
        MySQLStr = MySQLStr & "SELECT     32 AS Expr1, SC09034 "
        MySQLStr = MySQLStr & "FROM         SC090300 AS SC090300_9 WITH (NOLOCK) "
        MySQLStr = MySQLStr & "WHERE     (SC09001 = 'RUS') "
        MySQLStr = MySQLStr & "UNION "
        MySQLStr = MySQLStr & "SELECT     33 AS Expr1, SC09035 "
        MySQLStr = MySQLStr & "FROM         SC090300 AS SC090300_8 WITH (NOLOCK) "
        MySQLStr = MySQLStr & "WHERE     (SC09001 = 'RUS') "
        MySQLStr = MySQLStr & "UNION "
        MySQLStr = MySQLStr & "SELECT     34 AS Expr1, SC09036 "
        MySQLStr = MySQLStr & "FROM         SC090300 AS SC090300_7 WITH (NOLOCK) "
        MySQLStr = MySQLStr & "WHERE     (SC09001 = 'RUS') "
        MySQLStr = MySQLStr & "UNION "
        MySQLStr = MySQLStr & "SELECT     35 AS Expr1, SC09037 "
        MySQLStr = MySQLStr & "FROM         SC090300 AS SC090300_6 WITH (NOLOCK) "
        MySQLStr = MySQLStr & "WHERE     (SC09001 = 'RUS') "
        MySQLStr = MySQLStr & "UNION "
        MySQLStr = MySQLStr & "SELECT     36 AS Expr1, SC09038 "
        MySQLStr = MySQLStr & "FROM         SC090300 AS SC090300_5 WITH (NOLOCK) "
        MySQLStr = MySQLStr & "WHERE     (SC09001 = 'RUS') "
        MySQLStr = MySQLStr & "UNION "
        MySQLStr = MySQLStr & "SELECT     37 AS Expr1, SC09039 "
        MySQLStr = MySQLStr & "FROM         SC090300 AS SC090300_4 WITH (NOLOCK) "
        MySQLStr = MySQLStr & "WHERE     (SC09001 = 'RUS') "
        MySQLStr = MySQLStr & "UNION "
        MySQLStr = MySQLStr & "SELECT     38 AS Expr1, SC09040 "
        MySQLStr = MySQLStr & "FROM         SC090300 AS SC090300_3 WITH (NOLOCK) "
        MySQLStr = MySQLStr & "WHERE     (SC09001 = 'RUS') "
        MySQLStr = MySQLStr & "UNION "
        MySQLStr = MySQLStr & "SELECT     39 AS Expr1, SC09041 "
        MySQLStr = MySQLStr & "FROM         SC090300 AS SC090300_2 WITH (NOLOCK) "
        MySQLStr = MySQLStr & "WHERE     (SC09001 = 'RUS') "
        MySQLStr = MySQLStr & "UNION "
        MySQLStr = MySQLStr & "SELECT     40 AS Expr1, SC09042 "
        MySQLStr = MySQLStr & "FROM         SC090300 AS SC090300_1 WITH (NOLOCK)) AS View_1 ON tbl_SupplSearchItems.UOM = View_1.Expr1 "
        MySQLStr = MySQLStr & "WHERE (tbl_SupplSearchItems.SupplSearchID = " & DataGridView1.SelectedRows.Item(0).Cells("ID").Value & ") "
        Select Case MyCol
            Case 0
                MySQLStr = MySQLStr & "ORDER BY tbl_SupplSearchItems.ID "
            Case 2
                MySQLStr = MySQLStr & "ORDER BY SuppItemCode "
            Case 3
                MySQLStr = MySQLStr & "ORDER BY ItemName "
            Case 4
                MySQLStr = MySQLStr & "ORDER BY UOM "
            Case 5
                MySQLStr = MySQLStr & "ORDER BY tbl_SupplSearchItems.QTY "
            Case Else
                MySQLStr = MySQLStr & "ORDER BY tbl_SupplSearchItems.ID "
        End Select
        If MySort = 2 Then
            MySQLStr = MySQLStr & "Desc "
        End If
        InitMyConn(False)
        InitMyRec(False, MySQLStr)
        If Declarations.MyRec.BOF = True And Declarations.MyRec.EOF = True Then
            trycloseMyRec()
        Else
            Declarations.MyRec.MoveFirst()
            i = 0
            While Declarations.MyRec.EOF = False
                oSheet.getCellRangeByName("A" & CStr(i + 11)).Value = i + 1
                oSheet.getCellRangeByName("B" & CStr(i + 11)).String = Declarations.MyRec.Fields("ItemCode").Value
                oSheet.getCellRangeByName("C" & CStr(i + 11)).String = Declarations.MyRec.Fields("SuppItemCode").Value
                oSheet.getCellRangeByName("D" & CStr(i + 11)).String = Declarations.MyRec.Fields("ItemName").Value
                oSheet.getCellRangeByName("E" & CStr(i + 11)).String = Declarations.MyRec.Fields("UOM").Value
                oSheet.getCellRangeByName("F" & CStr(i + 11)).Value = Declarations.MyRec.Fields("QTY").Value
                'oSheet.getCellRangeByName("G" & CStr(i + 11)).Value = Declarations.MyRec.Fields("Price").Value
                oSheet.getCellRangeByName("H" & CStr(i + 11)).FormulaLocal = "=IF(F" & CStr(i + 11) & "*G" & CStr(i + 11) & " = 0;"""";F" & CStr(i + 11) & " * G" & CStr(i + 11) & ") "
                oSheet.getCellRangeByName("I" & CStr(i + 11)).Value = Declarations.MyRec.Fields("WeekQTY").Value

                i = i + 1
                Declarations.MyRec.MoveNext()
            End While
            trycloseMyRec()
        End If
        LOWrapText(oServiceManager, oDispatcher, oFrame, "A11:I" & CStr(11 + i - 1))
        LineFormat.LineStyle = 0
        LineFormat.LineWidth = 20
        oSheet.getCellRangeByName("A11:I" & CStr(11 + i - 1)).TopBorder = LineFormat
        oSheet.getCellRangeByName("A11:I" & CStr(11 + i - 1)).RightBorder = LineFormat
        oSheet.getCellRangeByName("A11:I" & CStr(11 + i - 1)).LeftBorder = LineFormat
        oSheet.getCellRangeByName("A11:I" & CStr(11 + i - 1)).BottomBorder = LineFormat
        '----������ �����
        LOSetCellProtection(oServiceManager, oDispatcher, oFrame, "A11:I500", False)
        '---�������� ���������� ����������
        oSheet.getCellRangeByName("N10").String = "������� ���������"
        oSheet.getCellRangeByName("N11").Value = 0
        oSheet.getCellRangeByName("O11").String = "pcs(��.)"
        oSheet.getCellRangeByName("N12").Value = 1
        oSheet.getCellRangeByName("O12").String = "m (�)"
        oSheet.getCellRangeByName("N13").Value = 2
        oSheet.getCellRangeByName("O13").String = "kg (��)"
        oSheet.getCellRangeByName("N14").Value = 3
        oSheet.getCellRangeByName("O14").String = "km (��)"
        oSheet.getCellRangeByName("N15").Value = 4
        oSheet.getCellRangeByName("O15").String = "litre (����)"
        oSheet.getCellRangeByName("N16").Value = 5
        oSheet.getCellRangeByName("O16").String = "pack (����.)"
        oSheet.getCellRangeByName("N17").Value = 6
        oSheet.getCellRangeByName("O17").String = "set (�����.)"
        oSheet.getCellRangeByName("N18").Value = 7
        oSheet.getCellRangeByName("O18").String = "pair (����)"
        oSheet.getCellRangeByName("N10:O18").CharColor = RGB(255, 255, 255) '---��� ����� �������� ������� - �� ����� ���� ��� �� RGB � BGR!!!! ������ ���� �����
        LOSetValidation(oSheet, "E11:E" & CStr(11 + i - 1), "=$O$11:$O$18")
        '----� ������ �����
        Dim args() As Object
        ReDim args(0)
        args(0) = oServiceManager.Bridge_GetStruct("com.sun.star.beans.PropertyValue")
        args(0).Name = "ToPoint"
        args(0).Value = "$A$1"
        oDispatcher.executeDispatch(oFrame, ".uno:GoToCell", "", 0, args)
        '----�������� �������
        LOPasswordProtect(oSheet, "!pass2022", True)
        '----���������
        oWorkBook.CurrentController.Frame.ContainerWindow.Visible = True
        oWorkBook.CurrentController.Frame.ContainerWindow.toFront()
    End Sub

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        '////////////////////////////////////////////////////////////////////////////////
        '//
        '// �������� ������������� ������� �� ������������
        '//
        '////////////////////////////////////////////////////////////////////////////////
        Dim MyRez As MsgBoxResult
        Dim MyTxt As String

        MyTxt = "��� ������� ������ ��� ���������� ������������ ���� Excel, � ������� ������� �� ������ 11 �������: " & Chr(13) & Chr(10)
        MyTxt = MyTxt & "� ������� A ������� ����� �� ������� (�� ������������ ��������), " & Chr(13) & Chr(10)
        MyTxt = MyTxt & "� ������� B ������� ��� ������ Scala. " & Chr(13) & Chr(10)
        MyTxt = MyTxt & "� ������� C ������� ��� ������ ����������. " & Chr(13) & Chr(10)
        MyTxt = MyTxt & "� ������� D ������� �������� ������ . " & Chr(13) & Chr(10)
        MyTxt = MyTxt & "� ������� E ������� ������� ��������� ������ . " & Chr(13) & Chr(10)
        MyTxt = MyTxt & "� ������� F ���������� ���������� ������ . " & Chr(13) & Chr(10)
        MyTxt = MyTxt & "� ������� G ���������� ���� ������ ��� ���. " & Chr(13) & Chr(10)
        MyTxt = MyTxt & "� ������� I ���������� ���� �������� � �������. " & Chr(13) & Chr(10)
        MyTxt = MyTxt & "���� ����� � ������� - ���� �������� = 0. " & Chr(13) & Chr(10)
        MyTxt = MyTxt & "������ ������ ���� ��������� ��� ���������. " & Chr(13) & Chr(10)
        MyTxt = MyTxt & "��� ������� ������ ���� ���������, ����� B � C:" & Chr(13) & Chr(10)
        MyTxt = MyTxt & "� ��� ����� ������� ��� ��� ������ Scala, " & Chr(13) & Chr(10)
        MyTxt = MyTxt & "��� ��� ������ ���������� (����� ����� ������� ���) " & Chr(13) & Chr(10)
        MyTxt = MyTxt & "� ��� ���� �������������� ���� Excel � �� ������ ������ ������?" & Chr(13) & Chr(10)
        MyRez = MsgBox(MyTxt, MsgBoxStyle.OkCancel, "��������!")

        If (MyRez = MsgBoxResult.Ok) Then
            If My.Settings.UseOffice = "LibreOffice" Then
                OpenFileDialog2.ShowDialog()
                If OpenFileDialog2.FileName.Equals("") = False Then
                    Declarations.ImportFileName = OpenFileDialog2.FileName
                    Me.Cursor = Cursors.WaitCursor
                    Me.Refresh()
                    System.Windows.Forms.Application.DoEvents()
                    Declarations.MyRequestNum = DataGridView1.SelectedRows.Item(0).Cells("ID").Value
                    ImportRequestDataFromLO()
                End If
            Else
                OpenFileDialog1.ShowDialog()
                If OpenFileDialog1.FileName.Equals("") = False Then
                    Declarations.ImportFileName = OpenFileDialog1.FileName
                    Me.Cursor = Cursors.WaitCursor
                    Me.Refresh()
                    System.Windows.Forms.Application.DoEvents()
                    Declarations.MyRequestNum = DataGridView1.SelectedRows.Item(0).Cells("ID").Value
                    ImportRequestDataFromExcel()
                End If
            End If
            Me.Cursor = Cursors.Default
            LoadItems()
            CheckItemButtons()
        End If
    End Sub

    Private Sub ComboBox1_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles ComboBox1.SelectedIndexChanged
        '////////////////////////////////////////////////////////////////////////////////
        '//
        '// ����� ������ - ��� ���������� � �����������
        '//
        '////////////////////////////////////////////////////////////////////////////////

        LoadProposal()
        CheckProposalButtons()
    End Sub

    Private Sub Button15_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button15.Click
        '////////////////////////////////////////////////////////////////////////////////
        '//
        '// �������� ����������� �� ����������� � ������������ � �����. ������
        '//
        '////////////////////////////////////////////////////////////////////////////////

        If My.Settings.UseOffice = "LibreOffice" Then
            If DataGridView3.SortedColumn Is Nothing Then
                ExportPropPurchToLO(0, 1)
            Else
                ExportPropPurchToLO(DataGridView3.SortedColumn.Index, DataGridView3.SortOrder)
            End If

        Else
            If DataGridView3.SortedColumn Is Nothing Then
                ExportPropPurchToExcel(0, 1)
            Else
                ExportPropPurchToExcel(DataGridView3.SortedColumn.Index, DataGridView3.SortOrder)
            End If
        End If
    End Sub

    Private Sub ExportPropPurchToExcel(ByVal MyCol As Integer, ByVal MySort As Integer)
        '////////////////////////////////////////////////////////////////////////////////
        '//
        '// �������� ����������� �� ����������� � ������������ � �����. ������ � Excel
        '//
        '////////////////////////////////////////////////////////////////////////////////
        Dim MySQLStr As String
        Dim SpecVersion As String               '--������ ������������
        Dim MyObj As Object
        Dim MyWRKBook As Object
        Dim i As Integer

        MyObj = CreateObject("Excel.Application")
        Dim oldCI As System.Globalization.CultureInfo = System.Threading.Thread.CurrentThread.CurrentCulture
        System.Threading.Thread.CurrentThread.CurrentCulture = New System.Globalization.CultureInfo("en-US")

        MyObj.SheetsInNewWorkbook = 1
        MyObj.ReferenceStyle = 1
        MyWRKBook = MyObj.Workbooks.Add

        '---���������
        '---������ ������������
        MySQLStr = "SELECT Version "
        MySQLStr = MySQLStr & "FROM tbl_VersionImportItemsFromExcel WITH (NOLOCK) "
        MySQLStr = MySQLStr & "WHERE (Name = N'������������ �����������') "
        InitMyConn(False)
        InitMyRec(False, MySQLStr)
        If Declarations.MyRec.BOF = True And Declarations.MyRec.EOF = True Then
            MsgBox("� Scala �� ����������� ������� ������ ����� Excel. ���������� � ��������������.", vbCritical, "��������!")
            trycloseMyRec()
            Exit Sub
        Else
            SpecVersion = Trim(Declarations.MyRec.Fields("Version").Value)
            trycloseMyRec()
        End If
        MyWRKBook.ActiveSheet.Range("A1") = SpecVersion
        With MyWRKBook.ActiveSheet.Range("A1").Font
            .Name = "Calibri"
            .Size = 9
            '.Color = -16776961
            .ColorIndex = 3
        End With

        MyWRKBook.ActiveSheet.Range("B2") = "Skandika"
        With MyWRKBook.ActiveSheet.Range("B2").Font
            .Name = "Calibri"
            .Size = 16
            .Color = 0
            .Bold = True
        End With

        MyWRKBook.ActiveSheet.Range("A4") = "OOO ""��������"""
        With MyWRKBook.ActiveSheet.Range("A4").Font
            .Name = "Tahoma"
            .Size = 9
            .Color = 0
            .Bold = True
        End With

        MyWRKBook.ActiveSheet.Range("A5") = "�����:"
        With MyWRKBook.ActiveSheet.Range("A5").Font
            .Name = "Tahoma"
            .Size = 9
            .Color = 0
            .Bold = True
        End With

        MyWRKBook.ActiveSheet.Range("B5:H6").MergeCells = True
        MyWRKBook.ActiveSheet.Range("B5") = "������,195027, ����� ���������, ������� ��., �.4 Tel: +7 (812) 325 2040, Fax: +7 (812) 325 2039, www.skandikagroup.ru"
        With MyWRKBook.ActiveSheet.Range("B5").Font
            .Name = "Tahoma"
            .Size = 9
            .Color = 0
            .Bold = True
        End With
        MyWRKBook.ActiveSheet.Range("B5:H6").WrapText = True
        MyWRKBook.ActiveSheet.Range("B5:H6").VerticalAlignment = -4160

        MyWRKBook.ActiveSheet.Range("D8") = "������������ ��������"
        With MyWRKBook.ActiveSheet.Range("D8").Font
            .Name = "Tahoma"
            .Size = 11.5
            .Color = 0
            .Bold = True
        End With

        '--��������� �������
        MyWRKBook.ActiveSheet.Range("A10") = "N �/�"
        MyWRKBook.ActiveSheet.Range("B10") = "��� ������ Scala"
        MyWRKBook.ActiveSheet.Range("C10") = "��� ������ ����������"
        MyWRKBook.ActiveSheet.Range("D10") = "������������ ������"
        MyWRKBook.ActiveSheet.Range("E10") = "�� ���������"
        MyWRKBook.ActiveSheet.Range("F10") = "���-��"
        MyWRKBook.ActiveSheet.Range("G10") = "���� ��� ���"
        MyWRKBook.ActiveSheet.Range("H10") = "����� ��� ���"
        MyWRKBook.ActiveSheet.Range("I10") = "���� �������� (���)"
        MyWRKBook.ActiveSheet.Range("J10") = "������������ ������"
        MyWRKBook.ActiveSheet.Range("K10") = "�������. ����������"
        MyWRKBook.ActiveSheet.Range("L10") = "��� ����������"
        MyWRKBook.ActiveSheet.Range("A10:I10").Select()
        MyWRKBook.ActiveSheet.Range("A10:I10").Borders(5).LineStyle = -4142
        MyWRKBook.ActiveSheet.Range("A10:I10").Borders(6).LineStyle = -4142
        With MyWRKBook.ActiveSheet.Range("A10:I10").Borders(7)
            .LineStyle = 1
            .Weight = 2
            .ColorIndex = -4105
        End With
        With MyWRKBook.ActiveSheet.Range("A10:I10").Borders(8)
            .LineStyle = 1
            .Weight = 2
            .ColorIndex = -4105
        End With
        With MyWRKBook.ActiveSheet.Range("A10:I10").Borders(9)
            .LineStyle = 1
            .Weight = 2
            .ColorIndex = -4105
        End With
        With MyWRKBook.ActiveSheet.Range("A10:I10").Borders(10)
            .LineStyle = 1
            .Weight = 2
            .ColorIndex = -4105
        End With
        With MyWRKBook.ActiveSheet.Range("A10:I10").Borders(11)
            .LineStyle = 1
            .Weight = 2
            .ColorIndex = -4105
        End With
        MyWRKBook.ActiveSheet.Range("A10:K10").WrapText = True
        MyWRKBook.ActiveSheet.Range("A10:K10").VerticalAlignment = -4108
        MyWRKBook.ActiveSheet.Range("A10:K10").HorizontalAlignment = -4108
        With MyWRKBook.ActiveSheet.Range("A10:K10").Font
            .Name = "Calibri"
            .Size = 9
            .Color = 0
            .Bold = False
        End With

        MyWRKBook.ActiveSheet.Columns("A:A").ColumnWidth = 4
        MyWRKBook.ActiveSheet.Columns("B:B").ColumnWidth = 10
        MyWRKBook.ActiveSheet.Columns("C:C").ColumnWidth = 10
        MyWRKBook.ActiveSheet.Columns("D:D").ColumnWidth = 30
        MyWRKBook.ActiveSheet.Columns("E:E").ColumnWidth = 5
        MyWRKBook.ActiveSheet.Columns("F:F").ColumnWidth = 8
        MyWRKBook.ActiveSheet.Columns("G:G").ColumnWidth = 10
        MyWRKBook.ActiveSheet.Columns("H:H").ColumnWidth = 12
        MyWRKBook.ActiveSheet.Columns("I:I").ColumnWidth = 12
        MyWRKBook.ActiveSheet.Columns("J:J").ColumnWidth = 30
        MyWRKBook.ActiveSheet.Columns("K:K").ColumnWidth = 30
        MyWRKBook.ActiveSheet.Columns("L:L").ColumnWidth = 30

        '--����� ����� ������������
        MySQLStr = "SELECT ISNULL(tbl_SupplSearch_PropItems.ItemCode, N'') AS ItemCode, ISNULL(tbl_SupplSearch_PropItems.ItemSuppCode, N'') AS SuppItemCode, "
        MySQLStr = MySQLStr & "CASE WHEN LTRIM(RTRIM(LTRIM(RTRIM(ISNULL(SC010300.SC01002, ''))) + LTRIM(RTRIM(ISNULL(SC010300.SC01003, ''))))) "
        MySQLStr = MySQLStr & "= '' THEN ISNULL(tbl_SupplSearch_PropItems.ItemName, '') ELSE LTRIM(RTRIM(LTRIM(RTRIM(ISNULL(SC010300.SC01002, ''))) "
        MySQLStr = MySQLStr & "+ LTRIM(RTRIM(ISNULL(SC010300.SC01003, ''))))) END AS ItemName, CASE WHEN ISNULL(View_1.txt, N'') "
        MySQLStr = MySQLStr & "= '��' THEN 'pcs(��.)' ELSE CASE WHEN ISNULL(View_1.txt, N'') = '�' THEN 'm (�)' ELSE CASE WHEN ISNULL(View_1.txt, N'') "
        MySQLStr = MySQLStr & "= '��' THEN 'kg (��)' ELSE CASE WHEN ISNULL(View_1.txt, N'') = '��' THEN 'km (��)' ELSE CASE WHEN ISNULL(View_1.txt, N'') "
        MySQLStr = MySQLStr & "= '�' THEN 'litre (����)' ELSE CASE WHEN ISNULL(View_1.txt, N'') = '����' THEN 'pack (����.)' ELSE CASE WHEN ISNULL(View_1.txt, N'') "
        MySQLStr = MySQLStr & "= '�����' THEN 'set (�����.)' ELSE CASE WHEN ISNULL(View_1.txt, N'') = '���' THEN 'pair (����)' END END END END END END END END AS UOM, "
        MySQLStr = MySQLStr & "tbl_SupplSearch_PropItems.QTY, tbl_SupplSearch_PropItems.Price, ISNULL(tbl_SupplSearch_PropItems.LeadTimeWeek, 1) AS WeekQTY, "
        MySQLStr = MySQLStr & "ISNULL(tbl_SupplSearch_PropItems.AlternateTo, '') AS AlternateTo, ISNULL(tbl_SupplSearch_PropItems.Comments, '') AS Comments, "
        MySQLStr = MySQLStr & "ISNULL(tbl_SupplSearch_PropItems.SupplierCode, '') AS SupplierCode "
        MySQLStr = MySQLStr & "FROM tbl_SupplSearch_PropItems LEFT OUTER JOIN "
        MySQLStr = MySQLStr & "SC010300 ON tbl_SupplSearch_PropItems.ItemCode = SC010300.SC01001 LEFT OUTER JOIN "
        MySQLStr = MySQLStr & "(SELECT     0 AS Expr1, SC09002 AS txt "
        MySQLStr = MySQLStr & "FROM          SC090300 WITH (NOLOCK) "
        MySQLStr = MySQLStr & "WHERE      (SC09001 = 'RUS') "
        MySQLStr = MySQLStr & "UNION "
        MySQLStr = MySQLStr & "SELECT     1 AS Expr1, SC09003 "
        MySQLStr = MySQLStr & "FROM         SC090300 AS SC090300_40 WITH (NOLOCK) "
        MySQLStr = MySQLStr & "WHERE     (SC09001 = 'RUS') "
        MySQLStr = MySQLStr & "UNION "
        MySQLStr = MySQLStr & "SELECT     2 AS Expr1, SC09004 "
        MySQLStr = MySQLStr & "FROM         SC090300 AS SC090300_39 WITH (NOLOCK) "
        MySQLStr = MySQLStr & "WHERE     (SC09001 = 'RUS') "
        MySQLStr = MySQLStr & "UNION "
        MySQLStr = MySQLStr & "SELECT     3 AS Expr1, SC09005 "
        MySQLStr = MySQLStr & "FROM         SC090300 AS SC090300_38 WITH (NOLOCK) "
        MySQLStr = MySQLStr & "WHERE     (SC09001 = 'RUS') "
        MySQLStr = MySQLStr & "UNION "
        MySQLStr = MySQLStr & "SELECT     4 AS Expr1, SC09006 "
        MySQLStr = MySQLStr & "FROM         SC090300 AS SC090300_37 WITH (NOLOCK) "
        MySQLStr = MySQLStr & "WHERE     (SC09001 = 'RUS') "
        MySQLStr = MySQLStr & "UNION "
        MySQLStr = MySQLStr & "SELECT     5 AS Expr1, SC09007 "
        MySQLStr = MySQLStr & "FROM         SC090300 AS SC090300_36 WITH (NOLOCK) "
        MySQLStr = MySQLStr & "WHERE     (SC09001 = 'RUS') "
        MySQLStr = MySQLStr & "UNION "
        MySQLStr = MySQLStr & "SELECT     6 AS Expr1, SC09008 "
        MySQLStr = MySQLStr & "FROM         SC090300 AS SC090300_35 WITH (NOLOCK) "
        MySQLStr = MySQLStr & "WHERE     (SC09001 = 'RUS') "
        MySQLStr = MySQLStr & "UNION "
        MySQLStr = MySQLStr & "SELECT     7 AS Expr1, SC09009 "
        MySQLStr = MySQLStr & "FROM         SC090300 AS SC090300_34 WITH (NOLOCK) "
        MySQLStr = MySQLStr & "WHERE     (SC09001 = 'RUS') "
        MySQLStr = MySQLStr & "UNION "
        MySQLStr = MySQLStr & "SELECT     8 AS Expr1, SC09010 "
        MySQLStr = MySQLStr & "FROM         SC090300 AS SC090300_33 WITH (NOLOCK) "
        MySQLStr = MySQLStr & "WHERE     (SC09001 = 'RUS') "
        MySQLStr = MySQLStr & "UNION "
        MySQLStr = MySQLStr & "SELECT     9 AS Expr1, SC09011 "
        MySQLStr = MySQLStr & "FROM         SC090300 AS SC090300_32 WITH (NOLOCK) "
        MySQLStr = MySQLStr & "WHERE     (SC09001 = 'RUS') "
        MySQLStr = MySQLStr & "UNION "
        MySQLStr = MySQLStr & "SELECT     10 AS Expr1, SC09012 "
        MySQLStr = MySQLStr & "FROM         SC090300 AS SC090300_31 WITH (NOLOCK) "
        MySQLStr = MySQLStr & "WHERE     (SC09001 = 'RUS') "
        MySQLStr = MySQLStr & "UNION "
        MySQLStr = MySQLStr & "SELECT     11 AS Expr1, SC09013 "
        MySQLStr = MySQLStr & "FROM         SC090300 AS SC090300_30 WITH (NOLOCK) "
        MySQLStr = MySQLStr & "WHERE     (SC09001 = 'RUS') "
        MySQLStr = MySQLStr & "UNION "
        MySQLStr = MySQLStr & "SELECT     12 AS Expr1, SC09014 "
        MySQLStr = MySQLStr & "FROM         SC090300 AS SC090300_29 WITH (NOLOCK) "
        MySQLStr = MySQLStr & "WHERE     (SC09001 = 'RUS') "
        MySQLStr = MySQLStr & "UNION "
        MySQLStr = MySQLStr & "SELECT     13 AS Expr1, SC09015 "
        MySQLStr = MySQLStr & "FROM         SC090300 AS SC090300_28 WITH (NOLOCK) "
        MySQLStr = MySQLStr & "WHERE     (SC09001 = 'RUS') "
        MySQLStr = MySQLStr & "UNION "
        MySQLStr = MySQLStr & "SELECT     14 AS Expr1, SC09016 "
        MySQLStr = MySQLStr & "FROM         SC090300 AS SC090300_27 WITH (NOLOCK) "
        MySQLStr = MySQLStr & "WHERE     (SC09001 = 'RUS') "
        MySQLStr = MySQLStr & "UNION "
        MySQLStr = MySQLStr & "SELECT     15 AS Expr1, SC09017 "
        MySQLStr = MySQLStr & "FROM         SC090300 AS SC090300_26 WITH (NOLOCK) "
        MySQLStr = MySQLStr & "WHERE     (SC09001 = 'RUS') "
        MySQLStr = MySQLStr & "UNION "
        MySQLStr = MySQLStr & "SELECT     16 AS Expr1, SC09018 "
        MySQLStr = MySQLStr & "FROM         SC090300 AS SC090300_25 WITH (NOLOCK) "
        MySQLStr = MySQLStr & "WHERE     (SC09001 = 'RUS') "
        MySQLStr = MySQLStr & "UNION "
        MySQLStr = MySQLStr & "SELECT     17 AS Expr1, SC09019 "
        MySQLStr = MySQLStr & "FROM         SC090300 AS SC090300_24 WITH (NOLOCK) "
        MySQLStr = MySQLStr & "WHERE     (SC09001 = 'RUS') "
        MySQLStr = MySQLStr & "UNION "
        MySQLStr = MySQLStr & "SELECT     18 AS Expr1, SC09020 "
        MySQLStr = MySQLStr & "FROM         SC090300 AS SC090300_23 WITH (NOLOCK) "
        MySQLStr = MySQLStr & "WHERE     (SC09001 = 'RUS') "
        MySQLStr = MySQLStr & "UNION "
        MySQLStr = MySQLStr & "SELECT     19 AS Expr1, SC09021 "
        MySQLStr = MySQLStr & "FROM         SC090300 AS SC090300_22 WITH (NOLOCK) "
        MySQLStr = MySQLStr & "WHERE     (SC09001 = 'RUS') "
        MySQLStr = MySQLStr & "UNION "
        MySQLStr = MySQLStr & "SELECT     20 AS Expr1, SC09022 "
        MySQLStr = MySQLStr & "FROM         SC090300 AS SC090300_21 WITH (NOLOCK) "
        MySQLStr = MySQLStr & "WHERE     (SC09001 = 'RUS') "
        MySQLStr = MySQLStr & "UNION "
        MySQLStr = MySQLStr & "SELECT     21 AS Expr1, SC09023 "
        MySQLStr = MySQLStr & "FROM         SC090300 AS SC090300_20 WITH (NOLOCK) "
        MySQLStr = MySQLStr & "WHERE     (SC09001 = 'RUS') "
        MySQLStr = MySQLStr & "UNION "
        MySQLStr = MySQLStr & "SELECT     22 AS Expr1, SC09024 "
        MySQLStr = MySQLStr & "FROM         SC090300 AS SC090300_19 WITH (NOLOCK) "
        MySQLStr = MySQLStr & "WHERE     (SC09001 = 'RUS') "
        MySQLStr = MySQLStr & "UNION "
        MySQLStr = MySQLStr & "SELECT     23 AS Expr1, SC09025 "
        MySQLStr = MySQLStr & "FROM         SC090300 AS SC090300_18 WITH (NOLOCK) "
        MySQLStr = MySQLStr & "WHERE     (SC09001 = 'RUS') "
        MySQLStr = MySQLStr & "UNION "
        MySQLStr = MySQLStr & "SELECT     24 AS Expr1, SC09026 "
        MySQLStr = MySQLStr & "FROM         SC090300 AS SC090300_17 WITH (NOLOCK) "
        MySQLStr = MySQLStr & "WHERE     (SC09001 = 'RUS') "
        MySQLStr = MySQLStr & "UNION "
        MySQLStr = MySQLStr & "SELECT     25 AS Expr1, SC09027 "
        MySQLStr = MySQLStr & "FROM         SC090300 AS SC090300_16 WITH (NOLOCK) "
        MySQLStr = MySQLStr & "WHERE     (SC09001 = 'RUS') "
        MySQLStr = MySQLStr & "UNION "
        MySQLStr = MySQLStr & "SELECT     26 AS Expr1, SC09028 "
        MySQLStr = MySQLStr & "FROM         SC090300 AS SC090300_15 WITH (NOLOCK) "
        MySQLStr = MySQLStr & "WHERE     (SC09001 = 'RUS') "
        MySQLStr = MySQLStr & "UNION "
        MySQLStr = MySQLStr & "SELECT     27 AS Expr1, SC09029 "
        MySQLStr = MySQLStr & "FROM         SC090300 AS SC090300_14 WITH (NOLOCK) "
        MySQLStr = MySQLStr & "WHERE     (SC09001 = 'RUS') "
        MySQLStr = MySQLStr & "UNION "
        MySQLStr = MySQLStr & "SELECT     28 AS Expr1, SC09030 "
        MySQLStr = MySQLStr & "FROM         SC090300 AS SC090300_13 WITH (NOLOCK) "
        MySQLStr = MySQLStr & "WHERE     (SC09001 = 'RUS') "
        MySQLStr = MySQLStr & "UNION "
        MySQLStr = MySQLStr & "SELECT     29 AS Expr1, SC09031 "
        MySQLStr = MySQLStr & "FROM         SC090300 AS SC090300_12 WITH (NOLOCK) "
        MySQLStr = MySQLStr & "WHERE     (SC09001 = 'RUS') "
        MySQLStr = MySQLStr & "UNION "
        MySQLStr = MySQLStr & "SELECT     30 AS Expr1, SC09032 "
        MySQLStr = MySQLStr & "FROM         SC090300 AS SC090300_11 WITH (NOLOCK) "
        MySQLStr = MySQLStr & "WHERE     (SC09001 = 'RUS') "
        MySQLStr = MySQLStr & "UNION "
        MySQLStr = MySQLStr & "SELECT     31 AS Expr1, SC09033 "
        MySQLStr = MySQLStr & "FROM         SC090300 AS SC090300_10 WITH (NOLOCK) "
        MySQLStr = MySQLStr & "WHERE     (SC09001 = 'RUS') "
        MySQLStr = MySQLStr & "UNION "
        MySQLStr = MySQLStr & "SELECT     32 AS Expr1, SC09034 "
        MySQLStr = MySQLStr & "FROM         SC090300 AS SC090300_9 WITH (NOLOCK) "
        MySQLStr = MySQLStr & "WHERE     (SC09001 = 'RUS') "
        MySQLStr = MySQLStr & "UNION "
        MySQLStr = MySQLStr & "SELECT     33 AS Expr1, SC09035 "
        MySQLStr = MySQLStr & "FROM         SC090300 AS SC090300_8 WITH (NOLOCK) "
        MySQLStr = MySQLStr & "WHERE     (SC09001 = 'RUS') "
        MySQLStr = MySQLStr & "UNION "
        MySQLStr = MySQLStr & "SELECT     34 AS Expr1, SC09036 "
        MySQLStr = MySQLStr & "FROM         SC090300 AS SC090300_7 WITH (NOLOCK) "
        MySQLStr = MySQLStr & "WHERE     (SC09001 = 'RUS') "
        MySQLStr = MySQLStr & "UNION "
        MySQLStr = MySQLStr & "SELECT     35 AS Expr1, SC09037 "
        MySQLStr = MySQLStr & "FROM         SC090300 AS SC090300_6 WITH (NOLOCK) "
        MySQLStr = MySQLStr & "WHERE     (SC09001 = 'RUS') "
        MySQLStr = MySQLStr & "UNION "
        MySQLStr = MySQLStr & "SELECT     36 AS Expr1, SC09038 "
        MySQLStr = MySQLStr & "FROM         SC090300 AS SC090300_5 WITH (NOLOCK) "
        MySQLStr = MySQLStr & "WHERE     (SC09001 = 'RUS') "
        MySQLStr = MySQLStr & "UNION "
        MySQLStr = MySQLStr & "SELECT     37 AS Expr1, SC09039 "
        MySQLStr = MySQLStr & "FROM         SC090300 AS SC090300_4 WITH (NOLOCK) "
        MySQLStr = MySQLStr & "WHERE     (SC09001 = 'RUS') "
        MySQLStr = MySQLStr & "UNION "
        MySQLStr = MySQLStr & "SELECT     38 AS Expr1, SC09040 "
        MySQLStr = MySQLStr & "FROM         SC090300 AS SC090300_3 WITH (NOLOCK) "
        MySQLStr = MySQLStr & "WHERE     (SC09001 = 'RUS') "
        MySQLStr = MySQLStr & "UNION "
        MySQLStr = MySQLStr & "SELECT     39 AS Expr1, SC09041 "
        MySQLStr = MySQLStr & "FROM         SC090300 AS SC090300_2 WITH (NOLOCK) "
        MySQLStr = MySQLStr & "WHERE     (SC09001 = 'RUS') "
        MySQLStr = MySQLStr & "UNION "
        MySQLStr = MySQLStr & "SELECT     40 AS Expr1, SC09042 "
        MySQLStr = MySQLStr & "FROM SC090300 AS SC090300_1 WITH (NOLOCK)) AS View_1 ON tbl_SupplSearch_PropItems.UOM = View_1.Expr1 "
        MySQLStr = MySQLStr & "WHERE (tbl_SupplSearch_PropItems.SupplSearchID = " & DataGridView1.SelectedRows.Item(0).Cells("ID").Value & ") "
        If ComboBox1.Text = "������������ ������������" Then
            MySQLStr = MySQLStr & "AND (tbl_SupplSearch_PropItems.IsSelected = 1) "
        End If
        If ComboBox1.Text = "��������� ���������" Then
            MySQLStr = MySQLStr & "AND (tbl_SupplSearch_PropItems.SelectedBySalesman = 1) "
        End If
        Select Case MyCol
            Case 0
                MySQLStr = MySQLStr & "ORDER BY tbl_SupplSearch_PropItems.ID "
            Case 4
                MySQLStr = MySQLStr & "ORDER BY ItemCode "
            Case 5
                MySQLStr = MySQLStr & "ORDER BY SuppItemCode "
            Case 6
                MySQLStr = MySQLStr & "ORDER BY ItemName "
            Case 8
                MySQLStr = MySQLStr & "ORDER BY UOM "
            Case 9
                MySQLStr = MySQLStr & "ORDER BY tbl_SupplSearch_PropItems.QTY "
            Case Else
                MySQLStr = MySQLStr & "ORDER BY tbl_SupplSearch_PropItems.ID "
        End Select
        If MySort = 2 Then
            MySQLStr = MySQLStr & "Desc "
        End If

        InitMyConn(False)
        InitMyRec(False, MySQLStr)
        If Declarations.MyRec.BOF = True And Declarations.MyRec.EOF = True Then
            trycloseMyRec()
        Else
            Declarations.MyRec.MoveFirst()
            i = 0
            While Declarations.MyRec.EOF = False
                MyWRKBook.ActiveSheet.Range("A" & CStr(i + 11)).NumberFormat = "@"
                MyWRKBook.ActiveSheet.Range("A" & CStr(i + 11)) = i + 1
                MyWRKBook.ActiveSheet.Range("B" & CStr(i + 11)).NumberFormat = "@"
                MyWRKBook.ActiveSheet.Range("B" & CStr(i + 11)) = Declarations.MyRec.Fields("ItemCode").Value
                MyWRKBook.ActiveSheet.Range("C" & CStr(i + 11)).NumberFormat = "@"
                MyWRKBook.ActiveSheet.Range("C" & CStr(i + 11)) = Declarations.MyRec.Fields("SuppItemCode").Value
                MyWRKBook.ActiveSheet.Range("D" & CStr(i + 11)) = Declarations.MyRec.Fields("ItemName").Value
                MyWRKBook.ActiveSheet.Range("E" & CStr(i + 11)) = Declarations.MyRec.Fields("UOM").Value
                MyWRKBook.ActiveSheet.Range("F" & CStr(i + 11)) = Declarations.MyRec.Fields("QTY").Value
                MyWRKBook.ActiveSheet.Range("G" & CStr(i + 11)) = Declarations.MyRec.Fields("Price").Value
                MyWRKBook.ActiveSheet.Range("I" & CStr(i + 11)) = Declarations.MyRec.Fields("WeekQTY").Value
                MyWRKBook.ActiveSheet.Range("J" & CStr(i + 11)) = Declarations.MyRec.Fields("AlternateTo").Value
                MyWRKBook.ActiveSheet.Range("K" & CStr(i + 11)) = Declarations.MyRec.Fields("Comments").Value
                MyWRKBook.ActiveSheet.Range("L" & CStr(i + 11)) = Declarations.MyRec.Fields("SupplierCode").Value

                i = i + 1
                Declarations.MyRec.MoveNext()
            End While
            trycloseMyRec()
        End If

        '---

        MyWRKBook.ActiveSheet.Range("A11:I11").Select()
        MyWRKBook.ActiveSheet.Range("A11:I11").Borders(5).LineStyle = -4142
        MyWRKBook.ActiveSheet.Range("A11:I11").Borders(6).LineStyle = -4142
        With MyWRKBook.ActiveSheet.Range("A11:I1011").Borders(7)
            .LineStyle = 1
            .Weight = 2
            .ColorIndex = -4105
        End With
        With MyWRKBook.ActiveSheet.Range("A11:I1011").Borders(8)
            .LineStyle = 1
            .Weight = 2
            .ColorIndex = -4105
        End With
        With MyWRKBook.ActiveSheet.Range("A11:I1011").Borders(9)
            .LineStyle = 1
            .Weight = 2
            .ColorIndex = -4105
        End With
        With MyWRKBook.ActiveSheet.Range("A11:I1011").Borders(10)
            .LineStyle = 1
            .Weight = 2
            .ColorIndex = -4105
        End With
        With MyWRKBook.ActiveSheet.Range("A11:I1011").Borders(11)
            .LineStyle = 1
            .Weight = 2
            .ColorIndex = -4105
        End With
        With MyWRKBook.ActiveSheet.Range("A11:I1011").Borders(12)
            .LineStyle = 1
            .Weight = 2
            .ColorIndex = -4105
        End With
        With MyWRKBook.ActiveSheet.Range("A11:I1011").Font
            .Name = "Calibri"
            .Size = 9
            .Color = 0
            .Bold = False
        End With
        With MyWRKBook.ActiveSheet.Range("H11:H1011")
            '.FormulaR1C1 = "=����(RC[-2]*RC[-1] = 0; """"; RC[-2]*RC[-1])"
            .FormulaR1C1 = "=IF(RC[-2]*RC[-1] = 0, """", RC[-2]*RC[-1])"
        End With
        System.Threading.Thread.CurrentThread.CurrentCulture = oldCI

        '---�������� ���������� ����������
        MyWRKBook.ActiveSheet.Range("N10") = "������� ���������"
        MyWRKBook.ActiveSheet.Range("N11") = 0
        MyWRKBook.ActiveSheet.Range("O11") = "pcs(��.)"
        MyWRKBook.ActiveSheet.Range("N12") = 1
        MyWRKBook.ActiveSheet.Range("O12") = "m (�)"
        MyWRKBook.ActiveSheet.Range("N13") = 2
        MyWRKBook.ActiveSheet.Range("O13") = "kg (��)"
        MyWRKBook.ActiveSheet.Range("N14") = 3
        MyWRKBook.ActiveSheet.Range("O14") = "km (��)"
        MyWRKBook.ActiveSheet.Range("N15") = 4
        MyWRKBook.ActiveSheet.Range("O15") = "litre (����)"
        MyWRKBook.ActiveSheet.Range("N16") = 5
        MyWRKBook.ActiveSheet.Range("O16") = "pack (����.)"
        MyWRKBook.ActiveSheet.Range("N17") = 6
        MyWRKBook.ActiveSheet.Range("O17") = "set (�����.)"
        MyWRKBook.ActiveSheet.Range("N18") = 7
        MyWRKBook.ActiveSheet.Range("O18") = "pair (����)"

        MyWRKBook.ActiveSheet.Range("N10:O18").Font.Color = 16777215
        'MyWRKBook.ActiveSheet.Range("N10:O18").Font.TintAndShade = 0
        MyWRKBook.ActiveSheet.Range("E11:E1011").Validation.Add(Type:=3, AlertStyle:=1, Operator:=1, Formula1:="=$O$11:$O$18")
        'MyWRKBook.ActiveSheet.Range("E11:E1011").Validation.Add(Type:=3, AlertStyle:=1, Operator:=1, Formula1:="=R11C15:R18C15")

        MyWRKBook.ActiveSheet.Cells.Locked = True
        MyWRKBook.ActiveSheet.Cells.FormulaHidden = True

        MyWRKBook.ActiveSheet.Range("A11:G1011").Locked = False
        MyWRKBook.ActiveSheet.Range("A11:G1011").FormulaHidden = False
        MyWRKBook.ActiveSheet.Range("I11:J1011").Locked = False
        MyWRKBook.ActiveSheet.Range("I11:J1011").FormulaHidden = False

        MyWRKBook.ActiveSheet.Protect(Password:="!pass2009", DrawingObjects:=True, Contents:=True, Scenarios:=True)

        'MyWRKBook.Application.PrintCommunication = True
        'MyWRKBook.ActiveSheet.PageSetup.PrintArea = "$A$1:$I$1011"
        'MyWRKBook.Application.PrintCommunication = False
        'MyWRKBook.ActiveSheet.PageSetup.FitToPagesWide = 1
        'MyWRKBook.ActiveSheet.PageSetup.FitToPagesTall = 0
        'MyWRKBook.Application.PrintCommunication = True

        MyWRKBook.ActiveSheet.Range("A11").Select()
        MyObj.Application.Visible = True
        MyWRKBook = Nothing
        MyObj = Nothing
        oldCI = Nothing
    End Sub

    Private Sub ExportPropPurchToLO(ByVal MyCol As Integer, ByVal MySort As Integer)
        '////////////////////////////////////////////////////////////////////////////////
        '//
        '// �������� ����������� �� ����������� � ������������ � �����. ������ � LibreOffice
        '//
        '////////////////////////////////////////////////////////////////////////////////
        Dim oServiceManager As Object
        Dim oDispatcher As Object
        Dim oDesktop As Object
        Dim oWorkBook As Object
        Dim oSheet As Object
        Dim oFrame As Object
        Dim MySQLStr As String
        Dim SpecVersion As String               '--������ ������������
        Dim i As Integer

        LOSetNotation(0)
        oServiceManager = CreateObject("com.sun.star.ServiceManager")
        oDesktop = oServiceManager.createInstance("com.sun.star.frame.Desktop")
        oDispatcher = oServiceManager.createInstance("com.sun.star.frame.DispatchHelper")
        Dim arg(1)
        arg(0) = mAkePropertyValue("Hidden", True, oServiceManager)
        arg(1) = mAkePropertyValue("MacroExecutionMode", 4, oServiceManager)
        oWorkBook = oDesktop.loadComponentFromURL("private:factory/scalc", "_blank", 0, arg)
        oSheet = oWorkBook.getSheets().getByIndex(0)
        oFrame = oWorkBook.getCurrentController.getFrame

        '---������ �������
        oSheet.getColumns().getByName("A").Width = 1390
        oSheet.getColumns().getByName("B").Width = 2280
        oSheet.getColumns().getByName("C").Width = 2570
        oSheet.getColumns().getByName("D").Width = 5590
        oSheet.getColumns().getByName("E").Width = 1150
        oSheet.getColumns().getByName("F").Width = 1770
        oSheet.getColumns().getByName("G").Width = 2190
        oSheet.getColumns().getByName("H").Width = 2260
        oSheet.getColumns().getByName("I").Width = 2260
        oSheet.getColumns().getByName("J").Width = 5590
        oSheet.getColumns().getByName("K").Width = 5590
        oSheet.getColumns().getByName("L").Width = 5590
        '---������ �����
        LOSetCellProtection(oServiceManager, oDispatcher, oFrame, "", True)
        '---���������
        '---������ ������������
        MySQLStr = "SELECT Version "
        MySQLStr = MySQLStr & "FROM tbl_VersionImportItemsFromExcel WITH (NOLOCK) "
        MySQLStr = MySQLStr & "WHERE (Name = N'������������ �����������') "
        InitMyConn(False)
        InitMyRec(False, MySQLStr)
        If Declarations.MyRec.BOF = True And Declarations.MyRec.EOF = True Then
            MsgBox("� Scala �� ����������� ������� ������ ����� Excel. ���������� � ��������������.", vbCritical, "��������!")
            trycloseMyRec()
            Exit Sub
        Else
            SpecVersion = Trim(Declarations.MyRec.Fields("Version").Value)
            trycloseMyRec()
        End If
        oSheet.getCellRangeByName("A1").String = SpecVersion
        LOFontSetFamilyName(oServiceManager, oDispatcher, oFrame, "A1", "Calibri")
        LOFontSetSize(oServiceManager, oDispatcher, oFrame, "A1", 11)
        oSheet.getCellRangeByName("A1").CharColor = RGB(61, 65, 239) '---��� ����� �������� ������� - �� ����� ���� ��� �� RGB � BGR!!!! ������ ���� �����

        oSheet.getCellRangeByName("B2").String = "Skandika"
        LOFontSetFamilyName(oServiceManager, oDispatcher, oFrame, "B2", "Calibri")
        LOFontSetBold(oServiceManager, oDispatcher, oFrame, "B2")
        LOFontSetSize(oServiceManager, oDispatcher, oFrame, "B2", 16)

        oSheet.getCellRangeByName("A4").String = "OOO ""��������"""
        LOFontSetFamilyName(oServiceManager, oDispatcher, oFrame, "A4", "Tahoma")
        LOFontSetBold(oServiceManager, oDispatcher, oFrame, "A4")
        LOFontSetSize(oServiceManager, oDispatcher, oFrame, "A4", 11)

        oSheet.getCellRangeByName("A5").String = "�����:"
        LOFontSetFamilyName(oServiceManager, oDispatcher, oFrame, "A5", "Tahoma")
        LOFontSetBold(oServiceManager, oDispatcher, oFrame, "A5")
        LOFontSetSize(oServiceManager, oDispatcher, oFrame, "A5", 9)

        LOMergeCells(oServiceManager, oDispatcher, oFrame, "B5:H6")
        oSheet.getCellRangeByName("B5").String = "������,195027, ����� ���������, ������� ��., �.4 Tel: +7 (812) 325 2040, Fax: +7 (812) 325 2039, www.skandikagroup.ru"
        LOFontSetFamilyName(oServiceManager, oDispatcher, oFrame, "B5", "Tahoma")
        LOFontSetBold(oServiceManager, oDispatcher, oFrame, "B5")
        LOFontSetSize(oServiceManager, oDispatcher, oFrame, "B5", 9)
        LOWrapText(oServiceManager, oDispatcher, oFrame, "B5:H6")

        oSheet.getCellRangeByName("D8").String = "������������ ��������"
        LOFontSetFamilyName(oServiceManager, oDispatcher, oFrame, "D8", "Tahoma")
        LOFontSetBold(oServiceManager, oDispatcher, oFrame, "D8")
        LOFontSetSize(oServiceManager, oDispatcher, oFrame, "D8", 11.5)

        '--��������� �������
        oSheet.getCellRangeByName("A10").String = "N �/�"
        oSheet.getCellRangeByName("B10").String = "��� ������ Scala"
        oSheet.getCellRangeByName("C10").String = "��� ������ ����������"
        oSheet.getCellRangeByName("D10").String = "������������ ������"
        oSheet.getCellRangeByName("E10").String = "�� ���������"
        oSheet.getCellRangeByName("F10").String = "���-��"
        oSheet.getCellRangeByName("G10").String = "���� ��� ���"
        oSheet.getCellRangeByName("H10").String = "����� ��� ���"
        oSheet.getCellRangeByName("I10").String = "���� �������� (���)"
        oSheet.getCellRangeByName("J10").String = "������������ ������"
        oSheet.getCellRangeByName("K10").String = "�������. ����������"
        oSheet.getCellRangeByName("L10").String = "��� ����������"
        LOFontSetFamilyName(oServiceManager, oDispatcher, oFrame, "A10:I10", "Calibri")
        LOFontSetSize(oServiceManager, oDispatcher, oFrame, "A10:I10", 9)
        LOWrapText(oServiceManager, oDispatcher, oFrame, "A10:I10")
        Dim LineFormat As Object
        LineFormat = oServiceManager.Bridge_GetStruct("com.sun.star.table.BorderLine2")
        LineFormat.LineStyle = 0
        LineFormat.LineWidth = 40
        oSheet.getCellRangeByName("A10:I10").TopBorder = LineFormat
        oSheet.getCellRangeByName("A10:I10").RightBorder = LineFormat
        oSheet.getCellRangeByName("A10:I10").LeftBorder = LineFormat
        oSheet.getCellRangeByName("A10:I10").BottomBorder = LineFormat
        oSheet.getCellRangeByName("A10:I10").VertJustify = 2
        oSheet.getCellRangeByName("A10:I10").HoriJustify = 2

        '--����� ����� ������������
        MySQLStr = "SELECT ISNULL(tbl_SupplSearch_PropItems.ItemCode, N'') AS ItemCode, ISNULL(tbl_SupplSearch_PropItems.ItemSuppCode, N'') AS SuppItemCode, "
        MySQLStr = MySQLStr & "CASE WHEN LTRIM(RTRIM(LTRIM(RTRIM(ISNULL(SC010300.SC01002, ''))) + LTRIM(RTRIM(ISNULL(SC010300.SC01003, ''))))) "
        MySQLStr = MySQLStr & "= '' THEN ISNULL(tbl_SupplSearch_PropItems.ItemName, '') ELSE LTRIM(RTRIM(LTRIM(RTRIM(ISNULL(SC010300.SC01002, ''))) "
        MySQLStr = MySQLStr & "+ LTRIM(RTRIM(ISNULL(SC010300.SC01003, ''))))) END AS ItemName, CASE WHEN ISNULL(View_1.txt, N'') "
        MySQLStr = MySQLStr & "= '��' THEN 'pcs(��.)' ELSE CASE WHEN ISNULL(View_1.txt, N'') = '�' THEN 'm (�)' ELSE CASE WHEN ISNULL(View_1.txt, N'') "
        MySQLStr = MySQLStr & "= '��' THEN 'kg (��)' ELSE CASE WHEN ISNULL(View_1.txt, N'') = '��' THEN 'km (��)' ELSE CASE WHEN ISNULL(View_1.txt, N'') "
        MySQLStr = MySQLStr & "= '�' THEN 'litre (����)' ELSE CASE WHEN ISNULL(View_1.txt, N'') = '����' THEN 'pack (����.)' ELSE CASE WHEN ISNULL(View_1.txt, N'') "
        MySQLStr = MySQLStr & "= '�����' THEN 'set (�����.)' ELSE CASE WHEN ISNULL(View_1.txt, N'') = '���' THEN 'pair (����)' END END END END END END END END AS UOM, "
        MySQLStr = MySQLStr & "tbl_SupplSearch_PropItems.QTY, tbl_SupplSearch_PropItems.Price, ISNULL(tbl_SupplSearch_PropItems.LeadTimeWeek, 1) AS WeekQTY, "
        MySQLStr = MySQLStr & "ISNULL(tbl_SupplSearch_PropItems.AlternateTo, '') AS AlternateTo, ISNULL(tbl_SupplSearch_PropItems.Comments, '') AS Comments, "
        MySQLStr = MySQLStr & "ISNULL(tbl_SupplSearch_PropItems.SupplierCode, '') AS SupplierCode "
        MySQLStr = MySQLStr & "FROM tbl_SupplSearch_PropItems LEFT OUTER JOIN "
        MySQLStr = MySQLStr & "SC010300 ON tbl_SupplSearch_PropItems.ItemCode = SC010300.SC01001 LEFT OUTER JOIN "
        MySQLStr = MySQLStr & "(SELECT     0 AS Expr1, SC09002 AS txt "
        MySQLStr = MySQLStr & "FROM          SC090300 WITH (NOLOCK) "
        MySQLStr = MySQLStr & "WHERE      (SC09001 = 'RUS') "
        MySQLStr = MySQLStr & "UNION "
        MySQLStr = MySQLStr & "SELECT     1 AS Expr1, SC09003 "
        MySQLStr = MySQLStr & "FROM         SC090300 AS SC090300_40 WITH (NOLOCK) "
        MySQLStr = MySQLStr & "WHERE     (SC09001 = 'RUS') "
        MySQLStr = MySQLStr & "UNION "
        MySQLStr = MySQLStr & "SELECT     2 AS Expr1, SC09004 "
        MySQLStr = MySQLStr & "FROM         SC090300 AS SC090300_39 WITH (NOLOCK) "
        MySQLStr = MySQLStr & "WHERE     (SC09001 = 'RUS') "
        MySQLStr = MySQLStr & "UNION "
        MySQLStr = MySQLStr & "SELECT     3 AS Expr1, SC09005 "
        MySQLStr = MySQLStr & "FROM         SC090300 AS SC090300_38 WITH (NOLOCK) "
        MySQLStr = MySQLStr & "WHERE     (SC09001 = 'RUS') "
        MySQLStr = MySQLStr & "UNION "
        MySQLStr = MySQLStr & "SELECT     4 AS Expr1, SC09006 "
        MySQLStr = MySQLStr & "FROM         SC090300 AS SC090300_37 WITH (NOLOCK) "
        MySQLStr = MySQLStr & "WHERE     (SC09001 = 'RUS') "
        MySQLStr = MySQLStr & "UNION "
        MySQLStr = MySQLStr & "SELECT     5 AS Expr1, SC09007 "
        MySQLStr = MySQLStr & "FROM         SC090300 AS SC090300_36 WITH (NOLOCK) "
        MySQLStr = MySQLStr & "WHERE     (SC09001 = 'RUS') "
        MySQLStr = MySQLStr & "UNION "
        MySQLStr = MySQLStr & "SELECT     6 AS Expr1, SC09008 "
        MySQLStr = MySQLStr & "FROM         SC090300 AS SC090300_35 WITH (NOLOCK) "
        MySQLStr = MySQLStr & "WHERE     (SC09001 = 'RUS') "
        MySQLStr = MySQLStr & "UNION "
        MySQLStr = MySQLStr & "SELECT     7 AS Expr1, SC09009 "
        MySQLStr = MySQLStr & "FROM         SC090300 AS SC090300_34 WITH (NOLOCK) "
        MySQLStr = MySQLStr & "WHERE     (SC09001 = 'RUS') "
        MySQLStr = MySQLStr & "UNION "
        MySQLStr = MySQLStr & "SELECT     8 AS Expr1, SC09010 "
        MySQLStr = MySQLStr & "FROM         SC090300 AS SC090300_33 WITH (NOLOCK) "
        MySQLStr = MySQLStr & "WHERE     (SC09001 = 'RUS') "
        MySQLStr = MySQLStr & "UNION "
        MySQLStr = MySQLStr & "SELECT     9 AS Expr1, SC09011 "
        MySQLStr = MySQLStr & "FROM         SC090300 AS SC090300_32 WITH (NOLOCK) "
        MySQLStr = MySQLStr & "WHERE     (SC09001 = 'RUS') "
        MySQLStr = MySQLStr & "UNION "
        MySQLStr = MySQLStr & "SELECT     10 AS Expr1, SC09012 "
        MySQLStr = MySQLStr & "FROM         SC090300 AS SC090300_31 WITH (NOLOCK) "
        MySQLStr = MySQLStr & "WHERE     (SC09001 = 'RUS') "
        MySQLStr = MySQLStr & "UNION "
        MySQLStr = MySQLStr & "SELECT     11 AS Expr1, SC09013 "
        MySQLStr = MySQLStr & "FROM         SC090300 AS SC090300_30 WITH (NOLOCK) "
        MySQLStr = MySQLStr & "WHERE     (SC09001 = 'RUS') "
        MySQLStr = MySQLStr & "UNION "
        MySQLStr = MySQLStr & "SELECT     12 AS Expr1, SC09014 "
        MySQLStr = MySQLStr & "FROM         SC090300 AS SC090300_29 WITH (NOLOCK) "
        MySQLStr = MySQLStr & "WHERE     (SC09001 = 'RUS') "
        MySQLStr = MySQLStr & "UNION "
        MySQLStr = MySQLStr & "SELECT     13 AS Expr1, SC09015 "
        MySQLStr = MySQLStr & "FROM         SC090300 AS SC090300_28 WITH (NOLOCK) "
        MySQLStr = MySQLStr & "WHERE     (SC09001 = 'RUS') "
        MySQLStr = MySQLStr & "UNION "
        MySQLStr = MySQLStr & "SELECT     14 AS Expr1, SC09016 "
        MySQLStr = MySQLStr & "FROM         SC090300 AS SC090300_27 WITH (NOLOCK) "
        MySQLStr = MySQLStr & "WHERE     (SC09001 = 'RUS') "
        MySQLStr = MySQLStr & "UNION "
        MySQLStr = MySQLStr & "SELECT     15 AS Expr1, SC09017 "
        MySQLStr = MySQLStr & "FROM         SC090300 AS SC090300_26 WITH (NOLOCK) "
        MySQLStr = MySQLStr & "WHERE     (SC09001 = 'RUS') "
        MySQLStr = MySQLStr & "UNION "
        MySQLStr = MySQLStr & "SELECT     16 AS Expr1, SC09018 "
        MySQLStr = MySQLStr & "FROM         SC090300 AS SC090300_25 WITH (NOLOCK) "
        MySQLStr = MySQLStr & "WHERE     (SC09001 = 'RUS') "
        MySQLStr = MySQLStr & "UNION "
        MySQLStr = MySQLStr & "SELECT     17 AS Expr1, SC09019 "
        MySQLStr = MySQLStr & "FROM         SC090300 AS SC090300_24 WITH (NOLOCK) "
        MySQLStr = MySQLStr & "WHERE     (SC09001 = 'RUS') "
        MySQLStr = MySQLStr & "UNION "
        MySQLStr = MySQLStr & "SELECT     18 AS Expr1, SC09020 "
        MySQLStr = MySQLStr & "FROM         SC090300 AS SC090300_23 WITH (NOLOCK) "
        MySQLStr = MySQLStr & "WHERE     (SC09001 = 'RUS') "
        MySQLStr = MySQLStr & "UNION "
        MySQLStr = MySQLStr & "SELECT     19 AS Expr1, SC09021 "
        MySQLStr = MySQLStr & "FROM         SC090300 AS SC090300_22 WITH (NOLOCK) "
        MySQLStr = MySQLStr & "WHERE     (SC09001 = 'RUS') "
        MySQLStr = MySQLStr & "UNION "
        MySQLStr = MySQLStr & "SELECT     20 AS Expr1, SC09022 "
        MySQLStr = MySQLStr & "FROM         SC090300 AS SC090300_21 WITH (NOLOCK) "
        MySQLStr = MySQLStr & "WHERE     (SC09001 = 'RUS') "
        MySQLStr = MySQLStr & "UNION "
        MySQLStr = MySQLStr & "SELECT     21 AS Expr1, SC09023 "
        MySQLStr = MySQLStr & "FROM         SC090300 AS SC090300_20 WITH (NOLOCK) "
        MySQLStr = MySQLStr & "WHERE     (SC09001 = 'RUS') "
        MySQLStr = MySQLStr & "UNION "
        MySQLStr = MySQLStr & "SELECT     22 AS Expr1, SC09024 "
        MySQLStr = MySQLStr & "FROM         SC090300 AS SC090300_19 WITH (NOLOCK) "
        MySQLStr = MySQLStr & "WHERE     (SC09001 = 'RUS') "
        MySQLStr = MySQLStr & "UNION "
        MySQLStr = MySQLStr & "SELECT     23 AS Expr1, SC09025 "
        MySQLStr = MySQLStr & "FROM         SC090300 AS SC090300_18 WITH (NOLOCK) "
        MySQLStr = MySQLStr & "WHERE     (SC09001 = 'RUS') "
        MySQLStr = MySQLStr & "UNION "
        MySQLStr = MySQLStr & "SELECT     24 AS Expr1, SC09026 "
        MySQLStr = MySQLStr & "FROM         SC090300 AS SC090300_17 WITH (NOLOCK) "
        MySQLStr = MySQLStr & "WHERE     (SC09001 = 'RUS') "
        MySQLStr = MySQLStr & "UNION "
        MySQLStr = MySQLStr & "SELECT     25 AS Expr1, SC09027 "
        MySQLStr = MySQLStr & "FROM         SC090300 AS SC090300_16 WITH (NOLOCK) "
        MySQLStr = MySQLStr & "WHERE     (SC09001 = 'RUS') "
        MySQLStr = MySQLStr & "UNION "
        MySQLStr = MySQLStr & "SELECT     26 AS Expr1, SC09028 "
        MySQLStr = MySQLStr & "FROM         SC090300 AS SC090300_15 WITH (NOLOCK) "
        MySQLStr = MySQLStr & "WHERE     (SC09001 = 'RUS') "
        MySQLStr = MySQLStr & "UNION "
        MySQLStr = MySQLStr & "SELECT     27 AS Expr1, SC09029 "
        MySQLStr = MySQLStr & "FROM         SC090300 AS SC090300_14 WITH (NOLOCK) "
        MySQLStr = MySQLStr & "WHERE     (SC09001 = 'RUS') "
        MySQLStr = MySQLStr & "UNION "
        MySQLStr = MySQLStr & "SELECT     28 AS Expr1, SC09030 "
        MySQLStr = MySQLStr & "FROM         SC090300 AS SC090300_13 WITH (NOLOCK) "
        MySQLStr = MySQLStr & "WHERE     (SC09001 = 'RUS') "
        MySQLStr = MySQLStr & "UNION "
        MySQLStr = MySQLStr & "SELECT     29 AS Expr1, SC09031 "
        MySQLStr = MySQLStr & "FROM         SC090300 AS SC090300_12 WITH (NOLOCK) "
        MySQLStr = MySQLStr & "WHERE     (SC09001 = 'RUS') "
        MySQLStr = MySQLStr & "UNION "
        MySQLStr = MySQLStr & "SELECT     30 AS Expr1, SC09032 "
        MySQLStr = MySQLStr & "FROM         SC090300 AS SC090300_11 WITH (NOLOCK) "
        MySQLStr = MySQLStr & "WHERE     (SC09001 = 'RUS') "
        MySQLStr = MySQLStr & "UNION "
        MySQLStr = MySQLStr & "SELECT     31 AS Expr1, SC09033 "
        MySQLStr = MySQLStr & "FROM         SC090300 AS SC090300_10 WITH (NOLOCK) "
        MySQLStr = MySQLStr & "WHERE     (SC09001 = 'RUS') "
        MySQLStr = MySQLStr & "UNION "
        MySQLStr = MySQLStr & "SELECT     32 AS Expr1, SC09034 "
        MySQLStr = MySQLStr & "FROM         SC090300 AS SC090300_9 WITH (NOLOCK) "
        MySQLStr = MySQLStr & "WHERE     (SC09001 = 'RUS') "
        MySQLStr = MySQLStr & "UNION "
        MySQLStr = MySQLStr & "SELECT     33 AS Expr1, SC09035 "
        MySQLStr = MySQLStr & "FROM         SC090300 AS SC090300_8 WITH (NOLOCK) "
        MySQLStr = MySQLStr & "WHERE     (SC09001 = 'RUS') "
        MySQLStr = MySQLStr & "UNION "
        MySQLStr = MySQLStr & "SELECT     34 AS Expr1, SC09036 "
        MySQLStr = MySQLStr & "FROM         SC090300 AS SC090300_7 WITH (NOLOCK) "
        MySQLStr = MySQLStr & "WHERE     (SC09001 = 'RUS') "
        MySQLStr = MySQLStr & "UNION "
        MySQLStr = MySQLStr & "SELECT     35 AS Expr1, SC09037 "
        MySQLStr = MySQLStr & "FROM         SC090300 AS SC090300_6 WITH (NOLOCK) "
        MySQLStr = MySQLStr & "WHERE     (SC09001 = 'RUS') "
        MySQLStr = MySQLStr & "UNION "
        MySQLStr = MySQLStr & "SELECT     36 AS Expr1, SC09038 "
        MySQLStr = MySQLStr & "FROM         SC090300 AS SC090300_5 WITH (NOLOCK) "
        MySQLStr = MySQLStr & "WHERE     (SC09001 = 'RUS') "
        MySQLStr = MySQLStr & "UNION "
        MySQLStr = MySQLStr & "SELECT     37 AS Expr1, SC09039 "
        MySQLStr = MySQLStr & "FROM         SC090300 AS SC090300_4 WITH (NOLOCK) "
        MySQLStr = MySQLStr & "WHERE     (SC09001 = 'RUS') "
        MySQLStr = MySQLStr & "UNION "
        MySQLStr = MySQLStr & "SELECT     38 AS Expr1, SC09040 "
        MySQLStr = MySQLStr & "FROM         SC090300 AS SC090300_3 WITH (NOLOCK) "
        MySQLStr = MySQLStr & "WHERE     (SC09001 = 'RUS') "
        MySQLStr = MySQLStr & "UNION "
        MySQLStr = MySQLStr & "SELECT     39 AS Expr1, SC09041 "
        MySQLStr = MySQLStr & "FROM         SC090300 AS SC090300_2 WITH (NOLOCK) "
        MySQLStr = MySQLStr & "WHERE     (SC09001 = 'RUS') "
        MySQLStr = MySQLStr & "UNION "
        MySQLStr = MySQLStr & "SELECT     40 AS Expr1, SC09042 "
        MySQLStr = MySQLStr & "FROM SC090300 AS SC090300_1 WITH (NOLOCK)) AS View_1 ON tbl_SupplSearch_PropItems.UOM = View_1.Expr1 "
        MySQLStr = MySQLStr & "WHERE (tbl_SupplSearch_PropItems.SupplSearchID = " & DataGridView1.SelectedRows.Item(0).Cells("ID").Value & ") "
        If ComboBox1.Text = "������������ ������������" Then
            MySQLStr = MySQLStr & "AND (tbl_SupplSearch_PropItems.IsSelected = 1) "
        End If
        If ComboBox1.Text = "��������� ���������" Then
            MySQLStr = MySQLStr & "AND (tbl_SupplSearch_PropItems.SelectedBySalesman = 1) "
        End If
        Select Case MyCol
            Case 0
                MySQLStr = MySQLStr & "ORDER BY tbl_SupplSearch_PropItems.ID "
            Case 4
                MySQLStr = MySQLStr & "ORDER BY ItemCode "
            Case 5
                MySQLStr = MySQLStr & "ORDER BY SuppItemCode "
            Case 6
                MySQLStr = MySQLStr & "ORDER BY ItemName "
            Case 8
                MySQLStr = MySQLStr & "ORDER BY UOM "
            Case 9
                MySQLStr = MySQLStr & "ORDER BY tbl_SupplSearch_PropItems.QTY "
            Case Else
                MySQLStr = MySQLStr & "ORDER BY tbl_SupplSearch_PropItems.ID "
        End Select
        If MySort = 2 Then
            MySQLStr = MySQLStr & "Desc "
        End If
        InitMyConn(False)
        InitMyRec(False, MySQLStr)
        If Declarations.MyRec.BOF = True And Declarations.MyRec.EOF = True Then
            trycloseMyRec()
        Else
            Declarations.MyRec.MoveFirst()
            i = 0
            While Declarations.MyRec.EOF = False
                oSheet.getCellRangeByName("A" & CStr(i + 11)).Value = i + 1
                oSheet.getCellRangeByName("B" & CStr(i + 11)).String = Declarations.MyRec.Fields("ItemCode").Value
                oSheet.getCellRangeByName("C" & CStr(i + 11)).String = Declarations.MyRec.Fields("SuppItemCode").Value
                oSheet.getCellRangeByName("D" & CStr(i + 11)).String = Declarations.MyRec.Fields("ItemName").Value
                oSheet.getCellRangeByName("E" & CStr(i + 11)).String = Declarations.MyRec.Fields("UOM").Value
                If Not IsDBNull(Declarations.MyRec.Fields("QTY").Value) Then
                    oSheet.getCellRangeByName("F" & CStr(i + 11)).Value = Declarations.MyRec.Fields("QTY").Value
                End If
                If Not IsDBNull(Declarations.MyRec.Fields("Price").Value) Then
                    oSheet.getCellRangeByName("G" & CStr(i + 11)).Value = Declarations.MyRec.Fields("Price").Value
                End If
                oSheet.getCellRangeByName("H" & CStr(i + 11)).FormulaLocal = "=IF(F" & CStr(i + 11) & "*G" & CStr(i + 11) & " = 0;"""";F" & CStr(i + 11) & " * G" & CStr(i + 11) & ") "
                If Not IsDBNull(Declarations.MyRec.Fields("WeekQTY").Value) Then
                    oSheet.getCellRangeByName("I" & CStr(i + 11)).Value = Declarations.MyRec.Fields("WeekQTY").Value
                End If
                oSheet.getCellRangeByName("J" & CStr(i + 11)).String = Declarations.MyRec.Fields("AlternateTo").Value
                oSheet.getCellRangeByName("K" & CStr(i + 11)).String = Declarations.MyRec.Fields("Comments").Value
                oSheet.getCellRangeByName("L" & CStr(i + 11)).String = Declarations.MyRec.Fields("SupplierCode").Value

                i = i + 1
                Declarations.MyRec.MoveNext()
            End While
            trycloseMyRec()
        End If
        LOWrapText(oServiceManager, oDispatcher, oFrame, "A11:I" & CStr(11 + i - 1))
        LineFormat.LineStyle = 0
        LineFormat.LineWidth = 20
        oSheet.getCellRangeByName("A11:I" & CStr(11 + i - 1)).TopBorder = LineFormat
        oSheet.getCellRangeByName("A11:I" & CStr(11 + i - 1)).RightBorder = LineFormat
        oSheet.getCellRangeByName("A11:I" & CStr(11 + i - 1)).LeftBorder = LineFormat
        oSheet.getCellRangeByName("A11:I" & CStr(11 + i - 1)).BottomBorder = LineFormat
        '----������ �����
        LOSetCellProtection(oServiceManager, oDispatcher, oFrame, "A11:I500", False)
        '---�������� ���������� ����������
        oSheet.getCellRangeByName("N10").String = "������� ���������"
        oSheet.getCellRangeByName("N11").Value = 0
        oSheet.getCellRangeByName("O11").String = "pcs(��.)"
        oSheet.getCellRangeByName("N12").Value = 1
        oSheet.getCellRangeByName("O12").String = "m (�)"
        oSheet.getCellRangeByName("N13").Value = 2
        oSheet.getCellRangeByName("O13").String = "kg (��)"
        oSheet.getCellRangeByName("N14").Value = 3
        oSheet.getCellRangeByName("O14").String = "km (��)"
        oSheet.getCellRangeByName("N15").Value = 4
        oSheet.getCellRangeByName("O15").String = "litre (����)"
        oSheet.getCellRangeByName("N16").Value = 5
        oSheet.getCellRangeByName("O16").String = "pack (����.)"
        oSheet.getCellRangeByName("N17").Value = 6
        oSheet.getCellRangeByName("O17").String = "set (�����.)"
        oSheet.getCellRangeByName("N18").Value = 7
        oSheet.getCellRangeByName("O18").String = "pair (����)"
        oSheet.getCellRangeByName("N10:O18").CharColor = RGB(255, 255, 255) '---��� ����� �������� ������� - �� ����� ���� ��� �� RGB � BGR!!!! ������ ���� �����
        LOSetValidation(oSheet, "E11:E" & CStr(11 + i - 1), "=$O$11:$O$18")
        '----� ������ �����
        Dim args() As Object
        ReDim args(0)
        args(0) = oServiceManager.Bridge_GetStruct("com.sun.star.beans.PropertyValue")
        args(0).Name = "ToPoint"
        args(0).Value = "$A$1"
        oDispatcher.executeDispatch(oFrame, ".uno:GoToCell", "", 0, args)
        '----�������� �������
        LOPasswordProtect(oSheet, "!pass2022", True)
        '----���������
        oWorkBook.CurrentController.Frame.ContainerWindow.Visible = True
        oWorkBook.CurrentController.Frame.ContainerWindow.toFront()
    End Sub



    Private Sub Button2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button2.Click
        '////////////////////////////////////////////////////////////////////////////////
        '//
        '// ������� ����������� �� ����������� � ������������ �����������
        '//
        '////////////////////////////////////////////////////////////////////////////////
        Dim MySQLStr As String
        Dim myValue As String
        Dim MyID As Integer
        Dim MyRez As Object

        MyID = DataGridView1.SelectedRows.Item(0).Cells("ID").Value
        myValue = ""
        Declarations.MyCPID = ""
        myValue = InputBox("������� ����� ������������� �����������, ���� ������. ���� ��� - ������� ""������""", "����� ��", "")
        If myValue = "" Then
            '---�������� ���� ������ ��
            MyCPList = New CPList
            MyCPList.ShowDialog()
        Else
            '---��������, ��� ����� �� ���� ��� ����������� ��������
            MySQLStr = "SELECT COUNT(OR01001) AS CC "
            MySQLStr = MySQLStr & "FROM tbl_OR010300 "
            MySQLStr = MySQLStr & "WHERE (OR01001 = N'" & Strings.Right("0000000000" & Trim(myValue), 10) & "') "
            MySQLStr = MySQLStr & "AND (OR01019 = N'" & Trim(Declarations.SalesmanCode) & "') "
            InitMyConn(False)
            InitMyRec(False, MySQLStr)
            If Declarations.MyRec.EOF = True And Declarations.MyRec.BOF = True Then
                trycloseMyRec()
                MsgBox("������������ ����������� " & Trim(myValue) & " ��� �������� " & Trim(Declarations.SalesmanCode) & " " & Declarations.SalesmanName & _
                    " �� �������. ������� ���������� ����� ��� �������������� ������� (������ ""������"" � ���� ����� ������ ��)", vbCritical, "��������!")
            Else
                If Declarations.MyRec.Fields("CC").Value = 0 Then
                    trycloseMyRec()
                    MsgBox("������������ ����������� " & Trim(myValue) & " ��� �������� " & Trim(Declarations.SalesmanCode) & " " & Declarations.SalesmanName & _
                        " �� �������. ������� ���������� ����� ��� �������������� ������� (������ ""������"" � ���� ����� ������ ��)", vbCritical, "��������!")
                Else
                    trycloseMyRec()
                    Declarations.MyCPID = Strings.Right("0000000000" & Trim(myValue), 10)
                End If
            End If
        End If

        If Declarations.MyCPID.Equals("") = False Then
            '-----������� ������ � ��
            MyRez = MsgBox("������� ������ ������ �� �������?", MsgBoxStyle.YesNo, "��������!")
            If MyRez = vbYes Then
                MySQLStr = "DELETE FROM tbl_OR030300 "
                MySQLStr = MySQLStr & "WHERE (OR03001 = N'" & Trim(Declarations.MyCPID) & "')"
                InitMyConn(False)
                Declarations.MyConn.Execute(MySQLStr)

                MySQLStr = "DELETE FROM tbl_OR170300 "
                MySQLStr = MySQLStr & "WHERE (OR17001 = N'" & Trim(Declarations.MyCPID) & "') "
                Declarations.MyConn.Execute(MySQLStr)
            End If

            '-----���������� �������
            MySQLStr = "exec spp_SupplSearch_MoveToCP N'" & Trim(Declarations.MyCPID) & "', " & MyID.ToString
            InitMyConn(False)
            Declarations.MyConn.Execute(MySQLStr)

            '-----�������� ������
            If DataGridView1.SelectedRows.Item(0).Cells("SearchStatusID").Value = 3 Then
                MyRez = MsgBox("������� ������ �� ����� ����������?", MsgBoxStyle.YesNo, "��������!")
                If MyRez = vbYes Then
                    myValue = ""
                    myValue = InputBox("������� ����������� (���� ����������)", "�����������", "")
                    MySQLStr = "UPDATE tbl_SupplSearch "
                    MySQLStr = MySQLStr + "SET SalesStatus = 3, "
                    MySQLStr = MySQLStr + "Comments = Comments + '  --" + Format(Now, "dd/MM/yyyy HH:mm") + "-->' + N'" & myValue & "' "
                    MySQLStr = MySQLStr + "WHERE (ID = " & MyID & ") "
                    InitMyConn(False)
                    Declarations.MyConn.Execute(MySQLStr)
                    LoadRequests()
                    For i As Integer = 0 To DataGridView1.Rows.Count - 1
                        If DataGridView1.Item("ID", i).Value = MyID Then
                            DataGridView1.CurrentCell = DataGridView1.Item("ID", i)
                            Exit For
                        End If
                    Next
                    CheckRequestButtons()
                End If
            End If

            '-----��������� ������ �� � ������ �� �����
            MySQLStr = "UPDATE tbl_SupplSearch "
            MySQLStr = MySQLStr & "SET CPNum = N'" & Trim(Declarations.MyCPID) & "' "
            MySQLStr = MySQLStr & "WHERE (ID = " & MyID & ") "
            InitMyConn(False)
            Declarations.MyConn.Execute(MySQLStr)


            MsgBox("��������� �������� ����� ����������� ���������� � ������������ ����������� ���������.", vbOKOnly, "��������!")
        Else
            MsgBox("��������� �������� ����� ����������� ���������� � ������������ ����������� �� ���������.", vbOKOnly, "��������!")
        End If



    End Sub

    Private Sub ComboBox3_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles ComboBox3.SelectedIndexChanged
        '////////////////////////////////////////////////////////////////////////////////
        '//
        '// ���������� ���������� � ���� ����� ����� - ������ � ������ / ������ �������������
        '//
        '////////////////////////////////////////////////////////////////////////////////

        LoadRequests()
        LoadItems()
        LoadProposal()
        CheckRequestButtons()
        CheckItemButtons()
        CheckProposalButtons()
    End Sub

    Private Sub Button20_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button20.Click
        '////////////////////////////////////////////////////////////////////////////////
        '//
        '// ������� ������ ������ / ����������� ����������
        '//
        '////////////////////////////////////////////////////////////////////////////////

        If FullInfoFlag = 0 Then
            FullInfoFlag = 1
            Button20.Text = "����������� ����������"
        Else
            FullInfoFlag = 0
            Button20.Text = "������ ����������"
        End If
        ChangeColumnsVisibility()
    End Sub

    Private Sub ChangeColumnsVisibility()
        '////////////////////////////////////////////////////////////////////////////////
        '//
        '// ��������� ��������� ������� � ����������� �� ����� FullInfoFlag
        '//
        '////////////////////////////////////////////////////////////////////////////////

        If FullInfoFlag = 0 Then
            'DataGridView1.Columns(2).Visible = False
            DataGridView1.Columns("CustomerContactName").Visible = False
            DataGridView1.Columns("CustomerPhone").Visible = False
            DataGridView1.Columns("CustomerEmail").Visible = False
        Else
            'DataGridView1.Columns(2).Visible = True
            DataGridView1.Columns("CustomerContactName").Visible = True
            DataGridView1.Columns("CustomerPhone").Visible = True
            DataGridView1.Columns("CustomerEmail").Visible = True
        End If
    End Sub

    Private Sub Button21_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button21.Click
        '////////////////////////////////////////////////////////////////////////////////
        '//
        '// ���� ����������� �������� � ������� �����������
        '//
        '////////////////////////////////////////////////////////////////////////////////

        Declarations.MyItemPropID = DataGridView3.SelectedRows.Item(0).Cells("ID").Value
        MySalesCommentsToProposal = New SalesCommentsToProposal
        MySalesCommentsToProposal.ShowDialog()
    End Sub


    Private Sub DataGridView3_CellMouseClick(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellMouseEventArgs) Handles DataGridView3.CellMouseClick
        '////////////////////////////////////////////////////////////////////////////////
        '//
        '// ����� ��������� ������ ������������� �������� �� ������� ��������
        '//
        '////////////////////////////////////////////////////////////////////////////////
        Dim MyID As Integer

        If e.Button = System.Windows.Forms.MouseButtons.Left Then
            '------������� � ������������  ��� ��������
            If e.ColumnIndex = 2 Then
                If e.RowIndex <> -1 Then
                    If Button17.Enabled = True Then
                        MyID = DataGridView3.SelectedRows.Item(0).Cells("ID").Value
                        ChangeReadyState(MyID, DataGridView3.SelectedRows.Item(0).Cells("SelectedBySalesman").Value, DataGridView3.SelectedRows.Item(0).Cells("ItemID").Value)
                        '---�������� ������
                        LoadProposal()
                        '---������� ������� ������� �������������
                        For i As Integer = 0 To DataGridView3.Rows.Count - 1
                            If Trim(DataGridView3.Item("ID", i).Value.ToString) = MyID Then
                                DataGridView3.CurrentCell = DataGridView3.Item("SelectedBySalesman", i)
                            End If
                        Next
                        CheckProposalButtons()
                    End If
                Else
                    If IsNothing(SortColumnNum) = True Then
                        SortColumnNum = e.ColumnIndex
                        SortColOrder = System.ComponentModel.ListSortDirection.Ascending
                    Else
                        If DataGridView3.Columns(SortColumnNum).Equals(DataGridView3.Columns(e.ColumnIndex)) Then
                            '------������� �� �� - ������ ����������
                            If SortColOrder = System.ComponentModel.ListSortDirection.Ascending Then
                                SortColOrder = System.ComponentModel.ListSortDirection.Descending
                            Else
                                SortColOrder = System.ComponentModel.ListSortDirection.Ascending
                            End If
                        Else
                            '------������� ����� - ���������� �� ������������
                            SortColumnNum = e.ColumnIndex
                            SortColOrder = System.ComponentModel.ListSortDirection.Ascending
                        End If
                    End If
                    SetSorting()
                    FormatDataGridView3()
                End If
            Else '------��������� �������
                If e.RowIndex = -1 Then
                    If IsNothing(SortColumnNum) = True Then
                        SortColumnNum = e.ColumnIndex
                        SortColOrder = System.ComponentModel.ListSortDirection.Ascending
                    Else
                        If DataGridView3.Columns(SortColumnNum).Equals(DataGridView3.Columns(e.ColumnIndex)) Then
                            '------������� �� �� - ������ ����������
                            If SortColOrder = System.ComponentModel.ListSortDirection.Ascending Then
                                SortColOrder = System.ComponentModel.ListSortDirection.Descending
                            Else
                                SortColOrder = System.ComponentModel.ListSortDirection.Ascending
                            End If
                        Else
                            '------������� ����� - ���������� �� ������������
                            SortColumnNum = e.ColumnIndex
                            SortColOrder = System.ComponentModel.ListSortDirection.Ascending
                        End If
                    End If
                    SetSorting()
                    FormatDataGridView3()
                End If
            End If
        End If
    End Sub

    Private Sub ChangeReadyState(ByVal MyID As Integer, ByVal CurrState As Boolean, ByVal ItemID As Integer)
        '////////////////////////////////////////////////////////////////////////////////
        '//
        '// ����� ��������� ����������� �������� � �����������
        '//
        '////////////////////////////////////////////////////////////////////////////////
        Dim MySQLStr As String

        MySQLStr = "UPDATE tbl_SupplSearch_PropItems "
        If CurrState = False Then
            MySQLStr = MySQLStr & "SET SelectedBySalesman = 1 "
        Else
            MySQLStr = MySQLStr & "SET SelectedBySalesman = 0 "
        End If
        MySQLStr = MySQLStr & "WHERE (ID = " & MyID.ToString & ") "
        InitMyConn(False)
        Declarations.MyConn.Execute(MySQLStr)
    End Sub

    Private Sub SetSorting()
        '////////////////////////////////////////////////////////////////////////////////
        '//
        '// ����� ���������� �������
        '//
        '////////////////////////////////////////////////////////////////////////////////

        If IsNothing(SortColumnNum) = False Then
            DataGridView3.Sort(DataGridView3.Columns(SortColumnNum), SortColOrder)
            If SortColOrder = System.ComponentModel.ListSortDirection.Ascending Then
                DataGridView3.Columns(SortColumnNum).HeaderCell.SortGlyphDirection = SortOrder.Ascending
            Else
                DataGridView3.Columns(SortColumnNum).HeaderCell.SortGlyphDirection = SortOrder.Descending
            End If
        End If
    End Sub

    Private Sub Button22_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button22.Click
        '////////////////////////////////////////////////////////////////////////////////
        '//
        '// ����������� ������� �� ��������
        '//
        '////////////////////////////////////////////////////////////////////////////////

        LoadRequests()
        LoadItems()
        LoadProposal()
        CheckRequestButtons()
        CheckItemButtons()
        CheckProposalButtons()
    End Sub

    Private Sub Button23_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button23.Click
        '////////////////////////////////////////////////////////////////////////////////
        '//
        '// ������ ������� �� ��������
        '//
        '////////////////////////////////////////////////////////////////////////////////

        TextBox1.Text = ""
        LoadRequests()
        LoadItems()
        LoadProposal()
        CheckRequestButtons()
        CheckItemButtons()
        CheckProposalButtons()
    End Sub

    Private Sub ToolStripMenuItem1_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles ToolStripMenuItem1.Click
        '////////////////////////////////////////////////////////////////////////////////
        '//
        '// ����� ������������ ���� ���� �����������
        '//
        '////////////////////////////////////////////////////////////////////////////////
        Dim MySQLStr As String
        Dim myValue As String
        Dim EmailStr As String

        myValue = ""
        myValue = InputBox("������� �����������", "�����������", "")
        If myValue <> "" Then
            MySQLStr = "UPDATE tbl_SupplSearch "
            MySQLStr = MySQLStr + "SET Comments = ISNULL(Comments, '') + '" + CStr(Chr(10) + Chr(13)) + "' + '--" + Format(Now, "dd/MM/yyyy HH:mm") + "-->' + N'" & myValue & "' "
            MySQLStr = MySQLStr + "WHERE (ID = " & Declarations.MyRequestNum & ") "
            InitMyConn(False)
            Declarations.MyConn.Execute(MySQLStr)
            DataGridView1.Rows(Declarations.MyRowIndex).Cells("Comments").Value = DataGridView1.Rows(Declarations.MyRowIndex).Cells("Comments").Value _
                & Chr(10) & Chr(13) & "--" & Format(Now, "dd/MM/yyyy HH:mm") & "-->" & myValue
            '--------------�������� �����
            EmailStr = GetEmailFromDB(Trim(Strings.Left(DataGridView1.Rows(Declarations.MyRowIndex).Cells("Searcher").Value.ToString, 4)))
            If EmailStr = "" Then
                MsgBox("��� ������������ " & DataGridView1.Rows(Declarations.MyRowIndex).Cells("Searcher").Value.ToString & " � �� �� �������� �����. ���������� � ��������������.", MsgBoxStyle.Critical, "��������!")
            Else
                SendCommentByEmail(DataGridView1.Rows(Declarations.MyRowIndex).Cells("ID").Value.ToString, DataGridView1.Rows(Declarations.MyRowIndex).Cells("StartDate").Value.ToString, _
                   EmailStr, DataGridView1.Rows(Declarations.MyRowIndex).Cells("Customer").Value.ToString, DataGridView1.Rows(Declarations.MyRowIndex).Cells("Salesman").Value.ToString, _
                   Trim(myValue))
            End If
        End If
    End Sub

    Private Sub ButtonPause_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ButtonPause.Click
        '////////////////////////////////////////////////////////////////////////////////
        '//
        '// ���������� ������ �� "�����"
        '//
        '////////////////////////////////////////////////////////////////////////////////
        Dim MySQLStr As String
        Dim myValue As String
        Dim MyID As Integer
        Dim EmailStr As String
        Dim RequestStatus As String

        MyID = DataGridView1.SelectedRows.Item(0).Cells("ID").Value
        myValue = ""
        myValue = InputBox("������� ����������� (���� ����������)", "�����������", "")
        MySQLStr = "UPDATE tbl_SupplSearch "
        MySQLStr = MySQLStr + "SET SalesStatus = 4, "
        MySQLStr = MySQLStr + "Comments = ISNULL(Comments, '') + '" + CStr(Chr(10) + Chr(13)) + "' + '--" + Format(Now, "dd/MM/yyyy HH:mm") + "-->' + N'" & myValue & "' "
        MySQLStr = MySQLStr + "WHERE (ID = " & MyID & ") "
        InitMyConn(False)
        Declarations.MyConn.Execute(MySQLStr)
        '--------------�������� �����
        EmailStr = GetEmailFromDB(Trim(Strings.Left(DataGridView1.SelectedRows.Item(0).Cells("Searcher").Value.ToString, 4)))
        If EmailStr = "" Then
            MsgBox("��� ������������ " & DataGridView1.SelectedRows.Item(0).Cells("Searcher").Value.ToString & " � �� �� �������� �����. ���������� � ��������������.", MsgBoxStyle.Critical, "��������!")
        Else
            RequestStatus = "����� �������������"
            SendInfoByEmail(DataGridView1.SelectedRows.Item(0).Cells("ID").Value.ToString(), DataGridView1.SelectedRows.Item(0).Cells("StartDate").Value.ToString(), _
               EmailStr, DataGridView1.SelectedRows.Item(0).Cells("Customer").Value.ToString(), DataGridView1.SelectedRows.Item(0).Cells("Salesman").Value.ToString(), _
               RequestStatus)
        End If
        '---------------------------
        LoadRequests()
        For i As Integer = 0 To DataGridView1.Rows.Count - 1
            If DataGridView1.Item("ID", i).Value = MyID Then
                DataGridView1.CurrentCell = DataGridView1.Item("ID", i)
                Exit For
            End If
        Next
        CheckRequestButtons()
    End Sub

    Private Sub ButtonContinue_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ButtonContinue.Click
        '////////////////////////////////////////////////////////////////////////////////
        '//
        '// ������������� ��������� ������
        '//
        '////////////////////////////////////////////////////////////////////////////////
        Dim MySQLStr As String
        Dim EmailStr As String
        Dim RequestStatus As String

        Declarations.MyRequestNum = DataGridView1.SelectedRows.Item(0).Cells("ID").Value
        Declarations.MyRez1 = 0
        MyRestoreSearch = New RestoreSearch
        MyRestoreSearch.ShowDialog()
        If Declarations.MyRez1 = 1 Then
            MySQLStr = "UPDATE tbl_SupplSearch "
            MySQLStr = MySQLStr + "SET SalesStatus = 0, "
            MySQLStr = MySQLStr + "SearchStatus = CASE WHEN SearchStatus = 1 THEN 2 ELSE SearchStatus END "
            MySQLStr = MySQLStr + "WHERE (ID = " & Declarations.MyRequestNum & ") "
            InitMyConn(False)
            Declarations.MyConn.Execute(MySQLStr)
            '--------------�������� �����
            EmailStr = GetEmailFromDB(Trim(Strings.Left(DataGridView1.SelectedRows.Item(0).Cells("Searcher").Value.ToString, 4)))
            If EmailStr = "" Then
                MsgBox("��� ������������ " & DataGridView1.SelectedRows.Item(0).Cells("Searcher").Value.ToString & " � �� �� �������� �����. ���������� � ��������������.", MsgBoxStyle.Critical, "��������!")
            Else
                RequestStatus = "������������� ��������� ������"
                SendInfoByEmail(DataGridView1.SelectedRows.Item(0).Cells("ID").Value.ToString(), DataGridView1.SelectedRows.Item(0).Cells("StartDate").Value.ToString(), _
                   EmailStr, DataGridView1.SelectedRows.Item(0).Cells("Customer").Value.ToString(), DataGridView1.SelectedRows.Item(0).Cells("Salesman").Value.ToString(), _
                   RequestStatus)
            End If
            '---------------------------
            LoadRequests()
            For i As Integer = 0 To DataGridView1.Rows.Count - 1
                If DataGridView1.Item("ID", i).Value = Declarations.MyRequestNum Then
                    DataGridView1.CurrentCell = DataGridView1.Item("ID", i)
                    Exit For
                End If
            Next
            CheckRequestButtons()
        End If
    End Sub

    Private Sub Button24_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button24.Click
        '////////////////////////////////////////////////////////////////////////////////////////
        '//
        '// �������� ���������� � �������� �� �����
        '//
        '////////////////////////////////////////////////////////////////////////////////////////

        Me.Cursor = Cursors.WaitCursor
        If My.Settings.UseOffice = "LibreOffice" Then
            ExportReqToLO()
        Else
            ExportReqToExcel()
        End If
        Me.Cursor = Cursors.Default
    End Sub

    Private Sub ExportReqToExcel()
        '////////////////////////////////////////////////////////////////////////////////////////
        '//
        '// �������� ���������� � �������� �� ����� � Excel
        '//
        '////////////////////////////////////////////////////////////////////////////////////////
        Dim MyObj As Object
        Dim MyWRKBook As Object
        Dim i As Integer

        MyObj = CreateObject("Excel.Application")
        MyObj.SheetsInNewWorkbook = 1
        MyWRKBook = MyObj.Workbooks.Add

        '---���������
        MyWRKBook.ActiveSheet.Range("B1") = "������ �������� �� ����� ���������� "
        MyWRKBook.ActiveSheet.Range("A3") = "N �������"
        MyWRKBook.ActiveSheet.Range("B3") = "���� �������"
        MyWRKBook.ActiveSheet.Range("C3") = "������"
        MyWRKBook.ActiveSheet.Range("D3") = "������� �������"
        MyWRKBook.ActiveSheet.Range("E3") = "������� ��������"
        MyWRKBook.ActiveSheet.Range("F3") = "EMail ��������"
        MyWRKBook.ActiveSheet.Range("G3") = "����������� ����"
        MyWRKBook.ActiveSheet.Range("H3") = "��������"
        MyWRKBook.ActiveSheet.Range("I3") = "������ ��������"
        MyWRKBook.ActiveSheet.Range("J3") = "����������� ��������"
        MyWRKBook.ActiveSheet.Range("K3") = "���������"
        MyWRKBook.ActiveSheet.Range("L3") = "������ ����������"
        MyWRKBook.ActiveSheet.Range("M3") = "����������� ����������"
        MyWRKBook.ActiveSheet.Range("N3") = "N ������� �������"
        MyWRKBook.ActiveSheet.Range("O3") = "N �� ��������"
        MyWRKBook.ActiveSheet.Range("P3") = "������� ������"

        MyWRKBook.ActiveSheet.Columns("A:A").ColumnWidth = 12
        MyWRKBook.ActiveSheet.Columns("B:B").ColumnWidth = 20
        MyWRKBook.ActiveSheet.Columns("C:C").ColumnWidth = 45
        MyWRKBook.ActiveSheet.Columns("D:D").ColumnWidth = 30
        MyWRKBook.ActiveSheet.Columns("E:E").ColumnWidth = 20
        MyWRKBook.ActiveSheet.Columns("F:F").ColumnWidth = 30
        MyWRKBook.ActiveSheet.Columns("G:G").ColumnWidth = 20
        MyWRKBook.ActiveSheet.Columns("H:H").ColumnWidth = 25
        MyWRKBook.ActiveSheet.Columns("I:I").ColumnWidth = 25
        MyWRKBook.ActiveSheet.Columns("J:J").ColumnWidth = 50
        MyWRKBook.ActiveSheet.Columns("K:K").ColumnWidth = 25
        MyWRKBook.ActiveSheet.Columns("L:L").ColumnWidth = 25
        MyWRKBook.ActiveSheet.Columns("M:M").ColumnWidth = 50
        MyWRKBook.ActiveSheet.Columns("N:N").ColumnWidth = 20
        MyWRKBook.ActiveSheet.Columns("O:O").ColumnWidth = 20
        MyWRKBook.ActiveSheet.Columns("P:P").ColumnWidth = 30

        '---�������������� ���������
        MyWRKBook.ActiveSheet.Range("A3:P3").Select()
        MyWRKBook.ActiveSheet.Range("A3:P3").Borders(5).LineStyle = -4142
        MyWRKBook.ActiveSheet.Range("A3:P3").Borders(6).LineStyle = -4142
        With MyWRKBook.ActiveSheet.Range("A3:P3").Borders(7)
            .LineStyle = 1
            .Weight = 4
            .ColorIndex = -4105
        End With
        With MyWRKBook.ActiveSheet.Range("A3:P3").Borders(8)
            .LineStyle = 1
            .Weight = 4
            .ColorIndex = -4105
        End With
        With MyWRKBook.ActiveSheet.Range("A3:P3").Borders(9)
            .LineStyle = 1
            .Weight = 4
            .ColorIndex = -4105
        End With
        With MyWRKBook.ActiveSheet.Range("A3:P3").Borders(10)
            .LineStyle = 1
            .Weight = 4
            .ColorIndex = -4105
        End With
        With MyWRKBook.ActiveSheet.Range("A3:P3").Borders(11)
            .LineStyle = 1
            .Weight = 4
            .ColorIndex = -4105
        End With
        With MyWRKBook.ActiveSheet.Range("A3:P3").Interior
            .ColorIndex = 36
            .Pattern = 1
            .PatternColorIndex = -4105
        End With
        MyWRKBook.ActiveSheet.Range("B1").Select()
        MyWRKBook.ActiveSheet.Range("B1").Font.Bold = True
        MyWRKBook.ActiveSheet.Range("A3:P3").Select()
        MyWRKBook.ActiveSheet.Range("A3:P3").Font.Bold = True

        '---�������
        For i = 0 To DataGridView1.Rows.Count - 1
            MyWRKBook.ActiveSheet.Range("A" & CStr(i + 4)).NumberFormat = "@"
            MyWRKBook.ActiveSheet.Range("A" & CStr(i + 4)) = DataGridView1.Item("ID", i).Value.ToString
            MyWRKBook.ActiveSheet.Range("B" & CStr(i + 4)) = Format(DataGridView1.Item("StartDate", i).Value, "dd/MM/yyyy  hh:mm")
            MyWRKBook.ActiveSheet.Range("C" & CStr(i + 4)) = DataGridView1.Item("Customer", i).Value.ToString
            MyWRKBook.ActiveSheet.Range("D" & CStr(i + 4)) = DataGridView1.Item("CustomerContactName", i).Value.ToString
            MyWRKBook.ActiveSheet.Range("E" & CStr(i + 4)) = DataGridView1.Item("CustomerPhone", i).Value.ToString
            MyWRKBook.ActiveSheet.Range("F" & CStr(i + 4)) = DataGridView1.Item("CustomerEmail", i).Value.ToString
            MyWRKBook.ActiveSheet.Range("G" & CStr(i + 4)) = Format(DataGridView1.Item("RequestDate", i).Value, "dd/MM/yyyy  hh:mm")
            MyWRKBook.ActiveSheet.Range("H" & CStr(i + 4)) = DataGridView1.Item("Salesman", i).Value.ToString
            MyWRKBook.ActiveSheet.Range("I" & CStr(i + 4)) = DataGridView1.Item("SalesStatus", i).Value.ToString
            MyWRKBook.ActiveSheet.Range("J" & CStr(i + 4)) = DataGridView1.Item("Comments", i).Value.ToString
            MyWRKBook.ActiveSheet.Range("K" & CStr(i + 4)) = DataGridView1.Item("Searcher", i).Value.ToString
            MyWRKBook.ActiveSheet.Range("L" & CStr(i + 4)) = DataGridView1.Item("SearchStatus", i).Value.ToString
            MyWRKBook.ActiveSheet.Range("M" & CStr(i + 4)) = DataGridView1.Item("SearcherComments", i).Value.ToString
            MyWRKBook.ActiveSheet.Range("N" & CStr(i + 4)) = DataGridView1.Item("CustomerRequestNum", i).Value.ToString
            MyWRKBook.ActiveSheet.Range("O" & CStr(i + 4)) = DataGridView1.Item("CPNum", i).Value.ToString
            MyWRKBook.ActiveSheet.Range("P" & CStr(i + 4)) = DataGridView1.Item("CancelReason", i).Value.ToString
        Next i

        '---������ �������
        MyWRKBook.ActiveSheet.Range("A4:P" & CStr(i + 4)).Select()
        With MyWRKBook.ActiveSheet.Range("A4:P" & CStr(i + 4))
            .HorizontalAlignment = 1
            .VerticalAlignment = -4107
            .WrapText = True
            .Orientation = 0
            .AddIndent = False
            .IndentLevel = 0
            .ShrinkToFit = False
            .ReadingOrder = -5002
            .MergeCells = False
        End With
        With MyWRKBook.ActiveSheet.Range("A4:P" & CStr(i + 4)).Borders(5)
            .LineStyle = -4142
        End With
        With MyWRKBook.ActiveSheet.Range("A4:P" & CStr(i + 4)).Borders(6)
            .LineStyle = -4142
        End With
        With MyWRKBook.ActiveSheet.Range("A4:P" & CStr(i + 4)).Borders(7)
            .LineStyle = 1
            .ColorIndex = 0
            .TintAndShade = 0
            .Weight = 2
        End With
        With MyWRKBook.ActiveSheet.Range("A4:P" & CStr(i + 4)).Borders(8)
            .LineStyle = 1
            .ColorIndex = 0
            .TintAndShade = 0
            .Weight = 4
        End With
        With MyWRKBook.ActiveSheet.Range("A4:P" & CStr(i + 4)).Borders(9)
            .LineStyle = 1
            .ColorIndex = 0
            .TintAndShade = 0
            .Weight = 2
        End With
        With MyWRKBook.ActiveSheet.Range("A4:P" & CStr(i + 4)).Borders(10)
            .LineStyle = 1
            .ColorIndex = 0
            .TintAndShade = 0
            .Weight = 2
        End With
        With MyWRKBook.ActiveSheet.Range("A4:P" & CStr(i + 4)).Borders(11)
            .LineStyle = 1
            .ColorIndex = 0
            .TintAndShade = 0
            .Weight = 2
        End With
        With MyWRKBook.ActiveSheet.Range("A4:P" & CStr(i + 4)).Borders(12)
            .LineStyle = 1
            .ColorIndex = 0
            .TintAndShade = 0
            .Weight = 2
        End With


        MyWRKBook.ActiveSheet.Range("A1").Select()
        MyObj.Application.Visible = True
        MyObj = Nothing
    End Sub

    Private Sub ExportReqToLO()
        '////////////////////////////////////////////////////////////////////////////////////////
        '//
        '// �������� ���������� � �������� �� ����� � LibreOffice
        '//
        '////////////////////////////////////////////////////////////////////////////////////////
        Dim oServiceManager As Object
        Dim oDispatcher As Object
        Dim oDesktop As Object
        Dim oWorkBook As Object
        Dim oSheet As Object
        Dim oFrame As Object

        LOSetNotation(0)
        oServiceManager = CreateObject("com.sun.star.ServiceManager")
        oDesktop = oServiceManager.createInstance("com.sun.star.frame.Desktop")
        oDispatcher = oServiceManager.createInstance("com.sun.star.frame.DispatchHelper")
        Dim arg(1)
        arg(0) = mAkePropertyValue("Hidden", True, oServiceManager)
        arg(1) = mAkePropertyValue("MacroExecutionMode", 4, oServiceManager)
        oWorkBook = oDesktop.loadComponentFromURL("private:factory/scalc", "_blank", 0, arg)
        oSheet = oWorkBook.getSheets().getByIndex(0)
        oFrame = oWorkBook.getCurrentController.getFrame

        '---������ �������
        oSheet.getColumns().getByName("A").Width = 2490
        oSheet.getColumns().getByName("B").Width = 4060
        oSheet.getColumns().getByName("C").Width = 8960
        oSheet.getColumns().getByName("D").Width = 6020
        oSheet.getColumns().getByName("E").Width = 4060
        oSheet.getColumns().getByName("F").Width = 6020
        oSheet.getColumns().getByName("G").Width = 4060
        oSheet.getColumns().getByName("H").Width = 5040
        oSheet.getColumns().getByName("I").Width = 5040
        oSheet.getColumns().getByName("J").Width = 9940
        oSheet.getColumns().getByName("K").Width = 5040
        oSheet.getColumns().getByName("L").Width = 5040
        oSheet.getColumns().getByName("M").Width = 9940
        oSheet.getColumns().getByName("N").Width = 4060
        oSheet.getColumns().getByName("O").Width = 4060
        oSheet.getColumns().getByName("P").Width = 6020

        '---�������
        oSheet.getCellRangeByName("B1").String = "������ �������� �� ����� ����������"
        LOFontSetFamilyName(oServiceManager, oDispatcher, oFrame, "B1", "Arial")
        LOFontSetBold(oServiceManager, oDispatcher, oFrame, "B1")
        LOFontSetSize(oServiceManager, oDispatcher, oFrame, "B1", 11)

        oSheet.getCellRangeByName("A3").String = "N �������"
        oSheet.getCellRangeByName("B3").String = "���� �������"
        oSheet.getCellRangeByName("C3").String = "������"
        oSheet.getCellRangeByName("D3").String = "������� �������"
        oSheet.getCellRangeByName("E3").String = "������� ��������"
        oSheet.getCellRangeByName("F3").String = "EMail ��������"
        oSheet.getCellRangeByName("G3").String = "����������� ����"
        oSheet.getCellRangeByName("H3").String = "��������"
        oSheet.getCellRangeByName("I3").String = "������ ��������"
        oSheet.getCellRangeByName("J3").String = "����������� ��������"
        oSheet.getCellRangeByName("K3").String = "���������"
        oSheet.getCellRangeByName("L3").String = "������ ����������"
        oSheet.getCellRangeByName("M3").String = "����������� ����������"
        oSheet.getCellRangeByName("N3").String = "N ������� �������"
        oSheet.getCellRangeByName("O3").String = "N �� ��������"
        oSheet.getCellRangeByName("P3").String = "������� ������"

        Dim i As Integer
        i = 3
        LOFontSetFamilyName(oServiceManager, oDispatcher, oFrame, "A" & CStr(i) & ":P" & CStr(i), "Arial")
        LOFontSetSize(oServiceManager, oDispatcher, oFrame, "A" & CStr(i) & ":P" & CStr(i), 8)
        LOFontSetBold(oServiceManager, oDispatcher, oFrame, "A" & CStr(i) & ":P" & CStr(i))
        oSheet.getCellRangeByName("A" & CStr(i) & ":P" & CStr(i)).CellBackColor = RGB(174, 249, 255)  '---��� ����� �������� ������� - �� ����� ���� ��� �� RGB � BGR!!!! ������ ���� �����
        LOSetBorders(oServiceManager, oSheet, "A" & CStr(i) & ":P" & CStr(i), 70, 70, RGB(0, 0, 0)) '---��� ����� �������� ������� - �� ����� ���� ��� �� RGB � BGR!!!! ������ ���� �����
        oSheet.getCellRangeByName("A" & CStr(i) & ":P" & CStr(i)).VertJustify = 2
        oSheet.getCellRangeByName("A" & CStr(i) & ":P" & CStr(i)).HoriJustify = 2
        LOWrapText(oServiceManager, oDispatcher, oFrame, "A" & CStr(i) & ":P" & CStr(i))

        '---�������
        For i = 0 To DataGridView1.Rows.Count - 1
            oSheet.getCellRangeByName("A" & CStr(i + 4)).String = Trim(DataGridView1.Item("ID", i).Value.ToString)
            oSheet.getCellRangeByName("B" & CStr(i + 4)).Value = DataGridView1.Item("StartDate", i).Value
            LOFormatCells(oServiceManager, oDispatcher, oFrame, "B" & CStr(i + 4) & ":B" & CStr(i + 4), 51)
            oSheet.getCellRangeByName("C" & CStr(i + 4)).String = Trim(DataGridView1.Item("Customer", i).Value.ToString)
            oSheet.getCellRangeByName("D" & CStr(i + 4)).String = Trim(DataGridView1.Item("CustomerContactName", i).Value.ToString)
            oSheet.getCellRangeByName("E" & CStr(i + 4)).String = Trim(DataGridView1.Item("CustomerPhone", i).Value.ToString)
            oSheet.getCellRangeByName("F" & CStr(i + 4)).String = Trim(DataGridView1.Item("CustomerEmail", i).Value.ToString)
            oSheet.getCellRangeByName("G" & CStr(i + 4)).Value = DataGridView1.Item("RequestDate", i).Value
            LOFormatCells(oServiceManager, oDispatcher, oFrame, "G" & CStr(i + 4) & ":G" & CStr(i + 4), 51)
            oSheet.getCellRangeByName("H" & CStr(i + 4)).String = Trim(DataGridView1.Item("Salesman", i).Value.ToString)
            oSheet.getCellRangeByName("I" & CStr(i + 4)).String = Trim(DataGridView1.Item("SalesStatus", i).Value.ToString)
            oSheet.getCellRangeByName("J" & CStr(i + 4)).String = Trim(DataGridView1.Item("Comments", i).Value.ToString)
            oSheet.getCellRangeByName("K" & CStr(i + 4)).String = Trim(DataGridView1.Item("Searcher", i).Value.ToString)
            oSheet.getCellRangeByName("L" & CStr(i + 4)).String = Trim(DataGridView1.Item("SearchStatus", i).Value.ToString)
            oSheet.getCellRangeByName("M" & CStr(i + 4)).String = Trim(DataGridView1.Item("SearcherComments", i).Value.ToString)
            oSheet.getCellRangeByName("N" & CStr(i + 4)).String = Trim(DataGridView1.Item("CustomerRequestNum", i).Value.ToString)
            oSheet.getCellRangeByName("O" & CStr(i + 4)).String = Trim(DataGridView1.Item("CPNum", i).Value.ToString)
            oSheet.getCellRangeByName("P" & CStr(i + 4)).String = Trim(DataGridView1.Item("CancelReason", i).Value.ToString)
        Next i

        '---������ �������
        LOFontSetFamilyName(oServiceManager, oDispatcher, oFrame, "A4:P" & CStr(i + 3), "Calibri")
        LOFontSetSize(oServiceManager, oDispatcher, oFrame, "A4:P" & CStr(i + 3), 9)
        LOWrapText(oServiceManager, oDispatcher, oFrame, "A4:P" & CStr(i + 3))
        Dim LineFormat As Object
        LineFormat = oServiceManager.Bridge_GetStruct("com.sun.star.table.BorderLine2")
        LineFormat.LineStyle = 0
        LineFormat.LineWidth = 20
        oSheet.getCellRangeByName("A4:P" & CStr(i + 3)).TopBorder = LineFormat
        oSheet.getCellRangeByName("A4:P" & CStr(i + 3)).RightBorder = LineFormat
        oSheet.getCellRangeByName("A4:P" & CStr(i + 3)).LeftBorder = LineFormat
        oSheet.getCellRangeByName("A4:P" & CStr(i + 3)).BottomBorder = LineFormat

        '---
        Dim args() As Object
        ReDim args(0)
        args(0) = oServiceManager.Bridge_GetStruct("com.sun.star.beans.PropertyValue")
        args(0).Name = "ToPoint"
        args(0).Value = "$A$1"
        oDispatcher.executeDispatch(oFrame, ".uno:GoToCell", "", 0, args)

        oWorkBook.CurrentController.Frame.ContainerWindow.Visible = True
        oWorkBook.CurrentController.Frame.ContainerWindow.toFront()
    End Sub

    Private Sub Button25_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button25.Click
        '////////////////////////////////////////////////////////////////////////////////
        '//
        '// �������� ����������� �� ����������� � ������������ c ��������������
        '//
        '////////////////////////////////////////////////////////////////////////////////
        If My.Settings.UseOffice = "LibreOffice" Then
            If DataGridView3.SortedColumn Is Nothing Then
                ExportPropPriCostToLO(0, 1)
            Else
                ExportPropPriCostToLO(DataGridView3.SortedColumn.Index, DataGridView3.SortOrder)
            End If

        Else
            If DataGridView3.SortedColumn Is Nothing Then
                ExportPropPriCostToExcel(0, 1)
            Else
                ExportPropPriCostToExcel(DataGridView3.SortedColumn.Index, DataGridView3.SortOrder)
            End If
        End If
    End Sub

    Private Sub ExportPropPriCostToExcel(ByVal MyCol As Integer, ByVal MySort As Integer)
        '////////////////////////////////////////////////////////////////////////////////
        '//
        '// �������� ����������� �� ����������� � ������������ c �������������� � Excel
        '//
        '////////////////////////////////////////////////////////////////////////////////
        Dim MySQLStr As String
        Dim SpecVersion As String               '--������ ������������
        Dim MyObj As Object
        Dim MyWRKBook As Object
        Dim i As Integer

        MyObj = CreateObject("Excel.Application")
        Dim oldCI As System.Globalization.CultureInfo = System.Threading.Thread.CurrentThread.CurrentCulture
        System.Threading.Thread.CurrentThread.CurrentCulture = New System.Globalization.CultureInfo("en-US")

        MyObj.SheetsInNewWorkbook = 1
        MyObj.ReferenceStyle = 1
        MyWRKBook = MyObj.Workbooks.Add

        '---���������
        '---������ ������������
        MySQLStr = "SELECT Version "
        MySQLStr = MySQLStr & "FROM tbl_VersionImportItemsFromExcel WITH (NOLOCK) "
        MySQLStr = MySQLStr & "WHERE (Name = N'������������ �����������') "
        InitMyConn(False)
        InitMyRec(False, MySQLStr)
        If Declarations.MyRec.BOF = True And Declarations.MyRec.EOF = True Then
            MsgBox("� Scala �� ����������� ������� ������ ����� Excel. ���������� � ��������������.", vbCritical, "��������!")
            trycloseMyRec()
            Exit Sub
        Else
            SpecVersion = Trim(Declarations.MyRec.Fields("Version").Value)
            trycloseMyRec()
        End If
        MyWRKBook.ActiveSheet.Range("A1") = SpecVersion
        With MyWRKBook.ActiveSheet.Range("A1").Font
            .Name = "Calibri"
            .Size = 9
            '.Color = -16776961
            .ColorIndex = 3
        End With

        MyWRKBook.ActiveSheet.Range("B2") = "Skandika"
        With MyWRKBook.ActiveSheet.Range("B2").Font
            .Name = "Calibri"
            .Size = 16
            .Color = 0
            .Bold = True
        End With

        MyWRKBook.ActiveSheet.Range("A4") = "OOO ""��������"""
        With MyWRKBook.ActiveSheet.Range("A4").Font
            .Name = "Tahoma"
            .Size = 9
            .Color = 0
            .Bold = True
        End With

        MyWRKBook.ActiveSheet.Range("A5") = "�����:"
        With MyWRKBook.ActiveSheet.Range("A5").Font
            .Name = "Tahoma"
            .Size = 9
            .Color = 0
            .Bold = True
        End With

        MyWRKBook.ActiveSheet.Range("B5:H6").MergeCells = True
        MyWRKBook.ActiveSheet.Range("B5") = "������,195027, ����� ���������, ������� ��., �.4 Tel: +7 (812) 325 2040, Fax: +7 (812) 325 2039, www.skandikagroup.ru"
        With MyWRKBook.ActiveSheet.Range("B5").Font
            .Name = "Tahoma"
            .Size = 9
            .Color = 0
            .Bold = True
        End With
        MyWRKBook.ActiveSheet.Range("B5:H6").WrapText = True
        MyWRKBook.ActiveSheet.Range("B5:H6").VerticalAlignment = -4160

        MyWRKBook.ActiveSheet.Range("D8") = "������������ ��������"
        With MyWRKBook.ActiveSheet.Range("D8").Font
            .Name = "Tahoma"
            .Size = 11.5
            .Color = 0
            .Bold = True
        End With

        '--��������� �������
        MyWRKBook.ActiveSheet.Range("A10") = "N �/�"
        MyWRKBook.ActiveSheet.Range("B10") = "��� ������ Scala"
        MyWRKBook.ActiveSheet.Range("C10") = "��� ������ ����������"
        MyWRKBook.ActiveSheet.Range("D10") = "������������ ������"
        MyWRKBook.ActiveSheet.Range("E10") = "�� ���������"
        MyWRKBook.ActiveSheet.Range("F10") = "���-��"
        MyWRKBook.ActiveSheet.Range("G10") = "���� ��� ���"
        MyWRKBook.ActiveSheet.Range("H10") = "����� ��� ���"
        MyWRKBook.ActiveSheet.Range("I10") = "���� �������� (���)"
        MyWRKBook.ActiveSheet.Range("J10") = "������������ ������"
        MyWRKBook.ActiveSheet.Range("K10") = "�������. ����������"
        MyWRKBook.ActiveSheet.Range("L10") = "��� ����������"
        MyWRKBook.ActiveSheet.Range("A10:I10").Select()
        MyWRKBook.ActiveSheet.Range("A10:I10").Borders(5).LineStyle = -4142
        MyWRKBook.ActiveSheet.Range("A10:I10").Borders(6).LineStyle = -4142
        With MyWRKBook.ActiveSheet.Range("A10:I10").Borders(7)
            .LineStyle = 1
            .Weight = 2
            .ColorIndex = -4105
        End With
        With MyWRKBook.ActiveSheet.Range("A10:I10").Borders(8)
            .LineStyle = 1
            .Weight = 2
            .ColorIndex = -4105
        End With
        With MyWRKBook.ActiveSheet.Range("A10:I10").Borders(9)
            .LineStyle = 1
            .Weight = 2
            .ColorIndex = -4105
        End With
        With MyWRKBook.ActiveSheet.Range("A10:I10").Borders(10)
            .LineStyle = 1
            .Weight = 2
            .ColorIndex = -4105
        End With
        With MyWRKBook.ActiveSheet.Range("A10:I10").Borders(11)
            .LineStyle = 1
            .Weight = 2
            .ColorIndex = -4105
        End With
        MyWRKBook.ActiveSheet.Range("A10:K10").WrapText = True
        MyWRKBook.ActiveSheet.Range("A10:K10").VerticalAlignment = -4108
        MyWRKBook.ActiveSheet.Range("A10:K10").HorizontalAlignment = -4108
        With MyWRKBook.ActiveSheet.Range("A10:K10").Font
            .Name = "Calibri"
            .Size = 9
            .Color = 0
            .Bold = False
        End With

        MyWRKBook.ActiveSheet.Columns("A:A").ColumnWidth = 4
        MyWRKBook.ActiveSheet.Columns("B:B").ColumnWidth = 10
        MyWRKBook.ActiveSheet.Columns("C:C").ColumnWidth = 10
        MyWRKBook.ActiveSheet.Columns("D:D").ColumnWidth = 30
        MyWRKBook.ActiveSheet.Columns("E:E").ColumnWidth = 5
        MyWRKBook.ActiveSheet.Columns("F:F").ColumnWidth = 8
        MyWRKBook.ActiveSheet.Columns("G:G").ColumnWidth = 10
        MyWRKBook.ActiveSheet.Columns("H:H").ColumnWidth = 12
        MyWRKBook.ActiveSheet.Columns("I:I").ColumnWidth = 12
        MyWRKBook.ActiveSheet.Columns("J:J").ColumnWidth = 30
        MyWRKBook.ActiveSheet.Columns("K:K").ColumnWidth = 30
        MyWRKBook.ActiveSheet.Columns("L:L").ColumnWidth = 30

        '--����� ����� ������������
        MySQLStr = "SELECT ISNULL(tbl_SupplSearch_PropItems.ItemCode, N'') AS ItemCode, ISNULL(tbl_SupplSearch_PropItems.ItemSuppCode, N'') AS SuppItemCode, "
        MySQLStr = MySQLStr & "CASE WHEN LTRIM(RTRIM(LTRIM(RTRIM(ISNULL(SC010300.SC01002, ''))) + LTRIM(RTRIM(ISNULL(SC010300.SC01003, ''))))) "
        MySQLStr = MySQLStr & "= '' THEN ISNULL(tbl_SupplSearch_PropItems.ItemName, '') ELSE LTRIM(RTRIM(LTRIM(RTRIM(ISNULL(SC010300.SC01002, ''))) "
        MySQLStr = MySQLStr & "+ LTRIM(RTRIM(ISNULL(SC010300.SC01003, ''))))) END AS ItemName, CASE WHEN ISNULL(View_1.txt, N'') "
        MySQLStr = MySQLStr & "= '��' THEN 'pcs(��.)' ELSE CASE WHEN ISNULL(View_1.txt, N'') = '�' THEN 'm (�)' ELSE CASE WHEN ISNULL(View_1.txt, N'') "
        MySQLStr = MySQLStr & "= '��' THEN 'kg (��)' ELSE CASE WHEN ISNULL(View_1.txt, N'') = '��' THEN 'km (��)' ELSE CASE WHEN ISNULL(View_1.txt, N'') "
        MySQLStr = MySQLStr & "= '�' THEN 'litre (����)' ELSE CASE WHEN ISNULL(View_1.txt, N'') = '����' THEN 'pack (����.)' ELSE CASE WHEN ISNULL(View_1.txt, N'') "
        MySQLStr = MySQLStr & "= '�����' THEN 'set (�����.)' ELSE CASE WHEN ISNULL(View_1.txt, N'') = '���' THEN 'pair (����)' END END END END END END END END AS UOM, "
        MySQLStr = MySQLStr & "tbl_SupplSearch_PropItems.QTY, tbl_SupplSearch_PropItems.Price * (100 + ISNULL(View_5.ShippingCost, 10)) / 100 AS PriCost, "
        MySQLStr = MySQLStr & "ISNULL(tbl_SupplSearch_PropItems.LeadTimeWeek, 1) AS WeekQTY, "
        MySQLStr = MySQLStr & "ISNULL(tbl_SupplSearch_PropItems.AlternateTo, '') AS AlternateTo, ISNULL(tbl_SupplSearch_PropItems.Comments, '') AS Comments, "
        MySQLStr = MySQLStr & "ISNULL(tbl_SupplSearch_PropItems.SupplierCode, '') AS SupplierCode "
        MySQLStr = MySQLStr & "FROM tbl_SupplSearch_PropItems LEFT OUTER JOIN "
        MySQLStr = MySQLStr & "(SELECT PL01001, ShippingCost "
        MySQLStr = MySQLStr & "FROM tbl_SupplierCard0300) AS View_5 ON tbl_SupplSearch_PropItems.SupplierCode = View_5.PL01001 LEFT OUTER JOIN "
        MySQLStr = MySQLStr & "SC010300 ON tbl_SupplSearch_PropItems.ItemCode = SC010300.SC01001 LEFT OUTER JOIN "
        MySQLStr = MySQLStr & "(SELECT     0 AS Expr1, SC09002 AS txt "
        MySQLStr = MySQLStr & "FROM          SC090300 WITH (NOLOCK) "
        MySQLStr = MySQLStr & "WHERE      (SC09001 = 'RUS') "
        MySQLStr = MySQLStr & "UNION "
        MySQLStr = MySQLStr & "SELECT     1 AS Expr1, SC09003 "
        MySQLStr = MySQLStr & "FROM         SC090300 AS SC090300_40 WITH (NOLOCK) "
        MySQLStr = MySQLStr & "WHERE     (SC09001 = 'RUS') "
        MySQLStr = MySQLStr & "UNION "
        MySQLStr = MySQLStr & "SELECT     2 AS Expr1, SC09004 "
        MySQLStr = MySQLStr & "FROM         SC090300 AS SC090300_39 WITH (NOLOCK) "
        MySQLStr = MySQLStr & "WHERE     (SC09001 = 'RUS') "
        MySQLStr = MySQLStr & "UNION "
        MySQLStr = MySQLStr & "SELECT     3 AS Expr1, SC09005 "
        MySQLStr = MySQLStr & "FROM         SC090300 AS SC090300_38 WITH (NOLOCK) "
        MySQLStr = MySQLStr & "WHERE     (SC09001 = 'RUS') "
        MySQLStr = MySQLStr & "UNION "
        MySQLStr = MySQLStr & "SELECT     4 AS Expr1, SC09006 "
        MySQLStr = MySQLStr & "FROM         SC090300 AS SC090300_37 WITH (NOLOCK) "
        MySQLStr = MySQLStr & "WHERE     (SC09001 = 'RUS') "
        MySQLStr = MySQLStr & "UNION "
        MySQLStr = MySQLStr & "SELECT     5 AS Expr1, SC09007 "
        MySQLStr = MySQLStr & "FROM         SC090300 AS SC090300_36 WITH (NOLOCK) "
        MySQLStr = MySQLStr & "WHERE     (SC09001 = 'RUS') "
        MySQLStr = MySQLStr & "UNION "
        MySQLStr = MySQLStr & "SELECT     6 AS Expr1, SC09008 "
        MySQLStr = MySQLStr & "FROM         SC090300 AS SC090300_35 WITH (NOLOCK) "
        MySQLStr = MySQLStr & "WHERE     (SC09001 = 'RUS') "
        MySQLStr = MySQLStr & "UNION "
        MySQLStr = MySQLStr & "SELECT     7 AS Expr1, SC09009 "
        MySQLStr = MySQLStr & "FROM         SC090300 AS SC090300_34 WITH (NOLOCK) "
        MySQLStr = MySQLStr & "WHERE     (SC09001 = 'RUS') "
        MySQLStr = MySQLStr & "UNION "
        MySQLStr = MySQLStr & "SELECT     8 AS Expr1, SC09010 "
        MySQLStr = MySQLStr & "FROM         SC090300 AS SC090300_33 WITH (NOLOCK) "
        MySQLStr = MySQLStr & "WHERE     (SC09001 = 'RUS') "
        MySQLStr = MySQLStr & "UNION "
        MySQLStr = MySQLStr & "SELECT     9 AS Expr1, SC09011 "
        MySQLStr = MySQLStr & "FROM         SC090300 AS SC090300_32 WITH (NOLOCK) "
        MySQLStr = MySQLStr & "WHERE     (SC09001 = 'RUS') "
        MySQLStr = MySQLStr & "UNION "
        MySQLStr = MySQLStr & "SELECT     10 AS Expr1, SC09012 "
        MySQLStr = MySQLStr & "FROM         SC090300 AS SC090300_31 WITH (NOLOCK) "
        MySQLStr = MySQLStr & "WHERE     (SC09001 = 'RUS') "
        MySQLStr = MySQLStr & "UNION "
        MySQLStr = MySQLStr & "SELECT     11 AS Expr1, SC09013 "
        MySQLStr = MySQLStr & "FROM         SC090300 AS SC090300_30 WITH (NOLOCK) "
        MySQLStr = MySQLStr & "WHERE     (SC09001 = 'RUS') "
        MySQLStr = MySQLStr & "UNION "
        MySQLStr = MySQLStr & "SELECT     12 AS Expr1, SC09014 "
        MySQLStr = MySQLStr & "FROM         SC090300 AS SC090300_29 WITH (NOLOCK) "
        MySQLStr = MySQLStr & "WHERE     (SC09001 = 'RUS') "
        MySQLStr = MySQLStr & "UNION "
        MySQLStr = MySQLStr & "SELECT     13 AS Expr1, SC09015 "
        MySQLStr = MySQLStr & "FROM         SC090300 AS SC090300_28 WITH (NOLOCK) "
        MySQLStr = MySQLStr & "WHERE     (SC09001 = 'RUS') "
        MySQLStr = MySQLStr & "UNION "
        MySQLStr = MySQLStr & "SELECT     14 AS Expr1, SC09016 "
        MySQLStr = MySQLStr & "FROM         SC090300 AS SC090300_27 WITH (NOLOCK) "
        MySQLStr = MySQLStr & "WHERE     (SC09001 = 'RUS') "
        MySQLStr = MySQLStr & "UNION "
        MySQLStr = MySQLStr & "SELECT     15 AS Expr1, SC09017 "
        MySQLStr = MySQLStr & "FROM         SC090300 AS SC090300_26 WITH (NOLOCK) "
        MySQLStr = MySQLStr & "WHERE     (SC09001 = 'RUS') "
        MySQLStr = MySQLStr & "UNION "
        MySQLStr = MySQLStr & "SELECT     16 AS Expr1, SC09018 "
        MySQLStr = MySQLStr & "FROM         SC090300 AS SC090300_25 WITH (NOLOCK) "
        MySQLStr = MySQLStr & "WHERE     (SC09001 = 'RUS') "
        MySQLStr = MySQLStr & "UNION "
        MySQLStr = MySQLStr & "SELECT     17 AS Expr1, SC09019 "
        MySQLStr = MySQLStr & "FROM         SC090300 AS SC090300_24 WITH (NOLOCK) "
        MySQLStr = MySQLStr & "WHERE     (SC09001 = 'RUS') "
        MySQLStr = MySQLStr & "UNION "
        MySQLStr = MySQLStr & "SELECT     18 AS Expr1, SC09020 "
        MySQLStr = MySQLStr & "FROM         SC090300 AS SC090300_23 WITH (NOLOCK) "
        MySQLStr = MySQLStr & "WHERE     (SC09001 = 'RUS') "
        MySQLStr = MySQLStr & "UNION "
        MySQLStr = MySQLStr & "SELECT     19 AS Expr1, SC09021 "
        MySQLStr = MySQLStr & "FROM         SC090300 AS SC090300_22 WITH (NOLOCK) "
        MySQLStr = MySQLStr & "WHERE     (SC09001 = 'RUS') "
        MySQLStr = MySQLStr & "UNION "
        MySQLStr = MySQLStr & "SELECT     20 AS Expr1, SC09022 "
        MySQLStr = MySQLStr & "FROM         SC090300 AS SC090300_21 WITH (NOLOCK) "
        MySQLStr = MySQLStr & "WHERE     (SC09001 = 'RUS') "
        MySQLStr = MySQLStr & "UNION "
        MySQLStr = MySQLStr & "SELECT     21 AS Expr1, SC09023 "
        MySQLStr = MySQLStr & "FROM         SC090300 AS SC090300_20 WITH (NOLOCK) "
        MySQLStr = MySQLStr & "WHERE     (SC09001 = 'RUS') "
        MySQLStr = MySQLStr & "UNION "
        MySQLStr = MySQLStr & "SELECT     22 AS Expr1, SC09024 "
        MySQLStr = MySQLStr & "FROM         SC090300 AS SC090300_19 WITH (NOLOCK) "
        MySQLStr = MySQLStr & "WHERE     (SC09001 = 'RUS') "
        MySQLStr = MySQLStr & "UNION "
        MySQLStr = MySQLStr & "SELECT     23 AS Expr1, SC09025 "
        MySQLStr = MySQLStr & "FROM         SC090300 AS SC090300_18 WITH (NOLOCK) "
        MySQLStr = MySQLStr & "WHERE     (SC09001 = 'RUS') "
        MySQLStr = MySQLStr & "UNION "
        MySQLStr = MySQLStr & "SELECT     24 AS Expr1, SC09026 "
        MySQLStr = MySQLStr & "FROM         SC090300 AS SC090300_17 WITH (NOLOCK) "
        MySQLStr = MySQLStr & "WHERE     (SC09001 = 'RUS') "
        MySQLStr = MySQLStr & "UNION "
        MySQLStr = MySQLStr & "SELECT     25 AS Expr1, SC09027 "
        MySQLStr = MySQLStr & "FROM         SC090300 AS SC090300_16 WITH (NOLOCK) "
        MySQLStr = MySQLStr & "WHERE     (SC09001 = 'RUS') "
        MySQLStr = MySQLStr & "UNION "
        MySQLStr = MySQLStr & "SELECT     26 AS Expr1, SC09028 "
        MySQLStr = MySQLStr & "FROM         SC090300 AS SC090300_15 WITH (NOLOCK) "
        MySQLStr = MySQLStr & "WHERE     (SC09001 = 'RUS') "
        MySQLStr = MySQLStr & "UNION "
        MySQLStr = MySQLStr & "SELECT     27 AS Expr1, SC09029 "
        MySQLStr = MySQLStr & "FROM         SC090300 AS SC090300_14 WITH (NOLOCK) "
        MySQLStr = MySQLStr & "WHERE     (SC09001 = 'RUS') "
        MySQLStr = MySQLStr & "UNION "
        MySQLStr = MySQLStr & "SELECT     28 AS Expr1, SC09030 "
        MySQLStr = MySQLStr & "FROM         SC090300 AS SC090300_13 WITH (NOLOCK) "
        MySQLStr = MySQLStr & "WHERE     (SC09001 = 'RUS') "
        MySQLStr = MySQLStr & "UNION "
        MySQLStr = MySQLStr & "SELECT     29 AS Expr1, SC09031 "
        MySQLStr = MySQLStr & "FROM         SC090300 AS SC090300_12 WITH (NOLOCK) "
        MySQLStr = MySQLStr & "WHERE     (SC09001 = 'RUS') "
        MySQLStr = MySQLStr & "UNION "
        MySQLStr = MySQLStr & "SELECT     30 AS Expr1, SC09032 "
        MySQLStr = MySQLStr & "FROM         SC090300 AS SC090300_11 WITH (NOLOCK) "
        MySQLStr = MySQLStr & "WHERE     (SC09001 = 'RUS') "
        MySQLStr = MySQLStr & "UNION "
        MySQLStr = MySQLStr & "SELECT     31 AS Expr1, SC09033 "
        MySQLStr = MySQLStr & "FROM         SC090300 AS SC090300_10 WITH (NOLOCK) "
        MySQLStr = MySQLStr & "WHERE     (SC09001 = 'RUS') "
        MySQLStr = MySQLStr & "UNION "
        MySQLStr = MySQLStr & "SELECT     32 AS Expr1, SC09034 "
        MySQLStr = MySQLStr & "FROM         SC090300 AS SC090300_9 WITH (NOLOCK) "
        MySQLStr = MySQLStr & "WHERE     (SC09001 = 'RUS') "
        MySQLStr = MySQLStr & "UNION "
        MySQLStr = MySQLStr & "SELECT     33 AS Expr1, SC09035 "
        MySQLStr = MySQLStr & "FROM         SC090300 AS SC090300_8 WITH (NOLOCK) "
        MySQLStr = MySQLStr & "WHERE     (SC09001 = 'RUS') "
        MySQLStr = MySQLStr & "UNION "
        MySQLStr = MySQLStr & "SELECT     34 AS Expr1, SC09036 "
        MySQLStr = MySQLStr & "FROM         SC090300 AS SC090300_7 WITH (NOLOCK) "
        MySQLStr = MySQLStr & "WHERE     (SC09001 = 'RUS') "
        MySQLStr = MySQLStr & "UNION "
        MySQLStr = MySQLStr & "SELECT     35 AS Expr1, SC09037 "
        MySQLStr = MySQLStr & "FROM         SC090300 AS SC090300_6 WITH (NOLOCK) "
        MySQLStr = MySQLStr & "WHERE     (SC09001 = 'RUS') "
        MySQLStr = MySQLStr & "UNION "
        MySQLStr = MySQLStr & "SELECT     36 AS Expr1, SC09038 "
        MySQLStr = MySQLStr & "FROM         SC090300 AS SC090300_5 WITH (NOLOCK) "
        MySQLStr = MySQLStr & "WHERE     (SC09001 = 'RUS') "
        MySQLStr = MySQLStr & "UNION "
        MySQLStr = MySQLStr & "SELECT     37 AS Expr1, SC09039 "
        MySQLStr = MySQLStr & "FROM         SC090300 AS SC090300_4 WITH (NOLOCK) "
        MySQLStr = MySQLStr & "WHERE     (SC09001 = 'RUS') "
        MySQLStr = MySQLStr & "UNION "
        MySQLStr = MySQLStr & "SELECT     38 AS Expr1, SC09040 "
        MySQLStr = MySQLStr & "FROM         SC090300 AS SC090300_3 WITH (NOLOCK) "
        MySQLStr = MySQLStr & "WHERE     (SC09001 = 'RUS') "
        MySQLStr = MySQLStr & "UNION "
        MySQLStr = MySQLStr & "SELECT     39 AS Expr1, SC09041 "
        MySQLStr = MySQLStr & "FROM         SC090300 AS SC090300_2 WITH (NOLOCK) "
        MySQLStr = MySQLStr & "WHERE     (SC09001 = 'RUS') "
        MySQLStr = MySQLStr & "UNION "
        MySQLStr = MySQLStr & "SELECT     40 AS Expr1, SC09042 "
        MySQLStr = MySQLStr & "FROM SC090300 AS SC090300_1 WITH (NOLOCK)) AS View_1 ON tbl_SupplSearch_PropItems.UOM = View_1.Expr1 "
        MySQLStr = MySQLStr & "WHERE (tbl_SupplSearch_PropItems.SupplSearchID = " & DataGridView1.SelectedRows.Item(0).Cells("ID").Value & ") "
        If ComboBox1.Text = "������������ ������������" Then
            MySQLStr = MySQLStr & "AND (tbl_SupplSearch_PropItems.IsSelected = 1) "
        End If
        If ComboBox1.Text = "��������� ���������" Then
            MySQLStr = MySQLStr & "AND (tbl_SupplSearch_PropItems.SelectedBySalesman = 1) "
        End If
        Select Case MyCol
            Case 0
                MySQLStr = MySQLStr & "ORDER BY tbl_SupplSearch_PropItems.ID "
            Case 4
                MySQLStr = MySQLStr & "ORDER BY ItemCode "
            Case 5
                MySQLStr = MySQLStr & "ORDER BY SuppItemCode "
            Case 6
                MySQLStr = MySQLStr & "ORDER BY ItemName "
            Case 8
                MySQLStr = MySQLStr & "ORDER BY UOM "
            Case 9
                MySQLStr = MySQLStr & "ORDER BY tbl_SupplSearch_PropItems.QTY "
            Case Else
                MySQLStr = MySQLStr & "ORDER BY tbl_SupplSearch_PropItems.ID "
        End Select
        If MySort = 2 Then
            MySQLStr = MySQLStr & "Desc "
        End If

        InitMyConn(False)
        InitMyRec(False, MySQLStr)
        If Declarations.MyRec.BOF = True And Declarations.MyRec.EOF = True Then
            trycloseMyRec()
        Else
            Declarations.MyRec.MoveFirst()
            i = 0
            While Declarations.MyRec.EOF = False
                MyWRKBook.ActiveSheet.Range("A" & CStr(i + 11)).NumberFormat = "@"
                MyWRKBook.ActiveSheet.Range("A" & CStr(i + 11)) = i + 1
                MyWRKBook.ActiveSheet.Range("B" & CStr(i + 11)).NumberFormat = "@"
                MyWRKBook.ActiveSheet.Range("B" & CStr(i + 11)) = Declarations.MyRec.Fields("ItemCode").Value
                MyWRKBook.ActiveSheet.Range("C" & CStr(i + 11)).NumberFormat = "@"
                MyWRKBook.ActiveSheet.Range("C" & CStr(i + 11)) = Declarations.MyRec.Fields("SuppItemCode").Value
                MyWRKBook.ActiveSheet.Range("D" & CStr(i + 11)) = Declarations.MyRec.Fields("ItemName").Value
                MyWRKBook.ActiveSheet.Range("E" & CStr(i + 11)) = Declarations.MyRec.Fields("UOM").Value
                MyWRKBook.ActiveSheet.Range("F" & CStr(i + 11)) = Declarations.MyRec.Fields("QTY").Value
                MyWRKBook.ActiveSheet.Range("G" & CStr(i + 11)) = Declarations.MyRec.Fields("PriCost").Value
                MyWRKBook.ActiveSheet.Range("I" & CStr(i + 11)) = Declarations.MyRec.Fields("WeekQTY").Value
                MyWRKBook.ActiveSheet.Range("J" & CStr(i + 11)) = Declarations.MyRec.Fields("AlternateTo").Value
                MyWRKBook.ActiveSheet.Range("K" & CStr(i + 11)) = Declarations.MyRec.Fields("Comments").Value
                MyWRKBook.ActiveSheet.Range("L" & CStr(i + 11)) = Declarations.MyRec.Fields("SupplierCode").Value

                i = i + 1
                Declarations.MyRec.MoveNext()
            End While
            trycloseMyRec()
        End If

        '---

        MyWRKBook.ActiveSheet.Range("A11:I11").Select()
        MyWRKBook.ActiveSheet.Range("A11:I11").Borders(5).LineStyle = -4142
        MyWRKBook.ActiveSheet.Range("A11:I11").Borders(6).LineStyle = -4142
        With MyWRKBook.ActiveSheet.Range("A11:I1011").Borders(7)
            .LineStyle = 1
            .Weight = 2
            .ColorIndex = -4105
        End With
        With MyWRKBook.ActiveSheet.Range("A11:I1011").Borders(8)
            .LineStyle = 1
            .Weight = 2
            .ColorIndex = -4105
        End With
        With MyWRKBook.ActiveSheet.Range("A11:I1011").Borders(9)
            .LineStyle = 1
            .Weight = 2
            .ColorIndex = -4105
        End With
        With MyWRKBook.ActiveSheet.Range("A11:I1011").Borders(10)
            .LineStyle = 1
            .Weight = 2
            .ColorIndex = -4105
        End With
        With MyWRKBook.ActiveSheet.Range("A11:I1011").Borders(11)
            .LineStyle = 1
            .Weight = 2
            .ColorIndex = -4105
        End With
        With MyWRKBook.ActiveSheet.Range("A11:I1011").Borders(12)
            .LineStyle = 1
            .Weight = 2
            .ColorIndex = -4105
        End With
        With MyWRKBook.ActiveSheet.Range("A11:I1011").Font
            .Name = "Calibri"
            .Size = 9
            .Color = 0
            .Bold = False
        End With
        With MyWRKBook.ActiveSheet.Range("H11:H1011")
            '.FormulaR1C1 = "=����(RC[-2]*RC[-1] = 0; """"; RC[-2]*RC[-1])"
            .FormulaR1C1 = "=IF(RC[-2]*RC[-1] = 0, """", RC[-2]*RC[-1])"
        End With
        System.Threading.Thread.CurrentThread.CurrentCulture = oldCI

        '---�������� ���������� ����������
        MyWRKBook.ActiveSheet.Range("N10") = "������� ���������"
        MyWRKBook.ActiveSheet.Range("N11") = 0
        MyWRKBook.ActiveSheet.Range("O11") = "pcs(��.)"
        MyWRKBook.ActiveSheet.Range("N12") = 1
        MyWRKBook.ActiveSheet.Range("O12") = "m (�)"
        MyWRKBook.ActiveSheet.Range("N13") = 2
        MyWRKBook.ActiveSheet.Range("O13") = "kg (��)"
        MyWRKBook.ActiveSheet.Range("N14") = 3
        MyWRKBook.ActiveSheet.Range("O14") = "km (��)"
        MyWRKBook.ActiveSheet.Range("N15") = 4
        MyWRKBook.ActiveSheet.Range("O15") = "litre (����)"
        MyWRKBook.ActiveSheet.Range("N16") = 5
        MyWRKBook.ActiveSheet.Range("O16") = "pack (����.)"
        MyWRKBook.ActiveSheet.Range("N17") = 6
        MyWRKBook.ActiveSheet.Range("O17") = "set (�����.)"
        MyWRKBook.ActiveSheet.Range("N18") = 7
        MyWRKBook.ActiveSheet.Range("O18") = "pair (����)"

        MyWRKBook.ActiveSheet.Range("N10:O18").Font.Color = 16777215
        'MyWRKBook.ActiveSheet.Range("N10:O18").Font.TintAndShade = 0
        MyWRKBook.ActiveSheet.Range("E11:E1011").Validation.Add(Type:=3, AlertStyle:=1, Operator:=1, Formula1:="=$O$11:$O$18")
        'MyWRKBook.ActiveSheet.Range("E11:E1011").Validation.Add(Type:=3, AlertStyle:=1, Operator:=1, Formula1:="=R11C15:R18C15")

        MyWRKBook.ActiveSheet.Cells.Locked = True
        MyWRKBook.ActiveSheet.Cells.FormulaHidden = True

        MyWRKBook.ActiveSheet.Range("A11:G1011").Locked = False
        MyWRKBook.ActiveSheet.Range("A11:G1011").FormulaHidden = False
        MyWRKBook.ActiveSheet.Range("I11:J1011").Locked = False
        MyWRKBook.ActiveSheet.Range("I11:J1011").FormulaHidden = False

        MyWRKBook.ActiveSheet.Protect(Password:="!pass2009", DrawingObjects:=True, Contents:=True, Scenarios:=True)

        'MyWRKBook.Application.PrintCommunication = True
        'MyWRKBook.ActiveSheet.PageSetup.PrintArea = "$A$1:$I$1011"
        'MyWRKBook.Application.PrintCommunication = False
        'MyWRKBook.ActiveSheet.PageSetup.FitToPagesWide = 1
        'MyWRKBook.ActiveSheet.PageSetup.FitToPagesTall = 0
        'MyWRKBook.Application.PrintCommunication = True

        MyWRKBook.ActiveSheet.Range("A11").Select()
        MyObj.Application.Visible = True
        MyWRKBook = Nothing
        MyObj = Nothing
        oldCI = Nothing
    End Sub

    Private Sub ExportPropPriCostToLO(ByVal MyCol As Integer, ByVal MySort As Integer)
        '////////////////////////////////////////////////////////////////////////////////
        '//
        '// �������� ����������� �� ����������� � ������������ c �������������� � LibreOffice
        '//
        '////////////////////////////////////////////////////////////////////////////////
        Dim oServiceManager As Object
        Dim oDispatcher As Object
        Dim oDesktop As Object
        Dim oWorkBook As Object
        Dim oSheet As Object
        Dim oFrame As Object
        Dim MySQLStr As String
        Dim SpecVersion As String               '--������ ������������
        Dim i As Integer

        LOSetNotation(0)
        oServiceManager = CreateObject("com.sun.star.ServiceManager")
        oDesktop = oServiceManager.createInstance("com.sun.star.frame.Desktop")
        oDispatcher = oServiceManager.createInstance("com.sun.star.frame.DispatchHelper")
        Dim arg(1)
        arg(0) = mAkePropertyValue("Hidden", True, oServiceManager)
        arg(1) = mAkePropertyValue("MacroExecutionMode", 4, oServiceManager)
        oWorkBook = oDesktop.loadComponentFromURL("private:factory/scalc", "_blank", 0, arg)
        oSheet = oWorkBook.getSheets().getByIndex(0)
        oFrame = oWorkBook.getCurrentController.getFrame

        '---������ �������
        oSheet.getColumns().getByName("A").Width = 1390
        oSheet.getColumns().getByName("B").Width = 2280
        oSheet.getColumns().getByName("C").Width = 2570
        oSheet.getColumns().getByName("D").Width = 5590
        oSheet.getColumns().getByName("E").Width = 1150
        oSheet.getColumns().getByName("F").Width = 1770
        oSheet.getColumns().getByName("G").Width = 2190
        oSheet.getColumns().getByName("H").Width = 2260
        oSheet.getColumns().getByName("I").Width = 2260
        oSheet.getColumns().getByName("J").Width = 5590
        oSheet.getColumns().getByName("K").Width = 5590
        oSheet.getColumns().getByName("L").Width = 5590
        '---������ �����
        LOSetCellProtection(oServiceManager, oDispatcher, oFrame, "", True)
        '---���������
        '---������ ������������
        MySQLStr = "SELECT Version "
        MySQLStr = MySQLStr & "FROM tbl_VersionImportItemsFromExcel WITH (NOLOCK) "
        MySQLStr = MySQLStr & "WHERE (Name = N'������������ �����������') "
        InitMyConn(False)
        InitMyRec(False, MySQLStr)
        If Declarations.MyRec.BOF = True And Declarations.MyRec.EOF = True Then
            MsgBox("� Scala �� ����������� ������� ������ ����� Excel. ���������� � ��������������.", vbCritical, "��������!")
            trycloseMyRec()
            Exit Sub
        Else
            SpecVersion = Trim(Declarations.MyRec.Fields("Version").Value)
            trycloseMyRec()
        End If
        oSheet.getCellRangeByName("A1").String = SpecVersion
        LOFontSetFamilyName(oServiceManager, oDispatcher, oFrame, "A1", "Calibri")
        LOFontSetSize(oServiceManager, oDispatcher, oFrame, "A1", 11)
        oSheet.getCellRangeByName("A1").CharColor = RGB(61, 65, 239) '---��� ����� �������� ������� - �� ����� ���� ��� �� RGB � BGR!!!! ������ ���� �����

        oSheet.getCellRangeByName("B2").String = "Skandika"
        LOFontSetFamilyName(oServiceManager, oDispatcher, oFrame, "B2", "Calibri")
        LOFontSetBold(oServiceManager, oDispatcher, oFrame, "B2")
        LOFontSetSize(oServiceManager, oDispatcher, oFrame, "B2", 16)

        oSheet.getCellRangeByName("A4").String = "OOO ""��������"""
        LOFontSetFamilyName(oServiceManager, oDispatcher, oFrame, "A4", "Tahoma")
        LOFontSetBold(oServiceManager, oDispatcher, oFrame, "A4")
        LOFontSetSize(oServiceManager, oDispatcher, oFrame, "A4", 11)

        oSheet.getCellRangeByName("A5").String = "�����:"
        LOFontSetFamilyName(oServiceManager, oDispatcher, oFrame, "A5", "Tahoma")
        LOFontSetBold(oServiceManager, oDispatcher, oFrame, "A5")
        LOFontSetSize(oServiceManager, oDispatcher, oFrame, "A5", 9)

        LOMergeCells(oServiceManager, oDispatcher, oFrame, "B5:H6")
        oSheet.getCellRangeByName("B5").String = "������,195027, ����� ���������, ������� ��., �.4 Tel: +7 (812) 325 2040, Fax: +7 (812) 325 2039, www.skandikagroup.ru"
        LOFontSetFamilyName(oServiceManager, oDispatcher, oFrame, "B5", "Tahoma")
        LOFontSetBold(oServiceManager, oDispatcher, oFrame, "B5")
        LOFontSetSize(oServiceManager, oDispatcher, oFrame, "B5", 9)
        LOWrapText(oServiceManager, oDispatcher, oFrame, "B5:H6")

        oSheet.getCellRangeByName("D8").String = "������������ ��������"
        LOFontSetFamilyName(oServiceManager, oDispatcher, oFrame, "D8", "Tahoma")
        LOFontSetBold(oServiceManager, oDispatcher, oFrame, "D8")
        LOFontSetSize(oServiceManager, oDispatcher, oFrame, "D8", 11.5)

        '--��������� �������
        oSheet.getCellRangeByName("A10").String = "N �/�"
        oSheet.getCellRangeByName("B10").String = "��� ������ Scala"
        oSheet.getCellRangeByName("C10").String = "��� ������ ����������"
        oSheet.getCellRangeByName("D10").String = "������������ ������"
        oSheet.getCellRangeByName("E10").String = "�� ���������"
        oSheet.getCellRangeByName("F10").String = "���-��"
        oSheet.getCellRangeByName("G10").String = "���� ��� ���"
        oSheet.getCellRangeByName("H10").String = "����� ��� ���"
        oSheet.getCellRangeByName("I10").String = "���� �������� (���)"
        oSheet.getCellRangeByName("J10").String = "������������ ������"
        oSheet.getCellRangeByName("K10").String = "�������. ����������"
        oSheet.getCellRangeByName("L10").String = "��� ����������"
        LOFontSetFamilyName(oServiceManager, oDispatcher, oFrame, "A10:I10", "Calibri")
        LOFontSetSize(oServiceManager, oDispatcher, oFrame, "A10:I10", 9)
        LOWrapText(oServiceManager, oDispatcher, oFrame, "A10:I10")
        Dim LineFormat As Object
        LineFormat = oServiceManager.Bridge_GetStruct("com.sun.star.table.BorderLine2")
        LineFormat.LineStyle = 0
        LineFormat.LineWidth = 40
        oSheet.getCellRangeByName("A10:I10").TopBorder = LineFormat
        oSheet.getCellRangeByName("A10:I10").RightBorder = LineFormat
        oSheet.getCellRangeByName("A10:I10").LeftBorder = LineFormat
        oSheet.getCellRangeByName("A10:I10").BottomBorder = LineFormat
        oSheet.getCellRangeByName("A10:I10").VertJustify = 2
        oSheet.getCellRangeByName("A10:I10").HoriJustify = 2

        '--����� ����� ������������
        MySQLStr = "SELECT ISNULL(tbl_SupplSearch_PropItems.ItemCode, N'') AS ItemCode, ISNULL(tbl_SupplSearch_PropItems.ItemSuppCode, N'') AS SuppItemCode, "
        MySQLStr = MySQLStr & "CASE WHEN LTRIM(RTRIM(LTRIM(RTRIM(ISNULL(SC010300.SC01002, ''))) + LTRIM(RTRIM(ISNULL(SC010300.SC01003, ''))))) "
        MySQLStr = MySQLStr & "= '' THEN ISNULL(tbl_SupplSearch_PropItems.ItemName, '') ELSE LTRIM(RTRIM(LTRIM(RTRIM(ISNULL(SC010300.SC01002, ''))) "
        MySQLStr = MySQLStr & "+ LTRIM(RTRIM(ISNULL(SC010300.SC01003, ''))))) END AS ItemName, CASE WHEN ISNULL(View_1.txt, N'') "
        MySQLStr = MySQLStr & "= '��' THEN 'pcs(��.)' ELSE CASE WHEN ISNULL(View_1.txt, N'') = '�' THEN 'm (�)' ELSE CASE WHEN ISNULL(View_1.txt, N'') "
        MySQLStr = MySQLStr & "= '��' THEN 'kg (��)' ELSE CASE WHEN ISNULL(View_1.txt, N'') = '��' THEN 'km (��)' ELSE CASE WHEN ISNULL(View_1.txt, N'') "
        MySQLStr = MySQLStr & "= '�' THEN 'litre (����)' ELSE CASE WHEN ISNULL(View_1.txt, N'') = '����' THEN 'pack (����.)' ELSE CASE WHEN ISNULL(View_1.txt, N'') "
        MySQLStr = MySQLStr & "= '�����' THEN 'set (�����.)' ELSE CASE WHEN ISNULL(View_1.txt, N'') = '���' THEN 'pair (����)' END END END END END END END END AS UOM, "
        MySQLStr = MySQLStr & "tbl_SupplSearch_PropItems.QTY, tbl_SupplSearch_PropItems.Price * (100 + ISNULL(View_5.ShippingCost, 10)) / 100 AS PriCost, "
        MySQLStr = MySQLStr & "ISNULL(tbl_SupplSearch_PropItems.LeadTimeWeek, 1) AS WeekQTY, "
        MySQLStr = MySQLStr & "ISNULL(tbl_SupplSearch_PropItems.AlternateTo, '') AS AlternateTo, ISNULL(tbl_SupplSearch_PropItems.Comments, '') AS Comments, "
        MySQLStr = MySQLStr & "ISNULL(tbl_SupplSearch_PropItems.SupplierCode, '') AS SupplierCode "
        MySQLStr = MySQLStr & "FROM tbl_SupplSearch_PropItems LEFT OUTER JOIN "
        MySQLStr = MySQLStr & "(SELECT PL01001, ShippingCost "
        MySQLStr = MySQLStr & "FROM tbl_SupplierCard0300) AS View_5 ON tbl_SupplSearch_PropItems.SupplierCode = View_5.PL01001 LEFT OUTER JOIN "
        MySQLStr = MySQLStr & "SC010300 ON tbl_SupplSearch_PropItems.ItemCode = SC010300.SC01001 LEFT OUTER JOIN "
        MySQLStr = MySQLStr & "(SELECT     0 AS Expr1, SC09002 AS txt "
        MySQLStr = MySQLStr & "FROM          SC090300 WITH (NOLOCK) "
        MySQLStr = MySQLStr & "WHERE      (SC09001 = 'RUS') "
        MySQLStr = MySQLStr & "UNION "
        MySQLStr = MySQLStr & "SELECT     1 AS Expr1, SC09003 "
        MySQLStr = MySQLStr & "FROM         SC090300 AS SC090300_40 WITH (NOLOCK) "
        MySQLStr = MySQLStr & "WHERE     (SC09001 = 'RUS') "
        MySQLStr = MySQLStr & "UNION "
        MySQLStr = MySQLStr & "SELECT     2 AS Expr1, SC09004 "
        MySQLStr = MySQLStr & "FROM         SC090300 AS SC090300_39 WITH (NOLOCK) "
        MySQLStr = MySQLStr & "WHERE     (SC09001 = 'RUS') "
        MySQLStr = MySQLStr & "UNION "
        MySQLStr = MySQLStr & "SELECT     3 AS Expr1, SC09005 "
        MySQLStr = MySQLStr & "FROM         SC090300 AS SC090300_38 WITH (NOLOCK) "
        MySQLStr = MySQLStr & "WHERE     (SC09001 = 'RUS') "
        MySQLStr = MySQLStr & "UNION "
        MySQLStr = MySQLStr & "SELECT     4 AS Expr1, SC09006 "
        MySQLStr = MySQLStr & "FROM         SC090300 AS SC090300_37 WITH (NOLOCK) "
        MySQLStr = MySQLStr & "WHERE     (SC09001 = 'RUS') "
        MySQLStr = MySQLStr & "UNION "
        MySQLStr = MySQLStr & "SELECT     5 AS Expr1, SC09007 "
        MySQLStr = MySQLStr & "FROM         SC090300 AS SC090300_36 WITH (NOLOCK) "
        MySQLStr = MySQLStr & "WHERE     (SC09001 = 'RUS') "
        MySQLStr = MySQLStr & "UNION "
        MySQLStr = MySQLStr & "SELECT     6 AS Expr1, SC09008 "
        MySQLStr = MySQLStr & "FROM         SC090300 AS SC090300_35 WITH (NOLOCK) "
        MySQLStr = MySQLStr & "WHERE     (SC09001 = 'RUS') "
        MySQLStr = MySQLStr & "UNION "
        MySQLStr = MySQLStr & "SELECT     7 AS Expr1, SC09009 "
        MySQLStr = MySQLStr & "FROM         SC090300 AS SC090300_34 WITH (NOLOCK) "
        MySQLStr = MySQLStr & "WHERE     (SC09001 = 'RUS') "
        MySQLStr = MySQLStr & "UNION "
        MySQLStr = MySQLStr & "SELECT     8 AS Expr1, SC09010 "
        MySQLStr = MySQLStr & "FROM         SC090300 AS SC090300_33 WITH (NOLOCK) "
        MySQLStr = MySQLStr & "WHERE     (SC09001 = 'RUS') "
        MySQLStr = MySQLStr & "UNION "
        MySQLStr = MySQLStr & "SELECT     9 AS Expr1, SC09011 "
        MySQLStr = MySQLStr & "FROM         SC090300 AS SC090300_32 WITH (NOLOCK) "
        MySQLStr = MySQLStr & "WHERE     (SC09001 = 'RUS') "
        MySQLStr = MySQLStr & "UNION "
        MySQLStr = MySQLStr & "SELECT     10 AS Expr1, SC09012 "
        MySQLStr = MySQLStr & "FROM         SC090300 AS SC090300_31 WITH (NOLOCK) "
        MySQLStr = MySQLStr & "WHERE     (SC09001 = 'RUS') "
        MySQLStr = MySQLStr & "UNION "
        MySQLStr = MySQLStr & "SELECT     11 AS Expr1, SC09013 "
        MySQLStr = MySQLStr & "FROM         SC090300 AS SC090300_30 WITH (NOLOCK) "
        MySQLStr = MySQLStr & "WHERE     (SC09001 = 'RUS') "
        MySQLStr = MySQLStr & "UNION "
        MySQLStr = MySQLStr & "SELECT     12 AS Expr1, SC09014 "
        MySQLStr = MySQLStr & "FROM         SC090300 AS SC090300_29 WITH (NOLOCK) "
        MySQLStr = MySQLStr & "WHERE     (SC09001 = 'RUS') "
        MySQLStr = MySQLStr & "UNION "
        MySQLStr = MySQLStr & "SELECT     13 AS Expr1, SC09015 "
        MySQLStr = MySQLStr & "FROM         SC090300 AS SC090300_28 WITH (NOLOCK) "
        MySQLStr = MySQLStr & "WHERE     (SC09001 = 'RUS') "
        MySQLStr = MySQLStr & "UNION "
        MySQLStr = MySQLStr & "SELECT     14 AS Expr1, SC09016 "
        MySQLStr = MySQLStr & "FROM         SC090300 AS SC090300_27 WITH (NOLOCK) "
        MySQLStr = MySQLStr & "WHERE     (SC09001 = 'RUS') "
        MySQLStr = MySQLStr & "UNION "
        MySQLStr = MySQLStr & "SELECT     15 AS Expr1, SC09017 "
        MySQLStr = MySQLStr & "FROM         SC090300 AS SC090300_26 WITH (NOLOCK) "
        MySQLStr = MySQLStr & "WHERE     (SC09001 = 'RUS') "
        MySQLStr = MySQLStr & "UNION "
        MySQLStr = MySQLStr & "SELECT     16 AS Expr1, SC09018 "
        MySQLStr = MySQLStr & "FROM         SC090300 AS SC090300_25 WITH (NOLOCK) "
        MySQLStr = MySQLStr & "WHERE     (SC09001 = 'RUS') "
        MySQLStr = MySQLStr & "UNION "
        MySQLStr = MySQLStr & "SELECT     17 AS Expr1, SC09019 "
        MySQLStr = MySQLStr & "FROM         SC090300 AS SC090300_24 WITH (NOLOCK) "
        MySQLStr = MySQLStr & "WHERE     (SC09001 = 'RUS') "
        MySQLStr = MySQLStr & "UNION "
        MySQLStr = MySQLStr & "SELECT     18 AS Expr1, SC09020 "
        MySQLStr = MySQLStr & "FROM         SC090300 AS SC090300_23 WITH (NOLOCK) "
        MySQLStr = MySQLStr & "WHERE     (SC09001 = 'RUS') "
        MySQLStr = MySQLStr & "UNION "
        MySQLStr = MySQLStr & "SELECT     19 AS Expr1, SC09021 "
        MySQLStr = MySQLStr & "FROM         SC090300 AS SC090300_22 WITH (NOLOCK) "
        MySQLStr = MySQLStr & "WHERE     (SC09001 = 'RUS') "
        MySQLStr = MySQLStr & "UNION "
        MySQLStr = MySQLStr & "SELECT     20 AS Expr1, SC09022 "
        MySQLStr = MySQLStr & "FROM         SC090300 AS SC090300_21 WITH (NOLOCK) "
        MySQLStr = MySQLStr & "WHERE     (SC09001 = 'RUS') "
        MySQLStr = MySQLStr & "UNION "
        MySQLStr = MySQLStr & "SELECT     21 AS Expr1, SC09023 "
        MySQLStr = MySQLStr & "FROM         SC090300 AS SC090300_20 WITH (NOLOCK) "
        MySQLStr = MySQLStr & "WHERE     (SC09001 = 'RUS') "
        MySQLStr = MySQLStr & "UNION "
        MySQLStr = MySQLStr & "SELECT     22 AS Expr1, SC09024 "
        MySQLStr = MySQLStr & "FROM         SC090300 AS SC090300_19 WITH (NOLOCK) "
        MySQLStr = MySQLStr & "WHERE     (SC09001 = 'RUS') "
        MySQLStr = MySQLStr & "UNION "
        MySQLStr = MySQLStr & "SELECT     23 AS Expr1, SC09025 "
        MySQLStr = MySQLStr & "FROM         SC090300 AS SC090300_18 WITH (NOLOCK) "
        MySQLStr = MySQLStr & "WHERE     (SC09001 = 'RUS') "
        MySQLStr = MySQLStr & "UNION "
        MySQLStr = MySQLStr & "SELECT     24 AS Expr1, SC09026 "
        MySQLStr = MySQLStr & "FROM         SC090300 AS SC090300_17 WITH (NOLOCK) "
        MySQLStr = MySQLStr & "WHERE     (SC09001 = 'RUS') "
        MySQLStr = MySQLStr & "UNION "
        MySQLStr = MySQLStr & "SELECT     25 AS Expr1, SC09027 "
        MySQLStr = MySQLStr & "FROM         SC090300 AS SC090300_16 WITH (NOLOCK) "
        MySQLStr = MySQLStr & "WHERE     (SC09001 = 'RUS') "
        MySQLStr = MySQLStr & "UNION "
        MySQLStr = MySQLStr & "SELECT     26 AS Expr1, SC09028 "
        MySQLStr = MySQLStr & "FROM         SC090300 AS SC090300_15 WITH (NOLOCK) "
        MySQLStr = MySQLStr & "WHERE     (SC09001 = 'RUS') "
        MySQLStr = MySQLStr & "UNION "
        MySQLStr = MySQLStr & "SELECT     27 AS Expr1, SC09029 "
        MySQLStr = MySQLStr & "FROM         SC090300 AS SC090300_14 WITH (NOLOCK) "
        MySQLStr = MySQLStr & "WHERE     (SC09001 = 'RUS') "
        MySQLStr = MySQLStr & "UNION "
        MySQLStr = MySQLStr & "SELECT     28 AS Expr1, SC09030 "
        MySQLStr = MySQLStr & "FROM         SC090300 AS SC090300_13 WITH (NOLOCK) "
        MySQLStr = MySQLStr & "WHERE     (SC09001 = 'RUS') "
        MySQLStr = MySQLStr & "UNION "
        MySQLStr = MySQLStr & "SELECT     29 AS Expr1, SC09031 "
        MySQLStr = MySQLStr & "FROM         SC090300 AS SC090300_12 WITH (NOLOCK) "
        MySQLStr = MySQLStr & "WHERE     (SC09001 = 'RUS') "
        MySQLStr = MySQLStr & "UNION "
        MySQLStr = MySQLStr & "SELECT     30 AS Expr1, SC09032 "
        MySQLStr = MySQLStr & "FROM         SC090300 AS SC090300_11 WITH (NOLOCK) "
        MySQLStr = MySQLStr & "WHERE     (SC09001 = 'RUS') "
        MySQLStr = MySQLStr & "UNION "
        MySQLStr = MySQLStr & "SELECT     31 AS Expr1, SC09033 "
        MySQLStr = MySQLStr & "FROM         SC090300 AS SC090300_10 WITH (NOLOCK) "
        MySQLStr = MySQLStr & "WHERE     (SC09001 = 'RUS') "
        MySQLStr = MySQLStr & "UNION "
        MySQLStr = MySQLStr & "SELECT     32 AS Expr1, SC09034 "
        MySQLStr = MySQLStr & "FROM         SC090300 AS SC090300_9 WITH (NOLOCK) "
        MySQLStr = MySQLStr & "WHERE     (SC09001 = 'RUS') "
        MySQLStr = MySQLStr & "UNION "
        MySQLStr = MySQLStr & "SELECT     33 AS Expr1, SC09035 "
        MySQLStr = MySQLStr & "FROM         SC090300 AS SC090300_8 WITH (NOLOCK) "
        MySQLStr = MySQLStr & "WHERE     (SC09001 = 'RUS') "
        MySQLStr = MySQLStr & "UNION "
        MySQLStr = MySQLStr & "SELECT     34 AS Expr1, SC09036 "
        MySQLStr = MySQLStr & "FROM         SC090300 AS SC090300_7 WITH (NOLOCK) "
        MySQLStr = MySQLStr & "WHERE     (SC09001 = 'RUS') "
        MySQLStr = MySQLStr & "UNION "
        MySQLStr = MySQLStr & "SELECT     35 AS Expr1, SC09037 "
        MySQLStr = MySQLStr & "FROM         SC090300 AS SC090300_6 WITH (NOLOCK) "
        MySQLStr = MySQLStr & "WHERE     (SC09001 = 'RUS') "
        MySQLStr = MySQLStr & "UNION "
        MySQLStr = MySQLStr & "SELECT     36 AS Expr1, SC09038 "
        MySQLStr = MySQLStr & "FROM         SC090300 AS SC090300_5 WITH (NOLOCK) "
        MySQLStr = MySQLStr & "WHERE     (SC09001 = 'RUS') "
        MySQLStr = MySQLStr & "UNION "
        MySQLStr = MySQLStr & "SELECT     37 AS Expr1, SC09039 "
        MySQLStr = MySQLStr & "FROM         SC090300 AS SC090300_4 WITH (NOLOCK) "
        MySQLStr = MySQLStr & "WHERE     (SC09001 = 'RUS') "
        MySQLStr = MySQLStr & "UNION "
        MySQLStr = MySQLStr & "SELECT     38 AS Expr1, SC09040 "
        MySQLStr = MySQLStr & "FROM         SC090300 AS SC090300_3 WITH (NOLOCK) "
        MySQLStr = MySQLStr & "WHERE     (SC09001 = 'RUS') "
        MySQLStr = MySQLStr & "UNION "
        MySQLStr = MySQLStr & "SELECT     39 AS Expr1, SC09041 "
        MySQLStr = MySQLStr & "FROM         SC090300 AS SC090300_2 WITH (NOLOCK) "
        MySQLStr = MySQLStr & "WHERE     (SC09001 = 'RUS') "
        MySQLStr = MySQLStr & "UNION "
        MySQLStr = MySQLStr & "SELECT     40 AS Expr1, SC09042 "
        MySQLStr = MySQLStr & "FROM SC090300 AS SC090300_1 WITH (NOLOCK)) AS View_1 ON tbl_SupplSearch_PropItems.UOM = View_1.Expr1 "
        MySQLStr = MySQLStr & "WHERE (tbl_SupplSearch_PropItems.SupplSearchID = " & DataGridView1.SelectedRows.Item(0).Cells("ID").Value & ") "
        If ComboBox1.Text = "������������ ������������" Then
            MySQLStr = MySQLStr & "AND (tbl_SupplSearch_PropItems.IsSelected = 1) "
        End If
        If ComboBox1.Text = "��������� ���������" Then
            MySQLStr = MySQLStr & "AND (tbl_SupplSearch_PropItems.SelectedBySalesman = 1) "
        End If
        Select Case MyCol
            Case 0
                MySQLStr = MySQLStr & "ORDER BY tbl_SupplSearch_PropItems.ID "
            Case 4
                MySQLStr = MySQLStr & "ORDER BY ItemCode "
            Case 5
                MySQLStr = MySQLStr & "ORDER BY SuppItemCode "
            Case 6
                MySQLStr = MySQLStr & "ORDER BY ItemName "
            Case 8
                MySQLStr = MySQLStr & "ORDER BY UOM "
            Case 9
                MySQLStr = MySQLStr & "ORDER BY tbl_SupplSearch_PropItems.QTY "
            Case Else
                MySQLStr = MySQLStr & "ORDER BY tbl_SupplSearch_PropItems.ID "
        End Select
        If MySort = 2 Then
            MySQLStr = MySQLStr & "Desc "
        End If
        InitMyConn(False)
        InitMyRec(False, MySQLStr)
        If Declarations.MyRec.BOF = True And Declarations.MyRec.EOF = True Then
            trycloseMyRec()
        Else
            Declarations.MyRec.MoveFirst()
            i = 0
            While Declarations.MyRec.EOF = False
                oSheet.getCellRangeByName("A" & CStr(i + 11)).Value = i + 1
                oSheet.getCellRangeByName("B" & CStr(i + 11)).String = Declarations.MyRec.Fields("ItemCode").Value
                oSheet.getCellRangeByName("C" & CStr(i + 11)).String = Declarations.MyRec.Fields("SuppItemCode").Value
                oSheet.getCellRangeByName("D" & CStr(i + 11)).String = Declarations.MyRec.Fields("ItemName").Value
                oSheet.getCellRangeByName("E" & CStr(i + 11)).String = Declarations.MyRec.Fields("UOM").Value
                If Not IsDBNull(Declarations.MyRec.Fields("QTY").Value) Then
                    oSheet.getCellRangeByName("F" & CStr(i + 11)).Value = Declarations.MyRec.Fields("QTY").Value
                End If
                If Not IsDBNull(Declarations.MyRec.Fields("PriCost").Value) Then
                    oSheet.getCellRangeByName("G" & CStr(i + 11)).Value = Declarations.MyRec.Fields("PriCost").Value
                End If
                oSheet.getCellRangeByName("H" & CStr(i + 11)).FormulaLocal = "=IF(F" & CStr(i + 11) & "*G" & CStr(i + 11) & " = 0;"""";F" & CStr(i + 11) & " * G" & CStr(i + 11) & ") "
                If Not IsDBNull(Declarations.MyRec.Fields("WeekQTY").Value) Then
                    oSheet.getCellRangeByName("I" & CStr(i + 11)).Value = Declarations.MyRec.Fields("WeekQTY").Value
                End If
                oSheet.getCellRangeByName("J" & CStr(i + 11)).String = Declarations.MyRec.Fields("AlternateTo").Value
                oSheet.getCellRangeByName("K" & CStr(i + 11)).String = Declarations.MyRec.Fields("Comments").Value
                oSheet.getCellRangeByName("L" & CStr(i + 11)).String = Declarations.MyRec.Fields("SupplierCode").Value

                i = i + 1
                Declarations.MyRec.MoveNext()
            End While
            trycloseMyRec()
        End If
        LOWrapText(oServiceManager, oDispatcher, oFrame, "A11:I" & CStr(11 + i - 1))
        LineFormat.LineStyle = 0
        LineFormat.LineWidth = 20
        oSheet.getCellRangeByName("A11:I" & CStr(11 + i - 1)).TopBorder = LineFormat
        oSheet.getCellRangeByName("A11:I" & CStr(11 + i - 1)).RightBorder = LineFormat
        oSheet.getCellRangeByName("A11:I" & CStr(11 + i - 1)).LeftBorder = LineFormat
        oSheet.getCellRangeByName("A11:I" & CStr(11 + i - 1)).BottomBorder = LineFormat
        '----������ �����
        LOSetCellProtection(oServiceManager, oDispatcher, oFrame, "A11:I500", False)
        '---�������� ���������� ����������
        oSheet.getCellRangeByName("N10").String = "������� ���������"
        oSheet.getCellRangeByName("N11").Value = 0
        oSheet.getCellRangeByName("O11").String = "pcs(��.)"
        oSheet.getCellRangeByName("N12").Value = 1
        oSheet.getCellRangeByName("O12").String = "m (�)"
        oSheet.getCellRangeByName("N13").Value = 2
        oSheet.getCellRangeByName("O13").String = "kg (��)"
        oSheet.getCellRangeByName("N14").Value = 3
        oSheet.getCellRangeByName("O14").String = "km (��)"
        oSheet.getCellRangeByName("N15").Value = 4
        oSheet.getCellRangeByName("O15").String = "litre (����)"
        oSheet.getCellRangeByName("N16").Value = 5
        oSheet.getCellRangeByName("O16").String = "pack (����.)"
        oSheet.getCellRangeByName("N17").Value = 6
        oSheet.getCellRangeByName("O17").String = "set (�����.)"
        oSheet.getCellRangeByName("N18").Value = 7
        oSheet.getCellRangeByName("O18").String = "pair (����)"
        oSheet.getCellRangeByName("N10:O18").CharColor = RGB(255, 255, 255) '---��� ����� �������� ������� - �� ����� ���� ��� �� RGB � BGR!!!! ������ ���� �����
        LOSetValidation(oSheet, "E11:E" & CStr(11 + i - 1), "=$O$11:$O$18")
        '----� ������ �����
        Dim args() As Object
        ReDim args(0)
        args(0) = oServiceManager.Bridge_GetStruct("com.sun.star.beans.PropertyValue")
        args(0).Name = "ToPoint"
        args(0).Value = "$A$1"
        oDispatcher.executeDispatch(oFrame, ".uno:GoToCell", "", 0, args)
        '----�������� �������
        LOPasswordProtect(oSheet, "!pass2022", True)
        '----���������
        oWorkBook.CurrentController.Frame.ContainerWindow.Visible = True
        oWorkBook.CurrentController.Frame.ContainerWindow.toFront()
    End Sub
End Class