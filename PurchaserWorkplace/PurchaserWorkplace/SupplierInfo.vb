Public Class SupplierInfo

    Private Sub SupplierInfo_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles Me.KeyDown
        '////////////////////////////////////////////////////////////////////////////////
        '//
        '// ������ �������� ���� �� ALT+F4
        '//
        '////////////////////////////////////////////////////////////////////////////////

        If e.KeyData = Keys.Alt + Keys.F4 Then
            e.Handled = True
        End If
    End Sub

    Private Sub SupplierInfo_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        '////////////////////////////////////////////////////////////////////////////////
        '//
        '// ��������� ���� ������������ ����������������� �������
        '//
        '////////////////////////////////////////////////////////////////////////////////
        Dim MySQLStr As String                        '������� ������
        Dim MyAdapter As SqlClient.SqlDataAdapter     '��� ������ �������
        Dim MyDs As New DataSet                       '

        '---������ �������
        MySQLStr = "SELECT SC23001 AS WHCode, SC23001 + ' ' + SC23002 AS WHName "
        MySQLStr = MySQLStr & "FROM SC230300 WITH(NOLOCK) "
        MySQLStr = MySQLStr & "WHERE (SC23006 = N'1') OR (LEFT(SC23006, 2) = N'TR') "
        MySQLStr = MySQLStr & "ORDER BY WHCode "

        Label9.Text = Declarations.MySupplierCode
        Label2.Text = Trim(MainForm.DataGridView1.SelectedRows.Item(0).Cells(1).Value.ToString())
        ComboBox1.SelectedText = "������ �������� (����������)"

        LoadConsolidatedOrders()
        CheckConsolidatedButtons()
    End Sub

    Private Sub LoadConsolidatedOrders()
        '////////////////////////////////////////////////////////////////////////////////
        '//
        '// �������� ���������� �� ���������� �������
        '//
        '////////////////////////////////////////////////////////////////////////////////
        Dim MySQLStr As String
        Dim MyAdapter As SqlClient.SqlDataAdapter     '��� ������ �������
        Dim MyDs As New DataSet

        MySQLStr = "spp_PurchaseWorkplace_TotalGroupOrdersPrep N'" & Declarations.MyWH & "', N'" & Declarations.MySupplierCode & "', "
        If ComboBox1.Text = "������ �������� (����������)" Then
            MySQLStr = MySQLStr & "1, "
        Else
            MySQLStr = MySQLStr & "0, "
        End If
        MySQLStr = MySQLStr & "1 "

        Try
            MyAdapter = New SqlClient.SqlDataAdapter(MySQLStr, Declarations.MyNETConnStr)
            MyAdapter.SelectCommand.CommandTimeout = 600
            MyAdapter.Fill(MyDs)
            DataGridView1.DataSource = MyDs.Tables(0)
        Catch ex As Exception
            MsgBox(ex.ToString)
        End Try


        '---���������
        DataGridView1.Columns(0).HeaderText = "N ������"
        DataGridView1.Columns(0).Width = 80
        DataGridView1.Columns(1).HeaderText = "���� ������"
        DataGridView1.Columns(1).Width = 80
        DataGridView1.Columns(2).HeaderText = "����� ������ ���"
        DataGridView1.Columns(2).Width = 130
        DataGridView1.Columns(3).HeaderText = "��������"
        DataGridView1.Columns(3).Width = 140
        DataGridView1.Columns(4).HeaderText = "���� ���������� ������"
        DataGridView1.Columns(4).Width = 80
        DataGridView1.Columns(5).HeaderText = "���� ������������� ������"
        DataGridView1.Columns(5).Width = 80
        DataGridView1.Columns(6).HeaderText = "N ������ ����������"
        DataGridView1.Columns(6).Width = 130
        DataGridView1.Columns(7).HeaderText = "�� ������"
        DataGridView1.Columns(7).Width = 100
        DataGridView1.Columns(7).Visible = False
        DataGridView1.Columns(8).HeaderText = "�������������� ���� ��������"
        DataGridView1.Columns(8).Width = 80
        DataGridView1.Columns(9).HeaderText = "���-�� ���������� �������"
        DataGridView1.Columns(9).Width = 90
        DataGridView1.Columns(10).HeaderText = "������� ������� ���� ��������"
        DataGridView1.Columns(10).Width = 80
        DataGridView1.Columns(11).HeaderText = "����������� ����"
        DataGridView1.Columns(11).Width = 80
        DataGridView1.Columns(12).HeaderText = "���������� ����������"
        DataGridView1.Columns(12).Width = 280
        DataGridView1.Columns(13).HeaderText = "�����������"
        DataGridView1.Columns(13).Width = 280
        DataGridView1.Columns(14).HeaderText = "���� �� ������"
        DataGridView1.Columns(14).Width = 150
        DataGridView1.Columns(15).HeaderText = "�����"
        DataGridView1.Columns(15).Width = 80
    End Sub

    Private Sub CheckConsolidatedButtons()
        '////////////////////////////////////////////////////////////////////////////////
        '//
        '// �������� ��������� ������ ������ � ����������� ��������
        '//
        '////////////////////////////////////////////////////////////////////////////////

        If DataGridView1.SelectedRows.Count = 0 Then
            Button19.Enabled = False
        Else
            If Trim(DataGridView1.SelectedRows.Item(0).Cells(14).Value.ToString()) <> "" Then
                Button19.Enabled = True
            Else
                Button19.Enabled = False
            End If
        End If
    End Sub

    Private Sub ComboBox1_SelectedValueChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles ComboBox1.SelectedValueChanged
        '////////////////////////////////////////////////////////////////////////////////
        '//
        '// ����� ������ - ��� ������ ��� ������ ��������
        '//
        '////////////////////////////////////////////////////////////////////////////////

        LoadConsolidatedOrders()
        CheckConsolidatedButtons()
    End Sub

    Private Sub DataGridView1_CellFormatting(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellFormattingEventArgs) Handles DataGridView1.CellFormatting
        '////////////////////////////////////////////////////////////////////////////////
        '//
        '// ��������� ���������� �� ���������� �������
        '//
        '////////////////////////////////////////////////////////////////////////////////
        Dim row As DataGridViewRow

        row = DataGridView1.Rows(e.RowIndex)
        If row.Cells(9).Value = 0 Then
            row.DefaultCellStyle.BackColor = Color.LightYellow
        Else
            If (Trim(row.Cells(4).Value.ToString) = "" Or Trim(row.Cells(5).Value.ToString) = "") And Trim(row.Cells(7).Value.ToString) = "" Then
                row.DefaultCellStyle.BackColor = Color.LightPink
            Else
                row.DefaultCellStyle.BackColor = Color.Empty
            End If
        End If
        If Trim(row.Cells(7).Value.ToString) <> "" Then
            If IsDBNull(row.Cells(8).Value) Then
                row.Cells(8).Style.BackColor = Color.Empty
            Else
                If row.Cells(8).Value < Now() Then
                    row.Cells(8).Style.BackColor = Color.LightPink
                ElseIf row.Cells(8).Value < DateAdd(DateInterval.Day, 3, Now()) Then
                    row.Cells(8).Style.BackColor = Color.Yellow
                Else
                    row.Cells(8).Style.BackColor = Color.Empty
                End If
            End If
            If IsDBNull(row.Cells(11).Value) Then
                row.Cells(11).Style.BackColor = Color.Empty
            Else
                If row.Cells(11).Value < DateAdd(DateInterval.Day, -2, Now()) Then
                    row.Cells(11).Style.BackColor = Color.Empty
                ElseIf row.Cells(11).Value <= DateAdd(DateInterval.Day, 3, Now()) Then
                    row.Cells(11).Style.BackColor = Color.Yellow
                Else
                    row.Cells(11).Style.BackColor = Color.Empty
                End If
            End If
        Else
            row.Cells(8).Style.BackColor = Color.Empty
            row.Cells(11).Style.BackColor = Color.Empty
        End If
    End Sub

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        '////////////////////////////////////////////////////////////////////////////////
        '//
        '// ���������� ���������� � ����
        '//
        '////////////////////////////////////////////////////////////////////////////////

        LoadConsolidatedOrders()
        CheckConsolidatedButtons()
    End Sub

    Private Sub Button19_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button19.Click
        '////////////////////////////////////////////////////////////////////////////////
        '//
        '// �������� ���������������� �����.
        '//
        '////////////////////////////////////////////////////////////////////////////////

        If System.IO.Directory.Exists(My.Settings.BillPath + Trim(DataGridView1.SelectedRows.Item(0).Cells(14).Value.ToString())) = False _
            And System.IO.File.Exists(My.Settings.BillPath + Trim(DataGridView1.SelectedRows.Item(0).Cells(14).Value.ToString())) = False Then
            MsgBox("���� ��� ������� " + My.Settings.BillPath + Trim(DataGridView1.SelectedRows.Item(0).Cells(14).Value.ToString()) + " �� ������.", MsgBoxStyle.Critical, "��������!")
        Else
            If System.IO.Directory.Exists(My.Settings.BillPath + Trim(DataGridView1.SelectedRows.Item(0).Cells(14).Value.ToString())) Then
                Try
                    Process.Start("explorer.exe", My.Settings.BillPath + Trim(DataGridView1.SelectedRows.Item(0).Cells(14).Value.ToString()))
                Catch ex As Exception
                    MsgBox(ex.Message, MsgBoxStyle.Critical, "��������!")
                End Try
            Else
                Try
                    Dim startInfo As New ProcessStartInfo("CMD.EXE")
                    startInfo.WindowStyle = ProcessWindowStyle.Hidden
                    startInfo.CreateNoWindow = True
                    startInfo.UseShellExecute = False
                    startInfo.Arguments = "/c " + """" + My.Settings.BillPath + Trim(DataGridView1.SelectedRows.Item(0).Cells(14).Value.ToString()) + """"
                    Process.Start(startInfo)
                Catch ex As Exception
                    MsgBox(ex.Message, MsgBoxStyle.Critical, "��������!")
                End Try
            End If
        End If
    End Sub

    Private Sub DataGridView1_SelectionChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles DataGridView1.SelectionChanged
        '////////////////////////////////////////////////////////////////////////////////
        '//
        '// ����� ������ ����������� ������
        '//
        '////////////////////////////////////////////////////////////////////////////////

        CheckConsolidatedButtons()
    End Sub

    Private Sub Button13_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button13.Click
        '////////////////////////////////////////////////////////////////////////////////
        '//
        '// ����� �� ����
        '//
        '////////////////////////////////////////////////////////////////////////////////

        Me.Close()
    End Sub
End Class