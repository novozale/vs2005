Public Class OrderDetails

    Private Sub Button13_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button13.Click
        '////////////////////////////////////////////////////////////////////////////////
        '//
        '// �����
        '//
        '////////////////////////////////////////////////////////////////////////////////

        Me.Close()
    End Sub

    Private Sub OrderDetails_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        '////////////////////////////////////////////////////////////////////////////////
        '//
        '// �������� ���� � ��������� ����������� �� ������
        '//
        '////////////////////////////////////////////////////////////////////////////////

        Label9.Text = Declarations.MyOrderNum
        LoadOrderDetails()
    End Sub

    Private Sub LoadOrderDetails()
        '////////////////////////////////////////////////////////////////////////////////
        '//
        '// �������� ��������� ���������� �� ������
        '//
        '////////////////////////////////////////////////////////////////////////////////
        Dim MySQLStr As String
        Dim MyAdapter As SqlClient.SqlDataAdapter     '��� ������ �����������
        Dim MyDs As New DataSet                       '

        MySQLStr = "Exec spp_Shipments_SalesmanWP_OrderInfo N'" & Trim(Declarations.MyOrderNum) & "' "

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
        DataGridView1.Columns(0).Width = 70
        DataGridView1.Columns(1).HeaderText = "��� ������"
        DataGridView1.Columns(1).Width = 70
        DataGridView1.Columns(2).HeaderText = "��������"
        DataGridView1.Columns(2).Width = 250
        DataGridView1.Columns(3).HeaderText = "���� ������"
        DataGridView1.Columns(3).Width = 60
        DataGridView1.Columns(4).HeaderText = "������ ������"
        DataGridView1.Columns(4).Width = 60
        DataGridView1.Columns(5).HeaderText = "���� (���)"
        DataGridView1.Columns(5).Width = 100
        DataGridView1.Columns(6).HeaderText = "����� ����"
        DataGridView1.Columns(6).Width = 50
        DataGridView1.Columns(7).HeaderText = "��������"
        DataGridView1.Columns(7).Width = 150
        DataGridView1.Columns(8).HeaderText = "������� ����"
        DataGridView1.Columns(8).Width = 80
        DataGridView1.Columns(9).HeaderText = "���� ��������"
        DataGridView1.Columns(9).Width = 80
        DataGridView1.Columns(10).HeaderText = "N ������ �� �������"
        DataGridView1.Columns(10).Width = 80
        DataGridView1.Columns(11).HeaderText = "��� ������ �� �������"
        DataGridView1.Columns(11).Width = 80
        DataGridView1.Columns(12).HeaderText = "���� ������ �� �������"
        DataGridView1.Columns(12).Width = 80
        DataGridView1.Columns(13).HeaderText = "����� ������ �� �������"
        DataGridView1.Columns(13).Width = 80
        DataGridView1.Columns(14).HeaderText = "����� ������ ���� ��������"
        DataGridView1.Columns(14).Width = 80
        DataGridView1.Columns(15).HeaderText = "������ ������ �� �������"
        DataGridView1.Columns(15).Width = 120
        DataGridView1.Columns(16).HeaderText = "N ������ �� ������ �����"
        DataGridView1.Columns(16).Width = 80
        DataGridView1.Columns(17).HeaderText = "���� ������� ������ �� ������ �����"
        DataGridView1.Columns(17).Width = 80
        DataGridView1.Columns(18).HeaderText = "������ ������ �� �����������"
        DataGridView1.Columns(18).Width = 120
        DataGridView1.Columns(19).HeaderText = "���� ��������"
        DataGridView1.Columns(19).Width = 80
    End Sub

    Private Sub DataGridView1_CellFormatting(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellFormattingEventArgs) Handles DataGridView1.CellFormatting
        '////////////////////////////////////////////////////////////////////////////////
        '//
        '// ��������� ����������
        '//
        '////////////////////////////////////////////////////////////////////////////////
        Dim row As DataGridViewRow

        row = DataGridView1.Rows(e.RowIndex)

        '---�������������
        If row.Cells(4).Value = 0 Then
            row.Cells(4).Style.BackColor = Color.LightPink
        Else
            If (row.Cells(4).Value < row.Cells(3).Value) Then
                row.Cells(4).Style.BackColor = Color.LightYellow
            Else
                row.Cells(4).Style.BackColor = Color.LightGreen
            End If
        End If

        '---���� ������� ������
        If IsDBNull(row.Cells(9).Value) = False And IsDBNull(row.Cells(19).Value) = False Then
            If row.Cells(9).Value > row.Cells(19).Value Then
                row.Cells(9).Style.BackColor = Color.LightYellow
            Else
                row.Cells(9).Style.BackColor = Color.White
            End If
        End If

        '---����������� ���� ������� ������
        If IsDBNull(row.Cells(14).Value) = False And IsDBNull(row.Cells(19).Value) = False Then
            If row.Cells(14).Value > row.Cells(19).Value Then
                row.Cells(14).Style.BackColor = Color.LightYellow
            Else
                row.Cells(14).Style.BackColor = Color.White
            End If
        End If
    End Sub
End Class