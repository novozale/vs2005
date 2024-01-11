Public Class SelectItemBySuppCode
    Public MyItemSuppCode As String                             '��� ������ ����������
    Public MyWindowFrom As String                               '�� ������ ���� �������


    Private Sub SelectItemBySuppCode_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles Me.KeyDown
        '////////////////////////////////////////////////////////////////////////////////
        '//
        '// ������ �������� ���� �� ALT+F4
        '// 
        '//
        '////////////////////////////////////////////////////////////////////////////////

        If e.KeyData = Keys.Alt + Keys.F4 Then
            e.Handled = True
        End If
    End Sub

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        '////////////////////////////////////////////////////////////////////////////////
        '//
        '// ����� �� ���� � ������� ���� Scala
        '// 
        '//
        '////////////////////////////////////////////////////////////////////////////////

        If MyWindowFrom = "AddToOrder" Then
            If Trim(DataGridView1.SelectedRows.Item(0).Cells(2).Value) = "" Then    '--���������� ���� ���
                MyAddToOrder.Label3.Text = "��������������� ���� � ������������� ����� ������ �� ������ ����� - ����� �� �������"
                MyAddToOrder.Label3.ForeColor = Color.Green
                MyAddToOrder.TextBox2.Text = "Unknown"
                MyAddToOrder.TextBox1.Text = ""
                MyAddToOrder.TextBox1.ReadOnly = False
                MyAddToOrder.ComboBox1.Text = ""
                MyAddToOrder.ComboBox1.Enabled = True
                MyAddToOrder.TextBox3.Text = 1
                MyAddToOrder.TextBox4.Text = DataGridView1.SelectedRows.Item(0).Cells(7).Value / Declarations.CurrencyValue
                MyAddToOrder.TextBox5.Text = DataGridView1.SelectedRows.Item(0).Cells(10).Value / Declarations.CurrencyValue
                MyAddToOrder.TextBox5.ReadOnly = False
                MyAddToOrder.TextBox6.Text = ""
                MyAddToOrder.TextBox7.Text = "1"
                MyAddToOrder.TextBox13.Text = DataGridView1.SelectedRows.Item(0).Cells(0).Value
                MyAddToOrder.TextBox13.Enabled = True
                MyAddToOrder.TextBox13.BackColor = Color.FromName("Window")
                MyAddToOrder.Button7.Enabled = True
                MyAddToOrder.TextBox14.Text = DataGridView1.SelectedRows.Item(0).Cells(1).Value
                MyAddToOrder.TextBox14.Enabled = False
                MyAddToOrder.TextBox14.BackColor = Color.FromName("ButtonFace")
                MyAddToOrder.Label16.Text = ""
                MyAddToOrder.Label17.Text = ""
                MyAddToOrder.Label19.Text = Math.Round(DataGridView1.SelectedRows.Item(0).Cells(9).Value, 2)
                MyAddToOrder.Label21.Text = Math.Round(DataGridView1.SelectedRows.Item(0).Cells(8).Value, 0)
                MyAddToOrder.LoadItemAddInfo(Trim(MyAddToOrder.TextBox2.Text))
            Else                                                                    '--���� � �����
                MyAddToOrder.TextBox2.Text = DataGridView1.SelectedRows.Item(0).Cells(2).Value
                MyAddToOrder.TextBox2Validation()
                MyAddToOrder.LoadItemAddInfo(Trim(MyAddToOrder.TextBox2.Text))
            End If
        ElseIf MyWindowFrom = "EditInOrder" Then
            If Trim(DataGridView1.SelectedRows.Item(0).Cells(2).Value) = "" Then    '--���������� ���� ���
                MyEditInOrder.Label3.Text = "��������������� ���� � ������������� ����� ������ �� ������ ����� - ����� �� �������"
                MyEditInOrder.Label3.ForeColor = Color.Green
                MyEditInOrder.TextBox1.Text = "Unknown"
                MyEditInOrder.TextBox2.Text = ""
                MyEditInOrder.TextBox1.ReadOnly = False
                MyEditInOrder.ComboBox1.Text = ""
                MyEditInOrder.ComboBox1.Enabled = True
                MyEditInOrder.TextBox3.Text = 1
                MyEditInOrder.TextBox4.Text = DataGridView1.SelectedRows.Item(0).Cells(7).Value / Declarations.CurrencyValue
                MyEditInOrder.TextBox5.Text = DataGridView1.SelectedRows.Item(0).Cells(10).Value / Declarations.CurrencyValue
                MyEditInOrder.TextBox5.ReadOnly = False
                MyEditInOrder.TextBox10.Text = ""
                MyEditInOrder.TextBox11.Text = "1"
                MyEditInOrder.TextBox15.Text = DataGridView1.SelectedRows.Item(0).Cells(0).Value
                MyEditInOrder.TextBox15.Enabled = True
                MyEditInOrder.TextBox15.BackColor = Color.FromName("Window")
                MyEditInOrder.Button9.Enabled = True
                MyEditInOrder.TextBox16.Text = DataGridView1.SelectedRows.Item(0).Cells(1).Value
                MyEditInOrder.TextBox16.Enabled = False
                MyEditInOrder.TextBox16.BackColor = Color.FromName("ButtonFace")
                MyEditInOrder.Label25.Text = ""
                MyEditInOrder.Label23.Text = ""
                MyEditInOrder.Label19.Text = Math.Round(DataGridView1.SelectedRows.Item(0).Cells(9).Value, 2)
                MyEditInOrder.Label21.Text = Math.Round(DataGridView1.SelectedRows.Item(0).Cells(8).Value, 0)
                MyEditInOrder.TextBox6.Text = 0
                MyEditInOrder.TextBox4Validation()
                MyEditInOrder.LoadItemAddInfo1(Trim(MyEditInOrder.TextBox1.Text))
            Else                                                                    '--���� � �����
                MyEditInOrder.TextBox1.Text = DataGridView1.SelectedRows.Item(0).Cells(2).Value
                MyEditInOrder.TextBox1Validation()
                MyEditInOrder.LoadItemAddInfo1(Trim(MyEditInOrder.TextBox1.Text))
            End If
        ElseIf MyWindowFrom = "Import" Then
            Declarations.MyItemCode = DataGridView1.SelectedRows.Item(0).Cells(2).Value
        End If

        Me.Close()
    End Sub

    Private Sub SelectItemBySuppCode_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        '////////////////////////////////////////////////////////////////////////////////
        '//
        '// �������� ����, �������� ������
        '//
        '////////////////////////////////////////////////////////////////////////////////
        Dim MySQLStr As String                        '������� ������
        Dim MyAdapter As SqlClient.SqlDataAdapter     '
        Dim MyDs As New DataSet                       '

        Label2.Text = Me.MyItemSuppCode
        Declarations.MyOperationResult = 0

        MySQLStr = "exec spp_SalesWorkplace4_SpecificationFilling "
        MySQLStr = MySQLStr & "N'', "                                   '--��� ������ � Scala
        MySQLStr = MySQLStr & "N'" & Trim(Me.MyItemSuppCode) & "'  "    '--��� ������ ����������

        Try
            MyAdapter = New SqlClient.SqlDataAdapter(MySQLStr, Declarations.MyNETConnStr)
            MyAdapter.SelectCommand.CommandTimeout = 600
            MyAdapter.Fill(MyDs)
            DataGridView1.DataSource = MyDs.Tables(0)
        Catch ex As Exception
            MsgBox(ex.ToString)
        End Try

        DataGridView1.Columns(0).HeaderText = "��� ����������"
        DataGridView1.Columns(0).Width = 100
        DataGridView1.Columns(1).HeaderText = "���������"
        DataGridView1.Columns(1).Width = 200
        DataGridView1.Columns(2).HeaderText = "��� ������ Scala"
        DataGridView1.Columns(2).Width = 100
        DataGridView1.Columns(3).HeaderText = "��� ������ ����������"
        DataGridView1.Columns(3).Width = 120
        DataGridView1.Columns(4).HeaderText = "�������� ������"
        DataGridView1.Columns(4).Width = 300
        DataGridView1.Columns(5).HeaderText = "������� ���������"
        DataGridView1.Columns(5).Width = 0
        DataGridView1.Columns(5).Visible = False
        DataGridView1.Columns(6).HeaderText = "�������� �� ������"
        DataGridView1.Columns(6).Width = 100
        DataGridView1.Columns(7).HeaderText = "����"
        DataGridView1.Columns(7).Width = 100
        DataGridView1.Columns(8).HeaderText = "���� �������� (���)"
        DataGridView1.Columns(8).Width = 80
        DataGridView1.Columns(9).HeaderText = "��������� � ��������"
        DataGridView1.Columns(9).Width = 80
        DataGridView1.Columns(10).HeaderText = "�������������"
        DataGridView1.Columns(10).Width = 100
        DataGridView1.Columns(11).HeaderText = "������� ���������"
        DataGridView1.Columns(11).Width = 80

    End Sub

    Private Sub Button2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button2.Click
        '////////////////////////////////////////////////////////////////////////////////
        '//
        '// ����� �� ���� ��� ������ ���� Scala
        '// 
        '//
        '////////////////////////////////////////////////////////////////////////////////

        If MyWindowFrom = "AddToOrder" Then
            MyAddToOrder.TextBox2.Text = ""                      '--��� Scala
            MyAddToOrder.TextBox2Validation()
            MyAddToOrder.LoadItemAddInfo(Trim(MyAddToOrder.TextBox2.Text))
            MyAddToOrder.TextBox1.Text = ""                      '--�������� ������
            MyAddToOrder.TextBox3.Text = ""                      '--����������
            MyAddToOrder.TextBox4.Text = ""                      '--����
            MyAddToOrder.TextBox5.Text = ""                      '--�������������
            MyAddToOrder.TextBox6.Text = ""                      '--���� �������� (���)
            MyAddToOrder.TextBox13.Text = ""
            MyAddToOrder.TextBox14.Text = ""
            MyAddToOrder.Label17.Text = ""
            MyAddToOrder.Label19.Text = ""
            MyAddToOrder.Label21.Text = ""
        ElseIf MyWindowFrom = "EditInOrder" Then
            MyEditInOrder.TextBox1.Text = ""
            MyEditInOrder.TextBox1Validation()
            MyEditInOrder.LoadItemAddInfo1(Trim(MyEditInOrder.TextBox1.Text))
            MyEditInOrder.TextBox2.Text = ""
            MyEditInOrder.TextBox3.Text = ""
            MyEditInOrder.TextBox4.Text = ""
            MyEditInOrder.TextBox5.Text = ""
            MyEditInOrder.TextBox6.Text = ""
            MyEditInOrder.TextBox7.Text = ""
            MyEditInOrder.TextBox10.Text = ""
            MyEditInOrder.Label23.Text = ""
            MyEditInOrder.Label19.Text = ""
            MyEditInOrder.Label21.Text = ""
        End If
        Me.Close()
    End Sub
End Class