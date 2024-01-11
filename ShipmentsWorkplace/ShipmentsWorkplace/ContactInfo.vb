Public Class ContactInfo

    Public StartParam As String

    Private Sub Button5_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button5.Click
        '////////////////////////////////////////////////////////////////////////////////
        '//
        '// ����� ��� ������
        '//
        '////////////////////////////////////////////////////////////////////////////////

        Me.Close()
    End Sub

    Private Sub ContactInfo_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles Me.KeyDown
        '/////////////////////////////////////////////////////////////////////////////////////
        '//
        '// ����� ��� ������ �� Escape
        '// 
        '/////////////////////////////////////////////////////////////////////////////////////

        If e.KeyData = Keys.Escape Then
            Me.Close()
        End If
    End Sub

    Private Sub ContactInfo_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        '/////////////////////////////////////////////////////////////////////////////////////
        '//
        '// ��� ������� ��������� ������ ���������
        '// 
        '/////////////////////////////////////////////////////////////////////////////////////

        If StartParam = "Contact" Then
            Label2.Text = Trim(MyShipment.LblCustomerCode.Text) + " " + Trim(MyShipment.LblCustomerName.Text)
        Else
            Label2.Text = Trim(MyShipmentsList.DataGridView1.SelectedRows.Item(0).Cells(1).Value.ToString()) + " " + Trim(MyShipmentsList.DataGridView1.SelectedRows.Item(0).Cells(2).Value.ToString())
        End If
        LoadData()
        CheckButtons()
    End Sub

    Private Sub LoadData()
        '/////////////////////////////////////////////////////////////////////////////////////
        '//
        '// �������� ������ ���������
        '// 
        '/////////////////////////////////////////////////////////////////////////////////////

        Dim MySQLStr As String                        '������� ������
        Dim MyAdapter As SqlClient.SqlDataAdapter     '
        Dim MyDs As New DataSet                       '

        If StartParam = "Contact" Then
            MySQLStr = "SELECT tbl_CRM_Contacts.ContactID, tbl_CRM_Contacts.CompanyID, tbl_CRM_Contacts.ContactName, tbl_CRM_Contacts.ContactPhone, "
            MySQLStr = MySQLStr & "tbl_CRM_Contacts.ContactEMail, ISNULL(tbl_CRM_Contacts.Comments,'') AS Comments, CASE WHEN tbl_CRM_Contacts.FromScala = 0 THEN '' ELSE 'X' END AS FromScala "
            MySQLStr = MySQLStr & "FROM tbl_CRM_Contacts WITH(NOLOCK) INNER JOIN "
            MySQLStr = MySQLStr & "tbl_CRM_Companies ON tbl_CRM_Contacts.CompanyID = tbl_CRM_Companies.CompanyID "
            MySQLStr = MySQLStr & "WHERE (tbl_CRM_Companies.ScalaCustomerCode = '" & Trim(MyShipment.LblCustomerCode.Text) & "') "
            MySQLStr = MySQLStr & "ORDER BY tbl_CRM_Contacts.ContactName "
        Else
            MySQLStr = "SELECT tbl_CRM_Contacts.ContactID, tbl_CRM_Contacts.CompanyID, tbl_CRM_Contacts.ContactName, tbl_CRM_Contacts.ContactPhone, "
            MySQLStr = MySQLStr & "tbl_CRM_Contacts.ContactEMail, ISNULL(tbl_CRM_Contacts.Comments,'') AS Comments, CASE WHEN tbl_CRM_Contacts.FromScala = 0 THEN '' ELSE 'X' END AS FromScala "
            MySQLStr = MySQLStr & "FROM tbl_CRM_Contacts WITH(NOLOCK) INNER JOIN "
            MySQLStr = MySQLStr & "tbl_CRM_Companies ON tbl_CRM_Contacts.CompanyID = tbl_CRM_Companies.CompanyID "
            MySQLStr = MySQLStr & "WHERE (tbl_CRM_Companies.ScalaCustomerCode = '" & Trim(MySendInfo.TextBox3.Text) & "') "
            MySQLStr = MySQLStr & "ORDER BY tbl_CRM_Contacts.ContactName "
        End If

        Try
            MyAdapter = New SqlClient.SqlDataAdapter(MySQLStr, Declarations.MyNETConnStr)
            MyAdapter.SelectCommand.CommandTimeout = 600
            MyAdapter.Fill(MyDs)
            DataGridView1.DataSource = MyDs.Tables(0)
        Catch ex As Exception
            MsgBox(ex.ToString)
        End Try

        DataGridView1.Columns(0).HeaderText = "ID"
        DataGridView1.Columns(0).Width = 40
        DataGridView1.Columns(0).Visible = False
        DataGridView1.Columns(1).HeaderText = "CID"
        DataGridView1.Columns(1).Width = 40
        DataGridView1.Columns(1).Visible = False
        DataGridView1.Columns(2).HeaderText = "���������� ����"
        DataGridView1.Columns(2).Width = 237
        DataGridView1.Columns(3).HeaderText = "�������"
        DataGridView1.Columns(3).Width = 150
        DataGridView1.Columns(4).HeaderText = "E-Mail"
        DataGridView1.Columns(4).Width = 150
        DataGridView1.Columns(5).HeaderText = "�����������"
        DataGridView1.Columns(5).Width = 150
        DataGridView1.Columns(6).HeaderText = "�� Scala"
        DataGridView1.Columns(6).Width = 40

    End Sub

    Private Sub CheckButtons()
        '/////////////////////////////////////////////////////////////////////////////////////
        '//
        '// ����������� ��������� ������
        '// 
        '/////////////////////////////////////////////////////////////////////////////////////

        If DataGridView1.SelectedRows.Count = 0 Then
            Button4.Enabled = False
        Else
            Button4.Enabled = True
        End If
    End Sub

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        '/////////////////////////////////////////////////////////////////////////////////////
        '//
        '// ���������� ������ ���������
        '// 
        '/////////////////////////////////////////////////////////////////////////////////////

        If StartParam = "Contact" Then
            Label2.Text = Trim(MyShipment.LblCustomerCode.Text) + " " + Trim(MyShipment.LblCustomerName.Text)
        Else
            Label2.Text = Trim(MyShipmentsList.DataGridView1.SelectedRows.Item(0).Cells(1).Value.ToString()) + " " + Trim(MyShipmentsList.DataGridView1.SelectedRows.Item(0).Cells(2).Value.ToString())
        End If
        LoadData()
        CheckButtons()
    End Sub

    Private Sub DataGridView1_CellDoubleClick(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs) Handles DataGridView1.CellDoubleClick
        '////////////////////////////////////////////////////////////////////////////////
        '//
        '// ����� �� ���� � ������� ��������
        '//
        '////////////////////////////////////////////////////////////////////////////////

        If DataGridView1.SelectedRows.Count <> 0 Then
            ContactSelect()
        End If
    End Sub

    Private Sub Button4_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button4.Click
        '////////////////////////////////////////////////////////////////////////////////
        '//
        '// ����� �� ���� � ������� ��������
        '//
        '////////////////////////////////////////////////////////////////////////////////

        ContactSelect()
    End Sub

    Private Sub ContactSelect()
        '////////////////////////////////////////////////////////////////////////////////
        '//
        '// ����� �� ���� � ������� ��������
        '//
        '////////////////////////////////////////////////////////////////////////////////

        If StartParam = "Contact" Then
            MyShipment.TextBox1.Text = "���������� ����: " + Trim(Me.DataGridView1.SelectedRows.Item(0).Cells(2).Value.ToString()) + " �������: " + Trim(Me.DataGridView1.SelectedRows.Item(0).Cells(3).Value.ToString()) + " E-Mail: " + Trim(Me.DataGridView1.SelectedRows.Item(0).Cells(4).Value.ToString())
        Else
            MySendInfo.TextBox2.Text = Trim(Me.DataGridView1.SelectedRows.Item(0).Cells(4).Value.ToString())
        End If
        Me.Close()
    End Sub

    Private Sub Button2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button2.Click
        '////////////////////////////////////////////////////////////////////////////////
        '//
        '// ����� �� ���� � ����������� ��������
        '//
        '////////////////////////////////////////////////////////////////////////////////

        ContactSelectAdd()
    End Sub

    Private Sub ContactSelectAdd()
        '////////////////////////////////////////////////////////////////////////////////
        '//
        '// ����� �� ���� � ����������� ��������
        '//
        '////////////////////////////////////////////////////////////////////////////////

        If StartParam = "Contact" Then
            If Trim(MyShipment.TextBox1.Text) = "" Then
                MyShipment.TextBox1.Text = "���������� ����: " + Trim(Me.DataGridView1.SelectedRows.Item(0).Cells(2).Value.ToString()) + " �������: " + Trim(Me.DataGridView1.SelectedRows.Item(0).Cells(3).Value.ToString()) + " E-Mail: " + Trim(Me.DataGridView1.SelectedRows.Item(0).Cells(4).Value.ToString())
            Else
                MyShipment.TextBox1.Text = MyShipment.TextBox1.Text + " " + "���������� ����: " + Trim(Me.DataGridView1.SelectedRows.Item(0).Cells(2).Value.ToString()) + " �������: " + Trim(Me.DataGridView1.SelectedRows.Item(0).Cells(3).Value.ToString()) + " E-Mail: " + Trim(Me.DataGridView1.SelectedRows.Item(0).Cells(4).Value.ToString())
            End If
        Else
            If Trim(MySendInfo.TextBox2.Text) = "" Then
                MySendInfo.TextBox2.Text = Trim(Me.DataGridView1.SelectedRows.Item(0).Cells(4).Value.ToString())
            Else
                MySendInfo.TextBox2.Text = MySendInfo.TextBox2.Text + "; " + Trim(Me.DataGridView1.SelectedRows.Item(0).Cells(4).Value.ToString())
            End If
        End If
        Me.Close()
    End Sub
End Class