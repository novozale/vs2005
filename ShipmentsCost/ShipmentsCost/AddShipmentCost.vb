Public Class AddShipmentCost

    Public StartParam As String


    Private Sub Button4_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button4.Click
        '////////////////////////////////////////////////////////////////////////////////
        '//
        '// �������� ���� ����� ��������� ��������
        '//
        '////////////////////////////////////////////////////////////////////////////////
        Dim MySQLStr As String

        '---�������� - ���� ��� �� ������ ������� - ������� ������ � ���������
        MySQLStr = "SELECT COUNT(ID) AS CC "
        MySQLStr = MySQLStr & "FROM tbl_ShipmentsCost_FactByInvoices WITH (NOLOCK) "
        MySQLStr = MySQLStr & "WHERE (DocID = N'" & Declarations.MyRecordID & "') "
        InitMyConn(False)
        InitMyRec(False, MySQLStr)
        If Declarations.MyRec.Fields("CC").Value = 0 Then
            trycloseMyRec()
            MySQLStr = "DELETE FROM tbl_ShipmentsCost_Fact "
            MySQLStr = MySQLStr & "WHERE (ID = N'" & Declarations.MyRecordID & "')"
            InitMyConn(False)
            Declarations.MyConn.Execute(MySQLStr)
        End If
        Me.Close()
    End Sub

    Private Sub AddShipmentCost_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles Me.KeyDown
        '/////////////////////////////////////////////////////////////////////////////////////
        '//
        '// ������ �������� �� ALT + F4
        '// 
        '/////////////////////////////////////////////////////////////////////////////////////

        If e.KeyData = Keys.Alt + Keys.F4 Then
            e.Handled = True
        End If
    End Sub

    Private Sub AddShipmentCost_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        '/////////////////////////////////////////////////////////////////////////////////////
        '//
        '// �������� ����� � �������� � �����
        '// 
        '/////////////////////////////////////////////////////////////////////////////////////
        Dim MyGUID As Guid
        Dim MySQLStr As String                        '������� ������
        Dim MyAdapter As SqlClient.SqlDataAdapter     '��� ������� ��������
        Dim MyDs As New DataSet

        '---������ ��������
        MySQLStr = "SELECT ID, Name "
        MySQLStr = MySQLStr & "FROM tbl_ShipmentsCost_ShipmentsType WITH (NOLOCK) "
        MySQLStr = MySQLStr & "UNION "
        MySQLStr = MySQLStr & "SELECT 0 AS ID, '' AS Name "
        MySQLStr = MySQLStr & "ORDER BY ID "
        Try
            MyAdapter = New SqlClient.SqlDataAdapter(MySQLStr, Declarations.MyNETConnStr)
            MyAdapter.SelectCommand.CommandTimeout = 600
            MyAdapter.Fill(MyDs)
            ComboBox2.DisplayMember = "Name" '��� �� ��� ����� ������������
            ComboBox2.ValueMember = "ID"   '��� �� ��� ����� ���������
            ComboBox2.DataSource = MyDs.Tables(0).DefaultView
        Catch ex As Exception
            MsgBox(ex.ToString)
        End Try

        '----� ����������� �� ���� - ��� ����� �������� ��� �������������� - ������ �������������� ������ � ���������� ��������
        If StartParam = "Create" Then      '---�������� ������ � ��������
            MyGUID = Guid.NewGuid
            Declarations.MyRecordID = MyGUID.ToString
            TextBox3.Enabled = True
            TextBox4.Enabled = True
            TextBox5.Enabled = True
            DateTimePicker1.Enabled = True
            TextBox1.Enabled = True
            ComboBox2.Enabled = True
        ElseIf StartParam = "Edit" Then    '---�������������� ������ � ��������
            Declarations.MyRecordID = Trim(MyShipmentCost.DataGridView1.SelectedRows.Item(0).Cells(0).Value.ToString())
            MySQLStr = "SELECT ID, PL01001, PL03002, DocDate, DocSumm, ShipmentsType "
            MySQLStr = MySQLStr & "FROM tbl_ShipmentsCost_Fact WITH (NOLOCK) "
            MySQLStr = MySQLStr & "WHERE (ID = '" & Declarations.MyRecordID & "') "
            InitMyConn(False)
            InitMyRec(False, MySQLStr)
            If Declarations.MyRec.EOF = True And Declarations.MyRec.BOF = True Then
                trycloseMyRec()
                MsgBox("������ ������ �� ����������, ��������, ������� ������ �������������. �������� ������ � ���� �� ������� ��������.", MsgBoxStyle.Critical, "��������!")
                Me.Close()
            Else
                Declarations.MyRec.MoveFirst()
                TextBox3.Text = Declarations.MyRec.Fields("PL01001").Value
                TextBox4.Text = Declarations.MyRec.Fields("PL03002").Value
                DateTimePicker1.Value = Declarations.MyRec.Fields("DocDate").Value
                TextBox1.Text = Declarations.MyRec.Fields("DocSumm").Value.ToString
                If Now < CDate("01/01/2019") Then
                    TextBox5.Text = CStr(Math.Round(Declarations.MyRec.Fields("DocSumm").Value * 1.18, 2))
                Else
                    TextBox5.Text = CStr(Math.Round(Declarations.MyRec.Fields("DocSumm").Value * 1.2, 2))
                End If
                ComboBox2.SelectedValue = Declarations.MyRec.Fields("ShipmentsType").Value
                trycloseMyRec()
            End If

            TextBox3.Enabled = True
            TextBox4.Enabled = True
            TextBox5.Enabled = True
            DateTimePicker1.Enabled = True
            TextBox1.Enabled = True
            ComboBox2.Enabled = True
        Else                               '---�������� ������
            Declarations.MyRecordID = Trim(MyShipmentCost.DataGridView1.SelectedRows.Item(0).Cells(0).Value.ToString())
            MySQLStr = "SELECT ID, PL01001, PL03002, DocDate, DocSumm, ShipmentsType "
            MySQLStr = MySQLStr & "FROM tbl_ShipmentsCost_Fact WITH (NOLOCK) "
            MySQLStr = MySQLStr & "WHERE (ID = '" & Declarations.MyRecordID & "') "
            InitMyConn(False)
            InitMyRec(False, MySQLStr)
            If Declarations.MyRec.EOF = True And Declarations.MyRec.BOF = True Then
                trycloseMyRec()
                MsgBox("������ ������ �� ����������, ��������, ������� ������ �������������. �������� ������ � ���� �� ������� ��������.", MsgBoxStyle.Critical, "��������!")
                Me.Close()
            Else
                Declarations.MyRec.MoveFirst()
                TextBox3.Text = Declarations.MyRec.Fields("PL01001").Value
                TextBox4.Text = Declarations.MyRec.Fields("PL03002").Value
                DateTimePicker1.Value = Declarations.MyRec.Fields("DocDate").Value
                TextBox1.Text = Declarations.MyRec.Fields("DocSumm").Value.ToString
                If Now < CDate("01/01/2019") Then
                    TextBox5.Text = CStr(Math.Round(Declarations.MyRec.Fields("DocSumm").Value * 1.18, 2))
                Else
                    TextBox5.Text = CStr(Math.Round(Declarations.MyRec.Fields("DocSumm").Value * 1.2, 2))
                End If
                ComboBox2.SelectedValue = Declarations.MyRec.Fields("ShipmentsType").Value
                trycloseMyRec()
            End If
            TextBox3.Enabled = False
            TextBox4.Enabled = False
            TextBox5.Enabled = False
            DateTimePicker1.Enabled = False
            TextBox1.Enabled = False
            ComboBox2.Enabled = False
        End If

        LoadInvoicesInfo()

        CheckButtonState()
    End Sub

    Private Sub CheckButtonState()
        '/////////////////////////////////////////////////////////////////////////////////////
        '//
        '// �������� � ����������� ��������� ������
        '// 
        '/////////////////////////////////////////////////////////////////////////////////////

        If StartParam = "View" Then
            Button5.Enabled = False
            Button2.Enabled = False
            Button1.Enabled = False
            Button3.Enabled = False
            Button4.Enabled = True
            Button6.Enabled = False
            Button7.Enabled = False
        Else
            If DataGridView1.SelectedRows.Count = 0 Then
                Button1.Enabled = False
            Else
                Button1.Enabled = True
            End If
            Button5.Enabled = True
            Button2.Enabled = True
            Button3.Enabled = True
            Button4.Enabled = True
            Button6.Enabled = True
            Button7.Enabled = True
        End If
    End Sub

    Private Sub LoadInvoicesInfo()
        '/////////////////////////////////////////////////////////////////////////////////////
        '//
        '// �������� ���������� �� ��������, �������� � ������
        '// 
        '/////////////////////////////////////////////////////////////////////////////////////
        Dim MySQLStr As String
        Dim MyAdapter4 As SqlClient.SqlDataAdapter    '��� ������ �����������
        Dim MyDs4 As New DataSet

        'MySQLStr = "SELECT ID, DocID, SL03002, InvoiceSumm, ShipmentCost "
        'MySQLStr = MySQLStr & "FROM tbl_ShipmentsCost_FactByInvoices WITH (NOLOCK) "
        'MySQLStr = MySQLStr & "WHERE (DocID = N'" & Declarations.MyRecordID & "')"
        'MySQLStr = MySQLStr & "ORDER BY SL03002 "
        MySQLStr = "SELECT tbl_ShipmentsCost_FactByInvoices.ID, tbl_ShipmentsCost_FactByInvoices.DocID, tbl_ShipmentsCost_FactByInvoices.SL03002, "
        MySQLStr = MySQLStr & "tbl_ShipmentsCost_FactByInvoices.InvoiceSumm, tbl_ShipmentsCost_FactByInvoices.ShipmentCost, tbl_ShipmentsCost_DocType.Name, "
        MySQLStr = MySQLStr & "ISNULL(tbl_ShipmentsCost_FactByInvoices.SupplierCode, '') AS SupplierCode, ISNULL(tbl_ShipmentsCost_FactByInvoices.DocYear, '') "
        MySQLStr = MySQLStr & "AS DocYear "
        MySQLStr = MySQLStr & "FROM tbl_ShipmentsCost_FactByInvoices WITH (NOLOCK) INNER JOIN "
        MySQLStr = MySQLStr & "tbl_ShipmentsCost_DocType WITH (NOLOCK) ON tbl_ShipmentsCost_FactByInvoices.DocType = tbl_ShipmentsCost_DocType.ID "
        MySQLStr = MySQLStr & "WHERE (tbl_ShipmentsCost_FactByInvoices.DocID = '" & Declarations.MyRecordID & "') "
        MySQLStr = MySQLStr & "ORDER BY tbl_ShipmentsCost_FactByInvoices.SL03002 "
        Try
            MyAdapter4 = New SqlClient.SqlDataAdapter(MySQLStr, Declarations.MyNETConnStr)
            MyAdapter4.SelectCommand.CommandTimeout = 600
            MyAdapter4.Fill(MyDs4)
            DataGridView1.DataSource = MyDs4.Tables(0)
        Catch ex As Exception
            MsgBox(ex.ToString)
        End Try

        DataGridView1.Columns(0).HeaderText = "ID"
        DataGridView1.Columns(0).Width = 0
        DataGridView1.Columns(0).Visible = False
        DataGridView1.Columns(1).HeaderText = "DocID"
        DataGridView1.Columns(1).Width = 0
        DataGridView1.Columns(1).Visible = False
        DataGridView1.Columns(2).HeaderText = "N ���������"
        DataGridView1.Columns(2).Width = 130
        DataGridView1.Columns(3).HeaderText = "����� ���-�� (���)"
        DataGridView1.Columns(3).Width = 120
        DataGridView1.Columns(4).HeaderText = "����� �������� ���-�� (���)"
        DataGridView1.Columns(4).Width = 120
        DataGridView1.Columns(5).HeaderText = "��������"
        DataGridView1.Columns(5).Width = 120
        DataGridView1.Columns(6).HeaderText = "��� ����������"
        DataGridView1.Columns(6).Width = 80
        DataGridView1.Columns(7).HeaderText = "��� ���������"
        DataGridView1.Columns(7).Width = 80

        MySQLStr = "SELECT SUM(InvoiceSumm) AS TotalSum "
        MySQLStr = MySQLStr & "FROM tbl_ShipmentsCost_FactByInvoices WITH (NOLOCK) "
        MySQLStr = MySQLStr & "WHERE (DocID = '" & Declarations.MyRecordID & "')"
        InitMyConn(False)
        InitMyRec(False, MySQLStr)
        If Declarations.MyRec.EOF = True And Declarations.MyRec.BOF = True Then
            TextBox2.Text = 0
        Else
            Declarations.MyRec.MoveFirst()
            If IsDBNull(Declarations.MyRec.Fields("TotalSum").Value) = True Then
                TextBox2.Text = 0
            Else
                TextBox2.Text = Declarations.MyRec.Fields("TotalSum").Value
            End If
        End If
        trycloseMyRec()
    End Sub

    Private Sub Button5_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button5.Click
        '/////////////////////////////////////////////////////////////////////////////////////
        '//
        '// �������� ���� ������ ����������
        '// 
        '/////////////////////////////////////////////////////////////////////////////////////

        MySupplierSelect = New SupplierSelect
        MySupplierSelect.MySrcWin = "AddShipmentCost"
        MySupplierSelect.ShowDialog()
    End Sub

    Private Sub TextBox3_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles TextBox3.KeyDown
        '////////////////////////////////////////////////////////////////////////////////
        '//
        '// ������� �� ��������� ���� �� Enter
        '//
        '////////////////////////////////////////////////////////////////////////////////

        If e.KeyData = Keys.Enter Then
            Me.SelectNextControl(sender, True, True, True, False)
        End If
    End Sub

    Private Sub TextBox3_Validating(ByVal sender As Object, ByVal e As System.ComponentModel.CancelEventArgs) Handles TextBox3.Validating
        '////////////////////////////////////////////////////////////////////////////////////////
        '//
        '// ������ ��� ���������� - ���������
        '//
        '////////////////////////////////////////////////////////////////////////////////////////
        Dim MySQLStr As String

        If TextBox3.Modified = True Then
            MySQLStr = "SELECT PL01002, PL01003 + ' ' + PL01004 + ' ' + PL01005 AS PL01003 "
            MySQLStr = MySQLStr & "FROM PL010300 WITH (NOLOCK) "
            MySQLStr = MySQLStr & "WHERE (PL01001 = N'" & Trim(TextBox3.Text) & "') "
            InitMyConn(False)
            InitMyRec(False, MySQLStr)
            If Declarations.MyRec.EOF = True And Declarations.MyRec.BOF = True Then
                trycloseMyRec()
                MsgBox("�� ����� �������� ��� ����������. ������� ���������� ��� �������������� �������.", vbCritical, "��������!")
                e.Cancel = True
                Exit Sub
            Else
                trycloseMyRec()
            End If
        End If
    End Sub

    Private Sub TextBox4_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles TextBox4.KeyDown
        '////////////////////////////////////////////////////////////////////////////////
        '//
        '// ������� �� ��������� ���� �� Enter
        '//
        '////////////////////////////////////////////////////////////////////////////////

        If e.KeyData = Keys.Enter Then
            Me.SelectNextControl(sender, True, True, True, False)
        End If
    End Sub

    Private Sub DateTimePicker1_CloseUp(ByVal sender As Object, ByVal e As System.EventArgs) Handles DateTimePicker1.CloseUp
        '////////////////////////////////////////////////////////////////////////////////
        '//
        '// ������� �� ��������� ���� �� ������
        '//
        '////////////////////////////////////////////////////////////////////////////////

        Me.SelectNextControl(sender, True, True, True, False)
    End Sub

    Private Sub DateTimePicker1_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles DateTimePicker1.KeyDown
        '////////////////////////////////////////////////////////////////////////////////
        '//
        '// ������� �� ��������� ���� �� Enter
        '//
        '////////////////////////////////////////////////////////////////////////////////

        If e.KeyData = Keys.Enter Then
            Me.SelectNextControl(sender, True, True, True, False)
        End If
    End Sub

    Private Sub TextBox1_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles TextBox1.KeyDown
        '////////////////////////////////////////////////////////////////////////////////
        '//
        '// ������� �� ��������� ���� �� Enter
        '//
        '////////////////////////////////////////////////////////////////////////////////

        If e.KeyData = Keys.Enter Then
            Me.SelectNextControl(sender, True, True, True, False)
        End If
    End Sub

    Private Sub TextBox1_Validated(ByVal sender As Object, ByVal e As System.EventArgs) Handles TextBox1.Validated
        '////////////////////////////////////////////////////////////////////////////////
        '//
        '// ���������� ���������� �������� �������� ����� �������������� ����� - ����� ���������
        '//
        '////////////////////////////////////////////////////////////////////////////////

        If Now < CDate("01/01/2019") Then
            TextBox5.Text = CStr(Math.Round(CDbl(Trim(TextBox1.Text)) * 1.18, 2))
        Else
            TextBox5.Text = CStr(Math.Round(CDbl(Trim(TextBox1.Text)) * 12, 2))
        End If
        TextBox1_ProcValidating()
    End Sub

    Private Sub TextBox1_ProcValidating()
        '////////////////////////////////////////////////////////////////////////////////
        '//
        '// ���������� ���������� �������� �������� ����� �������������� �����
        '//
        '////////////////////////////////////////////////////////////////////////////////
        Dim MySQLStr As String

        If Trim(TextBox1.Text) <> "" Then
            MySQLStr = "UPDATE tbl_ShipmentsCost_Fact "
            MySQLStr = MySQLStr & "SET DocSumm = " & Replace(Trim(TextBox1.Text), ",", ".")
            MySQLStr = MySQLStr & "WHERE (ID = '" & Declarations.MyRecordID & "') "
            InitMyConn(False)
            Declarations.MyConn.Execute(MySQLStr)


            ReCalculateShipment()
            LoadInvoicesInfo()
            CheckButtonState()
        End If
    End Sub

    Private Sub TextBox1_Validating(ByVal sender As Object, ByVal e As System.ComponentModel.CancelEventArgs) Handles TextBox1.Validating
        '////////////////////////////////////////////////////////////////////////////////
        '//
        '// �������� - �������� �� �������� �������
        '//
        '////////////////////////////////////////////////////////////////////////////////
        Dim MyRez As Double
        Dim aa As New System.Globalization.NumberFormatInfo

        If Trim(TextBox1.Text) <> "" Then
            If InStr(TextBox1.Text, aa.CurrentInfo.NumberGroupSeparator) Then
                MsgBox("� ���� ""����� �������� (���)"" ������ ���� ������� �����. ����������� ������� � ����� ����� - """ & aa.CurrentInfo.NumberDecimalSeparator & """ ", MsgBoxStyle.Critical, "��������!")
                e.Cancel = True
                Exit Sub
            Else
                Try
                    MyRez = TextBox1.Text
                Catch ex As Exception
                    MsgBox("� ���� ""����� �������� (���)"" ������ ���� ������� �����. ����������� ������� � ����� ����� - """ & aa.CurrentInfo.NumberDecimalSeparator & """ ", MsgBoxStyle.Critical, "��������!")
                    e.Cancel = True
                    Exit Sub
                End Try
            End If
        End If
        e.Cancel = False
    End Sub

    Private Sub ComboBox2_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles ComboBox2.KeyDown
        '////////////////////////////////////////////////////////////////////////////////
        '//
        '// ������� �� ��������� ���� �� Enter
        '//
        '////////////////////////////////////////////////////////////////////////////////

        If e.KeyData = Keys.Enter Then
            Me.SelectNextControl(sender, True, True, True, False)
        End If
    End Sub

    Private Sub ComboBox2_SelectedValueChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles ComboBox2.SelectedValueChanged
        '////////////////////////////////////////////////////////////////////////////////
        '//
        '// ������� �� ��������� ���� �� ������ ��������
        '//
        '////////////////////////////////////////////////////////////////////////////////

        Me.SelectNextControl(sender, True, True, True, False)
    End Sub

    Private Sub Button3_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button3.Click
        '////////////////////////////////////////////////////////////////////////////////
        '//
        '// ������� ������ ���������� ������
        '//
        '////////////////////////////////////////////////////////////////////////////////

        SaveData()
    End Sub

    Private Function SaveData()
        '////////////////////////////////////////////////////////////////////////////////
        '//
        '// ���������� ������
        '//
        '////////////////////////////////////////////////////////////////////////////////
        Dim MySQLStr As String

        If CheckDataFiling(True) = True Then
            '---�������� - ���� ��� �� ������ ������� - ������� ������ � ���������
            MySQLStr = "SELECT COUNT(ID) AS CC "
            MySQLStr = MySQLStr & "FROM tbl_ShipmentsCost_FactByInvoices WITH (NOLOCK)"
            MySQLStr = MySQLStr & "WHERE (DocID = N'" & Declarations.MyRecordID & "') "
            InitMyConn(False)
            InitMyRec(False, MySQLStr)
            If Declarations.MyRec.Fields("CC").Value = 0 Then
                trycloseMyRec()
                MySQLStr = "DELETE FROM tbl_ShipmentsCost_Fact "
                MySQLStr = MySQLStr & "WHERE (ID = N'" & Declarations.MyRecordID & "')"
                InitMyConn(False)
                Declarations.MyConn.Execute(MySQLStr)
            Else
                trycloseMyRec()
                If StartParam = "Create" Then
                    SaveNewData()
                ElseIf StartParam = "Edit" Then
                    UpdateData()
                End If
            End If
            Me.Close()
        End If
    End Function

    Private Function CheckDataFiling(ByVal AddCheck As Boolean) As Boolean
        '////////////////////////////////////////////////////////////////////////////////
        '//
        '// �������� ���������� ����� � ����
        '//
        '////////////////////////////////////////////////////////////////////////////////
        Dim MySQLStr As String

        If Trim(TextBox3.Text) = "" Then
            MsgBox("���� ""����������"" ������ ���� ���������", MsgBoxStyle.Critical, "��������")
            CheckDataFiling = False
            TextBox3.Select()
            Exit Function
        End If

        If Trim(TextBox4.Text) = "" Then
            MsgBox("���������� ������ ����� ���������.", MsgBoxStyle.Critical, "��������")
            CheckDataFiling = False
            TextBox4.Select()
            Exit Function
        End If

        If Trim(TextBox1.Text) = "" Then
            MsgBox("���������� ������ ����� �������� (���).", MsgBoxStyle.Critical, "��������")
            CheckDataFiling = False
            TextBox1.Select()
            Exit Function
        End If

        If ComboBox2.SelectedValue = 0 Then
            MsgBox("���������� ������� ������ ��������", MsgBoxStyle.Critical, "��������")
            CheckDataFiling = False
            ComboBox2.Select()
            Exit Function
        End If

        '---�������� - ����� ����� �������� �� ������ ���������� ��� ���� �� ����� ���� (� ����� ����)
        MySQLStr = "SELECT COUNT(ID) AS CC "
        MySQLStr = MySQLStr & "FROM tbl_ShipmentsCost_Fact WITH (NOLOCK) "
        MySQLStr = MySQLStr & "WHERE (ID <> '" & Declarations.MyRecordID & "') AND "
        MySQLStr = MySQLStr & "(PL01001 = N'" & TextBox3.Text & "') AND "
        MySQLStr = MySQLStr & "(PL03002 = N'" & TextBox4.Text & "') AND "
        MySQLStr = MySQLStr & "(DATEPART(yyyy, DocDate) = " & CStr(DatePart(DateInterval.Year, DateTimePicker1.Value)) & ") "
        InitMyConn(False)
        InitMyRec(False, MySQLStr)
        If Declarations.MyRec.Fields("CC").Value = 0 Then '---��� OK
            trycloseMyRec()
            CheckDataFiling = True
        Else
            trycloseMyRec()
            MsgBox("�������� ����� " & TextBox4.Text & " ��� ���������� " & TextBox3.Text & " ��� ������� � ������� � " & CStr(DatePart(DateInterval.Year, DateTimePicker1.Value)) & " ����.", MsgBoxStyle.Critical, "��������!")
            CheckDataFiling = False
            Exit Function
        End If

        CheckDataFiling = True

    End Function

    Private Sub SaveNewData()
        '////////////////////////////////////////////////////////////////////////////////
        '//
        '// ���������� ������ � ������ �������� ����� ������
        '//
        '////////////////////////////////////////////////////////////////////////////////
        Dim MySQLStr As String

        MySQLStr = "INSERT INTO tbl_ShipmentsCost_Fact "
        MySQLStr = MySQLStr & "(ID, PL01001, PL03002, DocDate, DocSumm, ShipmentsType) "
        MySQLStr = MySQLStr & "VALUES ('" & Declarations.MyRecordID & "', "
        MySQLStr = MySQLStr & "N'" & TextBox3.Text & "', "
        MySQLStr = MySQLStr & "N'" & TextBox4.Text & "', "
        MySQLStr = MySQLStr & "CONVERT(DATETIME, '" & DatePart(DateInterval.Day, DateTimePicker1.Value) & "/" & DatePart(DateInterval.Month, DateTimePicker1.Value) & "/" & DatePart(DateInterval.Year, DateTimePicker1.Value) & "', 103), "
        MySQLStr = MySQLStr & Replace(TextBox1.Text, ",", ".") & ", "
        MySQLStr = MySQLStr & CStr(ComboBox2.SelectedValue) & ") "
        InitMyConn(False)
        Declarations.MyConn.Execute(MySQLStr)
    End Sub

    Private Sub UpdateData()
        '////////////////////////////////////////////////////////////////////////////////
        '//
        '// ���������� ������ � ������ �������������� ������
        '//
        '////////////////////////////////////////////////////////////////////////////////
        Dim MySQLStr As String

        MySQLStr = "UPDATE tbl_ShipmentsCost_Fact "
        MySQLStr = MySQLStr & "SET PL01001 = N'" & TextBox3.Text & "', "
        MySQLStr = MySQLStr & "PL03002 = N'" & TextBox4.Text & "', "
        MySQLStr = MySQLStr & "DocDate = CONVERT(DATETIME, '" & DatePart(DateInterval.Day, DateTimePicker1.Value) & "/" & DatePart(DateInterval.Month, DateTimePicker1.Value) & "/" & DatePart(DateInterval.Year, DateTimePicker1.Value) & "', 103), "
        MySQLStr = MySQLStr & "DocSumm = " & Replace(TextBox1.Text, ",", ".") & ", "
        MySQLStr = MySQLStr & "ShipmentsType = " & CStr(ComboBox2.SelectedValue) & " "
        MySQLStr = MySQLStr & "WHERE (ID = '" & Declarations.MyRecordID & "')"
        InitMyConn(False)
        Declarations.MyConn.Execute(MySQLStr)
    End Sub

    Private Sub Button2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button2.Click
        '////////////////////////////////////////////////////////////////////////////////
        '//
        '// ������� ������ ���������� ������� �� ������� � ������ ������������
        '//
        '////////////////////////////////////////////////////////////////////////////////

        If StartParam = "Create" Then
            '---���� ������ ������ ��������� - �� ���� ������� ��������� ��������� � ���������
            If CheckDataFiling(False) = True Then
                SaveNewData()
                StartParam = "Edit"
                AddSalesInvoice()
            End If
        Else
            If CheckDataFiling(False) = True Then
                AddSalesInvoice()
            End If
        End If
        LoadInvoicesInfo()
        CheckButtonState()
    End Sub

    Private Sub AddSalesInvoice()
        '////////////////////////////////////////////////////////////////////////////////
        '//
        '// ���������� ������� �� ������� � ������ ������������
        '//
        '////////////////////////////////////////////////////////////////////////////////

        MyAddSInvoice = New AddSInvoice
        MyAddSInvoice.ShowDialog()
        ReCalculateShipment()
    End Sub

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        '////////////////////////////////////////////////////////////////////////////////
        '//
        '// �������� ������� �� ������� �� ������ ������������
        '//
        '////////////////////////////////////////////////////////////////////////////////
        Dim MySQLStr As String

        MySQLStr = "DELETE FROM tbl_ShipmentsCost_FactByInvoices "
        MySQLStr = MySQLStr & "WHERE (ID = '" & Trim(DataGridView1.SelectedRows.Item(0).Cells(0).Value.ToString()) & "')"
        InitMyConn(False)
        Declarations.MyConn.Execute(MySQLStr)
        ReCalculateShipment()
        LoadInvoicesInfo()
        CheckButtonState()
    End Sub

    Private Sub ReCalculateShipment()
        '////////////////////////////////////////////////////////////////////////////////
        '//
        '// ���������� ���������� �������� ��������, ���������� � ���� ��������
        '//
        '////////////////////////////////////////////////////////////////////////////////
        Dim MySQLStr As String
        Dim InvoicesSum As Double                   '����� ���� �������� � ��������
        Dim KProp As Double                         '�����������

        MySQLStr = "SELECT SUM(InvoiceSumm) AS TotalSum "
        MySQLStr = MySQLStr & "FROM tbl_ShipmentsCost_FactByInvoices WITH (NOLOCK) "
        MySQLStr = MySQLStr & "WHERE (DocID = '" & Declarations.MyRecordID & "')"
        InitMyConn(False)
        InitMyRec(False, MySQLStr)
        If Declarations.MyRec.EOF = True And Declarations.MyRec.BOF = True Then
            InvoicesSum = 0
        Else
            Declarations.MyRec.MoveFirst()
            If IsDBNull(Declarations.MyRec.Fields("TotalSum").Value) = True Then
                InvoicesSum = 0
            Else
                InvoicesSum = Declarations.MyRec.Fields("TotalSum").Value
            End If
        End If
        trycloseMyRec()
        TextBox2.Text = InvoicesSum

        If InvoicesSum = 0 Then
            KProp = 0
        Else
            KProp = CDbl(TextBox1.Text) / InvoicesSum
        End If

        MySQLStr = "UPDATE tbl_ShipmentsCost_FactByInvoices "
        MySQLStr = MySQLStr & "SET ShipmentCost = Round(InvoiceSumm * " & Replace(CStr(KProp), ",", ".") & ",3) "
        MySQLStr = MySQLStr & "WHERE (DocID = '" & Declarations.MyRecordID & "') "
        InitMyConn(False)
        Declarations.MyConn.Execute(MySQLStr)
    End Sub

    Private Sub Button6_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button6.Click
        '////////////////////////////////////////////////////////////////////////////////
        '//
        '// ������� ������ ���������� ������� �� ������� � ������ ������������
        '//
        '////////////////////////////////////////////////////////////////////////////////

        If StartParam = "Create" Then
            '---���� ������ ������ ��������� - �� ���� ������� ��������� ��������� � ���������
            If CheckDataFiling(False) = True Then
                SaveNewData()
                StartParam = "Edit"
                AddPurchaseInvoice()
            End If
        Else
            If CheckDataFiling(False) = True Then
                AddPurchaseInvoice()
            End If
        End If
        LoadInvoicesInfo()
        CheckButtonState()
    End Sub

    Private Sub Button7_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button7.Click
        '////////////////////////////////////////////////////////////////////////////////
        '//
        '// ������� ������ ���������� ������ �� ����������� � ������ ������������
        '//
        '////////////////////////////////////////////////////////////////////////////////

        If StartParam = "Create" Then
            '---���� ������ ������ ��������� - �� ���� ������� ��������� ��������� � ���������
            If CheckDataFiling(False) = True Then
                SaveNewData()
                StartParam = "Edit"
                AddRelocationOrder()
            End If
        Else
            If CheckDataFiling(False) = True Then
                AddRelocationOrder()
            End If
        End If
        LoadInvoicesInfo()
        CheckButtonState()
    End Sub

    Private Sub AddPurchaseInvoice()
        '////////////////////////////////////////////////////////////////////////////////
        '//
        '// ���������� ������� �� ������� � ������ ������������
        '//
        '////////////////////////////////////////////////////////////////////////////////

        MyAddPInvoice = New AddPInvoice
        MyAddPInvoice.ShowDialog()
        ReCalculateShipment()
    End Sub

    Private Sub AddRelocationOrder()
        '////////////////////////////////////////////////////////////////////////////////
        '//
        '// ���������� ������ �� ����������� � ������ ������������
        '//
        '////////////////////////////////////////////////////////////////////////////////

        MyAddRelOrder = New AddRelOrder
        MyAddRelOrder.ShowDialog()
        ReCalculateShipment()
    End Sub

    Private Sub TextBox5_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles TextBox5.KeyDown
        '////////////////////////////////////////////////////////////////////////////////
        '//
        '// ������� �� ��������� ���� �� Enter
        '//
        '////////////////////////////////////////////////////////////////////////////////

        If e.KeyData = Keys.Enter Then
            Me.SelectNextControl(sender, True, True, True, False)
        End If
    End Sub

    Private Sub TextBox5_Validated(ByVal sender As Object, ByVal e As System.EventArgs) Handles TextBox5.Validated
        '////////////////////////////////////////////////////////////////////////////////
        '//
        '// ����������� ����� ��� ��� ����� ��������
        '//
        '////////////////////////////////////////////////////////////////////////////////

        If Now < CDate("01/01/2019") Then
            TextBox1.Text = CStr(Math.Round(CDbl(Trim(TextBox5.Text)) / 1.18, 2))
        Else
            TextBox1.Text = CStr(Math.Round(CDbl(Trim(TextBox5.Text)) / 1.2, 2))
        End If
        TextBox1_ProcValidating()
    End Sub

    Private Sub TextBox5_Validating(ByVal sender As Object, ByVal e As System.ComponentModel.CancelEventArgs) Handles TextBox5.Validating
        '////////////////////////////////////////////////////////////////////////////////
        '//
        '// �������� - �������� �� �������� �������
        '//
        '////////////////////////////////////////////////////////////////////////////////
        Dim MyRez As Double
        Dim aa As New System.Globalization.NumberFormatInfo

        If Trim(TextBox5.Text) <> "" Then
            If InStr(TextBox5.Text, aa.CurrentInfo.NumberGroupSeparator) Then
                MsgBox("� ���� ""� ���"" ������ ���� ������� �����. ����������� ������� � ����� ����� - """ & aa.CurrentInfo.NumberDecimalSeparator & """ ", MsgBoxStyle.Critical, "��������!")
                e.Cancel = True
                Exit Sub
            Else
                Try
                    MyRez = TextBox5.Text
                Catch ex As Exception
                    MsgBox("� ���� ""� ���"" ������ ���� ������� �����. ����������� ������� � ����� ����� - """ & aa.CurrentInfo.NumberDecimalSeparator & """ ", MsgBoxStyle.Critical, "��������!")
                    e.Cancel = True
                    Exit Sub
                End Try
            End If
        End If
        e.Cancel = False
    End Sub
End Class