Public Class MainForm
    Public LoadFlag As Integer

    Private Sub Button13_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button13.Click
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

        ComboBoxAN.SelectedIndex = 0

        LoadFlag = 0
        Application.DoEvents()
        Windows.Forms.Cursor.Current = Cursors.WaitCursor
        '---�������� ������
        LoadShipments()
        LoadFreeOrders()
        '---�������� ��������� ������
        CheckSHButtonsState()
        CheckOrderButtonsState()
        Windows.Forms.Cursor.Current = Cursors.Default
    End Sub

    Public Sub LoadShipments()
        '/////////////////////////////////////////////////////////////////////////////////////
        '//
        '// �������� ������ �������� / ��������
        '// 
        '/////////////////////////////////////////////////////////////////////////////////////
        Dim MySQLStr As String
        Dim MyAdapter As SqlClient.SqlDataAdapter     '��� ������ �����������
        Dim MyDs As New DataSet                       '

        If LoadFlag = 0 Then
            If ComboBoxAN.Text = "��� �������� / ����������" Then
                MySQLStr = "EXEC spp_WEBShipments_ShipmentInfo N'" & Trim(ComboBox1.SelectedValue) & "', 0"
            Else
                MySQLStr = "EXEC spp_WEBShipments_ShipmentInfo N'" & Trim(ComboBox1.SelectedValue) & "', 1"
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
            DataGridView1.Columns(0).HeaderText = "N ���� ����"
            DataGridView1.Columns(0).Width = 50
            DataGridView1.Columns(1).HeaderText = "�������� �� ��������"
            DataGridView1.Columns(1).Width = 80
            DataGridView1.Columns(2).HeaderText = "N ������"
            DataGridView1.Columns(2).Width = 70
            DataGridView1.Columns(3).HeaderText = "N ������ � WEB"
            DataGridView1.Columns(3).Width = 70
            DataGridView1.Columns(4).HeaderText = "��� ���� ������"
            DataGridView1.Columns(4).Width = 70
            DataGridView1.Columns(5).HeaderText = "����������"
            DataGridView1.Columns(5).Width = 140
            DataGridView1.Columns(6).HeaderText = "��������"
            DataGridView1.Columns(6).Width = 100
            DataGridView1.Columns(7).HeaderText = "����� �� ��������"
            DataGridView1.Columns(7).Width = 80
            DataGridView1.Columns(7).DefaultCellStyle.Format = "n2"
            DataGridView1.Columns(8).HeaderText = "����� ������������"
            DataGridView1.Columns(8).Width = 100
            DataGridView1.Columns(8).DefaultCellStyle.Format = "n2"
            DataGridView1.Columns(9).HeaderText = "���������� ����������"
            DataGridView1.Columns(9).Width = 200
            DataGridView1.Columns(10).HeaderText = "����� ��������"
            DataGridView1.Columns(10).Width = 235
            DataGridView1.Columns(11).HeaderText = "�����������"
            DataGridView1.Columns(11).Width = 155
            DataGridView1.Columns(12).HeaderText = "������ �����"
            DataGridView1.Columns(12).Width = 60
            DataGridView1.Columns(13).HeaderText = "������ ������� - �����"
            DataGridView1.Columns(13).Width = 60
            DataGridView1.Columns(14).HeaderText = "������ ������� ����� (�����.)"
            DataGridView1.Columns(14).Width = 60
            DataGridView1.Columns(15).HeaderText = "����� ������ ���� ��������"
            DataGridView1.Columns(15).Width = 70
            DataGridView1.Columns(16).HeaderText = "������ �� ������"
            DataGridView1.Columns(16).Width = 60
            DataGridView1.Columns(17).HeaderText = "������ ����� �������"
            DataGridView1.Columns(17).Width = 60
            DataGridView1.Columns(18).HeaderText = "�������� ����� ������"
            DataGridView1.Columns(18).Width = 60
            DataGridView1.Columns(19).HeaderText = "����"
            DataGridView1.Columns(19).Width = 60
            DataGridView1.Columns(20).HeaderText = "���� � �����"
            DataGridView1.Columns(20).Visible = False

            FormatDataGridView1()
        End If
    End Sub

    Private Sub FormatDataGridView1()
        '////////////////////////////////////////////////////////////////////////////////
        '//
        '// ��������� ���������� �� ���������
        '//
        '////////////////////////////////////////////////////////////////////////////////
        Dim i As Integer

        For i = 0 To DataGridView1.Rows.Count - 1
            If DataGridView1.Rows(i).Cells(18).Value = 0 Then
                DataGridView1.Rows(i).DefaultCellStyle.BackColor = Color.White
                If DataGridView1.Rows(i).Cells(6).Value = "��������" Then
                    If DateDiff(DateInterval.Day, DataGridView1.Rows(i).Cells(15).Value, Now()) > 2 Then
                        DataGridView1.Rows(i).Cells(18).Style.BackColor = Color.LightPink
                    Else
                        DataGridView1.Rows(i).Cells(18).Style.BackColor = Color.White
                    End If

                Else            '---��������� ��� �������� � ������� ��������
                    If DateDiff(DateInterval.Day, DataGridView1.Rows(i).Cells(15).Value, Now()) > 7 Then
                        DataGridView1.Rows(i).Cells(18).Style.BackColor = Color.LightPink
                    Else
                        DataGridView1.Rows(i).Cells(18).Style.BackColor = Color.White
                    End If
                End If
            Else
                DataGridView1.Rows(i).DefaultCellStyle.BackColor = Color.LightGray
            End If
        Next
    End Sub

    Public Sub LoadFreeOrders()
        '/////////////////////////////////////////////////////////////////////////////////////
        '//
        '// �������� ������ ������� �� �������� / ��������
        '// 
        '/////////////////////////////////////////////////////////////////////////////////////
        Dim MySQLStr As String
        Dim MyAdapter As SqlClient.SqlDataAdapter     '��� ������ �����������
        Dim MyDs As New DataSet                       '

        If LoadFlag = 0 Then
            MySQLStr = "spp_WEBShipments_AvlOrders N'" & Trim(ComboBox1.SelectedValue) & "'"

            Try
                MyAdapter = New SqlClient.SqlDataAdapter(MySQLStr, Declarations.MyNETConnStr)
                MyAdapter.SelectCommand.CommandTimeout = 600
                MyAdapter.Fill(MyDs)
                DataGridView2.DataSource = MyDs.Tables(0)
            Catch ex As Exception
                MsgBox(ex.ToString)
            End Try

            '---���������
            DataGridView2.Columns(0).HeaderText = "N �-��"
            DataGridView2.Columns(0).Width = 90
            DataGridView2.Columns(1).HeaderText = "��� �-��"
            DataGridView2.Columns(1).Width = 60
            DataGridView2.Columns(2).HeaderText = "N �-�� WEB"
            DataGridView2.Columns(2).Width = 90
            DataGridView2.Columns(3).HeaderText = "������"
            DataGridView2.Columns(3).Width = 350
            DataGridView2.Columns(4).HeaderText = "��������"
            DataGridView2.Columns(4).Width = 200
            DataGridView2.Columns(5).HeaderText = "���� ��������"
            DataGridView2.Columns(5).DefaultCellStyle.Format = "dd/MM/yyyy"
            DataGridView2.Columns(5).Width = 100
            DataGridView2.Columns(6).HeaderText = "���� ���� ������� ������ �� �����"
            DataGridView2.Columns(6).Width = 100
            DataGridView2.Columns(6).DefaultCellStyle.Format = "dd/MM/yyyy"
            DataGridView2.Columns(7).HeaderText = "���� ������. ���� ������� ������ �� �����"
            DataGridView2.Columns(7).Width = 100
            DataGridView2.Columns(7).DefaultCellStyle.Format = "dd/MM/yyyy"
            DataGridView2.Columns(8).HeaderText = "����� �������� (�������)"
            DataGridView2.Columns(8).Width = 120
            DataGridView2.Columns(8).DefaultCellStyle.Format = "n2"
            DataGridView2.Columns(9).HeaderText = "����� ������"
            DataGridView2.Columns(9).Width = 120
            DataGridView2.Columns(9).DefaultCellStyle.Format = "n2"
            DataGridView2.Columns(10).HeaderText = "����� �������� �������"
            DataGridView2.Columns(10).Width = 120
            DataGridView2.Columns(10).DefaultCellStyle.Format = "n2"
            DataGridView2.Columns(11).HeaderText = "���������� �� ��������"
            DataGridView2.Columns(11).Width = 90

            FormatDataGridView2()
        End If
    End Sub

    Private Sub FormatDataGridView2()
        '////////////////////////////////////////////////////////////////////////////////
        '//
        '// ��������� ���������� �� ��������� �������
        '//
        '////////////////////////////////////////////////////////////////////////////////

        For Each row As DataGridViewRow In DataGridView2.Rows
            If Trim(row.Cells(11).Value.ToString) <> "" Then
                row.Cells(11).Style.BackColor = Color.LightGreen
            Else
                row.Cells(11).Style.BackColor = Color.LightPink
            End If
            If row.Cells(5).Value < Now Then
                row.Cells(5).Style.BackColor = Color.LightGreen
            Else
                row.Cells(5).Style.BackColor = Color.Empty
            End If
            If IsDBNull(row.Cells(6).Value) = False Then
                If row.Cells(5).Value < row.Cells(6).Value Then
                    row.Cells(6).Style.BackColor = Color.LightYellow
                Else
                    row.Cells(6).Style.BackColor = Color.Empty
                End If
            Else
                row.Cells(6).Style.BackColor = Color.Empty
            End If
            If IsDBNull(row.Cells(7).Value) = False Then
                If row.Cells(5).Value < row.Cells(7).Value Then
                    row.Cells(7).Style.BackColor = Color.LightYellow
                Else
                    row.Cells(7).Style.BackColor = Color.Empty
                End If
            Else
                row.Cells(7).Style.BackColor = Color.Empty
            End If
            If row.Cells(9).Value = 0 Then
                row.Cells(9).Style.BackColor = Color.LightPink
            Else
                row.Cells(9).Style.BackColor = Color.Empty
            End If
            If row.Cells(10).Value = 0 Then
                row.Cells(10).Style.BackColor = Color.LightPink
            Else
                If row.Cells(10).Value < row.Cells(9).Value Then
                    row.Cells(10).Style.BackColor = Color.LightYellow
                Else
                    row.Cells(10).Style.BackColor = Color.Empty
                End If
            End If
        Next
    End Sub

    Public Sub CheckSHButtonsState()
        '////////////////////////////////////////////////////////////////////////////////
        '//
        '// �������� � ����������� ��������� ������ �� ���������
        '//
        '////////////////////////////////////////////////////////////////////////////////

        If DataGridView1.SelectedRows.Count = 0 Then
            Button5.Enabled = False
            Button6.Enabled = False
        Else
            '--------------����������� �������
            If DataGridView1.SelectedRows.Item(0).Cells(17).Value = 0 Then  '---����������� ������� �� ����������
                If DataGridView1.SelectedRows.Item(0).Cells(18).Value = 0 Then '---�������� �� ������� (�� �����������)
                    Button5.Enabled = True
                Else
                    Button5.Enabled = False
                End If
            Else
                Button5.Enabled = False
            End If
            '--------------�������������� �������� ��������
            If DataGridView1.SelectedRows.Item(0).Cells(18).Value = 0 Then  '---�������� �� �����������
                Button6.Enabled = True
            Else
                Button6.Enabled = False
            End If
        End If
    End Sub

    Public Sub CheckOrderButtonsState()
        '////////////////////////////////////////////////////////////////////////////////
        '//
        '// �������� � ����������� ��������� ������ �� �������
        '//
        '////////////////////////////////////////////////////////////////////////////////

        If DataGridView2.SelectedRows.Count = 0 Then
            Button2.Enabled = False
            Button3.Enabled = False
            Button4.Enabled = False
        Else
            Button2.Enabled = True
            If DataGridView2.SelectedRows.Item(0).Cells(11).Value.ToString = "+" Then
                Button3.Enabled = False
            Else
                Button3.Enabled = True
            End If
            If (DataGridView2.SelectedRows.Item(0).Cells(11).Value.ToString = "+" _
                And DataGridView2.SelectedRows.Item(0).Cells(10).Value <> 0) Then
                Button4.Enabled = True
            Else
                Button4.Enabled = False
            End If
        End If

    End Sub

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        '////////////////////////////////////////////////////////////////////////////////
        '//
        '// ���������� ������
        '//
        '////////////////////////////////////////////////////////////////////////////////

        Windows.Forms.Cursor.Current = Cursors.WaitCursor
        '---�������� ������
        LoadShipments()
        LoadFreeOrders()
        '---�������� ��������� ������
        CheckSHButtonsState()
        CheckOrderButtonsState()
        Windows.Forms.Cursor.Current = Cursors.Default
    End Sub

    Private Sub ComboBox1_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles ComboBox1.SelectedIndexChanged
        '////////////////////////////////////////////////////////////////////////////////
        '//
        '// �������� ������ � ������������ � ��������� �������
        '//
        '////////////////////////////////////////////////////////////////////////////////

        If LoadFlag = 0 Then
            Windows.Forms.Cursor.Current = Cursors.WaitCursor
            LoadShipments()
            LoadFreeOrders()
            '---�������� ��������� ������
            CheckSHButtonsState()
            CheckOrderButtonsState()
            Windows.Forms.Cursor.Current = Cursors.Default
        End If
    End Sub

    Private Sub ComboBoxAN_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles ComboBoxAN.SelectedIndexChanged
        '////////////////////////////////////////////////////////////////////////////////
        '//
        '// �������� ������ � ������������ � ��������� ����������
        '//
        '////////////////////////////////////////////////////////////////////////////////

        If LoadFlag = 0 Then
            Windows.Forms.Cursor.Current = Cursors.WaitCursor
            LoadShipments()
            LoadFreeOrders()
            '---�������� ��������� ������
            CheckSHButtonsState()
            CheckOrderButtonsState()
            Windows.Forms.Cursor.Current = Cursors.Default
        End If
    End Sub

    Private Sub Button2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button2.Click
        '////////////////////////////////////////////////////////////////////////////////
        '//
        '// �������� ���� � ��������� ����������� �� ���������� ������ 
        '//
        '////////////////////////////////////////////////////////////////////////////////

        Declarations.MyOrderNum = Trim(DataGridView2.SelectedRows.Item(0).Cells(0).Value.ToString())
        Dim MyOrderDetails = New OrderDetails
        MyOrderDetails.ShowDialog()
    End Sub

    Private Sub Button3_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button3.Click
        '////////////////////////////////////////////////////////////////////////////////
        '//
        '// ������ ���������� �� �������� ��������� �������
        '//
        '////////////////////////////////////////////////////////////////////////////////

        If DataGridView2.SelectedRows.Item(0).Cells(11).Value.ToString = "" Then
            Windows.Forms.Cursor.Current = Cursors.WaitCursor
            ExecShippingAllovance(Trim(DataGridView2.SelectedRows.Item(0).Cells(0).Value.ToString))
            LoadFreeOrders()
            CheckOrderButtonsState()
            Windows.Forms.Cursor.Current = Cursors.Default
        End If

    End Sub

    Private Sub Button6_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button6.Click
        '////////////////////////////////////////////////////////////////////////////////
        '//
        '// �������� ��������, ���� ��� �� �� ����� �����������
        '//
        '////////////////////////////////////////////////////////////////////////////////
        Dim myrez As VariantType
        Dim MySQLStr As String

        myrez = MsgBox("�� �������, ��� ������ ������� ��������?" + Chr(13) + Chr(10) + "�������� ���� ��������� ������ � ��� ������, ���� ����� �� ����� ��������� ���������� ��������� �������� / ������.", MsgBoxStyle.YesNo, "��������!")
        If myrez = vbYes Then
            MySQLStr = "UPDATE tbl_Shipments_SalesmanWP_Details "
            MySQLStr = MySQLStr & "SET IsClosed = 1 "
            MySQLStr = MySQLStr & "WHERE (ShipmentsID = " & DataGridView1.SelectedRows.Item(0).Cells(0).Value.ToString & ")"
            InitMyConn(False)
            Declarations.MyConn.Execute(MySQLStr)

            Windows.Forms.Cursor.Current = Cursors.WaitCursor
            LoadShipments()
            LoadFreeOrders()
            '---�������� ��������� ������
            CheckSHButtonsState()
            CheckOrderButtonsState()
            Windows.Forms.Cursor.Current = Cursors.Default
        End If
    End Sub

    Private Sub Button5_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button5.Click
        '////////////////////////////////////////////////////////////////////////////////
        '//
        '// �������� ���� �������� ����������� �������� �� �������� / ���������� � ����������
        '//
        '////////////////////////////////////////////////////////////////////////////////

        MySendInfo = New SendInfo
        MySendInfo.ShowDialog()
        CheckSHButtonsState()
        CheckOrderButtonsState()
    End Sub

    Private Sub Button4_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button4.Click
        '////////////////////////////////////////////////////////////////////////////////
        '//
        '// �������� ���� �������� ��������
        '//
        '////////////////////////////////////////////////////////////////////////////////
        Dim MyMsg As String

        If (DataGridView2.SelectedRows.Item(0).Cells(11).Value.ToString = "+" _
                And DataGridView2.SelectedRows.Item(0).Cells(10).Value <> 0) Then
            Declarations.MyCustomerCode = GetFirstPartFromStr(DataGridView2.SelectedRows.Item(0).Cells(3).Value)
            Declarations.MyWH = Trim(ComboBox1.SelectedValue)
            Declarations.MyOrderNum = DataGridView2.SelectedRows.Item(0).Cells(0).Value
            Declarations.MyShipmentsID = 0
            MyOperationFlag = 0
            MyShipment = New Shipment
            MyShipment.ShowDialog()
            If MyOperationFlag <> 0 Then
                Application.DoEvents()
                Windows.Forms.Cursor.Current = Cursors.WaitCursor
                LoadShipments()
                LoadFreeOrders()
                CheckSHButtonsState()
                CheckOrderButtonsState()
                Windows.Forms.Cursor.Current = Cursors.Default
                '---���������� ������� ������������� ������
                For i As Integer = 0 To DataGridView1.Rows.Count - 1
                    If DataGridView1.Item(0, i).Value = Declarations.MyShipmentsID Then
                        DataGridView1.CurrentCell = DataGridView1.Item(0, i)
                        Exit For
                    End If
                Next
            End If
        Else
            If DataGridView2.SelectedRows.Item(0).Cells(11).Value.ToString <> "+" Then
                MyMsg = "����� �� ��� �������� � ��������. �������:" & Chr(13)
                MyMsg = MyMsg & "- ��� ���������� �� ��������" & Chr(13)
                MsgBox(MyMsg, MsgBoxStyle.Critical, "��������!")
            Else
                If DataGridView2.SelectedRows.Item(0).Cells(10).Value = 0 Then
                    MyMsg = "����� �� ��� �������� � ��������. �������:" & Chr(13)
                    MyMsg = MyMsg & "- � ������ ��� �������������� ���������" & Chr(13)
                    MsgBox(MyMsg, MsgBoxStyle.Critical, "��������!")
                End If
            End If
            
            End If
    End Sub

    Private Sub DataGridView2_ColumnHeaderMouseClick(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellMouseEventArgs) Handles DataGridView2.ColumnHeaderMouseClick
        '/////////////////////////////////////////////////////////////////////////////////////
        '//
        '// ������ �� ��������� ������� 
        '// 
        '/////////////////////////////////////////////////////////////////////////////////////

        FormatDataGridView2()
    End Sub

    Private Sub DataGridView2_SelectionChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles DataGridView2.SelectionChanged
        '////////////////////////////////////////////////////////////////////////////////
        '//
        '// ����� ��������� ������ � ��������� �������
        '//
        '////////////////////////////////////////////////////////////////////////////////

        CheckOrderButtonsState()
    End Sub

    Private Sub DataGridView1_ColumnHeaderMouseClick(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellMouseEventArgs) Handles DataGridView1.ColumnHeaderMouseClick
        '/////////////////////////////////////////////////////////////////////////////////////
        '//
        '// ������ �� ��������� ������� 
        '// 
        '/////////////////////////////////////////////////////////////////////////////////////

        FormatDataGridView1()
    End Sub

    Private Sub DataGridView1_SelectionChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles DataGridView1.SelectionChanged
        '////////////////////////////////////////////////////////////////////////////////
        '//
        '// ����� ��������� ������ � ���������
        '//
        '////////////////////////////////////////////////////////////////////////////////

        CheckSHButtonsState()
    End Sub
End Class
