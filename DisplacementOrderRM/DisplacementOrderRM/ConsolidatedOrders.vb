Public Class ConsolidatedOrders

    Private Sub Button13_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button13.Click
        '////////////////////////////////////////////////////////////////////////////////
        '//
        '// ����� �� ����
        '//
        '////////////////////////////////////////////////////////////////////////////////

        Me.Close()
    End Sub

    Private Sub ConsolidatedOrders_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles Me.KeyDown
        '////////////////////////////////////////////////////////////////////////////////
        '//
        '// ������ �������� ���� �� ALT+F4
        '//
        '////////////////////////////////////////////////////////////////////////////////

        If e.KeyData = Keys.Alt + Keys.F4 Then
            e.Handled = True
        End If
    End Sub

    Private Sub ConsolidatedOrders_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        '////////////////////////////////////////////////////////////////////////////////
        '//
        '// ��������� ���� ������������ ����������������� �������
        '//
        '////////////////////////////////////////////////////////////////////////////////

        Label2.Text = Declarations.WHToCode & " " & Trim(MainForm.DataGridView1.SelectedRows.Item(0).Cells(1).Value.ToString())
        Declarations.WHTo = Label2.Text
        Label4.Text = MainForm.ComboBox1.Text
        Declarations.WHFrom = Label4.Text
        ComboBox1.SelectedText = "������ �������� (����������)"

        LoadConsolidatedOrders()
        LoadIncludedOrders()
        LoadFreeOrders()
        CheckConsolidatedButtons()
        CheckRemoveButtons()
        CheckAddButtons()
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

        MySQLStr = "spp_DisplacementWorkplace_TotalGroupOrdersPrep N'" & Declarations.WHFromCode & "', N'" & Declarations.WHToCode & "', "
        If ComboBox1.Text = "������ �������� (����������)" Then
            MySQLStr = MySQLStr & "1 "
        Else
            MySQLStr = MySQLStr & "0 "
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
        DataGridView1.Columns(0).HeaderText = "N ������"
        DataGridView1.Columns(0).Width = 100
        DataGridView1.Columns(1).HeaderText = "���� ��������"
        DataGridView1.Columns(1).Width = 100
        DataGridView1.Columns(2).HeaderText = "����������� ��������"
        DataGridView1.Columns(2).Width = 100
        DataGridView1.Columns(3).HeaderText = "���� �������"
        DataGridView1.Columns(3).Width = 100
        DataGridView1.Columns(4).HeaderText = "����������� �������"
        DataGridView1.Columns(4).Width = 100
        DataGridView1.Columns(5).HeaderText = "���-�� ���������� �������"
        DataGridView1.Columns(5).Width = 100
        DataGridView1.Columns(6).HeaderText = "��������� �����"
        DataGridView1.Columns(6).Width = 200
        DataGridView1.Columns(7).HeaderText = "�� ������"
        DataGridView1.Columns(7).Width = 100
        DataGridView1.Columns(8).HeaderText = "N ���-�� �����������"
        DataGridView1.Columns(8).Width = 200
        DataGridView1.Columns(9).HeaderText = "�����������"
        DataGridView1.Columns(9).Width = 200
        DataGridView1.Columns(10).HeaderText = "������������"
        DataGridView1.Columns(10).Width = 0
        DataGridView1.Columns(10).Visible = False
        DataGridView1.Columns(11).HeaderText = "������������ ��� ����������"
        DataGridView1.Columns(11).Width = 0
        DataGridView1.Columns(11).Visible = False

    End Sub

    Private Sub LoadIncludedOrders()
        '////////////////////////////////////////////////////////////////////////////////
        '//
        '// �������� ���������� �� �������, ���������� � ��������� ����������
        '//
        '////////////////////////////////////////////////////////////////////////////////
        Dim MySQLStr As String
        Dim MyAdapter As SqlClient.SqlDataAdapter     '��� ������ �������
        Dim MyDs As New DataSet

        If DataGridView1.SelectedRows.Count <> 0 Then
            MySQLStr = "spp_DisplacementWorkplace_GroupOrdersPrep N'" & Trim(DataGridView1.SelectedRows.Item(0).Cells(0).Value.ToString()) & "'"
        Else
            MySQLStr = "spp_DisplacementWorkplace_GroupOrdersPrep N''"
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
        DataGridView2.Columns(0).HeaderText = "N ������"
        DataGridView2.Columns(0).Width = 75
        DataGridView2.Columns(1).HeaderText = "���� ��������"
        DataGridView2.Columns(1).Width = 75
        DataGridView2.Columns(2).HeaderText = "��������"
        DataGridView2.Columns(2).Width = 100
        DataGridView2.Columns(3).HeaderText = "���� �������"
        DataGridView2.Columns(3).Width = 150
        DataGridView2.Columns(4).HeaderText = "������"
        DataGridView2.Columns(4).Width = 75
        DataGridView2.Columns(5).HeaderText = "�������"
        DataGridView2.Columns(5).Width = 70
        DataGridView2.Columns(6).HeaderText = "����� �� �������"
        DataGridView2.Columns(6).Width = 100
        DataGridView2.Columns(7).HeaderText = "���������"
        DataGridView2.Columns(7).Width = 200


    End Sub

    Private Sub LoadFreeOrders()
        '////////////////////////////////////////////////////////////////////////////////
        '//
        '// �������� ���������� �� ��������� � ��������� � ���������� ����� �������
        '//
        '////////////////////////////////////////////////////////////////////////////////
        Dim MySQLStr As String
        Dim MyAdapter As SqlClient.SqlDataAdapter     '��� ������ �������
        Dim MyDs As New DataSet

        MySQLStr = "spp_DisplacementWorkplace_NonGroupOrdersPrep N'" & Declarations.WHFromCode & "', N'" & Declarations.WHToCode & "'"

        Try
            MyAdapter = New SqlClient.SqlDataAdapter(MySQLStr, Declarations.MyNETConnStr)
            MyAdapter.SelectCommand.CommandTimeout = 600
            MyAdapter.Fill(MyDs)
            DataGridView3.DataSource = MyDs.Tables(0)
        Catch ex As Exception
            MsgBox(ex.ToString)
        End Try

        '---���������
        DataGridView3.Columns(0).HeaderText = "N ������"
        DataGridView3.Columns(0).Width = 75
        DataGridView3.Columns(1).HeaderText = "���� ��������"
        DataGridView3.Columns(1).Width = 75
        DataGridView3.Columns(2).HeaderText = "��������"
        DataGridView3.Columns(2).Width = 100
        DataGridView3.Columns(3).HeaderText = "���� �������"
        DataGridView3.Columns(3).Width = 150
        DataGridView3.Columns(4).HeaderText = "����� �� �������"
        DataGridView3.Columns(4).Width = 100
        DataGridView3.Columns(5).HeaderText = "���������"
        DataGridView3.Columns(5).Width = 200
        DataGridView3.Columns(6).HeaderText = "������"
        DataGridView3.Columns(6).Width = 0
        DataGridView3.Columns(6).Visible = False

    End Sub

    Private Sub CheckConsolidatedButtons()
        '////////////////////////////////////////////////////////////////////////////////
        '//
        '// �������� ��������� ������ ������ � ����������� ��������
        '//
        '////////////////////////////////////////////////////////////////////////////////

        If DataGridView1.SelectedRows.Count = 0 Then
            Button2.Enabled = False
            Button7.Enabled = False
            Button6.Enabled = False
        Else
            Button2.Enabled = True
            If Trim(DataGridView1.SelectedRows.Item(0).Cells(10).Value.ToString()) = "1" Then
                If Declarations.MyPermission = True Then
                    Button7.Enabled = True
                Else
                    Button7.Enabled = False
                End If
            Else
                Button7.Enabled = True
            End If
            If DataGridView2.SelectedRows.Count = 0 Then
                Button6.Enabled = True
            Else
                Button6.Enabled = False
            End If
        End If
    End Sub

    Private Sub CheckRemoveButtons()
        '////////////////////////////////////////////////////////////////////////////////
        '//
        '// �������� ��������� ������ �������� ������ �� �����������
        '//
        '////////////////////////////////////////////////////////////////////////////////

        If DataGridView2.SelectedRows.Count = 0 Then
            Button11.Enabled = False
            Button12.Enabled = False
        Else
            If DataGridView1.SelectedRows.Count = 0 Then
                Button11.Enabled = False
                Button12.Enabled = False
            Else
                If Trim(DataGridView1.SelectedRows.Item(0).Cells(10).Value.ToString()) = "1" Then
                    If Declarations.MyPermission = True Then
                        Button11.Enabled = True
                        Button12.Enabled = True
                    Else
                        Button11.Enabled = False
                        Button12.Enabled = False
                    End If
                Else
                    Button11.Enabled = True
                    Button12.Enabled = True
                End If
            End If
        End If
    End Sub

    Private Sub CheckAddButtons()
        '////////////////////////////////////////////////////////////////////////////////
        '//
        '// �������� ��������� ������ ���������� ������ � ����������
        '//
        '////////////////////////////////////////////////////////////////////////////////

        If DataGridView1.SelectedRows.Count = 0 Or DataGridView3.SelectedRows.Count = 0 Then
            Button9.Enabled = False
            Button10.Enabled = False
        Else
            If DataGridView1.SelectedRows.Count = 0 Then
                Button9.Enabled = False
                Button10.Enabled = False
            Else
                If Trim(DataGridView1.SelectedRows.Item(0).Cells(11).Value.ToString()) = "1" Then
                    If Declarations.MyPermission = True Then
                        Button9.Enabled = True
                        Button10.Enabled = True
                    Else
                        Button9.Enabled = False
                        Button10.Enabled = False
                    End If
                Else
                    Button9.Enabled = True
                    Button10.Enabled = True
                End If
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
        LoadIncludedOrders()
        LoadFreeOrders()
        CheckConsolidatedButtons()
        CheckRemoveButtons()
        CheckAddButtons()
    End Sub

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        '////////////////////////////////////////////////////////////////////////////////
        '//
        '// ���������� ���������� � ����
        '//
        '////////////////////////////////////////////////////////////////////////////////

        LoadConsolidatedOrders()
        LoadIncludedOrders()
        LoadFreeOrders()
        CheckConsolidatedButtons()
        CheckRemoveButtons()
        CheckAddButtons()
    End Sub

    Private Sub DataGridView1_CellFormatting(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellFormattingEventArgs) Handles DataGridView1.CellFormatting
        '////////////////////////////////////////////////////////////////////////////////
        '//
        '// ��������� ���������� �� ���������� �������
        '//
        '////////////////////////////////////////////////////////////////////////////////
        Dim row As DataGridViewRow

        row = DataGridView1.Rows(e.RowIndex)
        If Trim(row.Cells(7).Value.ToString) <> "" Then  '---�������� �����
            row.DefaultCellStyle.BackColor = Color.Empty
        Else
            row.DefaultCellStyle.BackColor = Color.LightGreen
        End If
        If Trim(row.Cells(2).Value.ToString) <> "" Then  '---����������� ��������
            row.Cells(2).Style.BackColor = Color.LightPink
        Else
            row.Cells(2).Style.BackColor = Color.Empty
        End If
        If Trim(row.Cells(4).Value.ToString) <> "" Then  '---����������� �������
            row.Cells(4).Style.BackColor = Color.LightPink
        Else
            row.Cells(4).Style.BackColor = Color.Empty
        End If
    End Sub

    Private Sub DataGridView2_CellFormatting(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellFormattingEventArgs) Handles DataGridView2.CellFormatting
        '////////////////////////////////////////////////////////////////////////////////
        '//
        '// ��������� ���������� �� �������, ���������� � ����������
        '//
        '////////////////////////////////////////////////////////////////////////////////
        Dim row As DataGridViewRow

        row = DataGridView2.Rows(e.RowIndex)
        If Trim(row.Cells(2).Value.ToString) <> "�� ��������" Then
            row.DefaultCellStyle.BackColor = Color.LightGray
        Else
            row.DefaultCellStyle.BackColor = Color.Empty
        End If
    End Sub

    Private Sub DataGridView3_CellFormatting(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellFormattingEventArgs) Handles DataGridView3.CellFormatting
        '////////////////////////////////////////////////////////////////////////////////
        '//
        '// ��������� ���������� �� �������, �� ���������� � ����������
        '//
        '////////////////////////////////////////////////////////////////////////////////
        Dim row As DataGridViewRow

        row = DataGridView3.Rows(e.RowIndex)
        If Trim(row.Cells(6).Value.ToString) = "2" Then
            row.DefaultCellStyle.BackColor = Color.LightPink
        Else
            If Trim(row.Cells(2).Value.ToString) <> "�� ��������" Then
                row.DefaultCellStyle.BackColor = Color.LightGray
            Else
                row.DefaultCellStyle.BackColor = Color.Empty
            End If
        End If
        
    End Sub

    Private Sub DataGridView1_SelectionChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles DataGridView1.SelectionChanged
        '////////////////////////////////////////////////////////////////////////////////
        '//
        '// ����� ������ ����������� ������
        '//
        '////////////////////////////////////////////////////////////////////////////////

        LoadIncludedOrders()
        LoadFreeOrders()
        CheckConsolidatedButtons()
        CheckRemoveButtons()
        CheckAddButtons()
    End Sub

    Private Sub Button8_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button8.Click
        '////////////////////////////////////////////////////////////////////////////////
        '//
        '// �������� ������ ����������� ������
        '//
        '////////////////////////////////////////////////////////////////////////////////
        Dim MyID As String

        MyID = "0"
        MyEditConsolidatedOrder = New EditConsolidatedOrder
        MyEditConsolidatedOrder.StartParam = "Create"
        MyEditConsolidatedOrder.ShowDialog()
        LoadConsolidatedOrders()
        '---������� ������� ������� � ������������ ������� ������
        For i As Integer = 0 To DataGridView1.Rows.Count - 1
            If Trim(DataGridView1.Item(0, i).Value.ToString) > MyID Then
                MyID = DataGridView1.Item(0, i).Value.ToString
            End If
        Next
        For i As Integer = 0 To DataGridView1.Rows.Count - 1
            If Trim(DataGridView1.Item(0, i).Value.ToString) = MyID Then
                DataGridView1.CurrentCell = DataGridView1.Item(0, i)
            End If
        Next
        LoadIncludedOrders()
        LoadFreeOrders()
        CheckConsolidatedButtons()
        CheckRemoveButtons()
        CheckAddButtons()
    End Sub

    Private Sub Button7_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button7.Click
        '////////////////////////////////////////////////////////////////////////////////
        '//
        '// �������������� ����������� ������
        '//
        '////////////////////////////////////////////////////////////////////////////////

        Declarations.MyOrderID = Trim(DataGridView1.SelectedRows.Item(0).Cells(0).Value.ToString())
        MyEditConsolidatedOrder = New EditConsolidatedOrder
        MyEditConsolidatedOrder.StartParam = "Edit"
        MyEditConsolidatedOrder.ShowDialog()
        LoadConsolidatedOrders()
        '---������� ������� ������� ���������������
        For i As Integer = 0 To DataGridView1.Rows.Count - 1
            If Trim(DataGridView1.Item(0, i).Value.ToString) = Declarations.MyOrderID Then
                DataGridView1.CurrentCell = DataGridView1.Item(0, i)
            End If
        Next
        LoadIncludedOrders()
        LoadFreeOrders()
        CheckConsolidatedButtons()
        CheckRemoveButtons()
        CheckAddButtons()
    End Sub

    Private Sub Button6_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button6.Click
        '////////////////////////////////////////////////////////////////////////////////
        '//
        '// �������� ����������� ������
        '//
        '////////////////////////////////////////////////////////////////////////////////
        Dim MySQLStr As String
        Dim MyRez As MsgBoxResult

        MyRez = MsgBox("�� ������������� ������ ������� ���������� �����?", MsgBoxStyle.YesNo, "��������!")
        If MyRez = MsgBoxResult.Yes Then
            MySQLStr = "DELETE FROM tbl_DisplacementOrder_ShipmentInfo "
            MySQLStr = MySQLStr & "WHERE (ID = " & Trim(DataGridView1.SelectedRows.Item(0).Cells(0).Value.ToString()) & ") "
            InitMyConn(False)
            Declarations.MyConn.Execute(MySQLStr)
            LoadConsolidatedOrders()
            LoadIncludedOrders()
            LoadFreeOrders()
            CheckConsolidatedButtons()
            CheckRemoveButtons()
            CheckAddButtons()
        End If
    End Sub

    
    Private Sub Button9_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button9.Click
        '////////////////////////////////////////////////////////////////////////////////
        '//
        '// ��������� ������ � ���������� �����
        '//
        '////////////////////////////////////////////////////////////////////////////////
        Dim MySQLStr As String

        Declarations.MyOrderID = Trim(DataGridView1.SelectedRows.Item(0).Cells(0).Value.ToString())
        MySQLStr = "INSERT INTO tbl_DisplacementOrder_Shipment "
        MySQLStr = MySQLStr & "(ID, OrderNumber, ShipmentsNumber) "
        MySQLStr = MySQLStr & "VALUES (NEWID(), "
        MySQLStr = MySQLStr & "N'" & Trim(DataGridView3.SelectedRows.Item(0).Cells(0).Value.ToString()) & "', "
        MySQLStr = MySQLStr & Declarations.MyOrderID & ") "
        InitMyConn(False)
        Declarations.MyConn.Execute(MySQLStr)
        LoadConsolidatedOrders()
        '---������� ������� ������� ���������������
        For i As Integer = 0 To DataGridView1.Rows.Count - 1
            If Trim(DataGridView1.Item(0, i).Value.ToString) = Declarations.MyOrderID Then
                DataGridView1.CurrentCell = DataGridView1.Item(0, i)
            End If
        Next
        LoadIncludedOrders()
        LoadFreeOrders()
        CheckConsolidatedButtons()
        CheckRemoveButtons()
        CheckAddButtons()
    End Sub

    Private Sub Button11_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button11.Click
        '////////////////////////////////////////////////////////////////////////////////
        '//
        '// �������� ������ �� ����������� ������
        '//
        '////////////////////////////////////////////////////////////////////////////////
        Dim MySQLStr As String

        Declarations.MyOrderID = Trim(DataGridView1.SelectedRows.Item(0).Cells(0).Value.ToString())
        MySQLStr = "DELETE FROM tbl_DisplacementOrder_Shipment "
        MySQLStr = MySQLStr & "WHERE (OrderNumber = N'" & Trim(DataGridView2.SelectedRows.Item(0).Cells(0).Value.ToString()) & "') AND "
        MySQLStr = MySQLStr & "(ShipmentsNumber = " & Declarations.MyOrderID & ")"
        InitMyConn(False)
        Declarations.MyConn.Execute(MySQLStr)
        LoadConsolidatedOrders()
        '---������� ������� ������� ���������������
        For i As Integer = 0 To DataGridView1.Rows.Count - 1
            If Trim(DataGridView1.Item(0, i).Value.ToString) = Declarations.MyOrderID Then
                DataGridView1.CurrentCell = DataGridView1.Item(0, i)
            End If
        Next
        LoadIncludedOrders()
        LoadFreeOrders()
        CheckConsolidatedButtons()
        CheckRemoveButtons()
        CheckAddButtons()
    End Sub

    Private Sub Button12_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button12.Click
        '////////////////////////////////////////////////////////////////////////////////
        '//
        '// �������� ���� ������� �� ����������� ������
        '//
        '////////////////////////////////////////////////////////////////////////////////
        Dim MySQLStr As String

        Declarations.MyOrderID = Trim(DataGridView1.SelectedRows.Item(0).Cells(0).Value.ToString())
        MySQLStr = "DELETE FROM tbl_DisplacementOrder_Shipment "
        MySQLStr = MySQLStr & "WHERE (ShipmentsNumber = " & Declarations.MyOrderID & ")"
        InitMyConn(False)
        Declarations.MyConn.Execute(MySQLStr)
        LoadConsolidatedOrders()
        '---������� ������� ������� ���������������
        For i As Integer = 0 To DataGridView1.Rows.Count - 1
            If Trim(DataGridView1.Item(0, i).Value.ToString) = Declarations.MyOrderID Then
                DataGridView1.CurrentCell = DataGridView1.Item(0, i)
            End If
        Next
        LoadIncludedOrders()
        LoadFreeOrders()
        CheckConsolidatedButtons()
        CheckRemoveButtons()
        CheckAddButtons()
    End Sub

    Private Sub Button10_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button10.Click
        '////////////////////////////////////////////////////////////////////////////////
        '//
        '// ��������� ���� ��������� ������� � ���������� �����
        '//
        '////////////////////////////////////////////////////////////////////////////////
        Dim MySQLStr As String

        Declarations.MyOrderID = Trim(DataGridView1.SelectedRows.Item(0).Cells(0).Value.ToString())
        MySQLStr = "INSERT INTO tbl_DisplacementOrder_Shipment "
        MySQLStr = MySQLStr & "SELECT NEWID() AS ID, SC7C001, " & Declarations.MyOrderID & " AS Expr2 "
        MySQLStr = MySQLStr & "FROM SC7C0300 INNER JOIN "
        MySQLStr = MySQLStr & "tbl_DisplacementOrder ON SC7C0300.SC7C001 = tbl_DisplacementOrder.OrderNumber "
        MySQLStr = MySQLStr & "WHERE (SC7C010 = 1 OR SC7C010 = 2) AND "
        MySQLStr = MySQLStr & "(SC7C004 = N'" & Declarations.WHFromCode & "') AND "
        MySQLStr = MySQLStr & "(SC7C006 = N'" & Declarations.WHToCode & "') AND "
        MySQLStr = MySQLStr & "(tbl_DisplacementOrder.ReadyFlag = 1) AND "
        MySQLStr = MySQLStr & "(SC7C001 NOT IN (SELECT DISTINCT OrderNumber FROM tbl_DisplacementOrder_Shipment AS tbl_DisplacementOrder_Shipment_1)) "
        InitMyConn(False)
        Declarations.MyConn.Execute(MySQLStr)
        LoadConsolidatedOrders()
        '---������� ������� ������� ���������������
        For i As Integer = 0 To DataGridView1.Rows.Count - 1
            If Trim(DataGridView1.Item(0, i).Value.ToString) = Declarations.MyOrderID Then
                DataGridView1.CurrentCell = DataGridView1.Item(0, i)
            End If
        Next
        LoadIncludedOrders()
        LoadFreeOrders()
        CheckConsolidatedButtons()
        CheckRemoveButtons()
        CheckAddButtons()
    End Sub

    Private Sub Button2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button2.Click
        '////////////////////////////////////////////////////////////////////////////////
        '//
        '// ��������� �������� ����������� ������ 
        '//
        '////////////////////////////////////////////////////////////////////////////////

        If My.Settings.UseOffice = "LibreOffice" Then
            ExportToLO()
        Else
            ExportToExcel()
        End If
    End Sub
    Private Sub ExportToExcel()
        '////////////////////////////////////////////////////////////////////////////////
        '//
        '// ��������� �������� ����������� ������ � Excel
        '//
        '////////////////////////////////////////////////////////////////////////////////
        Dim MyObj As Object
        Dim MyWRKBook As Object
        Dim i As Integer                              '������� �����
        Dim ComOrder As String

        MyObj = CreateObject("Excel.Application")
        MyObj.SheetsInNewWorkbook = 1
        MyWRKBook = MyObj.Workbooks.Add

        ComOrder = Trim(DataGridView1.SelectedRows.Item(0).Cells(0).Value.ToString())

        ExportOrderHeaderToExcel(MyWRKBook, ComOrder, i)
        ExportOrderBodyToExcel(MyWRKBook, ComOrder, i)

        MyWRKBook.ActiveSheet.Range("A1").Select()
        MyObj.Application.Visible = True
        MyObj = Nothing
    End Sub

    Private Sub ExportToLO()
        '////////////////////////////////////////////////////////////////////////////////
        '//
        '// ��������� �������� ����������� ������ � LibreOffice
        '//
        '////////////////////////////////////////////////////////////////////////////////
        Dim i As Integer                              '������� �����
        Dim ComOrder As String
        Dim oServiceManager As Object
        Dim oDispatcher As Object
        Dim oDesktop As Object
        Dim oWorkBook As Object
        Dim oSheet As Object
        Dim Counter As Integer
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

        ComOrder = Trim(DataGridView1.SelectedRows.Item(0).Cells(0).Value.ToString())

        ExportOrderHeaderToLO(oSheet, oServiceManager, oWorkBook, oDispatcher, ComOrder, i)
        Counter = i
        ExportOrderBodyToLO(oSheet, oServiceManager, oWorkBook, oDispatcher, ComOrder, i)

        '-----������� �����
        LOFormatCells(oServiceManager, oDispatcher, oFrame, "D" & Counter & ":E" & i, 4)

        Dim args() As Object
        ReDim args(0)
        args(0) = oServiceManager.Bridge_GetStruct("com.sun.star.beans.PropertyValue")
        args(0).Name = "ToPoint"
        args(0).Value = "$A$1"
        oDispatcher.executeDispatch(oFrame, ".uno:GoToCell", "", 0, args)

        oWorkBook.CurrentController.Frame.ContainerWindow.Visible = True
        oWorkBook.CurrentController.Frame.ContainerWindow.toFront()
    End Sub

    Private Function mAkePropertyValue(ByVal cName, ByVal uValue, ByRef oServiceManager) As Object
        '/////////////////////////////////////////////////////////////////////////////////////
        '//
        '// ����������� ���������� ��� LO
        '// 
        '/////////////////////////////////////////////////////////////////////////////////////
        Dim oPropertyValue As Object

        oPropertyValue = oServiceManager.Bridge_getStruct("com.sun.star.beans.PropertyValue")
        oPropertyValue.Name = cName
        oPropertyValue.Value = uValue

        mAkePropertyValue = oPropertyValue
        oPropertyValue = Nothing
    End Function

    Private Sub ExportOrderHeaderToExcel(ByRef MyWRKBook As Object, ByVal ComOrder As String, ByRef i As Integer)
        '////////////////////////////////////////////////////////////////////////////////
        '//
        '// ��������� �������� ��������� ����������� ������ � Excel
        '//
        '////////////////////////////////////////////////////////////////////////////////

        MyWRKBook.ActiveSheet.Range("C1") = "������ ��������"
        MyWRKBook.ActiveSheet.Range("C1").Font.Name = "Arial"
        MyWRKBook.ActiveSheet.Range("C1").Font.Size = 12
        MyWRKBook.ActiveSheet.Range("C1").Font.Bold = True

        MyWRKBook.ActiveSheet.Range("D1:E1").MergeCells = True
        MyWRKBook.ActiveSheet.Range("D1") = Trim(DataGridView1.SelectedRows.Item(0).Cells(1).Value)
        MyWRKBook.ActiveSheet.Range("D1:E1").Font.Name = "Arial"
        MyWRKBook.ActiveSheet.Range("D1:E1").Font.Size = 8
        MyWRKBook.ActiveSheet.Range("D1:E1").Font.Bold = True
        MyWRKBook.ActiveSheet.Range("D1:E1").HorizontalAlignment = -4108

        MyWRKBook.ActiveSheet.Range("F1") = "N"
        MyWRKBook.ActiveSheet.Range("G1") = Trim(DataGridView1.SelectedRows.Item(0).Cells(0).Value)
        MyWRKBook.ActiveSheet.Range("F1:G1").Font.Name = "Arial"
        MyWRKBook.ActiveSheet.Range("F1:G1").Font.Size = 8
        MyWRKBook.ActiveSheet.Range("F1:G1").Font.Bold = True
        MyWRKBook.ActiveSheet.Range("F1:F1").HorizontalAlignment = -4108

        MyWRKBook.ActiveSheet.Range("B2") = "������:"
        MyWRKBook.ActiveSheet.Range("C2") = Declarations.WHFrom
        MyWRKBook.ActiveSheet.Range("B3") = "����:"
        MyWRKBook.ActiveSheet.Range("C3") = Declarations.WHTo
        MyWRKBook.ActiveSheet.Range("B2:C3").Font.Name = "Arial"
        MyWRKBook.ActiveSheet.Range("B2:C3").Font.Size = 8
        MyWRKBook.ActiveSheet.Range("B2:B3").Font.Bold = True

        MyWRKBook.ActiveSheet.Rows("4:4").RowHeight = 5

        MyWRKBook.ActiveSheet.Range("B5") = "������ �������� ���������:"
        MyWRKBook.ActiveSheet.Range("B5:B5").Font.Name = "Arial"
        MyWRKBook.ActiveSheet.Range("B5:B5").Font.Size = 8
        MyWRKBook.ActiveSheet.Range("B5:B5").Font.Bold = True

        MyWRKBook.ActiveSheet.Rows("6:6").RowHeight = 5

        MyWRKBook.ActiveSheet.Range("B7") = "N ������"
        MyWRKBook.ActiveSheet.Range("C7") = "���� ��������"
        MyWRKBook.ActiveSheet.Range("B7:C7").Font.Size = 7
        MyWRKBook.ActiveSheet.Range("B7:C7").WrapText = True
        MyWRKBook.ActiveSheet.Range("B7:C7").VerticalAlignment = -4108
        MyWRKBook.ActiveSheet.Range("B7:C7").HorizontalAlignment = -4108
        MyWRKBook.ActiveSheet.Range("B7:C7").Select()
        MyWRKBook.ActiveSheet.Range("B7:C7").Borders(5).LineStyle = -4142
        MyWRKBook.ActiveSheet.Range("B7:C7").Borders(6).LineStyle = -4142
        With MyWRKBook.ActiveSheet.Range("B7:C7").Borders(7)
            .LineStyle = 1
            .Weight = 4
            .ColorIndex = -4105
        End With
        With MyWRKBook.ActiveSheet.Range("B7:C7").Borders(8)
            .LineStyle = 1
            .Weight = 4
            .ColorIndex = -4105
        End With
        With MyWRKBook.ActiveSheet.Range("B7:C7").Borders(9)
            .LineStyle = 1
            .Weight = 4
            .ColorIndex = -4105
        End With
        With MyWRKBook.ActiveSheet.Range("B7:C7").Borders(10)
            .LineStyle = 1
            .Weight = 4
            .ColorIndex = -4105
        End With
        With MyWRKBook.ActiveSheet.Range("B7:C7").Borders(11)
            .LineStyle = 1
            .Weight = 4
            .ColorIndex = -4105
        End With
        With MyWRKBook.ActiveSheet.Range("B7:C7").Interior
            .ColorIndex = 36
            .Pattern = 1
            .PatternColorIndex = -4105
        End With

        i = 8
        For j As Integer = 0 To DataGridView2.Rows.Count - 1
            MyWRKBook.ActiveSheet.Range("B" & i) = DataGridView2.Item(0, j).Value
            MyWRKBook.ActiveSheet.Range("C" & i) = DataGridView2.Item(1, j).Value
            MyWRKBook.ActiveSheet.Range("B" & i & ":C" & i).Font.Name = "Arial"
            MyWRKBook.ActiveSheet.Range("B" & i & ":C" & i).Font.Size = 8
            i = i + 1
        Next

        MyWRKBook.ActiveSheet.Rows(i & ":" & i).RowHeight = 5
        i = i + 1

        MyWRKBook.ActiveSheet.Range("B" & i) = "������ �������� ���������:"
        MyWRKBook.ActiveSheet.Range("B" & i & ":B" & i).Font.Name = "Arial"
        MyWRKBook.ActiveSheet.Range("B" & i & ":B" & i).Font.Size = 8
        MyWRKBook.ActiveSheet.Range("B" & i & ":B" & i).Font.Bold = True
        i = i + 1

        MyWRKBook.ActiveSheet.Rows(i & ":" & i).RowHeight = 5
        i = i + 1

        MyWRKBook.ActiveSheet.Range("B" & i) = "���"
        MyWRKBook.ActiveSheet.Range("C" & i) = "��������"
        MyWRKBook.ActiveSheet.Range("D" & i) = "���-��"
        MyWRKBook.ActiveSheet.Range("E" & i) = "�� ���������"
        MyWRKBook.ActiveSheet.Range("F" & i) = "��. ���"
        MyWRKBook.ActiveSheet.Range("G" & i) = "������"
        MyWRKBook.ActiveSheet.Range("B" & i & ":G" & i).Font.Size = 7
        MyWRKBook.ActiveSheet.Range("B" & i & ":G" & i).WrapText = True
        MyWRKBook.ActiveSheet.Range("B" & i & ":G" & i).VerticalAlignment = -4108
        MyWRKBook.ActiveSheet.Range("B" & i & ":G" & i).HorizontalAlignment = -4108
        MyWRKBook.ActiveSheet.Range("B" & i & ":G" & i).Select()
        MyWRKBook.ActiveSheet.Range("B" & i & ":G" & i).Borders(5).LineStyle = -4142
        MyWRKBook.ActiveSheet.Range("B" & i & ":G" & i).Borders(6).LineStyle = -4142
        With MyWRKBook.ActiveSheet.Range("B" & i & ":G" & i).Borders(7)
            .LineStyle = 1
            .Weight = 4
            .ColorIndex = -4105
        End With
        With MyWRKBook.ActiveSheet.Range("B" & i & ":G" & i).Borders(8)
            .LineStyle = 1
            .Weight = 4
            .ColorIndex = -4105
        End With
        With MyWRKBook.ActiveSheet.Range("B" & i & ":G" & i).Borders(9)
            .LineStyle = 1
            .Weight = 4
            .ColorIndex = -4105
        End With
        With MyWRKBook.ActiveSheet.Range("B" & i & ":G" & i).Borders(10)
            .LineStyle = 1
            .Weight = 4
            .ColorIndex = -4105
        End With
        With MyWRKBook.ActiveSheet.Range("B" & i & ":G" & i).Borders(11)
            .LineStyle = 1
            .Weight = 4
            .ColorIndex = -4105
        End With
        With MyWRKBook.ActiveSheet.Range("B" & i & ":G" & i).Interior
            .ColorIndex = 36
            .Pattern = 1
            .PatternColorIndex = -4105
        End With
        MyWRKBook.ActiveSheet.Rows(i & ":" & i).RowHeight = 21
        i = i + 1

        MyWRKBook.ActiveSheet.Range("C" & i) = "N ������"
        MyWRKBook.ActiveSheet.Range("D" & i) = "���-�� �� ������"
        MyWRKBook.ActiveSheet.Range("E" & i) = "�� ��������� �� ������"
        MyWRKBook.ActiveSheet.Range("F" & i) = "��. ���"
        MyWRKBook.ActiveSheet.Range("G" & i) = "������"
        MyWRKBook.ActiveSheet.Range("C" & i & ":G" & i).Font.Size = 7
        MyWRKBook.ActiveSheet.Range("C" & i & ":G" & i).WrapText = True
        MyWRKBook.ActiveSheet.Range("C" & i & ":G" & i).VerticalAlignment = -4108
        MyWRKBook.ActiveSheet.Range("C" & i & ":G" & i).HorizontalAlignment = -4108
        MyWRKBook.ActiveSheet.Range("C" & i & ":G" & i).Select()
        MyWRKBook.ActiveSheet.Range("C" & i & ":G" & i).Borders(5).LineStyle = -4142
        MyWRKBook.ActiveSheet.Range("C" & i & ":G" & i).Borders(6).LineStyle = -4142
        With MyWRKBook.ActiveSheet.Range("C" & i & ":G" & i).Borders(7)
            .LineStyle = 1
            .Weight = 4
            .ColorIndex = -4105
        End With
        With MyWRKBook.ActiveSheet.Range("C" & i & ":G" & i).Borders(8)
            .LineStyle = 1
            .Weight = 4
            .ColorIndex = -4105
        End With
        With MyWRKBook.ActiveSheet.Range("C" & i & ":G" & i).Borders(9)
            .LineStyle = 1
            .Weight = 4
            .ColorIndex = -4105
        End With
        With MyWRKBook.ActiveSheet.Range("C" & i & ":G" & i).Borders(10)
            .LineStyle = 1
            .Weight = 4
            .ColorIndex = -4105
        End With
        With MyWRKBook.ActiveSheet.Range("C" & i & ":G" & i).Borders(11)
            .LineStyle = 1
            .Weight = 4
            .ColorIndex = -4105
        End With
        With MyWRKBook.ActiveSheet.Range("C" & i & ":G" & i).Interior
            .ColorIndex = 36
            .Pattern = 1
            .PatternColorIndex = -4105
        End With
        MyWRKBook.ActiveSheet.Rows(i & ":" & i).RowHeight = 29
        i = i + 1

        MyWRKBook.ActiveSheet.Rows(i & ":" & i).RowHeight = 1
        i = i + 1

        MyWRKBook.ActiveSheet.Columns("A:A").ColumnWidth = 1
        MyWRKBook.ActiveSheet.Columns("B:B").ColumnWidth = 10
        MyWRKBook.ActiveSheet.Columns("C:C").ColumnWidth = 40
        MyWRKBook.ActiveSheet.Columns("D:D").ColumnWidth = 7
        MyWRKBook.ActiveSheet.Columns("E:E").ColumnWidth = 7
        MyWRKBook.ActiveSheet.Columns("F:F").ColumnWidth = 6
        MyWRKBook.ActiveSheet.Columns("G:G").ColumnWidth = 8
    End Sub

    Private Sub ExportOrderHeaderToLO(ByRef oSheet As Object, ByRef oServiceManager As Object, ByRef oWorkBook As Object, _
        ByRef oDispatcher As Object, ByVal ComOrder As String, ByRef i As Integer)
        '////////////////////////////////////////////////////////////////////////////////
        '//
        '// ��������� �������� ��������� ����������� ������ � LibreOffice
        '//
        '////////////////////////////////////////////////////////////////////////////////
        Dim oFrame As Object

        oFrame = oWorkBook.getCurrentController.getFrame
        '-----������ �������
        oSheet.getColumns().getByName("A").Width = 400
        oSheet.getColumns().getByName("B").Width = 2000
        oSheet.getColumns().getByName("C").Width = 8000
        oSheet.getColumns().getByName("D").Width = 1400
        oSheet.getColumns().getByName("E").Width = 1400
        oSheet.getColumns().getByName("F").Width = 1200
        oSheet.getColumns().getByName("G").Width = 1600

        oSheet.getCellRangeByName("C1").String = "������ ��������"
        LOMergeCells(oServiceManager, oDispatcher, oFrame, "D1:E1")
        oSheet.getCellRangeByName("D1").String = Trim(DataGridView1.SelectedRows.Item(0).Cells(1).Value)
        oSheet.getCellRangeByName("F1").String = "N"
        oSheet.getCellRangeByName("G1").String = Trim(DataGridView1.SelectedRows.Item(0).Cells(0).Value)
        oSheet.getCellRangeByName("D1:G1").HoriJustify = 2
        LOFontSetFamilyName(oServiceManager, oDispatcher, oFrame, "A1:G1", "Arial")
        LOFontSetBold(oServiceManager, oDispatcher, oFrame, "A1:G1")
        LOFontSetSize(oServiceManager, oDispatcher, oFrame, "C1:C1", 12)
        LOFontSetSize(oServiceManager, oDispatcher, oFrame, "D1:G1", 8)


        oSheet.getCellRangeByName("B2").String = "������:"
        oSheet.getCellRangeByName("C2").String = Declarations.WHFrom
        oSheet.getCellRangeByName("B3").String = "����:"
        oSheet.getCellRangeByName("C3").String = Declarations.WHTo
        LOFontSetFamilyName(oServiceManager, oDispatcher, oFrame, "B2:C3", "Arial")
        LOFontSetSize(oServiceManager, oDispatcher, oFrame, "B2:C3", 8)
        LOFontSetBold(oServiceManager, oDispatcher, oFrame, "B2:B3")
        oSheet.getCellRangeByName("A4").Rows.Height = 200

        oSheet.getCellRangeByName("B5").String = "������ �������� ���������:"
        LOFontSetFamilyName(oServiceManager, oDispatcher, oFrame, "B5", "Arial")
        LOFontSetSize(oServiceManager, oDispatcher, oFrame, "B5", 8)
        LOFontSetBold(oServiceManager, oDispatcher, oFrame, "B5")
        oSheet.getCellRangeByName("A6").Rows.Height = 200

        oSheet.getCellRangeByName("B7").String = "N ������"
        oSheet.getCellRangeByName("C7").String = "���� ��������"
        LOFontSetFamilyName(oServiceManager, oDispatcher, oFrame, "B7:C7", "Arial")
        LOFontSetSize(oServiceManager, oDispatcher, oFrame, "B7:C7", 7)
        oSheet.getCellRangeByName("B7:C7").CellBackColor = 16775598
        Dim LineFormat As Object
        LineFormat = oServiceManager.Bridge_GetStruct("com.sun.star.table.BorderLine2")
        LineFormat.LineStyle = 0
        LineFormat.LineWidth = 70
        oSheet.getCellRangeByName("B7:C7").TopBorder = LineFormat
        oSheet.getCellRangeByName("B7:C7").RightBorder = LineFormat
        oSheet.getCellRangeByName("B7:C7").LeftBorder = LineFormat
        oSheet.getCellRangeByName("B7:C7").BottomBorder = LineFormat
        oSheet.getCellRangeByName("B7:C7").VertJustify = 2
        oSheet.getCellRangeByName("B7:C7").HoriJustify = 2
        LOWrapText(oServiceManager, oDispatcher, oFrame, "B7:C7")

        i = 8
        For j As Integer = 0 To DataGridView2.Rows.Count - 1
            oSheet.getCellRangeByName("B" & i).String = DataGridView2.Item(0, j).Value
            oSheet.getCellRangeByName("C" & i).Value = DataGridView2.Item(1, j).Value
            i = i + 1
        Next
        LOFontSetFamilyName(oServiceManager, oDispatcher, oFrame, "B8:C" & i, "Arial")
        LOFontSetSize(oServiceManager, oDispatcher, oFrame, "B8:C" & i, 8)
        LOFormatCells(oServiceManager, oDispatcher, oFrame, "C8:C" & i, 36)

        oSheet.getCellRangeByName("A" & i).Rows.Height = 200
        i = i + 1

        oSheet.getCellRangeByName("B" & i).String = "������ �������� ���������:"
        LOFontSetFamilyName(oServiceManager, oDispatcher, oFrame, "B" & i, "Arial")
        LOFontSetSize(oServiceManager, oDispatcher, oFrame, "B" & i, 8)
        LOFontSetBold(oServiceManager, oDispatcher, oFrame, "B" & i)
        i = i + 1

        oSheet.getCellRangeByName("A" & i).Rows.Height = 200
        i = i + 1
        
        oSheet.getCellRangeByName("B" & i).String = "���"
        oSheet.getCellRangeByName("C" & i).String = "��������"
        oSheet.getCellRangeByName("D" & i).String = "���-��"
        oSheet.getCellRangeByName("E" & i).String = "�� ���������"
        oSheet.getCellRangeByName("F" & i).String = "��. ���"
        oSheet.getCellRangeByName("G" & i).String = "������"
        LOFontSetFamilyName(oServiceManager, oDispatcher, oFrame, "B" & i & ":G" & i, "Arial")
        LOFontSetSize(oServiceManager, oDispatcher, oFrame, "B" & i & ":G" & i, 7)
        oSheet.getCellRangeByName("B" & i & ":G" & i).CellBackColor = 16775598
        LineFormat.LineStyle = 0
        LineFormat.LineWidth = 70
        oSheet.getCellRangeByName("B" & i & ":G" & i).TopBorder = LineFormat
        oSheet.getCellRangeByName("B" & i & ":G" & i).RightBorder = LineFormat
        oSheet.getCellRangeByName("B" & i & ":G" & i).LeftBorder = LineFormat
        oSheet.getCellRangeByName("B" & i & ":G" & i).BottomBorder = LineFormat
        oSheet.getCellRangeByName("B" & i & ":G" & i).VertJustify = 2
        oSheet.getCellRangeByName("B" & i & ":G" & i).HoriJustify = 2
        LOWrapText(oServiceManager, oDispatcher, oFrame, "B" & i & ":G" & i)
        i = i + 1

        oSheet.getCellRangeByName("C" & i).String = "N ������"
        oSheet.getCellRangeByName("D" & i).String = "���-�� �� ������"
        oSheet.getCellRangeByName("E" & i).String = "�� ��������� �� ������"
        oSheet.getCellRangeByName("F" & i).String = "��. ���"
        oSheet.getCellRangeByName("G" & i).String = "������"
        LOFontSetFamilyName(oServiceManager, oDispatcher, oFrame, "C" & i & ":G" & i, "Arial")
        LOFontSetSize(oServiceManager, oDispatcher, oFrame, "C" & i & ":G" & i, 7)
        oSheet.getCellRangeByName("C" & i & ":G" & i).CellBackColor = 16775598
        LineFormat.LineStyle = 0
        LineFormat.LineWidth = 70
        oSheet.getCellRangeByName("C" & i & ":G" & i).TopBorder = LineFormat
        oSheet.getCellRangeByName("C" & i & ":G" & i).RightBorder = LineFormat
        oSheet.getCellRangeByName("C" & i & ":G" & i).LeftBorder = LineFormat
        oSheet.getCellRangeByName("C" & i & ":G" & i).BottomBorder = LineFormat
        oSheet.getCellRangeByName("C" & i & ":G" & i).VertJustify = 2
        oSheet.getCellRangeByName("C" & i & ":G" & i).HoriJustify = 2
        LOWrapText(oServiceManager, oDispatcher, oFrame, "C" & i & ":G" & i)
        i = i + 1
    End Sub

    Private Sub ExportOrderBodyToExcel(ByRef MyWRKBook As Object, ByVal ComOrder As String, ByRef i As Integer)
        '////////////////////////////////////////////////////////////////////////////////
        '//
        '// ��������� �������� ���� ����������� ������ � Excel
        '//
        '////////////////////////////////////////////////////////////////////////////////
        Dim MySQLStr As String
        Dim OldCode As String                                 '���������� ��� ������ (��� ��������� - �������� ��� ���)

        MySQLStr = "IF exists(select * from tempdb..sysobjects where "
        MySQLStr = MySQLStr & "id = object_id(N'tempdb..#_MyShipment')  "
        MySQLStr = MySQLStr & "and xtype = N'U') "
        MySQLStr = MySQLStr & "DROP TABLE #_MyShipment "
        InitMyConn(False)
        Declarations.MyConn.Execute(MySQLStr)

        MySQLStr = "CREATE TABLE #_MyShipment( "
        MySQLStr = MySQLStr & "[StockCode] [nvarchar](35), "
        MySQLStr = MySQLStr & "[StockName] [nvarchar](80) NULL, "
        MySQLStr = MySQLStr & "[WHFrom][nvarchar](6), "
        MySQLStr = MySQLStr & "[QTYOrdered] float NULL, "
        MySQLStr = MySQLStr & "[QTYOrderedRest] float NULL, "
        MySQLStr = MySQLStr & "[BatchID] [nvarchar](20), "
        MySQLStr = MySQLStr & "[QTYOrderedBatch] float, "
        MySQLStr = MySQLStr & "[QTYOrderedBatchRest] float, "
        MySQLStr = MySQLStr & "[UOM] int, "
        MySQLStr = MySQLStr & "[UOMName][nvarchar](30) NULL,  "
        MySQLStr = MySQLStr & "[BINNumber] [nvarchar](6) NULL) "
        InitMyConn(False)
        Declarations.MyConn.Execute(MySQLStr)

        MySQLStr = "EXEC spp_DisplacementWorkplace_PickingListPrep N'" & ComOrder & "' "
        InitMyConn(False)
        Declarations.MyConn.Execute(MySQLStr)

        MySQLStr = "SELECT * FROM #_MyShipment "
        MySQLStr = MySQLStr & "ORDER BY StockCode, "
        MySQLStr = MySQLStr & "BatchID "
        InitMyConn(False)
        InitMyRec(False, MySQLStr)
        If Declarations.MyRec.BOF = True And Declarations.MyRec.EOF = True Then
            trycloseMyRec()
        Else
            Declarations.MyRec.MoveFirst()
            OldCode = ""
            While Declarations.MyRec.EOF = False
                If Trim(Declarations.MyRec.Fields("StockCode").Value) <> OldCode Then  '---����� ������ �� ���� � �����
                    '-----��� ������
                    MyWRKBook.ActiveSheet.Range("B" & CStr(i)).NumberFormat = "@"
                    MyWRKBook.ActiveSheet.Range("B" & CStr(i)) = Declarations.MyRec.Fields("StockCode").Value
                    OldCode = Trim(Declarations.MyRec.Fields("StockCode").Value)
                    '-----�������� ������
                    MyWRKBook.ActiveSheet.Range("C" & CStr(i)) = Declarations.MyRec.Fields("StockName").Value
                    '-----���������� � ����������� ���������� ������
                    MyWRKBook.ActiveSheet.Range("D" & CStr(i)) = Declarations.MyRec.Fields("QTYOrdered").Value
                    '-----������������� ���������� ������
                    MyWRKBook.ActiveSheet.Range("E" & CStr(i)) = Declarations.MyRec.Fields("QTYOrderedRest").Value
                    '-----������� ���������
                    MyWRKBook.ActiveSheet.Range("F" & CStr(i)).NumberFormat = "@"
                    MyWRKBook.ActiveSheet.Range("F" & CStr(i)) = Declarations.MyRec.Fields("UOMName").Value
                    '-----������ ��������
                    MyWRKBook.ActiveSheet.Range("G" & CStr(i)).NumberFormat = "@"
                    MyWRKBook.ActiveSheet.Range("G" & CStr(i)) = Declarations.MyRec.Fields("BINNumber").Value
                    '------------------------------��������������-------------------------------------
                    MyWRKBook.ActiveSheet.Range("B" & CStr(i) & ":G" & CStr(i)).Select()
                    MyWRKBook.ActiveSheet.Range("B" & CStr(i) & ":G" & CStr(i)).Borders(5).LineStyle = -4142
                    MyWRKBook.ActiveSheet.Range("B" & CStr(i) & ":G" & CStr(i)).Borders(6).LineStyle = -4142
                    With MyWRKBook.ActiveSheet.Range("B" & CStr(i) & ":G" & CStr(i)).Borders(7)
                        .LineStyle = 1
                        .Weight = 2
                        .ColorIndex = -4105
                    End With
                    With MyWRKBook.ActiveSheet.Range("B" & CStr(i) & ":G" & CStr(i)).Borders(8)
                        .LineStyle = 1
                        .Weight = 2
                        .ColorIndex = -4105
                    End With
                    With MyWRKBook.ActiveSheet.Range("B" & CStr(i) & ":G" & CStr(i)).Borders(9)
                        .LineStyle = 1
                        .Weight = 2
                        .ColorIndex = -4105
                    End With
                    With MyWRKBook.ActiveSheet.Range("B" & CStr(i) & ":G" & CStr(i)).Borders(10)
                        .LineStyle = 1
                        .Weight = 2
                        .ColorIndex = -4105
                    End With
                    With MyWRKBook.ActiveSheet.Range("B" & CStr(i) & ":G" & CStr(i)).Borders(11)
                        .LineStyle = 1
                        .Weight = 2
                        .ColorIndex = -4105
                    End With
                    With MyWRKBook.ActiveSheet.Range("B" & CStr(i) & ":G" & CStr(i)).Borders(12)
                        .LineStyle = 1
                        .Weight = 2
                        .ColorIndex = -4105
                    End With

                    MyWRKBook.ActiveSheet.Range("B" & CStr(i) & ":G" & CStr(i)).Select()
                    With MyWRKBook.ActiveSheet.Range("B" & CStr(i) & ":G" & CStr(i)).Font
                        .Name = "Arial"
                        .Size = 7
                    End With
                    With MyWRKBook.ActiveSheet.Range("B" & CStr(i) & ":G" & CStr(i)).Interior
                        .Color = 15132390
                        .Pattern = 1
                        .PatternColorIndex = -4105
                    End With
                    MyWRKBook.ActiveSheet.Range("B" & CStr(i) & ":C" & CStr(i)).Font.Bold = True
                    MyWRKBook.ActiveSheet.Range("C" & CStr(i) & ":C" & CStr(i)).WrapText = True
                    '------------------------------����� ��������������--------------------------------

                    i = i + 1
                End If
                '-------����� ���������� �� �������
                '-----����� ������
                MyWRKBook.ActiveSheet.Range("C" & CStr(i)).NumberFormat = "@"
                MyWRKBook.ActiveSheet.Range("C" & CStr(i)) = Declarations.MyRec.Fields("BatchID").Value
                '-----���������� � ����������� ���������� ������ �� ������
                MyWRKBook.ActiveSheet.Range("D" & CStr(i)) = Declarations.MyRec.Fields("QTYOrderedBatch").Value
                '-----������������� ���������� ������ �� ������
                MyWRKBook.ActiveSheet.Range("E" & CStr(i)) = Declarations.MyRec.Fields("QTYOrderedBatchRest").Value
                '-----������� ���������
                MyWRKBook.ActiveSheet.Range("F" & CStr(i)).NumberFormat = "@"
                MyWRKBook.ActiveSheet.Range("F" & CStr(i)) = Declarations.MyRec.Fields("UOMName").Value
                '-----������ ��������
                MyWRKBook.ActiveSheet.Range("G" & CStr(i)).NumberFormat = "@"
                MyWRKBook.ActiveSheet.Range("G" & CStr(i)) = Declarations.MyRec.Fields("BINNumber").Value
                '------------------------------��������������-------------------------------------
                MyWRKBook.ActiveSheet.Range("C" & CStr(i) & ":G" & CStr(i)).Select()
                With MyWRKBook.ActiveSheet.Range("C" & CStr(i) & ":G" & CStr(i)).Font
                    .Name = "Arial"
                    .Size = 7
                End With
                '------------------------------����� ��������������--------------------------------

                i = i + 1
                Declarations.MyRec.MoveNext()
            End While
        End If
    End Sub

    Private Sub ExportOrderBodyToLO(ByRef oSheet As Object, ByRef oServiceManager As Object, ByRef oWorkBook As Object, _
        ByRef oDispatcher As Object, ByVal ComOrder As String, ByRef i As Integer)
        '////////////////////////////////////////////////////////////////////////////////
        '//
        '// ��������� �������� ���� ����������� ������ � LibreOffice
        '//
        '////////////////////////////////////////////////////////////////////////////////
        Dim MySQLStr As String
        Dim OldCode As String                                 '���������� ��� ������ (��� ��������� - �������� ��� ���)
        Dim oFrame As Object
        Dim LineFormat As Object

        LineFormat = oServiceManager.Bridge_GetStruct("com.sun.star.table.BorderLine2")
        oFrame = oWorkBook.getCurrentController.getFrame

        MySQLStr = "IF exists(select * from tempdb..sysobjects where "
        MySQLStr = MySQLStr & "id = object_id(N'tempdb..#_MyShipment')  "
        MySQLStr = MySQLStr & "and xtype = N'U') "
        MySQLStr = MySQLStr & "DROP TABLE #_MyShipment "
        InitMyConn(False)
        Declarations.MyConn.Execute(MySQLStr)

        MySQLStr = "CREATE TABLE #_MyShipment( "
        MySQLStr = MySQLStr & "[StockCode] [nvarchar](35), "
        MySQLStr = MySQLStr & "[StockName] [nvarchar](80) NULL, "
        MySQLStr = MySQLStr & "[WHFrom][nvarchar](6), "
        MySQLStr = MySQLStr & "[QTYOrdered] float NULL, "
        MySQLStr = MySQLStr & "[QTYOrderedRest] float NULL, "
        MySQLStr = MySQLStr & "[BatchID] [nvarchar](20), "
        MySQLStr = MySQLStr & "[QTYOrderedBatch] float, "
        MySQLStr = MySQLStr & "[QTYOrderedBatchRest] float, "
        MySQLStr = MySQLStr & "[UOM] int, "
        MySQLStr = MySQLStr & "[UOMName][nvarchar](30) NULL,  "
        MySQLStr = MySQLStr & "[BINNumber] [nvarchar](6) NULL) "
        InitMyConn(False)
        Declarations.MyConn.Execute(MySQLStr)

        MySQLStr = "EXEC spp_DisplacementWorkplace_PickingListPrep N'" & ComOrder & "' "
        InitMyConn(False)
        Declarations.MyConn.Execute(MySQLStr)

        MySQLStr = "SELECT * FROM #_MyShipment "
        MySQLStr = MySQLStr & "ORDER BY StockCode, "
        MySQLStr = MySQLStr & "BatchID "
        InitMyConn(False)
        InitMyRec(False, MySQLStr)
        If Declarations.MyRec.BOF = True And Declarations.MyRec.EOF = True Then
            trycloseMyRec()
        Else
            Declarations.MyRec.MoveFirst()
            OldCode = ""
            While Declarations.MyRec.EOF = False
                If Trim(Declarations.MyRec.Fields("StockCode").Value) <> OldCode Then  '---����� ������ �� ���� � �����
                    '-----��� ������
                    oSheet.getCellRangeByName("B" & i).String = Declarations.MyRec.Fields("StockCode").Value
                    OldCode = Trim(Declarations.MyRec.Fields("StockCode").Value)
                    '-----�������� ������
                    oSheet.getCellRangeByName("C" & i).String = Declarations.MyRec.Fields("StockName").Value
                    '-----���������� � ����������� ���������� ������
                    oSheet.getCellRangeByName("D" & i).Value = Declarations.MyRec.Fields("QTYOrdered").Value
                    '-----������������� ���������� ������
                    oSheet.getCellRangeByName("E" & i).Value = Declarations.MyRec.Fields("QTYOrderedRest").Value
                    '-----������� ���������
                    oSheet.getCellRangeByName("F" & i).String = Declarations.MyRec.Fields("UOMName").Value
                    '-----������ ��������
                    oSheet.getCellRangeByName("G" & i).String = Declarations.MyRec.Fields("BINNumber").Value
                    '----�������
                    LOFontSetFamilyName(oServiceManager, oDispatcher, oFrame, "B" & i & ":G" & i, "Arial")
                    LineFormat.LineStyle = 0
                    LineFormat.LineWidth = 20
                    oSheet.getCellRangeByName("B" & i & ":G" & i).TopBorder = LineFormat
                    oSheet.getCellRangeByName("B" & i & ":G" & i).RightBorder = LineFormat
                    oSheet.getCellRangeByName("B" & i & ":G" & i).LeftBorder = LineFormat
                    oSheet.getCellRangeByName("B" & i & ":G" & i).BottomBorder = LineFormat
                    LOFontSetSize(oServiceManager, oDispatcher, oFrame, "B" & i & ":G" & i, 7)
                    LOWrapText(oServiceManager, oDispatcher, oFrame, "B" & i & ":G" & i)
                    oSheet.getCellRangeByName("B" & i & ":G" & i).CellBackColor = 14540253
                    i = i + 1
                End If
                '-----����� ������
                oSheet.getCellRangeByName("C" & CStr(i)).String = Declarations.MyRec.Fields("BatchID").Value
                '-----���������� � ����������� ���������� ������ �� ������
                oSheet.getCellRangeByName("D" & CStr(i)).Value = Declarations.MyRec.Fields("QTYOrderedBatch").Value
                '-----������������� ���������� ������ �� ������
                oSheet.getCellRangeByName("E" & CStr(i)).Value = Declarations.MyRec.Fields("QTYOrderedBatchRest").Value
                '-----������� ���������
                oSheet.getCellRangeByName("F" & CStr(i)).String = Declarations.MyRec.Fields("UOMName").Value
                '-----������ ��������
                oSheet.getCellRangeByName("G" & CStr(i)).String = Declarations.MyRec.Fields("BINNumber").Value
                LOFontSetFamilyName(oServiceManager, oDispatcher, oFrame, "C" & i & ":G" & i, "Arial")
                LineFormat.LineStyle = 0
                LineFormat.LineWidth = 20
                oSheet.getCellRangeByName("C" & i & ":G" & i).TopBorder = LineFormat
                oSheet.getCellRangeByName("C" & i & ":G" & i).RightBorder = LineFormat
                oSheet.getCellRangeByName("C" & i & ":G" & i).LeftBorder = LineFormat
                oSheet.getCellRangeByName("C" & i & ":G" & i).BottomBorder = LineFormat
                LOFontSetSize(oServiceManager, oDispatcher, oFrame, "C" & i & ":G" & i, 7)
                LOWrapText(oServiceManager, oDispatcher, oFrame, "C" & i & ":G" & i)
                i = i + 1
                Declarations.MyRec.MoveNext()
            End While
        End If
    End Sub
End Class