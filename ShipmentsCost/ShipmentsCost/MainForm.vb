Imports System.Runtime.InteropServices

Public Class MainForm
    <DllImport( _
        "user32.dll", _
        CharSet:=CharSet.Auto, _
        CallingConvention:=CallingConvention.StdCall _
    )> _
    Public Shared Function SetWindowPos( _
        ByVal hWnd As IntPtr, _
        ByVal hWndInsertAfter As IntPtr, _
        ByVal X As Int32, _
        ByVal Y As Int32, _
        ByVal cx As Int32, _
        ByVal cy As Int32, _
        ByVal uFlags As Int32) _
        As Boolean
    End Function
    Dim LoadFlag As Integer

    Private Sub MainForm_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        '/////////////////////////////////////////////////////////////////////////////////////
        '//
        '// ��� ������� ���������� ��������� - ���, ��������, ������������ � �.�.
        '// ����� ���� ������� ����� ����� �� ��������
        '/////////////////////////////////////////////////////////////////////////////////////

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

        LoadFlag = 1
        '---����� ������ � ����
        '---���������� ComboBox
        BuildComboBox()
        LoadFlag = 0

        '---����� - �����
        BuildPriceList()

        CheckButtons()
    End Sub

    Private Sub BuildComboBox()
        '/////////////////////////////////////////////////////////////////////////////////////
        '//
        '// ���������� Combobox ����������
        '// 
        '/////////////////////////////////////////////////////////////////////////////////////
        Dim MySQLStr As String                        '������� ������
        Dim MyAdapter As SqlClient.SqlDataAdapter     '
        Dim MyAdapter1 As SqlClient.SqlDataAdapter    '
        Dim MyAdapter2 As SqlClient.SqlDataAdapter    '
        Dim MyDs As New DataSet                       '
        Dim MyDs1 As New DataSet                       '
        Dim MyDs2 As New DataSet                       '

        '---������ �������
        MySQLStr = "SELECT SC23001, SC23001 + ' ' + SC23002 AS SC23002 "
        MySQLStr = MySQLStr & "FROM SC230300 WITH (NOLOCK) "
        MySQLStr = MySQLStr & "WHERE (SC23006 = N'1') "
        MySQLStr = MySQLStr & "ORDER BY SC23001"
        InitMyConn(False)
        Try
            MyAdapter = New SqlClient.SqlDataAdapter(MySQLStr, Declarations.MyNETConnStr)
            MyAdapter.SelectCommand.CommandTimeout = 600
            MyAdapter.Fill(MyDs)
            ComboBox1.DisplayMember = "SC23002" '��� �� ��� ����� ������������
            ComboBox1.ValueMember = "SC23001"   '��� �� ��� ����� ���������
            ComboBox1.DataSource = MyDs.Tables(0).DefaultView
        Catch ex As Exception
            MsgBox(ex.ToString)
        End Try

        '---������ ����� ��������
        MySQLStr = "SELECT ID, CONVERT(nvarchar, ID) + ' ' + Name AS Name "
        MySQLStr = MySQLStr & "FROM tbl_ShipmentsCost_ShipmentsType WITH (NOLOCK) "
        MySQLStr = MySQLStr & "ORDER BY ID"
        InitMyConn(False)
        Try
            MyAdapter1 = New SqlClient.SqlDataAdapter(MySQLStr, Declarations.MyNETConnStr)
            MyAdapter1.SelectCommand.CommandTimeout = 600
            MyAdapter1.Fill(MyDs1)
            ComboBox2.DisplayMember = "Name" '��� �� ��� ����� ������������
            ComboBox2.ValueMember = "ID"   '��� �� ��� ����� ���������
            ComboBox2.DataSource = MyDs1.Tables(0).DefaultView
        Catch ex As Exception
            MsgBox(ex.ToString)
        End Try

        '---������ ����� ������
        MySQLStr = "SELECT ID, CONVERT(nvarchar, ID) + ' ' + Name AS Name "
        MySQLStr = MySQLStr & "FROM tbl_ShipmentsCost_CostType WITH (NOLOCK) "
        MySQLStr = MySQLStr & "ORDER BY ID"
        InitMyConn(False)
        Try
            MyAdapter2 = New SqlClient.SqlDataAdapter(MySQLStr, Declarations.MyNETConnStr)
            MyAdapter2.SelectCommand.CommandTimeout = 600
            MyAdapter2.Fill(MyDs2)
            ComboBox3.DisplayMember = "Name" '��� �� ��� ����� ������������
            ComboBox3.ValueMember = "ID"   '��� �� ��� ����� ���������
            ComboBox3.DataSource = MyDs2.Tables(0).DefaultView
        Catch ex As Exception
            MsgBox(ex.ToString)
        End Try

    End Sub

    Private Sub BuildPriceList()
        '/////////////////////////////////////////////////////////////////////////////////////
        '//
        '// ����� ����� - �����
        '// 
        '/////////////////////////////////////////////////////////////////////////////////////
        Dim MySQLStr As String                        '������� ������
        Dim MyAdapter As SqlClient.SqlDataAdapter     '
        Dim MyDs As New DataSet

        MySQLStr = "SELECT tbl_ShipmentsCost_Price.Destination,"
        MySQLStr = MySQLStr & "tbl_ShipmentsCost_PriceType.Name, "
        MySQLStr = MySQLStr & "tbl_ShipmentsCost_Price.PriceFrom, "
        MySQLStr = MySQLStr & "tbl_ShipmentsCost_Price.PriceTo, "
        MySQLStr = MySQLStr & "tbl_ShipmentsCost_Price.PriceVal, "
        MySQLStr = MySQLStr & "tbl_ShipmentsCost_Price.MinCost "
        MySQLStr = MySQLStr & "FROM tbl_ShipmentsCost_Price WITH (NOLOCK) INNER JOIN "
        MySQLStr = MySQLStr & "tbl_ShipmentsCost_PriceType ON "
        MySQLStr = MySQLStr & "tbl_ShipmentsCost_Price.PriceType = tbl_ShipmentsCost_PriceType.ID "
        MySQLStr = MySQLStr & "WHERE (tbl_ShipmentsCost_Price.WHNum = N'" & ComboBox1.SelectedValue & "') AND "
        MySQLStr = MySQLStr & "(tbl_ShipmentsCost_Price.ShipmentsType = " & ComboBox2.SelectedValue & ") AND "
        MySQLStr = MySQLStr & "(tbl_ShipmentsCost_Price.CostType = " & ComboBox3.SelectedValue & ") "
        MySQLStr = MySQLStr & "ORDER BY tbl_ShipmentsCost_Price.Destination, "
        MySQLStr = MySQLStr & "tbl_ShipmentsCost_Price.PriceType, "
        MySQLStr = MySQLStr & "tbl_ShipmentsCost_Price.PriceFrom "

        InitMyConn(False)

        Try
            MyAdapter = New SqlClient.SqlDataAdapter(MySQLStr, Declarations.MyNETConnStr)
            MyAdapter.SelectCommand.CommandTimeout = 600
            MyAdapter.Fill(MyDs)
            DataGridView1.DataSource = MyDs.Tables(0)

            DataGridView1.Columns(0).HeaderText = "����� ����������"
            DataGridView1.Columns(0).Width = 210
            DataGridView1.Columns(1).HeaderText = "��� ����� �����"
            DataGridView1.Columns(1).Width = 100
            If ComboBox3.SelectedValue = 1 Then     '---�� ����
                DataGridView1.Columns(2).HeaderText = "������� � ����"
                DataGridView1.Columns(2).Width = 130
                DataGridView1.Columns(3).HeaderText = "�� ���"
                DataGridView1.Columns(3).Width = 130
                DataGridView1.Columns(4).HeaderText = "���� �� �� (���)"
                DataGridView1.Columns(4).Width = 130
                DataGridView1.Columns(5).HeaderText = "���. ���� (���)"
                DataGridView1.Columns(5).Width = 130
            Else                                    '---�� ������
                DataGridView1.Columns(2).HeaderText = "������� � ������"
                DataGridView1.Columns(2).Width = 130
                DataGridView1.Columns(3).HeaderText = "�� �����"
                DataGridView1.Columns(3).Width = 130
                DataGridView1.Columns(4).HeaderText = "���� �� ��� � (���)"
                DataGridView1.Columns(4).Width = 130
                DataGridView1.Columns(5).HeaderText = "���. ���� (���)"
                DataGridView1.Columns(5).Width = 130
            End If


            DataGridView1.SelectionMode = DataGridViewSelectionMode.FullRowSelect

        Catch ex As Exception
            MsgBox(ex.ToString)
        End Try

    End Sub

    Private Sub CheckButtons()
        '/////////////////////////////////////////////////////////////////////////////////////
        '//
        '// �������� ���������� ������
        '// 
        '/////////////////////////////////////////////////////////////////////////////////////

        If DataGridView1.SelectedRows.Count = 0 Then
            Button3.Enabled = False
            Button4.Enabled = False
        Else
            Button3.Enabled = True
            Button4.Enabled = True
        End If

    End Sub

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        '/////////////////////////////////////////////////////////////////////////////////////
        '//
        '// ����� �� ���������
        '// 
        '/////////////////////////////////////////////////////////////////////////////////////

        Application.Exit()
    End Sub

    Private Sub ComboBox1_SelectedValueChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles ComboBox1.SelectedValueChanged
        '////////////////////////////////////////////////////////////////////////////////
        '//
        '// ����� ������ � ComboBox1 - ������������� ������
        '//
        '////////////////////////////////////////////////////////////////////////////////

        If LoadFlag <> 1 Then
            BuildPriceList()
            CheckButtons()
        End If
    End Sub

    Private Sub ComboBox2_SelectedValueChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles ComboBox2.SelectedValueChanged
        '////////////////////////////////////////////////////////////////////////////////
        '//
        '// ����� ���� �������� � ComboBox2 - ������������� ������
        '//
        '////////////////////////////////////////////////////////////////////////////////

        If LoadFlag <> 1 Then
            BuildPriceList()
            CheckButtons()
        End If
    End Sub

    Private Sub ComboBox3_SelectedValueChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles ComboBox3.SelectedValueChanged
        '////////////////////////////////////////////////////////////////////////////////
        '//
        '// ����� ���� ������ � ComboBox3 - ������������� ������
        '//
        '////////////////////////////////////////////////////////////////////////////////

        If LoadFlag <> 1 Then
            BuildPriceList()
            CheckButtons()
        End If
    End Sub

    Private Sub Button2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button2.Click
        '////////////////////////////////////////////////////////////////////////////////
        '//
        '// ���������� �������� ����� - �����
        '//
        '////////////////////////////////////////////////////////////////////////////////

        AddPriceValue()
    End Sub

    Private Sub AddPriceValue()
        '////////////////////////////////////////////////////////////////////////////////
        '//
        '// ������� ���������� �������� ����� - �����
        '//
        '////////////////////////////////////////////////////////////////////////////////
        Dim MySQLStr As String                        '������� ������

        Declarations.MySuccess = False
        MyAddPriceValue = New AddPriceValue
        MyAddPriceValue.ShowDialog()
        If Declarations.MySuccess = False Then
            Exit Sub
        Else
            MySQLStr = "INSERT INTO tbl_ShipmentsCost_Price "
            MySQLStr = MySQLStr & "(ID, WHNum, ShipmentsType, CostType, Destination, PriceType, PriceFrom, PriceTo, PriceVal, MinCost) "
            MySQLStr = MySQLStr & "VALUES (NEWID(), "
            MySQLStr = MySQLStr & "N'" & ComboBox1.SelectedValue & "', "
            MySQLStr = MySQLStr & ComboBox2.SelectedValue & ", "
            MySQLStr = MySQLStr & ComboBox3.SelectedValue & ", "
            MySQLStr = MySQLStr & "N'" & Declarations.Destination & "', "
            MySQLStr = MySQLStr & Declarations.PriceType & ", "
            MySQLStr = MySQLStr & Replace(CStr(Declarations.PriceFrom), ",", ".") & ", "
            MySQLStr = MySQLStr & Replace(CStr(Declarations.PriceTo), ",", ".") & ", "
            MySQLStr = MySQLStr & Replace(CStr(Declarations.PriceVal), ",", ".") & ", "
            MySQLStr = MySQLStr & Replace(CStr(Declarations.MinCost), ",", ".") & ")"
            InitMyConn(False)
            Declarations.MyConn.Execute(MySQLStr)

            BuildPriceList()
            CheckButtons()
        End If
    End Sub

    Private Sub Button3_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button3.Click
        '////////////////////////////////////////////////////////////////////////////////
        '//
        '// �������� �������� ����� - �����
        '//
        '////////////////////////////////////////////////////////////////////////////////

        DeletePriceValue()
    End Sub

    Private Sub DeletePriceValue()
        '////////////////////////////////////////////////////////////////////////////////
        '//
        '// ������� �������� �������� ����� - �����
        '//
        '////////////////////////////////////////////////////////////////////////////////

        Dim MySQLStr As String                        '������� ������

        '---�������� �������� 
        MySQLStr = "DELETE FROM tbl_ShipmentsCost_Price "
        MySQLStr = MySQLStr & "WHERE (WHNum = N'" & ComboBox1.SelectedValue & "') AND "
        MySQLStr = MySQLStr & "(ShipmentsType = " & ComboBox2.SelectedValue & ") AND "
        MySQLStr = MySQLStr & "(CostType = " & ComboBox3.SelectedValue & ") AND "
        MySQLStr = MySQLStr & "(Destination = N'" & Trim(DataGridView1.SelectedRows.Item(0).Cells(0).Value.ToString()) & "') AND "
        MySQLStr = MySQLStr & "(PriceFrom = " & Replace(Trim(DataGridView1.SelectedRows.Item(0).Cells(2).Value.ToString), ",", ".") & ") AND "
        MySQLStr = MySQLStr & "(PriceTo = " & Replace(Trim(DataGridView1.SelectedRows.Item(0).Cells(3).Value.ToString), ",", ".") & ")"

        InitMyConn(False)
        Declarations.MyConn.Execute(MySQLStr)

        BuildPriceList()
        CheckButtons()
    End Sub

    Private Sub Button4_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button4.Click
        '////////////////////////////////////////////////////////////////////////////////
        '//
        '// �������������� �������� ����� - �����
        '//
        '////////////////////////////////////////////////////////////////////////////////

        EditPriceValue()
    End Sub

    Private Sub EditPriceValue()
        '////////////////////////////////////////////////////////////////////////////////
        '//
        '// ������� �������������� �������� ����� - �����
        '//
        '////////////////////////////////////////////////////////////////////////////////

        Dim MySQLStr As String                        '������� ������

        Declarations.MySuccess = False
        MyEditPriceValue = New EditPriceValue
        MyEditPriceValue.ShowDialog()
        If Declarations.MySuccess = False Then
            Exit Sub
        Else
            MySQLStr = "Update tbl_ShipmentsCost_Price "
            MySQLStr = MySQLStr & "SET PriceVal = " & Replace(Declarations.PriceVal, ",", ".") & ", "
            MySQLStr = MySQLStr & "MinCost = " & Replace(Declarations.MinCost, ",", ".") & " "
            MySQLStr = MySQLStr & "WHERE (WHNum = N'" & ComboBox1.SelectedValue & "') AND "
            MySQLStr = MySQLStr & "(ShipmentsType = " & ComboBox2.SelectedValue & ") AND "
            MySQLStr = MySQLStr & "(CostType = " & ComboBox3.SelectedValue & ") AND "
            MySQLStr = MySQLStr & "(Destination = N'" & Trim(DataGridView1.SelectedRows.Item(0).Cells(0).Value.ToString()) & "') AND "
            MySQLStr = MySQLStr & "(PriceFrom = " & Replace(Trim(DataGridView1.SelectedRows.Item(0).Cells(2).Value.ToString), ",", ".") & ") AND "
            MySQLStr = MySQLStr & "(PriceTo = " & Replace(Trim(DataGridView1.SelectedRows.Item(0).Cells(3).Value.ToString), ",", ".") & ")"

            InitMyConn(False)
            Declarations.MyConn.Execute(MySQLStr)

            BuildPriceList()
            CheckButtons()
        End If
    End Sub

    Private Sub Button6_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button6.Click
        '////////////////////////////////////////////////////////////////////////////////
        '//
        '// �������� � Excel �������� ����� - ����� �� ��������
        '//
        '////////////////////////////////////////////////////////////////////////////////

        Me.Cursor = Cursors.WaitCursor
        If My.Settings.UseOffice = "LibreOffice" Then
            UploadPriceInfoToLO()
        Else
            UploadPriceInfoToExcel()
        End If

        Me.Cursor = Cursors.Default
    End Sub

    Private Sub UploadPriceInfoToExcel()
        '////////////////////////////////////////////////////////////////////////////////
        '//
        '// ��������� �������� � Excel �������� ����� - ����� �� ��������
        '//
        '////////////////////////////////////////////////////////////////////////////////
        Dim WHList(,) As String         '������ �������
        Dim SHTypeList(,) As String     '������ ����� ��������
        Dim CostTypeList(,) As String   '������ ����� ����� - ������ (�� ����, ������...)
        Dim MySQLStr As String
        Dim i As Integer
        Dim j As Integer
        Dim k As Integer
        Dim StrNum As Double      '����� ������
        Dim MyObj As Object       'Excel
        Dim MyWRKBook As Object   '�����

        MySQLStr = "SELECT SC23001, SC23002 "
        MySQLStr = MySQLStr & "FROM SC230300 WITH (NOLOCK) "
        MySQLStr = MySQLStr & "WHERE (SC23006 = N'1') "
        MySQLStr = MySQLStr & "ORDER BY SC23001 "
        InitMyConn(False)
        InitMyRec(False, MySQLStr)
        If (Declarations.MyRec.EOF = True And Declarations.MyRec.BOF = True) Then
            trycloseMyRec()
            MsgBox("������ ��������� ���������� �� ���� ������. ���������� � ��������������", MsgBoxStyle.Critical, "��������!")
            Exit Sub
        Else
            Declarations.MyRec.MoveFirst()
            i = 0
            While Declarations.MyRec.EOF <> True
                ReDim Preserve WHList(1, i)
                WHList(0, i) = Declarations.MyRec.Fields("SC23001").Value.ToString()
                WHList(1, i) = Declarations.MyRec.Fields("SC23002").Value.ToString()
                i = i + 1
                Declarations.MyRec.MoveNext()
            End While
            trycloseMyRec()
        End If

        MyObj = CreateObject("Excel.Application")
        MyObj.SheetsInNewWorkbook = 1
        MyWRKBook = MyObj.Workbooks.Add

        '---����� ���������
        UploadCommonHeaderExcel(MyWRKBook)
        StrNum = 5
        For i = 0 To WHList.GetUpperBound(1)
            '---��������� ������
            StrNum = UploadWHHeaderExcel(MyWRKBook, WHList(0, i), WHList(1, i), StrNum)

            MySQLStr = "SELECT ID, Name "
            MySQLStr = MySQLStr & "FROM tbl_ShipmentsCost_ShipmentsType WITH (NOLOCK) "
            MySQLStr = MySQLStr & "ORDER BY ID "
            InitMyConn(False)
            InitMyRec(False, MySQLStr)
            If (Declarations.MyRec.EOF = True And Declarations.MyRec.BOF = True) Then
                trycloseMyRec()
                MsgBox("������ ��������� ���������� �� ���� ������ - tbl_ShipmentsCost_ShipmentsType. ���������� � ��������������", MsgBoxStyle.Critical, "��������!")
                Exit Sub
            Else
                Declarations.MyRec.MoveFirst()
                j = 0
                While Declarations.MyRec.EOF <> True
                    ReDim Preserve SHTypeList(1, j)
                    SHTypeList(0, j) = Declarations.MyRec.Fields("ID").Value.ToString()
                    SHTypeList(1, j) = Declarations.MyRec.Fields("Name").Value.ToString()
                    j = j + 1
                    Declarations.MyRec.MoveNext()
                End While
                trycloseMyRec()
            End If

            For j = 0 To SHTypeList.GetUpperBound(1)
                '---��������� ���� ��������
                StrNum = UploadSHTypeHeaderExcel(MyWRKBook, SHTypeList(0, j), SHTypeList(1, j), StrNum)

                MySQLStr = "SELECT ID, Name "
                MySQLStr = MySQLStr & "FROM  tbl_ShipmentsCost_CostType WITH (NOLOCK) "
                MySQLStr = MySQLStr & "ORDER BY ID "
                InitMyConn(False)
                InitMyRec(False, MySQLStr)
                If (Declarations.MyRec.EOF = True And Declarations.MyRec.BOF = True) Then
                    trycloseMyRec()
                    MsgBox("������ ��������� ���������� �� ���� ������ - tbl_ShipmentsCost_CostType. ���������� � ��������������", MsgBoxStyle.Critical, "��������!")
                    Exit Sub
                Else
                    Declarations.MyRec.MoveFirst()
                    k = 0
                    While Declarations.MyRec.EOF <> True
                        ReDim Preserve CostTypeList(1, k)
                        CostTypeList(0, k) = Declarations.MyRec.Fields("ID").Value.ToString()
                        CostTypeList(1, k) = Declarations.MyRec.Fields("Name").Value.ToString()
                        k = k + 1
                        Declarations.MyRec.MoveNext()
                    End While
                    trycloseMyRec()
                End If

                For k = 0 To CostTypeList.GetUpperBound(1)
                    '---��������� ���� ����� - ����� (�� ���, �����...)
                    StrNum = UploadCostTypeHeaderExcel(MyWRKBook, CostTypeList(0, k), CostTypeList(1, k), StrNum)
                    '---�������� ����� ����� - �����
                    StrNum = UploadRowsExcel(MyWRKBook, WHList(0, i), SHTypeList(0, j), CostTypeList(0, k), StrNum)
                Next


            Next

            'StrNum = UploadWHRows(MyWRKBook, WHList(0, i), StrNum)
        Next

        MyWRKBook.ActiveSheet.Range("A1").Select()
        MyObj.Application.Visible = True
        MyObj = Nothing
    End Sub

    Private Sub UploadPriceInfoToLO()
        '////////////////////////////////////////////////////////////////////////////////
        '//
        '// ��������� �������� � LibreOffice �������� ����� - ����� �� ��������
        '//
        '////////////////////////////////////////////////////////////////////////////////
        Dim WHList(,) As String         '������ �������
        Dim SHTypeList(,) As String     '������ ����� ��������
        Dim CostTypeList(,) As String   '������ ����� ����� - ������ (�� ����, ������...)
        Dim MySQLStr As String
        Dim i As Integer
        Dim j As Integer
        Dim k As Integer
        Dim StrNum As Integer      '����� ������
        Dim oServiceManager As Object
        Dim oDispatcher As Object
        Dim oDesktop As Object
        Dim oWorkBook As Object
        Dim oSheet As Object
        Dim oFrame As Object

        MySQLStr = "SELECT SC23001, SC23002 "
        MySQLStr = MySQLStr & "FROM SC230300 WITH (NOLOCK) "
        MySQLStr = MySQLStr & "WHERE (SC23006 = N'1') "
        MySQLStr = MySQLStr & "ORDER BY SC23001 "
        InitMyConn(False)
        InitMyRec(False, MySQLStr)
        If (Declarations.MyRec.EOF = True And Declarations.MyRec.BOF = True) Then
            trycloseMyRec()
            MsgBox("������ ��������� ���������� �� ���� ������. ���������� � ��������������", MsgBoxStyle.Critical, "��������!")
            Exit Sub
        Else
            Declarations.MyRec.MoveFirst()
            i = 0
            While Declarations.MyRec.EOF <> True
                ReDim Preserve WHList(1, i)
                WHList(0, i) = Declarations.MyRec.Fields("SC23001").Value.ToString()
                WHList(1, i) = Declarations.MyRec.Fields("SC23002").Value.ToString()
                i = i + 1
                Declarations.MyRec.MoveNext()
            End While
            trycloseMyRec()
        End If

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

        '---����� ���������
        UploadCommonHeaderLO(oServiceManager, oDispatcher, oDesktop, oWorkBook, _
            oSheet, oFrame)
        StrNum = 5

        For i = 0 To WHList.GetUpperBound(1)
            '---��������� ������
            StrNum = UploadWHHeaderLO(oServiceManager, oDispatcher, oDesktop, oWorkBook, _
            oSheet, oFrame, WHList(0, i), WHList(1, i), StrNum)

            '---���� �������� ��� ������
            MySQLStr = "SELECT ID, Name "
            MySQLStr = MySQLStr & "FROM tbl_ShipmentsCost_ShipmentsType WITH (NOLOCK) "
            MySQLStr = MySQLStr & "ORDER BY ID "
            InitMyConn(False)
            InitMyRec(False, MySQLStr)
            If (Declarations.MyRec.EOF = True And Declarations.MyRec.BOF = True) Then
                trycloseMyRec()
                MsgBox("������ ��������� ���������� �� ���� ������ - tbl_ShipmentsCost_ShipmentsType. ���������� � ��������������", MsgBoxStyle.Critical, "��������!")
                Exit Sub
            Else
                Declarations.MyRec.MoveFirst()
                j = 0
                While Declarations.MyRec.EOF <> True
                    ReDim Preserve SHTypeList(1, j)
                    SHTypeList(0, j) = Declarations.MyRec.Fields("ID").Value.ToString()
                    SHTypeList(1, j) = Declarations.MyRec.Fields("Name").Value.ToString()
                    j = j + 1
                    Declarations.MyRec.MoveNext()
                End While
                trycloseMyRec()
            End If

            For j = 0 To SHTypeList.GetUpperBound(1)
                StrNum = UploadSHTypeHeaderLO(oServiceManager, oDispatcher, oDesktop, oWorkBook, _
                oSheet, oFrame, SHTypeList(0, j), SHTypeList(1, j), StrNum)

                '---���� ����� - ����� (�� ���, �����...)
                MySQLStr = "SELECT ID, Name "
                MySQLStr = MySQLStr & "FROM  tbl_ShipmentsCost_CostType WITH (NOLOCK) "
                MySQLStr = MySQLStr & "ORDER BY ID "
                InitMyConn(False)
                InitMyRec(False, MySQLStr)
                If (Declarations.MyRec.EOF = True And Declarations.MyRec.BOF = True) Then
                    trycloseMyRec()
                    MsgBox("������ ��������� ���������� �� ���� ������ - tbl_ShipmentsCost_CostType. ���������� � ��������������", MsgBoxStyle.Critical, "��������!")
                    Exit Sub
                Else
                    Declarations.MyRec.MoveFirst()
                    k = 0
                    While Declarations.MyRec.EOF <> True
                        ReDim Preserve CostTypeList(1, k)
                        CostTypeList(0, k) = Declarations.MyRec.Fields("ID").Value.ToString()
                        CostTypeList(1, k) = Declarations.MyRec.Fields("Name").Value.ToString()
                        k = k + 1
                        Declarations.MyRec.MoveNext()
                    End While
                    trycloseMyRec()
                End If

                For k = 0 To CostTypeList.GetUpperBound(1)
                    '---��������� ���� ����� - ����� (�� ���, �����...)
                    StrNum = UploadCostTypeHeaderLO(oServiceManager, oDispatcher, oDesktop, oWorkBook, _
                        oSheet, oFrame, CostTypeList(0, k), CostTypeList(1, k), StrNum)
                    '---�������� ����� ����� - �����
                    StrNum = UploadRowsLO(oServiceManager, oDispatcher, oDesktop, oWorkBook, _
                        oSheet, oFrame, WHList(0, i), SHTypeList(0, j), CostTypeList(0, k), StrNum)
                Next k
            Next j
        Next i

        Dim args() As Object
        ReDim args(0)
        args(0) = oServiceManager.Bridge_GetStruct("com.sun.star.beans.PropertyValue")
        args(0).Name = "ToPoint"
        args(0).Value = "$A$1"
        oDispatcher.executeDispatch(oFrame, ".uno:GoToCell", "", 0, args)

        oWorkBook.CurrentController.Frame.ContainerWindow.Visible = True
        oWorkBook.CurrentController.Frame.ContainerWindow.toFront()
    End Sub

    Private Function UploadCommonHeaderExcel(ByRef MyWRKBook As Object)
        '////////////////////////////////////////////////////////////////////////////////////////
        '//
        '// ��������� �������� � Excel ������ ��������� 
        '//
        '////////////////////////////////////////////////////////////////////////////////////////

        MyWRKBook.ActiveSheet.Range("B1") = "���������� � ����� - ������ �� �������� "
        MyWRKBook.ActiveSheet.Range("B2") = "������� � �������� ������� �������� �������� " & Now
        MyWRKBook.ActiveSheet.Range("B1:B2").Select()
        MyWRKBook.ActiveSheet.Range("B1:B2").Font.Bold = True

        '--- � ������� �����
        MyWRKBook.ActiveSheet.Columns("A:A").ColumnWidth = 10
        MyWRKBook.ActiveSheet.Columns("B:B").ColumnWidth = 20
        MyWRKBook.ActiveSheet.Columns("C:C").ColumnWidth = 20
        MyWRKBook.ActiveSheet.Columns("D:D").ColumnWidth = 20
        MyWRKBook.ActiveSheet.Columns("E:E").ColumnWidth = 20
        MyWRKBook.ActiveSheet.Columns("F:F").ColumnWidth = 20
        MyWRKBook.ActiveSheet.Columns("G:G").ColumnWidth = 20
        MyWRKBook.ActiveSheet.Columns("H:H").ColumnWidth = 20
        MyWRKBook.ActiveSheet.Columns("I:I").ColumnWidth = 20
    End Function

    Private Function UploadCommonHeaderLO(ByRef oServiceManager As Object, ByRef oDispatcher As Object, _
        ByRef oDesktop As Object, ByRef oWorkBook As Object, ByRef oSheet As Object, _
        ByRef oFrame As Object)
        '////////////////////////////////////////////////////////////////////////////////////////
        '//
        '// ��������� �������� � LibreOffice ������ ��������� 
        '//
        '////////////////////////////////////////////////////////////////////////////////////////

        '---������ �������
        oSheet.getColumns().getByName("A").Width = 2030
        oSheet.getColumns().getByName("B").Width = 4060
        oSheet.getColumns().getByName("C").Width = 4060
        oSheet.getColumns().getByName("D").Width = 4060
        oSheet.getColumns().getByName("E").Width = 4060
        oSheet.getColumns().getByName("F").Width = 4060
        oSheet.getColumns().getByName("G").Width = 4060
        oSheet.getColumns().getByName("H").Width = 4060
        oSheet.getColumns().getByName("I").Width = 4060

        oSheet.getCellRangeByName("B1").String = "���������� � ����� - ������ �� ��������"
        oSheet.getCellRangeByName("B2").String = "������� � �������� ������� �������� �������� " & Now
        LOFontSetFamilyName(oServiceManager, oDispatcher, oFrame, "B1:B2", "Arial")
        LOFontSetBold(oServiceManager, oDispatcher, oFrame, "B1:B2")
        LOFontSetSize(oServiceManager, oDispatcher, oFrame, "B1:B2", 11)
    End Function

    Private Function UploadWHHeaderExcel(ByVal MyWRKBook As Object, ByVal WHCode As String, ByVal WHName As String, ByVal StrNum As Double) As Double
        '////////////////////////////////////////////////////////////////////////////////////////
        '//
        '// ��������� �������� � Excel ��������� �� ������ ������ 
        '//
        '////////////////////////////////////////////////////////////////////////////////////////

        MyWRKBook.ActiveSheet.Range("A" & StrNum) = "'" & WHCode
        MyWRKBook.ActiveSheet.Range("B" & StrNum) = WHName

        MyWRKBook.ActiveSheet.Range("A" & StrNum & ":B" & StrNum).Select()
        MyWRKBook.ActiveSheet.Range("A" & StrNum & ":B" & StrNum).Borders(5).LineStyle = -4142
        MyWRKBook.ActiveSheet.Range("A" & StrNum & ":B" & StrNum).Borders(6).LineStyle = -4142
        With MyWRKBook.ActiveSheet.Range("A" & StrNum & ":B" & StrNum).Borders(7)
            .LineStyle = 1
            .Weight = 4
            .ColorIndex = -4105
        End With
        With MyWRKBook.ActiveSheet.Range("A" & StrNum & ":B" & StrNum).Borders(8)
            .LineStyle = 1
            .Weight = 4
            .ColorIndex = -4105
        End With
        With MyWRKBook.ActiveSheet.Range("A" & StrNum & ":B" & StrNum).Borders(9)
            .LineStyle = 1
            .Weight = 4
            .ColorIndex = -4105
        End With
        With MyWRKBook.ActiveSheet.Range("A" & StrNum & ":B" & StrNum).Borders(10)
            .LineStyle = 1
            .Weight = 4
            .ColorIndex = -4105
        End With
        With MyWRKBook.ActiveSheet.Range("A" & StrNum & ":B" & StrNum).Borders(11)
            .LineStyle = 1
            .Weight = 4
            .ColorIndex = -4105
        End With
        With MyWRKBook.ActiveSheet.Range("A" & StrNum & ":B" & StrNum).Interior
            .ColorIndex = 36
            .Pattern = 1
            .PatternColorIndex = -4105
        End With

        StrNum = StrNum + 2

        Return StrNum
    End Function

    Private Function UploadWHHeaderLO(ByRef oServiceManager As Object, ByRef oDispatcher As Object, _
        ByRef oDesktop As Object, ByRef oWorkBook As Object, ByRef oSheet As Object, _
        ByRef oFrame As Object, ByVal WHCode As String, ByVal WHName As String, ByVal StrNum As Integer) As Integer
        '////////////////////////////////////////////////////////////////////////////////////////
        '//
        '// ��������� �������� � LibreOffice ��������� �� ������ ������ 
        '//
        '////////////////////////////////////////////////////////////////////////////////////////

        oSheet.getCellRangeByName("A" & StrNum).String = WHCode
        oSheet.getCellRangeByName("B" & StrNum).String = WHName
        LOFontSetFamilyName(oServiceManager, oDispatcher, oFrame, "A" & StrNum & ":B" & StrNum, "Arial")
        LOFontSetSize(oServiceManager, oDispatcher, oFrame, "A" & StrNum & ":B" & StrNum, 10)
        LOFontSetBold(oServiceManager, oDispatcher, oFrame, "A" & StrNum & ":B" & StrNum)
        oSheet.getCellRangeByName("A" & StrNum & ":B" & StrNum).CellBackColor = RGB(174, 249, 255)  '---��� ����� �������� ������� - �� ����� ���� ��� �� RGB � BGR!!!! ������ ���� �����
        LOSetBorders(oServiceManager, oSheet, "A" & StrNum & ":B" & StrNum, 70, 70, RGB(0, 0, 0)) '---��� ����� �������� ������� - �� ����� ���� ��� �� RGB � BGR!!!! ������ ���� �����
        oSheet.getCellRangeByName("A" & StrNum & ":B" & StrNum).VertJustify = 2
        oSheet.getCellRangeByName("A" & StrNum & ":B" & StrNum).HoriJustify = 2
        LOWrapText(oServiceManager, oDispatcher, oFrame, "A" & StrNum & ":B" & StrNum)
        StrNum = StrNum + 1

        Return StrNum
    End Function

    Private Function UploadSHTypeHeaderExcel(ByVal MyWRKBook As Object, ByVal SHTypeCode As String, ByVal SHTypeName As String, ByVal StrNum As Double) As Double
        '////////////////////////////////////////////////////////////////////////////////////////
        '//
        '// ��������� �������� � Excel ��������� �� ������ ���� �������� 
        '//
        '////////////////////////////////////////////////////////////////////////////////////////

        MyWRKBook.ActiveSheet.Range("B" & StrNum) = "'" & SHTypeCode
        MyWRKBook.ActiveSheet.Range("C" & StrNum) = SHTypeName

        MyWRKBook.ActiveSheet.Range("B" & StrNum & ":C" & StrNum).Select()
        MyWRKBook.ActiveSheet.Range("B" & StrNum & ":C" & StrNum).Borders(5).LineStyle = -4142
        MyWRKBook.ActiveSheet.Range("B" & StrNum & ":C" & StrNum).Borders(6).LineStyle = -4142
        With MyWRKBook.ActiveSheet.Range("B" & StrNum & ":C" & StrNum).Borders(7)
            .LineStyle = 1
            .Weight = 4
            .ColorIndex = -4105
        End With
        With MyWRKBook.ActiveSheet.Range("B" & StrNum & ":C" & StrNum).Borders(8)
            .LineStyle = 1
            .Weight = 4
            .ColorIndex = -4105
        End With
        With MyWRKBook.ActiveSheet.Range("B" & StrNum & ":C" & StrNum).Borders(9)
            .LineStyle = 1
            .Weight = 4
            .ColorIndex = -4105
        End With
        With MyWRKBook.ActiveSheet.Range("B" & StrNum & ":C" & StrNum).Borders(10)
            .LineStyle = 1
            .Weight = 4
            .ColorIndex = -4105
        End With
        With MyWRKBook.ActiveSheet.Range("B" & StrNum & ":C" & StrNum).Borders(11)
            .LineStyle = 1
            .Weight = 4
            .ColorIndex = -4105
        End With
        With MyWRKBook.ActiveSheet.Range("B" & StrNum & ":C" & StrNum).Interior
            .ColorIndex = 35
            .Pattern = 1
            .PatternColorIndex = -4105
        End With

        StrNum = StrNum + 2

        Return StrNum
    End Function

    Private Function UploadSHTypeHeaderLO(ByRef oServiceManager As Object, ByRef oDispatcher As Object, _
        ByRef oDesktop As Object, ByRef oWorkBook As Object, ByRef oSheet As Object, _
        ByRef oFrame As Object, ByVal SHTypeCode As String, ByVal SHTypeName As String, ByVal StrNum As Integer) As Integer
        '////////////////////////////////////////////////////////////////////////////////////////
        '//
        '// ��������� �������� � LibreOffice ��������� �� ������ ���� �������� 
        '//
        '////////////////////////////////////////////////////////////////////////////////////////

        oSheet.getCellRangeByName("B" & StrNum).String = SHTypeCode
        oSheet.getCellRangeByName("C" & StrNum).String = SHTypeName
        LOFontSetFamilyName(oServiceManager, oDispatcher, oFrame, "B" & StrNum & ":C" & StrNum, "Arial")
        LOFontSetSize(oServiceManager, oDispatcher, oFrame, "B" & StrNum & ":C" & StrNum, 10)
        LOFontSetBold(oServiceManager, oDispatcher, oFrame, "B" & StrNum & ":C" & StrNum)
        oSheet.getCellRangeByName("B" & StrNum & ":C" & StrNum).CellBackColor = RGB(204, 255, 204)  '---��� ����� �������� ������� - �� ����� ���� ��� �� RGB � BGR!!!! ������ ���� �����
        LOSetBorders(oServiceManager, oSheet, "B" & StrNum & ":C" & StrNum, 70, 70, RGB(0, 0, 0)) '---��� ����� �������� ������� - �� ����� ���� ��� �� RGB � BGR!!!! ������ ���� �����
        oSheet.getCellRangeByName("B" & StrNum & ":C" & StrNum).VertJustify = 2
        oSheet.getCellRangeByName("B" & StrNum & ":C" & StrNum).HoriJustify = 2
        LOWrapText(oServiceManager, oDispatcher, oFrame, "B" & StrNum & ":C" & StrNum)
        StrNum = StrNum + 1

        Return StrNum
    End Function

    Private Function UploadCostTypeHeaderExcel(ByVal MyWRKBook As Object, ByVal CostTypeCode As String, ByVal CostTypeName As String, ByVal StrNum As Double) As Double
        '////////////////////////////////////////////////////////////////////////////////////////
        '//
        '// ��������� �������� � Excel ��������� �� ������ ���� �������� 
        '//
        '////////////////////////////////////////////////////////////////////////////////////////

        MyWRKBook.ActiveSheet.Range("C" & StrNum) = "'" & CostTypeCode
        MyWRKBook.ActiveSheet.Range("D" & StrNum) = "�� " & CostTypeName

        MyWRKBook.ActiveSheet.Range("C" & StrNum & ":D" & StrNum).Select()
        MyWRKBook.ActiveSheet.Range("C" & StrNum & ":D" & StrNum).Borders(5).LineStyle = -4142
        MyWRKBook.ActiveSheet.Range("C" & StrNum & ":D" & StrNum).Borders(6).LineStyle = -4142
        With MyWRKBook.ActiveSheet.Range("C" & StrNum & ":D" & StrNum).Borders(7)
            .LineStyle = 1
            .Weight = 4
            .ColorIndex = -4105
        End With
        With MyWRKBook.ActiveSheet.Range("C" & StrNum & ":D" & StrNum).Borders(8)
            .LineStyle = 1
            .Weight = 4
            .ColorIndex = -4105
        End With
        With MyWRKBook.ActiveSheet.Range("C" & StrNum & ":D" & StrNum).Borders(9)
            .LineStyle = 1
            .Weight = 4
            .ColorIndex = -4105
        End With
        With MyWRKBook.ActiveSheet.Range("C" & StrNum & ":D" & StrNum).Borders(10)
            .LineStyle = 1
            .Weight = 4
            .ColorIndex = -4105
        End With
        With MyWRKBook.ActiveSheet.Range("C" & StrNum & ":D" & StrNum).Borders(11)
            .LineStyle = 1
            .Weight = 4
            .ColorIndex = -4105
        End With
        With MyWRKBook.ActiveSheet.Range("C" & StrNum & ":D" & StrNum).Interior
            .ColorIndex = 34
            .Pattern = 1
            .PatternColorIndex = -4105
        End With

        StrNum = StrNum + 1

        '---�������� ��������� ��� ����� ����� - �����
        MyWRKBook.ActiveSheet.Range("D" & StrNum) = "����� ����������"
        MyWRKBook.ActiveSheet.Range("E" & StrNum) = "��� ����� - �����"
        If CostTypeCode = 1 Then
            MyWRKBook.ActiveSheet.Range("F" & StrNum) = "������� � ����"
            MyWRKBook.ActiveSheet.Range("G" & StrNum) = "�� ���"
            MyWRKBook.ActiveSheet.Range("H" & StrNum) = "���� �� �� (���)"
            MyWRKBook.ActiveSheet.Range("I" & StrNum) = "���. ���� (���)"
        Else
            MyWRKBook.ActiveSheet.Range("F" & StrNum) = "������� � ������"
            MyWRKBook.ActiveSheet.Range("G" & StrNum) = "�� �����"
            MyWRKBook.ActiveSheet.Range("H" & StrNum) = "���� �� ��� � (���)"
            MyWRKBook.ActiveSheet.Range("I" & StrNum) = "���. ���� (���)"
        End If

        MyWRKBook.ActiveSheet.Range("D" & StrNum & ":I" & StrNum).Select()
        MyWRKBook.ActiveSheet.Range("D" & StrNum & ":I" & StrNum).HorizontalAlignment = -4108
        MyWRKBook.ActiveSheet.Range("D" & StrNum & ":I" & StrNum).VerticalAlignment = -4108
        MyWRKBook.ActiveSheet.Range("D" & StrNum & ":I" & StrNum).WrapText = True
        MyWRKBook.ActiveSheet.Range("D" & StrNum & ":I" & StrNum).Borders(5).LineStyle = -4142
        MyWRKBook.ActiveSheet.Range("D" & StrNum & ":I" & StrNum).Borders(6).LineStyle = -4142
        With MyWRKBook.ActiveSheet.Range("D" & StrNum & ":I" & StrNum).Borders(7)
            .LineStyle = 1
            .Weight = 4
            .ColorIndex = -4105
        End With
        With MyWRKBook.ActiveSheet.Range("D" & StrNum & ":I" & StrNum).Borders(8)
            .LineStyle = 1
            .Weight = 4
            .ColorIndex = -4105
        End With
        With MyWRKBook.ActiveSheet.Range("D" & StrNum & ":I" & StrNum).Borders(9)
            .LineStyle = 1
            .Weight = 4
            .ColorIndex = -4105
        End With
        With MyWRKBook.ActiveSheet.Range("D" & StrNum & ":I" & StrNum).Borders(10)
            .LineStyle = 1
            .Weight = 4
            .ColorIndex = -4105
        End With
        With MyWRKBook.ActiveSheet.Range("D" & StrNum & ":I" & StrNum).Borders(11)
            .LineStyle = 1
            .Weight = 4
            .ColorIndex = -4105
        End With
        With MyWRKBook.ActiveSheet.Range("D" & StrNum & ":I" & StrNum).Interior
            .ColorIndex = 15
            .Pattern = 1
            .PatternColorIndex = -4105
        End With

        StrNum = StrNum + 1

        Return StrNum
    End Function

    Private Function UploadCostTypeHeaderLO(ByRef oServiceManager As Object, ByRef oDispatcher As Object, _
        ByRef oDesktop As Object, ByRef oWorkBook As Object, ByRef oSheet As Object, _
        ByRef oFrame As Object, ByVal CostTypeCode As String, ByVal CostTypeName As String, ByVal StrNum As Integer) As Integer
        '////////////////////////////////////////////////////////////////////////////////////////
        '//
        '// ��������� �������� � LibreOffice ��������� �� ������ ���� �������� 
        '//
        '////////////////////////////////////////////////////////////////////////////////////////

        oSheet.getCellRangeByName("C" & StrNum).String = CostTypeCode
        oSheet.getCellRangeByName("D" & StrNum).String = CostTypeName
        LOFontSetFamilyName(oServiceManager, oDispatcher, oFrame, "C" & StrNum & ":D" & StrNum, "Arial")
        LOFontSetSize(oServiceManager, oDispatcher, oFrame, "C" & StrNum & ":D" & StrNum, 10)
        LOFontSetBold(oServiceManager, oDispatcher, oFrame, "C" & StrNum & ":D" & StrNum)
        oSheet.getCellRangeByName("C" & StrNum & ":D" & StrNum).CellBackColor = RGB(255, 255, 204)  '---��� ����� �������� ������� - �� ����� ���� ��� �� RGB � BGR!!!! ������ ���� �����
        LOSetBorders(oServiceManager, oSheet, "C" & StrNum & ":D" & StrNum, 70, 70, RGB(0, 0, 0)) '---��� ����� �������� ������� - �� ����� ���� ��� �� RGB � BGR!!!! ������ ���� �����
        oSheet.getCellRangeByName("C" & StrNum & ":D" & StrNum).VertJustify = 2
        oSheet.getCellRangeByName("C" & StrNum & ":D" & StrNum).HoriJustify = 2
        LOWrapText(oServiceManager, oDispatcher, oFrame, "C" & StrNum & ":D" & StrNum)
        StrNum = StrNum + 1

        '---�������� ��������� ��� ����� ����� - �����
        oSheet.getCellRangeByName("D" & StrNum).String = "����� ����������"
        oSheet.getCellRangeByName("E" & StrNum).String = "��� ����� - �����"
        If CostTypeCode = 1 Then
            oSheet.getCellRangeByName("F" & StrNum).String = "������� � ����"
            oSheet.getCellRangeByName("G" & StrNum).String = "�� ���"
            oSheet.getCellRangeByName("H" & StrNum).String = "���� �� �� (���)"
            oSheet.getCellRangeByName("I" & StrNum).String = "���. ���� (���)"
        Else
            oSheet.getCellRangeByName("F" & StrNum).String = "������� � ������"
            oSheet.getCellRangeByName("G" & StrNum).String = "�� �����"
            oSheet.getCellRangeByName("H" & StrNum).String = "���� �� ��� � (���)"
            oSheet.getCellRangeByName("I" & StrNum).String = "���. ���� (���)"
        End If
        LOFontSetFamilyName(oServiceManager, oDispatcher, oFrame, "D" & StrNum & ":I" & StrNum, "Arial")
        LOFontSetSize(oServiceManager, oDispatcher, oFrame, "D" & StrNum & ":I" & StrNum, 10)
        LOFontSetBold(oServiceManager, oDispatcher, oFrame, "D" & StrNum & ":I" & StrNum)
        oSheet.getCellRangeByName("D" & StrNum & ":I" & StrNum).CellBackColor = RGB(192, 192, 192)  '---��� ����� �������� ������� - �� ����� ���� ��� �� RGB � BGR!!!! ������ ���� �����
        LOSetBorders(oServiceManager, oSheet, "D" & StrNum & ":I" & StrNum, 70, 70, RGB(0, 0, 0)) '---��� ����� �������� ������� - �� ����� ���� ��� �� RGB � BGR!!!! ������ ���� �����
        oSheet.getCellRangeByName("D" & StrNum & ":I" & StrNum).VertJustify = 2
        oSheet.getCellRangeByName("D" & StrNum & ":I" & StrNum).HoriJustify = 2
        LOWrapText(oServiceManager, oDispatcher, oFrame, "D" & StrNum & ":I" & StrNum)
        StrNum = StrNum + 1

        Return StrNum
    End Function

    Private Function UploadRowsExcel(ByVal MyWRKBook As Object, ByVal WHCode As String, ByVal SHTypeCode As String, ByVal CostTypeCode As String, ByVal StrNum As Double) As Double
        '////////////////////////////////////////////////////////////////////////////////////////
        '//
        '// ��������� �������� � Excel ����� ����� - ����� 
        '//
        '////////////////////////////////////////////////////////////////////////////////////////
        Dim MySQLStr As String
        Dim aa As New System.Globalization.NumberFormatInfo
        Dim MySep As String
        Dim MyDig As String

        MySep = aa.CurrentInfo.NumberGroupSeparator
        MyDig = aa.CurrentInfo.NumberDecimalSeparator

        MySQLStr = "SELECT tbl_ShipmentsCost_Price.Destination,"
        MySQLStr = MySQLStr & "tbl_ShipmentsCost_PriceType.Name, "
        MySQLStr = MySQLStr & "tbl_ShipmentsCost_Price.PriceFrom, "
        MySQLStr = MySQLStr & "tbl_ShipmentsCost_Price.PriceTo, "
        MySQLStr = MySQLStr & "tbl_ShipmentsCost_Price.PriceVal, "
        MySQLStr = MySQLStr & "tbl_ShipmentsCost_Price.MinCost "
        MySQLStr = MySQLStr & "FROM tbl_ShipmentsCost_Price WITH (NOLOCK) INNER JOIN "
        MySQLStr = MySQLStr & "tbl_ShipmentsCost_PriceType ON "
        MySQLStr = MySQLStr & "tbl_ShipmentsCost_Price.PriceType = tbl_ShipmentsCost_PriceType.ID "
        MySQLStr = MySQLStr & "WHERE (tbl_ShipmentsCost_Price.WHNum = N'" & WHCode & "') AND "
        MySQLStr = MySQLStr & "(tbl_ShipmentsCost_Price.ShipmentsType = " & SHTypeCode & ") AND "
        MySQLStr = MySQLStr & "(tbl_ShipmentsCost_Price.CostType = " & CostTypeCode & ") "
        MySQLStr = MySQLStr & "ORDER BY tbl_ShipmentsCost_Price.Destination, "
        MySQLStr = MySQLStr & "tbl_ShipmentsCost_Price.PriceType, "
        MySQLStr = MySQLStr & "tbl_ShipmentsCost_Price.PriceFrom "
        InitMyConn(False)
        InitMyRec(False, MySQLStr)
        If (Declarations.MyRec.EOF = True And Declarations.MyRec.BOF = True) Then
            trycloseMyRec()
        Else
            Declarations.MyRec.MoveFirst()
            While Declarations.MyRec.EOF <> True
                MyWRKBook.ActiveSheet.Range("D" & StrNum) = "'" & Declarations.MyRec.Fields("Destination").Value
                MyWRKBook.ActiveSheet.Range("E" & StrNum) = "'" & Declarations.MyRec.Fields("Name").Value
                MyWRKBook.ActiveSheet.Range("F" & StrNum).NumberFormat = "#" & MySep & "##0" & MyDig & "00"
                MyWRKBook.ActiveSheet.Range("F" & StrNum) = Declarations.MyRec.Fields("PriceFrom").Value
                MyWRKBook.ActiveSheet.Range("G" & StrNum).NumberFormat = "#" & MySep & "##0" & MyDig & "00"
                MyWRKBook.ActiveSheet.Range("G" & StrNum) = Declarations.MyRec.Fields("PriceTo").Value
                MyWRKBook.ActiveSheet.Range("H" & StrNum).NumberFormat = "#" & MySep & "##0" & MyDig & "00"
                MyWRKBook.ActiveSheet.Range("H" & StrNum) = Declarations.MyRec.Fields("PriceVal").Value
                MyWRKBook.ActiveSheet.Range("I" & StrNum).NumberFormat = "#" & MySep & "##0" & MyDig & "00"
                MyWRKBook.ActiveSheet.Range("I" & StrNum) = Declarations.MyRec.Fields("MinCost").Value

                '---������������ ���� �� ������������� �����
                If Declarations.MyRec.Fields("Name").Value = "�� 100 ����������" Then
                    MyWRKBook.ActiveSheet.Range("D" & StrNum & ":I" & StrNum).Select()
                    With MyWRKBook.ActiveSheet.Range("D" & StrNum & ":I" & StrNum).Interior
                        .ColorIndex = 19
                        .Pattern = 1
                        .PatternColorIndex = -4105
                    End With
                End If

                StrNum = StrNum + 1
                Declarations.MyRec.MoveNext()
            End While
            trycloseMyRec()
        End If

        StrNum = StrNum + 2
        Return StrNum
    End Function

    Private Function UploadRowsLO(ByRef oServiceManager As Object, ByRef oDispatcher As Object, _
        ByRef oDesktop As Object, ByRef oWorkBook As Object, ByRef oSheet As Object, _
        ByRef oFrame As Object, ByVal WHCode As String, ByVal SHTypeCode As String, _
        ByVal CostTypeCode As String, ByVal StrNum As Integer) As Integer
        '////////////////////////////////////////////////////////////////////////////////////////
        '//
        '// ��������� �������� � LibreOffice ����� ����� - ����� 
        '//
        '////////////////////////////////////////////////////////////////////////////////////////
        Dim MySQLStr As String

        MySQLStr = "SELECT tbl_ShipmentsCost_Price.Destination,"
        MySQLStr = MySQLStr & "tbl_ShipmentsCost_PriceType.Name, "
        MySQLStr = MySQLStr & "tbl_ShipmentsCost_Price.PriceFrom, "
        MySQLStr = MySQLStr & "tbl_ShipmentsCost_Price.PriceTo, "
        MySQLStr = MySQLStr & "tbl_ShipmentsCost_Price.PriceVal, "
        MySQLStr = MySQLStr & "tbl_ShipmentsCost_Price.MinCost "
        MySQLStr = MySQLStr & "FROM tbl_ShipmentsCost_Price WITH (NOLOCK) INNER JOIN "
        MySQLStr = MySQLStr & "tbl_ShipmentsCost_PriceType ON "
        MySQLStr = MySQLStr & "tbl_ShipmentsCost_Price.PriceType = tbl_ShipmentsCost_PriceType.ID "
        MySQLStr = MySQLStr & "WHERE (tbl_ShipmentsCost_Price.WHNum = N'" & WHCode & "') AND "
        MySQLStr = MySQLStr & "(tbl_ShipmentsCost_Price.ShipmentsType = " & SHTypeCode & ") AND "
        MySQLStr = MySQLStr & "(tbl_ShipmentsCost_Price.CostType = " & CostTypeCode & ") "
        MySQLStr = MySQLStr & "ORDER BY tbl_ShipmentsCost_Price.Destination, "
        MySQLStr = MySQLStr & "tbl_ShipmentsCost_Price.PriceType, "
        MySQLStr = MySQLStr & "tbl_ShipmentsCost_Price.PriceFrom "
        InitMyConn(False)
        InitMyRec(False, MySQLStr)
        If (Declarations.MyRec.EOF = True And Declarations.MyRec.BOF = True) Then
            trycloseMyRec()
        Else
            Declarations.MyRec.MoveFirst()
            While Declarations.MyRec.EOF <> True
                oSheet.getCellRangeByName("D" & StrNum).String = Declarations.MyRec.Fields("Destination").Value
                oSheet.getCellRangeByName("E" & StrNum).String = Declarations.MyRec.Fields("Name").Value
                oSheet.getCellRangeByName("F" & StrNum).Value = Declarations.MyRec.Fields("PriceFrom").Value
                oSheet.getCellRangeByName("G" & StrNum).Value = Declarations.MyRec.Fields("PriceTo").Value
                oSheet.getCellRangeByName("H" & StrNum).Value = Declarations.MyRec.Fields("PriceVal").Value
                oSheet.getCellRangeByName("I" & StrNum).Value = Declarations.MyRec.Fields("MinCost").Value

                LOFormatCells(oServiceManager, oDispatcher, oFrame, "F" & StrNum & ":I" & StrNum, 4)

                '---������������ ���� �� ������������� �����
                If Declarations.MyRec.Fields("Name").Value = "�� 100 ����������" Then
                    LOSetBGColor(oServiceManager, oDispatcher, oFrame, oWorkBook, oSheet, "D" & StrNum & ":I" & StrNum, RGB(222, 241, 235)) '---��� ����� �������� ������� - �� ����� ���� ��� �� RGB � BGR!!!! ������ ���� �����
                End If

                StrNum = StrNum + 1
                Declarations.MyRec.MoveNext()
            End While
        End If

        StrNum = StrNum + 2
        Return StrNum
    End Function

    Private Sub DataGridView1_CellFormatting(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellFormattingEventArgs) Handles DataGridView1.CellFormatting
        '////////////////////////////////////////////////////////////////////////////////
        '//
        '// ��������� ������� �� �������� ��������� (�� �������������)
        '//
        '////////////////////////////////////////////////////////////////////////////////

        Dim row As DataGridViewRow

        row = DataGridView1.Rows(e.RowIndex)
        If (row.Cells(1).Value = "�� 100 ����������") Then
            row.DefaultCellStyle.BackColor = Color.LightSkyBlue
        End If

    End Sub

    Private Sub Button5_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button5.Click
        '////////////////////////////////////////////////////////////////////////////////////////
        '//
        '// ��������� �������� �� Excel ����� ����� �� ������ ������ �� ������ ���� ���������� �� ������ ���� ������ (�� ���, �����...) 
        '//
        '////////////////////////////////////////////////////////////////////////////////////////
        Dim MyRez As MsgBoxResult
        Dim MyTxt As String

        MyTxt = "��� ������� ������ ��� ���������� ����������� ���� Excel, � ������� � ������ D1 ������� ����� ������ (� �������������� 0), " & Chr(13) & Chr(10)
        MyTxt = MyTxt & "� ������ D2 ������� ��� �������� (1 - �����������, 2 - �.�. 3- ��������), " & Chr(13) & Chr(10)
        MyTxt = MyTxt & "� ������ D3 ������� ��� ����� - ����� (1 - �� ����, 2 - �� ������). " & Chr(13) & Chr(10)
        MyTxt = MyTxt & "������ � ����� Excel ������ ���������� � 6 ������, � ������� '�'. ������ ������ ���� ��������� ��� ���������. " & Chr(13) & Chr(10)
        MyTxt = MyTxt & "� ������� '�' ������ ������������� ������ ���������� (��� ����� '������� �� �������' ��� �������� �� 100 �� �� �������) " & Chr(13) & Chr(10)
        MyTxt = MyTxt & "� �������� 'D, 'E', 'F' � 'G' ������ ������������� �������� '���� ����� - �����', '������� �', '��' � '����'. " & Chr(13) & Chr(10)
        MyTxt = MyTxt & "��� ������� ������ ���� ���������." & Chr(13) & Chr(10)
        MyTxt = MyTxt & "� ��� ���� �������������� ���� Excel � �� ������ ������ ������?" & Chr(13) & Chr(10)
        MyRez = MsgBox(MyTxt, MsgBoxStyle.OkCancel, "��������!")
        If (MyRez = MsgBoxResult.Ok) Then
            If My.Settings.UseOffice = "LibreOffice" Then
                ImportDataFromLO()
            Else
                ImportDataFromExcel()
            End If

            BuildPriceList()
            CheckButtons()

            SetWindowPos(Me.Handle.ToInt32, -2, 0, 0, 0, 0, &H3)
            Me.Cursor = Cursors.Default
        Else

        End If
    End Sub

    Private Sub ImportDataFromExcel()
        '////////////////////////////////////////////////////////////////////////////////////////
        '//
        '// �������� �� Excel ����� ����� �� ������ ������ �� ������ ���� ���������� �� ������ ���� ������ (�� ���, �����...)
        '//
        '////////////////////////////////////////////////////////////////////////////////////////
        Dim appXLSRC As Object
        Dim MyWH As String              '---�����
        Dim MySHType As Integer         '---��� ��������
        Dim MyCostType As Integer       '---��� ������ (�� ����, ������...)
        Dim MyDestination As String     '---����� ���������� ��� "������� �� �������"
        Dim MyPriceType As Integer      '---��� ������ 0 - �������������, 1 - �� 100 ��
        Dim MyPriceFrom As Double       '---����� �...
        Dim MyPriceTo As Double         '---����� ��...
        Dim MyPriceVal As Double        '---���������� �����
        Dim MyMinVal As Double          '---����������� �����, ������������� � ����� ��������
        Dim StrCnt As Double
        Dim MySQLStr As String

        OpenFileDialog1.ShowDialog()
        If (OpenFileDialog1.FileName = "") Then
        Else
            Me.Cursor = Cursors.WaitCursor
            Me.Refresh()
            System.Windows.Forms.Application.DoEvents()

            appXLSRC = CreateObject("Excel.Application")
            appXLSRC.Workbooks.Open(OpenFileDialog1.FileName)

            '---�������� ��� � Excel ���������� ����� � ��� �� ���� � Scala
            MyWH = appXLSRC.Worksheets(1).Range("D1").Value
            '---��������� ��� � Excel ���������� ��� ������
            If MyWH = Nothing Then
                MsgBox("� ������������� ����� Excel � ������ 'D1' �� ���������� ��� ������ � Scala ", MsgBoxStyle.Critical, "��������!")
                appXLSRC.DisplayAlerts = 0
                appXLSRC.Workbooks.Close()
                appXLSRC.DisplayAlerts = 1
                appXLSRC.Quit()
                appXLSRC = Nothing
                Exit Sub
            End If
            '---��������� ��� ���� ����� ���� � Scala
            MySQLStr = "SELECT COUNT(SC23001) AS CC "
            MySQLStr = MySQLStr & "FROM SC230300 WITH (NOLOCK) "
            MySQLStr = MySQLStr & "WHERE (SC23006 = N'1') "
            MySQLStr = MySQLStr & "AND (SC23001 = N'" & MyWH & "')"
            InitMyConn(False)
            InitMyRec(False, MySQLStr)
            If (Declarations.MyRec.Fields("CC").Value = 0) Then
                MsgBox("� ������������� ����� Excel � ������ 'D1' ���������� �������� ��� ������ � Scala ", MsgBoxStyle.Critical, "��������!")
                appXLSRC.DisplayAlerts = 0
                appXLSRC.Workbooks.Close()
                appXLSRC.DisplayAlerts = 1
                appXLSRC.Quit()
                appXLSRC = Nothing
                trycloseMyRec()
                Exit Sub
            End If
            trycloseMyRec()

            '---�������� ��� � Excel ���������� ��� �������� � ��� �� ���� � Scala
            MySHType = appXLSRC.Worksheets(1).Range("D2").Value
            '---��������� ��� � Excel ���������� ��� ��������
            If MySHType = Nothing Then
                MsgBox("� ������������� ����� Excel � ������ 'D2' �� ���������� ��� �������� ", MsgBoxStyle.Critical, "��������!")
                appXLSRC.DisplayAlerts = 0
                appXLSRC.Workbooks.Close()
                appXLSRC.DisplayAlerts = 1
                appXLSRC.Quit()
                appXLSRC = Nothing
                Exit Sub
            End If
            '---��������� ��� ���� ��� �������� ���� � Scala
            MySQLStr = "SELECT COUNT(ID) AS CC "
            MySQLStr = MySQLStr & "FROM tbl_ShipmentsCost_ShipmentsType WITH (NOLOCK) "
            MySQLStr = MySQLStr & "WHERE (ID = " & MySHType & ")"
            InitMyConn(False)
            InitMyRec(False, MySQLStr)
            If (Declarations.MyRec.Fields("CC").Value = 0) Then
                MsgBox("� ������������� ����� Excel � ������ 'D2' ���������� �������� ��� �������� ", MsgBoxStyle.Critical, "��������!")
                appXLSRC.DisplayAlerts = 0
                appXLSRC.Workbooks.Close()
                appXLSRC.DisplayAlerts = 1
                appXLSRC.Quit()
                appXLSRC = Nothing
                trycloseMyRec()
                Exit Sub
            End If
            trycloseMyRec()

            '---�������� ��� � Excel ���������� ��� ����� ����� (�� ����, ������...) � ��� �� ���� � Scala
            MyCostType = appXLSRC.Worksheets(1).Range("D3").Value
            '---��������� ��� � Excel ���������� ��� ����� ����� (�� ����, ������...)
            If MyCostType = Nothing Then
                MsgBox("� ������������� ����� Excel � ������ 'D3' �� ���������� ��� ����� ����� (�� ����, ������...) ", MsgBoxStyle.Critical, "��������!")
                appXLSRC.DisplayAlerts = 0
                appXLSRC.Workbooks.Close()
                appXLSRC.DisplayAlerts = 1
                appXLSRC.Quit()
                appXLSRC = Nothing
                Exit Sub
            End If
            '---��������� ��� ���� ��� ����� ����� (�� ����, ������...) ���� � Scala
            MySQLStr = "SELECT COUNT(ID) AS CC "
            MySQLStr = MySQLStr & "FROM tbl_ShipmentsCost_CostType WITH (NOLOCK) "
            MySQLStr = MySQLStr & "WHERE (ID = " & MyCostType & ")"
            InitMyConn(False)
            InitMyRec(False, MySQLStr)
            If (Declarations.MyRec.Fields("CC").Value = 0) Then
                MsgBox("� ������������� ����� Excel � ������ 'D3' ���������� �������� ��� ����� ����� (�� ����, ������...) ", MsgBoxStyle.Critical, "��������!")
                appXLSRC.DisplayAlerts = 0
                appXLSRC.Workbooks.Close()
                appXLSRC.DisplayAlerts = 1
                appXLSRC.Quit()
                appXLSRC = Nothing
                trycloseMyRec()
                Exit Sub
            End If
            trycloseMyRec()

            '---�������� ������ �������� �� ������� (��� ������� ������, ���� �������� � ���� ����� - �����)
            MySQLStr = "DELETE FROM  tbl_ShipmentsCost_Price "
            MySQLStr = MySQLStr & "WHERE (WHNum = N'" & MyWH & "') AND "
            MySQLStr = MySQLStr & "(ShipmentsType = " & MySHType & ") AND "
            MySQLStr = MySQLStr & "(CostType = " & MyCostType & ") "
            InitMyConn(False)
            Declarations.MyConn.Execute(MySQLStr)

            StrCnt = 6
            While Not appXLSRC.Worksheets(1).Range("C" & StrCnt).Value = Nothing
                MyDestination = appXLSRC.Worksheets(1).Range("C" & StrCnt).Value.ToString
                If (appXLSRC.Worksheets(1).Range("D" & StrCnt).Value = Nothing And Not TypeOf (appXLSRC.Worksheets(1).Range("D" & StrCnt).Value) Is Double) Then
                    MsgBox("������ D" & StrCnt & " �������� ���� ����� ����� (0 - ������������� ��� 1 - �� 100 �� �� �������) ������ ���� ���������.", MsgBoxStyle.Critical, "��������!")
                Else
                    If (Not TypeOf (appXLSRC.Worksheets(1).Range("D" & StrCnt).Value) Is Double) Then
                        MsgBox("������ D" & StrCnt & " �������� ���� ����� ����� (0 - ������������� ��� 1 - �� 100 �� �� �������) ������ ���� �������� ���������.", MsgBoxStyle.Critical, "��������!")
                    Else
                        MyPriceType = appXLSRC.Worksheets(1).Range("D" & StrCnt).Value
                        If MyPriceType <> 0 And MyPriceType <> 1 Then
                            MsgBox("������ D" & StrCnt & " �������� ���� ����� ����� (0 - ������������� ��� 1 - �� 100 �� �� �������) ������ ���� 0 ��� 1.", MsgBoxStyle.Critical, "��������!")
                        Else
                            If (appXLSRC.Worksheets(1).Range("E" & StrCnt).Value = Nothing And Not TypeOf (appXLSRC.Worksheets(1).Range("E" & StrCnt).Value) Is Double) Then
                                MsgBox("������ E" & StrCnt & " �������� '������� �...' ������ ���� ���������", MsgBoxStyle.Critical, "��������!")
                            Else
                                If (Not TypeOf (appXLSRC.Worksheets(1).Range("E" & StrCnt).Value) Is Double) Then
                                    MsgBox("������ E" & StrCnt & " �������� '������� �...' ������ ���� ��������� �������� ���������.", MsgBoxStyle.Critical, "��������!")
                                Else
                                    MyPriceFrom = appXLSRC.Worksheets(1).Range("E" & StrCnt).Value
                                    If (appXLSRC.Worksheets(1).Range("F" & StrCnt).Value = Nothing And Not TypeOf (appXLSRC.Worksheets(1).Range("F" & StrCnt).Value) Is Double) Then
                                        MsgBox("������ F" & StrCnt & " �������� '��...' ������ ���� ���������", MsgBoxStyle.Critical, "��������!")
                                    Else
                                        If (Not TypeOf (appXLSRC.Worksheets(1).Range("F" & StrCnt).Value) Is Double) Then
                                            MsgBox("������ F" & StrCnt & " �������� '��...' ������ ���� ��������� �������� ���������.", MsgBoxStyle.Critical, "��������!")
                                        Else
                                            MyPriceTo = appXLSRC.Worksheets(1).Range("F" & StrCnt).Value
                                            If (appXLSRC.Worksheets(1).Range("G" & StrCnt).Value = Nothing And Not TypeOf (appXLSRC.Worksheets(1).Range("G" & StrCnt).Value) Is Double) Then
                                                MsgBox("������ G" & StrCnt & " �������� ���� ������ ���� ���������", MsgBoxStyle.Critical, "��������!")
                                            Else
                                                If (Not TypeOf (appXLSRC.Worksheets(1).Range("G" & StrCnt).Value) Is Double) Then
                                                    MsgBox("������ G" & StrCnt & " �������� ���� ������ ���� ��������� �������� ���������.", MsgBoxStyle.Critical, "��������!")
                                                Else
                                                    MyPriceVal = appXLSRC.Worksheets(1).Range("G" & StrCnt).Value
                                                    If (appXLSRC.Worksheets(1).Range("H" & StrCnt).Value = Nothing And Not TypeOf (appXLSRC.Worksheets(1).Range("H" & StrCnt).Value) Is Double) Then
                                                        MsgBox("������ H" & StrCnt & " �������� ����������� ���� ������ ���� ���������", MsgBoxStyle.Critical, "��������!")
                                                    Else
                                                        If (Not TypeOf (appXLSRC.Worksheets(1).Range("H" & StrCnt).Value) Is Double) Then
                                                            MsgBox("������ H" & StrCnt & " �������� ���� ������ ���� ��������� �������� ���������.", MsgBoxStyle.Critical, "��������!")
                                                        Else
                                                            MyMinVal = appXLSRC.Worksheets(1).Range("H" & StrCnt).Value
                                                            UpdateDBInfo(MyWH, MySHType, MyCostType, MyDestination, MyPriceType, MyPriceFrom, MyPriceTo, MyPriceVal, MyMinVal)
                                                        End If
                                                    End If
                                                End If
                                            End If
                                        End If
                                    End If
                                End If
                            End If
                        End If
                    End If
                End If
                StrCnt = StrCnt + 1
            End While

            appXLSRC.DisplayAlerts = 0
            appXLSRC.Workbooks.Close()
            appXLSRC.DisplayAlerts = 1
            appXLSRC.Quit()
            appXLSRC = Nothing
            MsgBox("������ ����� ����� ����������.", MsgBoxStyle.OkOnly, "��������!")
        End If
    End Sub

    Private Sub ImportDataFromLO()
        '////////////////////////////////////////////////////////////////////////////////////////
        '//
        '// �������� �� LibreOffice ����� ����� �� ������ ������ �� ������ ���� ���������� �� ������ ���� ������ (�� ���, �����...)
        '//
        '////////////////////////////////////////////////////////////////////////////////////////
        Dim MyWH As String              '---�����
        Dim MySHType As Integer         '---��� ��������
        Dim MyCostType As Integer       '---��� ������ (�� ����, ������...)
        Dim MyDestination As String     '---����� ���������� ��� "������� �� �������"
        Dim MyPriceType As Integer      '---��� ������ 0 - �������������, 1 - �� 100 ��
        Dim MyPriceFrom As Double       '---����� �...
        Dim MyPriceTo As Double         '---����� ��...
        Dim MyPriceVal As Double        '---���������� �����
        Dim MyMinVal As Double          '---����������� �����, ������������� � ����� ��������
        Dim StrCnt As Integer
        Dim MySQLStr As String
        Dim oServiceManager As Object
        Dim oDispatcher As Object
        Dim oDesktop As Object
        Dim oWorkBook As Object
        Dim oSheet As Object
        Dim oFileName As String

        OpenFileDialog2.ShowDialog()
        If (OpenFileDialog2.FileName = "") Then
        Else
            Me.Cursor = Cursors.WaitCursor
            Me.Refresh()
            System.Windows.Forms.Application.DoEvents()

            oServiceManager = CreateObject("com.sun.star.ServiceManager")
            oDesktop = oServiceManager.createInstance("com.sun.star.frame.Desktop")
            oDispatcher = oServiceManager.createInstance("com.sun.star.frame.DispatchHelper")
            oFileName = Replace(OpenFileDialog2.FileName, "\", "/")
            oFileName = "file:///" + oFileName
            Dim arg(1)
            arg(0) = mAkePropertyValue("Hidden", True, oServiceManager)
            arg(1) = mAkePropertyValue("MacroExecutionMode", 4, oServiceManager)
            oWorkBook = oDesktop.loadComponentFromURL(oFileName, "_blank", 0, arg)
            oSheet = oWorkBook.getSheets().getByIndex(0)

            '---�������� ��� � Excel ���������� ����� � ��� �� ���� � Scala
            MyWH = oSheet.getCellRangeByName("D1").String
            '---��������� ��� � Excel ���������� ��� ������
            If MyWH.Equals("") Then
                MsgBox("� ������������� ����� Excel � ������ 'D1' �� ���������� ��� ������ � Scala ", MsgBoxStyle.Critical, "��������!")
                oWorkBook.Close(True)
                Exit Sub
            End If

            '---��������� ��� ���� ����� ���� � Scala
            MySQLStr = "SELECT COUNT(SC23001) AS CC "
            MySQLStr = MySQLStr & "FROM SC230300 WITH (NOLOCK) "
            MySQLStr = MySQLStr & "WHERE (SC23006 = N'1') "
            MySQLStr = MySQLStr & "AND (SC23001 = N'" & MyWH & "')"
            InitMyConn(False)
            InitMyRec(False, MySQLStr)
            If (Declarations.MyRec.Fields("CC").Value = 0) Then
                MsgBox("� ������������� ����� Excel � ������ 'D1' ���������� �������� ��� ������ � Scala ", MsgBoxStyle.Critical, "��������!")
                oWorkBook.Close(True)
                trycloseMyRec()
                Exit Sub
            End If
            trycloseMyRec()

            '---�������� ��� � Excel ���������� ��� �������� � ��� �� ���� � Scala
            MySHType = oSheet.getCellRangeByName("D2").String
            '---��������� ��� � Excel ���������� ��� ��������
            If MySHType.Equals("") Then
                MsgBox("� ������������� ����� Excel � ������ 'D2' �� ���������� ��� �������� ", MsgBoxStyle.Critical, "��������!")
                oWorkBook.Close(True)
                Exit Sub
            End If
            '---��������� ��� ���� ��� �������� ���� � Scala
            MySQLStr = "SELECT COUNT(ID) AS CC "
            MySQLStr = MySQLStr & "FROM tbl_ShipmentsCost_ShipmentsType WITH (NOLOCK) "
            MySQLStr = MySQLStr & "WHERE (ID = " & MySHType & ")"
            InitMyConn(False)
            InitMyRec(False, MySQLStr)
            If (Declarations.MyRec.Fields("CC").Value = 0) Then
                MsgBox("� ������������� ����� Excel � ������ 'D2' ���������� �������� ��� �������� ", MsgBoxStyle.Critical, "��������!")
                oWorkBook.Close(True)
                trycloseMyRec()
                Exit Sub
            End If
            trycloseMyRec()

            '---�������� ��� � Excel ���������� ��� ����� ����� (�� ����, ������...) � ��� �� ���� � Scala
            MyCostType = oSheet.getCellRangeByName("D3").String
            '---��������� ��� � Excel ���������� ��� ����� ����� (�� ����, ������...)
            If MyCostType.Equals("") Then
                MsgBox("� ������������� ����� Excel � ������ 'D3' �� ���������� ��� ����� ����� (�� ����, ������...) ", MsgBoxStyle.Critical, "��������!")
                oWorkBook.Close(True)
                Exit Sub
            End If
            '---��������� ��� ���� ��� ����� ����� (�� ����, ������...) ���� � Scala
            MySQLStr = "SELECT COUNT(ID) AS CC "
            MySQLStr = MySQLStr & "FROM tbl_ShipmentsCost_CostType WITH (NOLOCK) "
            MySQLStr = MySQLStr & "WHERE (ID = " & MyCostType & ")"
            InitMyConn(False)
            InitMyRec(False, MySQLStr)
            If (Declarations.MyRec.Fields("CC").Value = 0) Then
                MsgBox("� ������������� ����� Excel � ������ 'D3' ���������� �������� ��� ����� ����� (�� ����, ������...) ", MsgBoxStyle.Critical, "��������!")
                oWorkBook.Close(True)
                trycloseMyRec()
                Exit Sub
            End If
            trycloseMyRec()

            '---�������� ������ �������� �� ������� (��� ������� ������, ���� �������� � ���� ����� - �����)
            MySQLStr = "DELETE FROM  tbl_ShipmentsCost_Price "
            MySQLStr = MySQLStr & "WHERE (WHNum = N'" & MyWH & "') AND "
            MySQLStr = MySQLStr & "(ShipmentsType = " & MySHType & ") AND "
            MySQLStr = MySQLStr & "(CostType = " & MyCostType & ") "
            InitMyConn(False)
            Declarations.MyConn.Execute(MySQLStr)

            StrCnt = 6
            While Not oSheet.getCellRangeByName("C" & StrCnt).String.Equals("")
                '---����� ����������
                MyDestination = oSheet.getCellRangeByName("C" & StrCnt).String
                '---��� ����� �����
                Try
                    MyPriceType = oSheet.getCellRangeByName("D" & StrCnt).Value
                    If MyPriceType <> 0 And MyPriceType <> 1 Then
                        MsgBox("������ D" & StrCnt & " �������� ���� ����� ����� (0 - ������������� ��� 1 - �� 100 �� �� �������) ������ ���� 0 ��� 1.", MsgBoxStyle.Critical, "��������!")
                        oWorkBook.Close(True)
                        Exit Sub
                    End If
                Catch ex As Exception
                    MsgBox("������ D" & StrCnt & " �������� ���� ����� ����� (0 - ������������� ��� 1 - �� 100 �� �� �������) ������ ���� �������� ���������.", MsgBoxStyle.Critical, "��������!")
                    oWorkBook.Close(True)
                    Exit Sub
                End Try

                '---�������� "������� �"
                Try
                    MyPriceFrom = oSheet.getCellRangeByName("E" & StrCnt).Value
                Catch ex As Exception
                    MsgBox("������ E" & StrCnt & " �������� '������� �...' ������ ���� ��������� �������� ���������.", MsgBoxStyle.Critical, "��������!")
                    oWorkBook.Close(True)
                    Exit Sub
                End Try

                '---�������� "��"
                Try
                    MyPriceTo = oSheet.getCellRangeByName("F" & StrCnt).Value
                Catch ex As Exception
                    MsgBox("������ F" & StrCnt & " �������� '��...' ������ ���� ��������� �������� ���������.", MsgBoxStyle.Critical, "��������!")
                    oWorkBook.Close(True)
                    Exit Sub
                End Try

                '---�������� ����
                Try
                    MyPriceVal = oSheet.getCellRangeByName("G" & StrCnt).Value
                Catch ex As Exception
                    MsgBox("������ G" & StrCnt & " �������� ���� ������ ���� ��������� �������� ���������.", MsgBoxStyle.Critical, "��������!")
                    oWorkBook.Close(True)
                    Exit Sub
                End Try

                '---�������� ������������ ����
                Try
                    MyMinVal = oSheet.getCellRangeByName("H" & StrCnt).Value
                Catch ex As Exception
                    MsgBox("������ H" & StrCnt & " �������� ����������� ���� ������ ���� ��������� �������� ���������.", MsgBoxStyle.Critical, "��������!")
                    oWorkBook.Close(True)
                    Exit Sub
                End Try

                UpdateDBInfo(MyWH, MySHType, MyCostType, MyDestination, MyPriceType, MyPriceFrom, MyPriceTo, MyPriceVal, MyMinVal)

                StrCnt = StrCnt + 1
            End While
            oWorkBook.Close(True)
            MsgBox("������ ����� ����� ����������.", MsgBoxStyle.OkOnly, "��������!")
        End If
    End Sub

    Private Sub UpdateDBInfo(ByVal MyWH As String, ByVal MySHType As Integer, ByVal MyCostType As Integer, ByVal MyDestination As String, _
        ByVal MyPriceType As Integer, ByVal MyPriceFrom As Double, ByVal MyPriceTo As Double, ByVal MyPriceVal As Double, ByVal MyMinVal As Double)
        '////////////////////////////////////////////////////////////////////////////////////////
        '//
        '// ��������� ���������� �� ����� ������ ����� - �����
        '//
        '////////////////////////////////////////////////////////////////////////////////////////
        Dim MySQLStr As String

        MySQLStr = "INSERT INTO tbl_ShipmentsCost_Price "
        MySQLStr = MySQLStr & "(ID, WHNum, ShipmentsType, CostType, Destination, PriceType, PriceFrom, PriceTo, PriceVal, MinCost) "
        MySQLStr = MySQLStr & "VALUES (NEWID(), "
        MySQLStr = MySQLStr & "N'" & MyWH & "', "
        MySQLStr = MySQLStr & Replace(CStr(MySHType), ",", ".") & ", "
        MySQLStr = MySQLStr & Replace(CStr(MyCostType), ",", ".") & ", "
        MySQLStr = MySQLStr & "N'" & MyDestination & "', "
        MySQLStr = MySQLStr & Replace(CStr(MyPriceType), ",", ".") & ", "
        MySQLStr = MySQLStr & Replace(CStr(MyPriceFrom), ",", ".") & ", "
        MySQLStr = MySQLStr & Replace(CStr(MyPriceTo), ",", ".") & ", "
        MySQLStr = MySQLStr & Replace(CStr(MyPriceVal), ",", ".") & ", "
        MySQLStr = MySQLStr & Replace(CStr(MyMinVal), ",", ".") & ") "
        InitMyConn(False)
        Declarations.MyConn.Execute(MySQLStr)
    End Sub

    Private Sub Button7_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button7.Click
        '////////////////////////////////////////////////////////////////////////////////////////
        '//
        '// ���� �������� ��������� �������� �� ��������
        '//
        '////////////////////////////////////////////////////////////////////////////////////////

        MyShipmentCost = New ShipmentCost
        MyShipmentCost.ShowDialog()
    End Sub
End Class
