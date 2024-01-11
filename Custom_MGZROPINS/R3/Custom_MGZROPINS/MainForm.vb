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

    Private Sub MainForm_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        '/////////////////////////////////////////////////////////////////////////////////////
        '//
        '// ��� ������� ���������� ��������� - ���, ��������, ������������ � �.�.
        '// ����� ���� ������� ������ ����������� ������� ������������
        '/////////////////////////////////////////////////////////////////////////////////////
        Dim MyRez As MsgBoxResult
        Dim MyTxt As String

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

        '---����� ������ � ����
        '---������ �������
        BuildWHList()

        '---������ ������� �� ������� ���, ROP � ��������� ����� ��������� �������������
        BuildAutoItemList()

        '---������ ������� �� ������� ���, ROP � ��������� ����� ������������ �������
        BuildManualItemList()

        CheckButtons()
        System.Windows.Forms.Application.DoEvents()
        If (Button10.Enabled = True) Then
            MyTxt = "�� ����� � �������� ������������ ��������������� ��������� �������� ���, ROP � ���������� ������" & Chr(13) & Chr(10)
            MyTxt = MyTxt & "� ��������� ����������� ���� ��������� ����� ������. ��� ���������� ������ �������� � �������" & Chr(13) & Chr(10)
            MyTxt = MyTxt & "������������� ����������� ��� ���� ������� ������������� ������������� ���, ROP � ��������� �����" & Chr(13) & Chr(10)
            MyTxt = MyTxt & "� ��������� ������ ��� ���� ������� ����� ������������� ����." & Chr(13) & Chr(10)
            MyTxt = MyTxt & "���������� ����������?"
            MyRez = MsgBox(MyTxt, MsgBoxStyle.OkCancel, "��������!")
            If (MyRez = MsgBoxResult.Ok) Then
                Me.Cursor = Cursors.WaitCursor
                ReCalculate_Partial()
                Me.Cursor = Cursors.Default
            End If
        End If

    End Sub

    Private Sub BuildWHList()
        '/////////////////////////////////////////////////////////////////////////////////////
        '//
        '// ����� � Combobox ������ ������� � ����� �������
        '// 
        '/////////////////////////////////////////////////////////////////////////////////////
        Dim MySQLStr As String                        '������� ������
        Dim MyAdapter As SqlClient.SqlDataAdapter     '
        Dim MyDs As New DataSet                       '


        MySQLStr = "SELECT SC23001, SC23001 + ' ' + SC23002 AS SC23002 "
        MySQLStr = MySQLStr & "FROM SC230300 WITH(NOLOCK) "
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

    End Sub

    Private Sub BuildAutoItemList()
        '/////////////////////////////////////////////////////////////////////////////////////
        '//
        '// ����� � DataGridView1 ������ ������� �� ������� ���, ROP � ��������� ����� ��������� �������������
        '// 
        '/////////////////////////////////////////////////////////////////////////////////////
        Dim MySQLStr As String                        '������� ������
        Dim MyAdapter As SqlClient.SqlDataAdapter     '
        Dim MyDs As New DataSet                       '

        MySQLStr = "SELECT tbl_ForecastOrderR3_Main.Code, "
        MySQLStr = MySQLStr & "tbl_ForecastOrderR3_Main.Name, "
        MySQLStr = MySQLStr & "tbl_ForecastOrderR3_Main.ABC, "
        MySQLStr = MySQLStr & "tbl_ForecastOrderR3_Main.XYZ, "
        MySQLStr = MySQLStr & "tbl_ForecastOrderR3_Main.LT, "
        MySQLStr = MySQLStr & "tbl_ForecastOrderR3_Main.OI, "
        MySQLStr = MySQLStr & "ROUND(CONVERT(float, tbl_ForecastOrderR3_Main.MGZ), 3) AS MGZ, "
        MySQLStr = MySQLStr & "ROUND(CONVERT(float, ISNULL(View_1.MGZ, tbl_ForecastOrderR3_Main.MGZ)), 3) AS MGZ_OLD, "
        MySQLStr = MySQLStr & "ROUND(CONVERT(float, tbl_ForecastOrderR3_Main.ROP), 3) AS ROP, "
        MySQLStr = MySQLStr & "ROUND(CONVERT(float, ISNULL(View_1.ROP, tbl_ForecastOrderR3_Main.ROP)), 3) AS ROP_OLD, "
        MySQLStr = MySQLStr & "ROUND(CONVERT(float, tbl_ForecastOrderR3_Main.InshuranceLVL), 3) AS InshuranceLVL, "
        MySQLStr = MySQLStr & "ROUND(CONVERT(float, ISNULL(View_1.InshuranceLVL, tbl_ForecastOrderR3_Main.InshuranceLVL)), 3) AS InshuranceLVL_OLD "
        MySQLStr = MySQLStr & "FROM tbl_ForecastOrderR3_Main WITH(NOLOCK) LEFT OUTER JOIN "
        MySQLStr = MySQLStr & "(SELECT tbl_ForecastOrderR3_Main_History.Code, "
        MySQLStr = MySQLStr & "tbl_ForecastOrderR3_Main_History.MGZ, "
        MySQLStr = MySQLStr & "tbl_ForecastOrderR3_Main_History.ROP, "
        MySQLStr = MySQLStr & "tbl_ForecastOrderR3_Main_History.InshuranceLVL "
        MySQLStr = MySQLStr & "FROM tbl_ForecastOrderR3_Main_History WITH(NOLOCK) INNER JOIN "
        MySQLStr = MySQLStr & "(SELECT Code, WarNo, MAX(Date) AS Expr1 "
        MySQLStr = MySQLStr & "FROM (SELECT tbl_ForecastOrderR3_Main_History_2.Code, "
        MySQLStr = MySQLStr & "tbl_ForecastOrderR3_Main_History_2.WarNo, "
        MySQLStr = MySQLStr & "tbl_ForecastOrderR3_Main_History_2.Date "
        MySQLStr = MySQLStr & "FROM tbl_ForecastOrderR3_Main_History AS tbl_ForecastOrderR3_Main_History_2 WITH(NOLOCK) LEFT OUTER JOIN "
        MySQLStr = MySQLStr & "(SELECT Code, WarNo, MAX(Date) AS Expr1 "
        MySQLStr = MySQLStr & "FROM  tbl_ForecastOrderR3_Main_History AS tbl_ForecastOrderR3_Main_History_1 WITH(NOLOCK) "
        MySQLStr = MySQLStr & "WHERE (WarNo = N'" & ComboBox1.SelectedValue & "') "
        MySQLStr = MySQLStr & "GROUP BY Code, WarNo) AS View_2 ON tbl_ForecastOrderR3_Main_History_2.Code = View_2.Code AND "
        MySQLStr = MySQLStr & "tbl_ForecastOrderR3_Main_History_2.WarNo = View_2.WarNo AND "
        MySQLStr = MySQLStr & "tbl_ForecastOrderR3_Main_History_2.Date = View_2.Expr1 "
        MySQLStr = MySQLStr & "WHERE (tbl_ForecastOrderR3_Main_History_2.WarNo = N'" & ComboBox1.SelectedValue & "') AND "
        MySQLStr = MySQLStr & "(View_2.Expr1 IS NULL)) AS View_3 "
        MySQLStr = MySQLStr & "GROUP BY Code, WarNo) AS View_4 ON "
        MySQLStr = MySQLStr & "tbl_ForecastOrderR3_Main_History.Code = View_4.Code AND "
        MySQLStr = MySQLStr & "tbl_ForecastOrderR3_Main_History.WarNo = View_4.WarNo AND "
        MySQLStr = MySQLStr & "tbl_ForecastOrderR3_Main_History.Date = View_4.Expr1) AS View_1 ON "
        MySQLStr = MySQLStr & "tbl_ForecastOrderR3_Main.Code = View_1.Code "
        MySQLStr = MySQLStr & "WHERE (tbl_ForecastOrderR3_Main.WarNo = N'" & ComboBox1.SelectedValue & "') AND "
        MySQLStr = MySQLStr & "(tbl_ForecastOrderR3_Main.WHass = - 1) AND "
        MySQLStr = MySQLStr & "(tbl_ForecastOrderR3_Main.Code NOT IN "
        MySQLStr = MySQLStr & "(SELECT Code "
        MySQLStr = MySQLStr & "FROM tbl_ForecastOrderR2_CustomMGZROPINS WITH(NOLOCK) "
        MySQLStr = MySQLStr & "WHERE (WH = N'" & ComboBox1.SelectedValue & "'))) "
        MySQLStr = MySQLStr & "ORDER BY tbl_ForecastOrderR3_Main.Code "

        InitMyConn(False)

        Try
            MyAdapter = New SqlClient.SqlDataAdapter(MySQLStr, Declarations.MyNETConnStr)
            MyAdapter.SelectCommand.CommandTimeout = 600
            MyAdapter.Fill(MyDs)
            DataGridView1.DataSource = MyDs.Tables(0)

            DataGridView1.Columns(0).HeaderText = "ID"
            DataGridView1.Columns(0).Width = 80
            DataGridView1.Columns(1).HeaderText = "�����"
            DataGridView1.Columns(1).Width = 200
            DataGridView1.Columns(2).HeaderText = "ABC"
            DataGridView1.Columns(2).Width = 50
            DataGridView1.Columns(3).HeaderText = "XYZ"
            DataGridView1.Columns(3).Width = 50
            DataGridView1.Columns(4).HeaderText = "LT"
            DataGridView1.Columns(4).Width = 50
            DataGridView1.Columns(5).HeaderText = "OI"
            DataGridView1.Columns(5).Width = 50
            DataGridView1.Columns(6).HeaderText = "���"
            DataGridView1.Columns(6).Width = 60
            DataGridView1.Columns(7).HeaderText = "��� ������"
            DataGridView1.Columns(7).Width = 60
            DataGridView1.Columns(8).HeaderText = "ROP"
            DataGridView1.Columns(8).Width = 60
            DataGridView1.Columns(9).HeaderText = "ROP ������"
            DataGridView1.Columns(9).Width = 60
            DataGridView1.Columns(10).HeaderText = "����� �������"
            DataGridView1.Columns(10).Width = 60
            DataGridView1.Columns(11).HeaderText = "����� ������� ������"
            DataGridView1.Columns(11).Width = 60

            DataGridView1.SelectionMode = DataGridViewSelectionMode.FullRowSelect

        Catch ex As Exception
            MsgBox(ex.ToString)
        End Try
    End Sub

    Private Sub BuildManualItemList()
        '/////////////////////////////////////////////////////////////////////////////////////
        '//
        '// ����� � DataGridView1 ������ ������� �� ������� ���, ROP � ��������� ����� ������������ �������
        '// 
        '/////////////////////////////////////////////////////////////////////////////////////
        Dim MySQLStr As String                        '������� ������
        Dim MyAdapter As SqlClient.SqlDataAdapter     '
        Dim MyDs As New DataSet                       '

        MySQLStr = "SELECT View_8.Code, "
        MySQLStr = MySQLStr & "View_9.Name, "
        MySQLStr = MySQLStr & "View_9.ABC, "
        MySQLStr = MySQLStr & "View_9.XYZ, "
        MySQLStr = MySQLStr & "View_9.LT, "
        MySQLStr = MySQLStr & "View_9.OI, "
        MySQLStr = MySQLStr & "ROUND(CONVERT(float, View_8.MGZ), 3) AS HMGZ, "
        MySQLStr = MySQLStr & "ROUND(CONVERT(float, ISNULL(View_7.MGZ, View_8.MGZ)), 3) AS HMGZ_OLD, "
        MySQLStr = MySQLStr & "View_5.MGZ, "
        MySQLStr = MySQLStr & "View_5.MGZ_OLD, "
        MySQLStr = MySQLStr & "ROUND(CONVERT(float, View_8.ROP), 3) AS HROP, "
        MySQLStr = MySQLStr & "ROUND(CONVERT(float, ISNULL(View_7.ROP, View_8.ROP)), 3) AS HROP_OLD, "
        MySQLStr = MySQLStr & "View_5.ROP, "
        MySQLStr = MySQLStr & "View_5.ROP_OLD, "
        MySQLStr = MySQLStr & "ROUND(CONVERT(float, View_8.IshuranceLVL), 3) AS HIshuranceLVL, "
        MySQLStr = MySQLStr & "ROUND(CONVERT(float, ISNULL(View_7.IshuranceLVL, View_8.IshuranceLVL)), 3) AS HIshuranceLVL_OLD, "
        MySQLStr = MySQLStr & "View_5.InshuranceLVL, "
        MySQLStr = MySQLStr & "View_5.InshuranceLVL_OLD "
        MySQLStr = MySQLStr & "FROM (SELECT Code, Name, ABC, XYZ, LT, OI "
        MySQLStr = MySQLStr & "FROM tbl_ForecastOrderR3_Main WITH(NOLOCK) "
        MySQLStr = MySQLStr & "WHERE (WarNo = N'" & ComboBox1.SelectedValue & "') "
        MySQLStr = MySQLStr & "AND (WHass = - 1)) AS View_9 RIGHT OUTER JOIN "
        MySQLStr = MySQLStr & "(SELECT Code, MGZ, ROP, IshuranceLVL "
        MySQLStr = MySQLStr & "FROM tbl_ForecastOrderR2_CustomMGZROPINS WITH(NOLOCK) "
        MySQLStr = MySQLStr & "WHERE (WH = N'" & ComboBox1.SelectedValue & "')) AS View_8 ON View_9.Code = View_8.Code LEFT OUTER JOIN "
        MySQLStr = MySQLStr & "(SELECT tbl_ForecastOrderR3_Main_1.Code, "
        MySQLStr = MySQLStr & "tbl_ForecastOrderR3_Main_1.Name, "
        MySQLStr = MySQLStr & "tbl_ForecastOrderR3_Main_1.ABC, "
        MySQLStr = MySQLStr & "tbl_ForecastOrderR3_Main_1.XYZ, "
        MySQLStr = MySQLStr & "tbl_ForecastOrderR3_Main_1.LT, "
        MySQLStr = MySQLStr & "tbl_ForecastOrderR3_Main_1.OI, "
        MySQLStr = MySQLStr & "ROUND(CONVERT(float, tbl_ForecastOrderR3_Main_1.MGZ), 3) AS MGZ, "
        MySQLStr = MySQLStr & "ROUND(CONVERT(float, ISNULL(View_1.MGZ, tbl_ForecastOrderR3_Main_1.MGZ)), 3) AS MGZ_OLD, "
        MySQLStr = MySQLStr & "ROUND(CONVERT(float, tbl_ForecastOrderR3_Main_1.ROP), 3) AS ROP, "
        MySQLStr = MySQLStr & "ROUND(CONVERT(float, ISNULL(View_1.ROP, tbl_ForecastOrderR3_Main_1.ROP)), 3) AS ROP_OLD, "
        MySQLStr = MySQLStr & "ROUND(CONVERT(float, tbl_ForecastOrderR3_Main_1.InshuranceLVL), 3) AS InshuranceLVL, "
        MySQLStr = MySQLStr & "ROUND(CONVERT(float, ISNULL(View_1.InshuranceLVL, tbl_ForecastOrderR3_Main_1.InshuranceLVL)), 3) AS InshuranceLVL_OLD "
        MySQLStr = MySQLStr & "FROM tbl_ForecastOrderR3_Main AS tbl_ForecastOrderR3_Main_1 WITH(NOLOCK) LEFT OUTER JOIN "
        MySQLStr = MySQLStr & "(SELECT tbl_ForecastOrderR3_Main_History.Code, tbl_ForecastOrderR3_Main_History.MGZ, "
        MySQLStr = MySQLStr & "tbl_ForecastOrderR3_Main_History.ROP, tbl_ForecastOrderR3_Main_History.InshuranceLVL "
        MySQLStr = MySQLStr & "FROM tbl_ForecastOrderR3_Main_History WITH(NOLOCK) INNER JOIN "
        MySQLStr = MySQLStr & "(SELECT Code, WarNo, MAX(Date) AS Expr1 "
        MySQLStr = MySQLStr & "FROM (SELECT tbl_ForecastOrderR3_Main_History_2.Code, tbl_ForecastOrderR3_Main_History_2.WarNo, "
        MySQLStr = MySQLStr & "tbl_ForecastOrderR3_Main_History_2.Date "
        MySQLStr = MySQLStr & "FROM tbl_ForecastOrderR3_Main_History AS tbl_ForecastOrderR3_Main_History_2 WITH(NOLOCK) LEFT OUTER JOIN "
        MySQLStr = MySQLStr & "(SELECT Code, WarNo, MAX(Date) AS Expr1 "
        MySQLStr = MySQLStr & "FROM tbl_ForecastOrderR3_Main_History AS tbl_ForecastOrderR3_Main_History_1 WITH(NOLOCK) "
        MySQLStr = MySQLStr & "WHERE (WarNo = N'" & ComboBox1.SelectedValue & "') "
        MySQLStr = MySQLStr & "GROUP BY Code, WarNo) AS View_2 ON tbl_ForecastOrderR3_Main_History_2.Code = View_2.Code AND "
        MySQLStr = MySQLStr & "tbl_ForecastOrderR3_Main_History_2.WarNo = View_2.WarNo AND "
        MySQLStr = MySQLStr & "tbl_ForecastOrderR3_Main_History_2.Date = View_2.Expr1 "
        MySQLStr = MySQLStr & "WHERE (tbl_ForecastOrderR3_Main_History_2.WarNo = N'" & ComboBox1.SelectedValue & "') "
        MySQLStr = MySQLStr & "AND (View_2.Expr1 IS NULL)) AS View_3 "
        MySQLStr = MySQLStr & "GROUP BY Code, WarNo) AS View_4 ON tbl_ForecastOrderR3_Main_History.Code = View_4.Code AND "
        MySQLStr = MySQLStr & "tbl_ForecastOrderR3_Main_History.WarNo = View_4.WarNo "
        MySQLStr = MySQLStr & "AND tbl_ForecastOrderR3_Main_History.Date = View_4.Expr1) AS View_1 ON "
        MySQLStr = MySQLStr & "tbl_ForecastOrderR3_Main_1.Code = View_1.Code "
        MySQLStr = MySQLStr & "WHERE (tbl_ForecastOrderR3_Main_1.WarNo = N'" & ComboBox1.SelectedValue & "') "
        MySQLStr = MySQLStr & "AND (tbl_ForecastOrderR3_Main_1.WHass = - 1)) AS View_5 ON "
        MySQLStr = MySQLStr & "View_8.Code = View_5.Code LEFT OUTER JOIN "
        MySQLStr = MySQLStr & "(SELECT tbl_ForecastOrderR2_CustomMGZROPINS_History.Code, tbl_ForecastOrderR2_CustomMGZROPINS_History.MGZ, "
        MySQLStr = MySQLStr & "tbl_ForecastOrderR2_CustomMGZROPINS_History.ROP, "
        MySQLStr = MySQLStr & "tbl_ForecastOrderR2_CustomMGZROPINS_History.IshuranceLVL "
        MySQLStr = MySQLStr & "FROM tbl_ForecastOrderR2_CustomMGZROPINS_History WITH(NOLOCK) INNER JOIN "
        MySQLStr = MySQLStr & "(SELECT Code, MAX(DateFrom) AS DateFrom "
        MySQLStr = MySQLStr & "FROM tbl_ForecastOrderR2_CustomMGZROPINS_History AS tbl_ForecastOrderR2_CustomMGZROPINS_History_1 WITH(NOLOCK) "
        MySQLStr = MySQLStr & "WHERE (DateTo <> Convert(datetime,'31/12/9999',103)) AND (WH = N'" & ComboBox1.SelectedValue & "') "
        MySQLStr = MySQLStr & "GROUP BY Code) AS View_6 ON tbl_ForecastOrderR2_CustomMGZROPINS_History.Code = View_6.Code AND "
        MySQLStr = MySQLStr & "tbl_ForecastOrderR2_CustomMGZROPINS_History.DateFrom = View_6.DateFrom "
        MySQLStr = MySQLStr & "WHERE (tbl_ForecastOrderR2_CustomMGZROPINS_History.WH = N'" & ComboBox1.SelectedValue & "') "
        MySQLStr = MySQLStr & "GROUP BY tbl_ForecastOrderR2_CustomMGZROPINS_History.Code, "
        MySQLStr = MySQLStr & "tbl_ForecastOrderR2_CustomMGZROPINS_History.MGZ, "
        MySQLStr = MySQLStr & "tbl_ForecastOrderR2_CustomMGZROPINS_History.ROP, "
        MySQLStr = MySQLStr & "tbl_ForecastOrderR2_CustomMGZROPINS_History.IshuranceLVL) AS View_7 ON "
        MySQLStr = MySQLStr & "View_8.Code = View_7.Code "
        MySQLStr = MySQLStr & "Order By View_8.Code "

        InitMyConn(False)

        Try
            MyAdapter = New SqlClient.SqlDataAdapter(MySQLStr, Declarations.MyNETConnStr)
            MyAdapter.SelectCommand.CommandTimeout = 600
            MyAdapter.Fill(MyDs)
            DataGridView2.DataSource = MyDs.Tables(0)

            DataGridView2.Columns(0).HeaderText = "ID"
            DataGridView2.Columns(0).Width = 80
            DataGridView2.Columns(1).HeaderText = "�����"
            DataGridView2.Columns(1).Width = 200
            DataGridView2.Columns(2).HeaderText = "ABC"
            DataGridView2.Columns(2).Width = 50
            DataGridView2.Columns(3).HeaderText = "XYZ"
            DataGridView2.Columns(3).Width = 50
            DataGridView2.Columns(4).HeaderText = "LT"
            DataGridView2.Columns(4).Width = 50
            DataGridView2.Columns(5).HeaderText = "OI"
            DataGridView2.Columns(5).Width = 50
            DataGridView2.Columns(6).HeaderText = "����. ���"
            DataGridView2.Columns(6).Width = 60
            DataGridView2.Columns(7).HeaderText = "����. ��� ������"
            DataGridView2.Columns(7).Width = 60
            DataGridView2.Columns(8).HeaderText = "���"
            DataGridView2.Columns(8).Width = 60
            DataGridView2.Columns(9).HeaderText = "��� ������"
            DataGridView2.Columns(9).Width = 60
            DataGridView2.Columns(10).HeaderText = "����. ROP"
            DataGridView2.Columns(10).Width = 60
            DataGridView2.Columns(11).HeaderText = "����. ROP ������"
            DataGridView2.Columns(11).Width = 60
            DataGridView2.Columns(12).HeaderText = "ROP"
            DataGridView2.Columns(12).Width = 60
            DataGridView2.Columns(13).HeaderText = "ROP ������"
            DataGridView2.Columns(13).Width = 60
            DataGridView2.Columns(14).HeaderText = "����. ����� �������"
            DataGridView2.Columns(14).Width = 60
            DataGridView2.Columns(15).HeaderText = "����. ����� ������� ������"
            DataGridView2.Columns(15).Width = 60
            DataGridView2.Columns(16).HeaderText = "����� �������"
            DataGridView2.Columns(16).Width = 60
            DataGridView2.Columns(17).HeaderText = "����� ������� ������"
            DataGridView2.Columns(17).Width = 60

            DataGridView2.SelectionMode = DataGridViewSelectionMode.FullRowSelect
        Catch ex As Exception
            MsgBox(ex.ToString)
        End Try
    End Sub

    Private Sub CheckButtons()
        '/////////////////////////////////////////////////////////////////////////////////////
        '//
        '// ����������� ��������� ������
        '// 
        '/////////////////////////////////////////////////////////////////////////////////////
        Dim MySQLStr As String                        '������� ������

        If DataGridView1.SelectedRows.Count = 0 Then
            Button2.Enabled = False
        Else
            Button2.Enabled = True
        End If
        If DataGridView2.SelectedRows.Count = 0 Then
            Button3.Enabled = False
            Button4.Enabled = False
        Else
            Button3.Enabled = True
            Button4.Enabled = True
        End If

        MySQLStr = "SELECT COUNT(IsNew) AS CC "
        MySQLStr = MySQLStr & "FROM tbl_WHCharacteristicsHistory WITH(NOLOCK) "
        MySQLStr = MySQLStr & "WHERE (IsNew = 1)"
        InitMyConn(False)
        InitMyRec(False, MySQLStr)
        If (Declarations.MyRec.Fields("CC").Value = 0) Then
            Button10.Enabled = False
        Else
            Button10.Enabled = True
        End If
        trycloseMyRec()
    End Sub

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        '/////////////////////////////////////////////////////////////////////////////////////
        '//
        '// ����� �� ���������
        '// 
        '/////////////////////////////////////////////////////////////////////////////////////

        Application.Exit()
    End Sub

    Private Sub DataGridView1_CellFormatting(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellFormattingEventArgs) Handles DataGridView1.CellFormatting
        '////////////////////////////////////////////////////////////////////////////////
        '//
        '// ��������� �������, � ������� ��� ������� (�����������) ����� ��� � 1.5 ���� � �������� �������
        '//
        '////////////////////////////////////////////////////////////////////////////////
        Dim row As DataGridViewRow

        row = DataGridView1.Rows(e.RowIndex)
        If (row.Cells(6).Value / row.Cells(7).Value) > 1.5 Or (row.Cells(7).Value / row.Cells(6).Value) > 1.5 _
            Or (row.Cells(8).Value / row.Cells(9).Value) > 1.5 Or (row.Cells(9).Value / row.Cells(8).Value) > 1.5 _
            Or (row.Cells(10).Value / row.Cells(11).Value) > 1.5 Or (row.Cells(11).Value / row.Cells(10).Value) > 1.5 _
            Or row.Cells(6).Value = 0 Or row.Cells(8).Value = 0 Or row.Cells(10).Value = 0 Then
            row.DefaultCellStyle.BackColor = Color.Yellow
        End If
    End Sub

    Private Sub ComboBox1_SelectedValueChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles ComboBox1.SelectedValueChanged
        '////////////////////////////////////////////////////////////////////////////////
        '//
        '// ����� ������ � ComboBox1 - ������������� ������
        '//
        '////////////////////////////////////////////////////////////////////////////////

        '---������ ������� �� ������� ���, ROP � ��������� ����� ��������� �������������
        BuildAutoItemList()

        '---������ ������� �� ������� ���, ROP � ��������� ����� ������������ �������
        BuildManualItemList()

        CheckButtons()
    End Sub

    Private Sub DataGridView2_CellFormatting(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellFormattingEventArgs) Handles DataGridView2.CellFormatting
        '////////////////////////////////////////////////////////////////////////////////
        '//
        '// ��������� �������, � ������� ��� ������� (�����������) ����� ��� � 1.5 ���� � �������� �������
        '//
        '////////////////////////////////////////////////////////////////////////////////
        Dim row As DataGridViewRow

        row = DataGridView2.Rows(e.RowIndex)
        If row.Cells(8).Value = 0 Or row.Cells(12).Value = 0 Or row.Cells(16).Value = 0 Then
            row.DefaultCellStyle.BackColor = Color.Yellow
        End If

    End Sub

    Private Sub Button16_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button16.Click
        '////////////////////////////////////////////////////////////////////////////////
        '//
        '// ����� ������� ������ �� ���� � ������ ������
        '//
        '////////////////////////////////////////////////////////////////////////////////

        If Trim(TextBox2.Text) = "" Then
        Else
            Windows.Forms.Cursor.Current = Cursors.WaitCursor
            For i As Integer = 0 To DataGridView1.Rows.Count - 1
                If StrComp(UCase(Trim(TextBox2.Text)), Microsoft.VisualBasic.Left(UCase(Trim(DataGridView1.Item(0, i).Value.ToString)), Len(UCase(Trim(TextBox2.Text)))), 1) = 0 Then
                    DataGridView1.CurrentCell = DataGridView1.Item(0, i)
                    Windows.Forms.Cursor.Current = Cursors.Default
                    Exit Sub
                End If
            Next i
            Exit Sub
            Windows.Forms.Cursor.Current = Cursors.Default
        End If
    End Sub

    Private Sub Button9_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button9.Click
        '////////////////////////////////////////////////////////////////////////////////
        '//
        '// ����� ������� ����������� �� �������� ������
        '//
        '////////////////////////////////////////////////////////////////////////////////

        If Trim(TextBox2.Text) = "" And Trim(TextBox3.Text) = "" Then
            MsgBox("���������� ������ �������� ������", MsgBoxStyle.OkOnly, "��������!")
            TextBox2.Select()
        Else
            Windows.Forms.Cursor.Current = Cursors.WaitCursor
            For i As Integer = 0 To DataGridView1.Rows.Count - 1
                If InStr(UCase(Trim(DataGridView1.Item(0, i).Value.ToString)), UCase(Trim(TextBox2.Text))) <> 0 And InStr(UCase(Trim(DataGridView1.Item(0, i).Value.ToString)), UCase(Trim(TextBox3.Text))) <> 0 _
                    Or InStr(UCase(Trim(DataGridView1.Item(1, i).Value.ToString)), UCase(Trim(TextBox2.Text))) <> 0 And InStr(UCase(Trim(DataGridView1.Item(1, i).Value.ToString)), UCase(Trim(TextBox3.Text))) <> 0 Then
                    DataGridView1.CurrentCell = DataGridView1.Item(0, i)
                    Windows.Forms.Cursor.Current = Cursors.Default
                    Exit Sub
                End If
            Next
            Windows.Forms.Cursor.Current = Cursors.Default
        End If
    End Sub

    Private Sub Button8_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button8.Click
        '////////////////////////////////////////////////////////////////////////////////
        '//
        '// ����� ���������� ����������� �� �������� ������
        '//
        '////////////////////////////////////////////////////////////////////////////////
        Dim MyRez As Object

        If Trim(TextBox2.Text) = "" And Trim(TextBox3.Text) = "" Then
            MsgBox("���������� ������ �������� ������", MsgBoxStyle.OkOnly, "��������!")
            TextBox2.Select()
        Else
            Windows.Forms.Cursor.Current = Cursors.WaitCursor
            For i As Integer = DataGridView1.CurrentCellAddress.Y + 1 To DataGridView1.Rows.Count
                If i = DataGridView1.Rows.Count Then
                    MyRez = MsgBox("����� ����� �� ����� ������. ������ �������?", MsgBoxStyle.YesNo, "��������!")
                    If MyRez = 6 Then
                        i = 0
                    Else
                        Windows.Forms.Cursor.Current = Cursors.Default
                        Exit Sub
                    End If
                End If
                If DataGridView1.Rows.Count = 0 Then
                Else
                    If InStr(UCase(Trim(DataGridView1.Item(0, i).Value.ToString)), UCase(Trim(TextBox2.Text))) <> 0 And InStr(UCase(Trim(DataGridView1.Item(0, i).Value.ToString)), UCase(Trim(TextBox3.Text))) <> 0 _
                        Or InStr(UCase(Trim(DataGridView1.Item(1, i).Value.ToString)), UCase(Trim(TextBox2.Text))) <> 0 And InStr(UCase(Trim(DataGridView1.Item(1, i).Value.ToString)), UCase(Trim(TextBox3.Text))) <> 0 Then
                        DataGridView1.CurrentCell = DataGridView1.Item(0, i)
                        Windows.Forms.Cursor.Current = Cursors.Default
                        Exit Sub
                    End If
                End If
            Next i
            Windows.Forms.Cursor.Current = Cursors.Default
        End If
    End Sub

    Private Sub Button7_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button7.Click
        '////////////////////////////////////////////////////////////////////////////////
        '//
        '// ����� ���� ���������� �� �������� ������� � ��������� ����
        '//
        '////////////////////////////////////////////////////////////////////////////////

        If Trim(TextBox2.Text) = "" And Trim(TextBox3.Text) = "" Then
            MsgBox("���������� ������ �������� ������", MsgBoxStyle.OkOnly, "��������!")
            TextBox2.Select()
        Else
            MyItemSelectList = New ItemSelectList
            MyItemSelectList.ShowDialog()
        End If
    End Sub

    Private Sub Button2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button2.Click
        '////////////////////////////////////////////////////////////////////////////////
        '//
        '// ���������� ������ �������� ���, ROP � ���������� ������
        '//
        '////////////////////////////////////////////////////////////////////////////////

        AddItemCustomInfo()
    End Sub

    Private Sub Button3_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button3.Click
        '////////////////////////////////////////////////////////////////////////////////
        '//
        '// �������� (��������) ������ �������� ���, ROP � ���������� ������
        '//
        '////////////////////////////////////////////////////////////////////////////////

        RemoveItemCustomInfo()
    End Sub

    Private Sub Button4_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button4.Click
        '////////////////////////////////////////////////////////////////////////////////
        '//
        '// �������������� ������ �������� ���, ROP � ���������� ������
        '//
        '////////////////////////////////////////////////////////////////////////////////

        EditItemCustomInfo()
    End Sub

    Private Sub Button5_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button5.Click
        '////////////////////////////////////////////////////////////////////////////////
        '//
        '// �������� � Excel �������� ���, ROP � ���������� ������
        '//
        '////////////////////////////////////////////////////////////////////////////////

        Me.Cursor = Cursors.WaitCursor
        UploadItemInfo()
        Me.Cursor = Cursors.Default
    End Sub

    Private Sub AddItemCustomInfo()
        '////////////////////////////////////////////////////////////////////////////////////////
        '//
        '// ��������� ���������� ������ �������� ���, ROP � ���������� ������
        '//
        '////////////////////////////////////////////////////////////////////////////////////////
        Dim MySQLStr As String                        '������� ������

        Declarations.MySuccess = False
        MyAddCustom = New AddCustom
        MyAddCustom.ShowDialog()
        If Declarations.MySuccess = False Then
            Exit Sub
        Else '---���������� ������ �������� ���, ROP � ���������� ������
            '---�������� ������ �������� � �������
            MySQLStr = "Update tbl_ForecastOrderR2_CustomMGZROPINS_History "
            MySQLStr = MySQLStr & "SET DateTo = GETDATE() "
            MySQLStr = MySQLStr & "WHERE (DateTo = Convert(datetime,'31/12/9999',103)) "
            MySQLStr = MySQLStr & "AND (WH = N'" & ComboBox1.SelectedValue & "') "
            MySQLStr = MySQLStr & "AND (Code = N'" & Trim(DataGridView1.SelectedRows.Item(0).Cells(0).Value.ToString()) & "') "
            InitMyConn(False)
            Declarations.MyConn.Execute(MySQLStr)

            '---��������� ����� ��������
            '---� ������� �������
            MySQLStr = "INSERT INTO tbl_ForecastOrderR2_CustomMGZROPINS "
            MySQLStr = MySQLStr & "(ID, Code, WH, MGZ, ROP, IshuranceLVL) "
            MySQLStr = MySQLStr & "VALUES (NEWID(), "
            MySQLStr = MySQLStr & "N'" & Trim(DataGridView1.SelectedRows.Item(0).Cells(0).Value.ToString()) & "', "
            MySQLStr = MySQLStr & "N'" & ComboBox1.SelectedValue & "', "
            MySQLStr = MySQLStr & Replace(CStr(Declarations.MyMGZ), ",", ".") & ", "
            MySQLStr = MySQLStr & Replace(CStr(Declarations.MyROP), ",", ".") & ", "
            MySQLStr = MySQLStr & Replace(CStr(Declarations.MyInsuranceLVL), ",", ".") & ") "
            InitMyConn(False)
            Declarations.MyConn.Execute(MySQLStr)

            '---� �������
            MySQLStr = "INSERT INTO tbl_ForecastOrderR2_CustomMGZROPINS_History "
            MySQLStr = MySQLStr & "(ID, Code, WH, MGZ, ROP, IshuranceLVL, UserID, DateFrom, DateTo) "
            MySQLStr = MySQLStr & "VALUES (NEWID(), "
            MySQLStr = MySQLStr & "N'" & Trim(DataGridView1.SelectedRows.Item(0).Cells(0).Value.ToString()) & "', "
            MySQLStr = MySQLStr & "N'" & ComboBox1.SelectedValue & "', "
            MySQLStr = MySQLStr & Replace(CStr(Declarations.MyMGZ), ",", ".") & ", "
            MySQLStr = MySQLStr & Replace(CStr(Declarations.MyROP), ",", ".") & ", "
            MySQLStr = MySQLStr & Replace(CStr(Declarations.MyInsuranceLVL), ",", ".") & ", "
            MySQLStr = MySQLStr & " N'" & Declarations.UserCode & "', "
            MySQLStr = MySQLStr & "GETDATE(), Convert(datetime,'31/12/9999',103)) "
            InitMyConn(False)
            Declarations.MyConn.Execute(MySQLStr)

            '---������ ������� �� ������� ���, ROP � ��������� ����� ��������� �������������
            BuildAutoItemList()
            '---������ ������� �� ������� ���, ROP � ��������� ����� ������������ �������
            BuildManualItemList()
            CheckButtons()
        End If
    End Sub

    Private Sub RemoveItemCustomInfo()
        '////////////////////////////////////////////////////////////////////////////////////////
        '//
        '// ��������� �������� (��������) ������ �������� ���, ROP � ���������� ������
        '//
        '////////////////////////////////////////////////////////////////////////////////////////
        Dim MySQLStr As String                        '������� ������

        '---�������� ������ �������� � �������
        MySQLStr = "Update tbl_ForecastOrderR2_CustomMGZROPINS_History "
        MySQLStr = MySQLStr & "SET DateTo = GETDATE() "
        MySQLStr = MySQLStr & "WHERE (DateTo = Convert(datetime,'31/12/9999',103)) "
        MySQLStr = MySQLStr & "AND (WH = N'" & ComboBox1.SelectedValue & "') "
        MySQLStr = MySQLStr & "AND (Code = N'" & Trim(DataGridView2.SelectedRows.Item(0).Cells(0).Value.ToString()) & "') "
        InitMyConn(False)
        Declarations.MyConn.Execute(MySQLStr)

        '---�������� �������� �� ������� �������
        MySQLStr = "DELETE FROM tbl_ForecastOrderR2_CustomMGZROPINS "
        MySQLStr = MySQLStr & "WHERE (Code = N'" & Trim(DataGridView2.SelectedRows.Item(0).Cells(0).Value.ToString()) & "') "
        MySQLStr = MySQLStr & "AND (WH = N'" & ComboBox1.SelectedValue & "')"
        InitMyConn(False)
        Declarations.MyConn.Execute(MySQLStr)

        '---������ ������� �� ������� ���, ROP � ��������� ����� ��������� �������������
        BuildAutoItemList()
        '---������ ������� �� ������� ���, ROP � ��������� ����� ������������ �������
        BuildManualItemList()
        CheckButtons()
    End Sub

    Private Sub EditItemCustomInfo()
        '////////////////////////////////////////////////////////////////////////////////////////
        '//
        '// ��������� �������������� ������ �������� ���, ROP � ���������� ������
        '//
        '////////////////////////////////////////////////////////////////////////////////////////
        Dim MySQLStr As String                        '������� ������

        Declarations.MySuccess = False
        MyEditCustom = New EditCustom
        MyEditCustom.ShowDialog()
        If Declarations.MySuccess = False Then
            Exit Sub
        Else '---��������� ������ �������� ���, ROP � ���������� ������
            '---�������� ������ �������� � �������
            MySQLStr = "Update tbl_ForecastOrderR2_CustomMGZROPINS_History "
            MySQLStr = MySQLStr & "SET DateTo = GETDATE() "
            MySQLStr = MySQLStr & "WHERE (DateTo = Convert(datetime,'31/12/9999',103)) "
            MySQLStr = MySQLStr & "AND (WH = N'" & ComboBox1.SelectedValue & "') "
            MySQLStr = MySQLStr & "AND (Code = N'" & Trim(DataGridView2.SelectedRows.Item(0).Cells(0).Value.ToString()) & "') "
            InitMyConn(False)
            Declarations.MyConn.Execute(MySQLStr)

            '---���������� ��������
            '---� ������� �������
            MySQLStr = "UPDATE tbl_ForecastOrderR2_CustomMGZROPINS "
            MySQLStr = MySQLStr & "SET MGZ = " & Replace(CStr(Declarations.MyMGZ), ",", ".") & ", "
            MySQLStr = MySQLStr & "ROP = " & Replace(CStr(Declarations.MyROP), ",", ".") & ", "
            MySQLStr = MySQLStr & "IshuranceLVL = " & Replace(CStr(Declarations.MyInsuranceLVL), ",", ".") & " "
            MySQLStr = MySQLStr & "WHERE (Code = N'" & Trim(DataGridView2.SelectedRows.Item(0).Cells(0).Value.ToString()) & "') "
            MySQLStr = MySQLStr & "AND (WH = N'" & ComboBox1.SelectedValue & "') "
            InitMyConn(False)
            Declarations.MyConn.Execute(MySQLStr)

            '---� �������
            MySQLStr = "INSERT INTO tbl_ForecastOrderR2_CustomMGZROPINS_History "
            MySQLStr = MySQLStr & "(ID, Code, WH, MGZ, ROP, IshuranceLVL, UserID, DateFrom, DateTo) "
            MySQLStr = MySQLStr & "VALUES (NEWID(), "
            MySQLStr = MySQLStr & "N'" & Trim(DataGridView2.SelectedRows.Item(0).Cells(0).Value.ToString()) & "', "
            MySQLStr = MySQLStr & "N'" & ComboBox1.SelectedValue & "', "
            MySQLStr = MySQLStr & Replace(CStr(Declarations.MyMGZ), ",", ".") & ", "
            MySQLStr = MySQLStr & Replace(CStr(Declarations.MyROP), ",", ".") & ", "
            MySQLStr = MySQLStr & Replace(CStr(Declarations.MyInsuranceLVL), ",", ".") & ", "
            MySQLStr = MySQLStr & " N'" & Declarations.UserCode & "', "
            MySQLStr = MySQLStr & "GETDATE(), Convert(datetime,'31/12/9999',103)) "
            InitMyConn(False)
            Declarations.MyConn.Execute(MySQLStr)

            '---������ ������� �� ������� ���, ROP � ��������� ����� ��������� �������������
            BuildAutoItemList()
            '---������ ������� �� ������� ���, ROP � ��������� ����� ������������ �������
            BuildManualItemList()
            CheckButtons()
        End If
    End Sub

    Private Sub UploadItemInfo()
        '////////////////////////////////////////////////////////////////////////////////////////
        '//
        '// ��������� �������� � Excel �������� ���, ROP � ���������� ������
        '//
        '////////////////////////////////////////////////////////////////////////////////////////
        Dim WHList(,) As String
        Dim MySQLStr As String
        Dim i As Integer
        Dim StrNum As Double      '����� ������
        Dim MyObj As Object       'Excel
        Dim MyWRKBook As Object   '�����

        MySQLStr = "SELECT SC23001, SC23002 "
        MySQLStr = MySQLStr & "FROM SC230300 WITH(NOLOCK) "
        MySQLStr = MySQLStr & "WHERE (SC23006 = N'1') "
        MySQLStr = MySQLStr & "ORDER BY SC23001 "
        InitMyConn(False)
        InitMyRec(False, MySQLStr)
        If (Declarations.MyRec.EOF = True And Declarations.MyRec.BOF = True) Then
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

        UploadCommonHeader(MyWRKBook)
        StrNum = 5
        For i = 0 To WHList.GetUpperBound(1)
            StrNum = UploadWHHeader(MyWRKBook, WHList(0, i), WHList(1, i), StrNum)
            StrNum = UploadWHRows(MyWRKBook, WHList(0, i), StrNum)
        Next

        MyWRKBook.ActiveSheet.Range("A1").Select()
        MyObj.Application.Visible = True
        MyObj = Nothing
    End Sub

    Private Function UploadCommonHeader(ByVal MyWRKBook As Object)
        '////////////////////////////////////////////////////////////////////////////////////////
        '//
        '// ��������� �������� � Excel ������ ��������� 
        '//
        '////////////////////////////////////////////////////////////////////////////////////////

        MyWRKBook.ActiveSheet.Range("B1") = "���������� � ���, ROP � ������ ���������� ������ "
        MyWRKBook.ActiveSheet.Range("B2") = "���������� ������������ �� ������� �� " & Now
        MyWRKBook.ActiveSheet.Range("B1:B2").Select()
        MyWRKBook.ActiveSheet.Range("B1:B2").Font.Bold = True

        '--- � ������� �����
        MyWRKBook.ActiveSheet.Columns("A:A").ColumnWidth = 10
        MyWRKBook.ActiveSheet.Columns("B:B").ColumnWidth = 16
        MyWRKBook.ActiveSheet.Columns("C:C").ColumnWidth = 30
        MyWRKBook.ActiveSheet.Columns("D:D").ColumnWidth = 10
        MyWRKBook.ActiveSheet.Columns("E:E").ColumnWidth = 10
        MyWRKBook.ActiveSheet.Columns("F:F").ColumnWidth = 10
        MyWRKBook.ActiveSheet.Columns("G:G").ColumnWidth = 10
        MyWRKBook.ActiveSheet.Columns("H:H").ColumnWidth = 10
        MyWRKBook.ActiveSheet.Columns("I:I").ColumnWidth = 10
        MyWRKBook.ActiveSheet.Columns("J:J").ColumnWidth = 10
        MyWRKBook.ActiveSheet.Columns("K:K").ColumnWidth = 10
        MyWRKBook.ActiveSheet.Columns("L:L").ColumnWidth = 10
        MyWRKBook.ActiveSheet.Columns("M:M").ColumnWidth = 10
        MyWRKBook.ActiveSheet.Columns("N:N").ColumnWidth = 10
        MyWRKBook.ActiveSheet.Columns("O:O").ColumnWidth = 10
        MyWRKBook.ActiveSheet.Columns("P:P").ColumnWidth = 10
        MyWRKBook.ActiveSheet.Columns("Q:Q").ColumnWidth = 10
        MyWRKBook.ActiveSheet.Columns("R:R").ColumnWidth = 10
        MyWRKBook.ActiveSheet.Columns("S:S").ColumnWidth = 10
    End Function

    Private Function UploadWHHeader(ByVal MyWRKBook As Object, ByVal WHCode As String, ByVal WHName As String, ByVal StrNum As Double) As Double
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

        '---� ��������� ��� �����
        MyWRKBook.ActiveSheet.Range("B" & StrNum) = "ID"
        MyWRKBook.ActiveSheet.Range("C" & StrNum) = "�������� ������"
        MyWRKBook.ActiveSheet.Range("D" & StrNum) = "ABC"
        MyWRKBook.ActiveSheet.Range("E" & StrNum) = "XYZ"
        MyWRKBook.ActiveSheet.Range("F" & StrNum) = "����� ��������"
        MyWRKBook.ActiveSheet.Range("G" & StrNum) = "����� ����� ��������"
        MyWRKBook.ActiveSheet.Range("H" & StrNum) = "������ ���"
        MyWRKBook.ActiveSheet.Range("I" & StrNum) = "���"
        MyWRKBook.ActiveSheet.Range("J" & StrNum) = "������ ROP"
        MyWRKBook.ActiveSheet.Range("K" & StrNum) = "ROP"
        MyWRKBook.ActiveSheet.Range("L" & StrNum) = "������ ����� �������"
        MyWRKBook.ActiveSheet.Range("M" & StrNum) = "����� �������"
        MyWRKBook.ActiveSheet.Range("N" & StrNum) = "������ ���� ���"
        MyWRKBook.ActiveSheet.Range("O" & StrNum) = "���� ���"
        MyWRKBook.ActiveSheet.Range("P" & StrNum) = "������ ���� ROP"
        MyWRKBook.ActiveSheet.Range("Q" & StrNum) = "���� ROP"
        MyWRKBook.ActiveSheet.Range("R" & StrNum) = "������ ���� ����� �������"
        MyWRKBook.ActiveSheet.Range("S" & StrNum) = "���� ����� �������"

        MyWRKBook.ActiveSheet.Range("B" & StrNum & ":S" & StrNum).Select()
        MyWRKBook.ActiveSheet.Range("B" & StrNum & ":S" & StrNum).HorizontalAlignment = -4108
        MyWRKBook.ActiveSheet.Range("B" & StrNum & ":S" & StrNum).VerticalAlignment = -4108
        MyWRKBook.ActiveSheet.Range("B" & StrNum & ":S" & StrNum).WrapText = True
        MyWRKBook.ActiveSheet.Range("B" & StrNum & ":S" & StrNum).Borders(5).LineStyle = -4142
        MyWRKBook.ActiveSheet.Range("B" & StrNum & ":S" & StrNum).Borders(6).LineStyle = -4142
        With MyWRKBook.ActiveSheet.Range("B" & StrNum & ":S" & StrNum).Borders(7)
            .LineStyle = 1
            .Weight = 4
            .ColorIndex = -4105
        End With
        With MyWRKBook.ActiveSheet.Range("B" & StrNum & ":S" & StrNum).Borders(8)
            .LineStyle = 1
            .Weight = 4
            .ColorIndex = -4105
        End With
        With MyWRKBook.ActiveSheet.Range("B" & StrNum & ":S" & StrNum).Borders(9)
            .LineStyle = 1
            .Weight = 4
            .ColorIndex = -4105
        End With
        With MyWRKBook.ActiveSheet.Range("B" & StrNum & ":S" & StrNum).Borders(10)
            .LineStyle = 1
            .Weight = 4
            .ColorIndex = -4105
        End With
        With MyWRKBook.ActiveSheet.Range("B" & StrNum & ":S" & StrNum).Borders(11)
            .LineStyle = 1
            .Weight = 4
            .ColorIndex = -4105
        End With
        With MyWRKBook.ActiveSheet.Range("B" & StrNum & ":S" & StrNum).Interior
            .ColorIndex = 35
            .Pattern = 1
            .PatternColorIndex = -4105
        End With

        StrNum = StrNum + 1
        Return StrNum
    End Function

    Private Function UploadWHRows(ByVal MyWRKBook As Object, ByVal WHCode As String, ByVal StrNum As Double) As Double
        '////////////////////////////////////////////////////////////////////////////////////////
        '//
        '// ��������� �������� � Excel ����� �� ������ ������ 
        '//
        '////////////////////////////////////////////////////////////////////////////////////////
        Dim MySQLStr As String
        Dim aa As New System.Globalization.NumberFormatInfo
        Dim MySep As String
        Dim MyDig As String

        MySep = aa.CurrentInfo.NumberGroupSeparator
        MyDig = aa.CurrentInfo.NumberDecimalSeparator

        'MySQLStr = "SELECT View_1.Code, "
        'MySQLStr = MySQLStr & "View_1.Name, "
        'MySQLStr = MySQLStr & "View_1.ABC, "
        'MySQLStr = MySQLStr & "View_1.XYZ, "
        'MySQLStr = MySQLStr & "View_1.LT, "
        'MySQLStr = MySQLStr & "View_1.OI, "
        'MySQLStr = MySQLStr & "View_1.MGZ, "
        'MySQLStr = MySQLStr & "View_1.ROP, "
        'MySQLStr = MySQLStr & "View_1.InshuranceLVL, "
        'MySQLStr = MySQLStr & "ISNULL(View_2.MGZ, 0) AS CustMGZ, "
        'MySQLStr = MySQLStr & "ISNULL(View_2.ROP, 0) AS CustROP, "
        'MySQLStr = MySQLStr & "ISNULL(View_2.IshuranceLVL, 0) AS CustInshuranceLVL, "
        'MySQLStr = MySQLStr & "ISNULL(View_2.IsCustom, 0) AS IsCustom "
        'MySQLStr = MySQLStr & "FROM (SELECT Code, "
        'MySQLStr = MySQLStr & "Name, "
        'MySQLStr = MySQLStr & "ABC, "
        'MySQLStr = MySQLStr & "XYZ, "
        'MySQLStr = MySQLStr & "LT, "
        'MySQLStr = MySQLStr & "OI, "
        'MySQLStr = MySQLStr & "ROUND(CONVERT(float, MGZ), 3) AS MGZ, "
        'MySQLStr = MySQLStr & "ROUND(CONVERT(float, ROP), 3) AS ROP, "
        'MySQLStr = MySQLStr & "ROUND(CONVERT(float, InshuranceLVL), 3) AS InshuranceLVL "
        'MySQLStr = MySQLStr & "FROM tbl_ForecastOrderR2_Main "
        'MySQLStr = MySQLStr & "WHERE (WarNo = N'" & WHCode & "') "
        'MySQLStr = MySQLStr & "AND (WHass = - 1)) AS View_1 LEFT OUTER JOIN "
        'MySQLStr = MySQLStr & "(SELECT Code, "
        'MySQLStr = MySQLStr & "ROUND(CONVERT(float, MGZ), 3) AS MGZ, "
        'MySQLStr = MySQLStr & "ROUND(CONVERT(float, ROP), 3) AS ROP, "
        'MySQLStr = MySQLStr & "ROUND(Convert(float, IshuranceLVL), 3) AS IshuranceLVL, "
        'MySQLStr = MySQLStr & "1 AS IsCustom "
        'MySQLStr = MySQLStr & "FROM tbl_ForecastOrderR2_CustomMGZROPINS "
        'MySQLStr = MySQLStr & "WHERE (WH = N'" & WHCode & "')) AS View_2 ON View_1.Code = View_2.Code "
        'MySQLStr = MySQLStr & "ORDER BY View_1.Code "

        MySQLStr = "SELECT View_2_1.Code, View_2_1.Name, View_2_1.ABC, View_2_1.XYZ, View_2_1.LT, View_2_1.OI, View_2_1.MGZ_OLD, "
        MySQLStr = MySQLStr & "View_2_1.MGZ, View_2_1.ROP_OLD, View_2_1.ROP, View_2_1.InshuranceLVL_OLD, View_2_1.InshuranceLVL, "
        MySQLStr = MySQLStr & "View_3_2.HMGZ_OLD, View_3_2.HMGZ, View_3_2.HROP_OLD, View_3_2.HROP, "
        MySQLStr = MySQLStr & "View_3_2.HIshuranceLVL_OLD, View_3_2.HIshuranceLVL, ISNULL(View_3_2.IsCustom, 0) AS IsCustom "
        MySQLStr = MySQLStr & "FROM (SELECT TOP (100) PERCENT tbl_ForecastOrderR3_Main.Code, tbl_ForecastOrderR3_Main.Name, "
        MySQLStr = MySQLStr & "tbl_ForecastOrderR3_Main.ABC, tbl_ForecastOrderR3_Main.XYZ, tbl_ForecastOrderR3_Main.LT, "
        MySQLStr = MySQLStr & "tbl_ForecastOrderR3_Main.OI, ROUND(CONVERT(float, tbl_ForecastOrderR3_Main.MGZ), 3) AS MGZ, "
        MySQLStr = MySQLStr & "ROUND(CONVERT(float, ISNULL(View_1.MGZ, tbl_ForecastOrderR3_Main.MGZ)), 3) AS MGZ_OLD, "
        MySQLStr = MySQLStr & "ROUND(CONVERT(float, tbl_ForecastOrderR3_Main.ROP), 3) AS ROP, "
        MySQLStr = MySQLStr & "ROUND(CONVERT(float, ISNULL(View_1.ROP, tbl_ForecastOrderR3_Main.ROP)), 3) AS ROP_OLD, "
        MySQLStr = MySQLStr & "ROUND(CONVERT(float, tbl_ForecastOrderR3_Main.InshuranceLVL), 3) AS InshuranceLVL, "
        MySQLStr = MySQLStr & "ROUND(CONVERT(float, ISNULL(View_1.InshuranceLVL, tbl_ForecastOrderR3_Main.InshuranceLVL)), 3) AS InshuranceLVL_OLD "
        MySQLStr = MySQLStr & "FROM tbl_ForecastOrderR3_Main WITH(NOLOCK) LEFT OUTER JOIN "
        MySQLStr = MySQLStr & "(SELECT tbl_ForecastOrderR3_Main_History.Code, tbl_ForecastOrderR3_Main_History.MGZ, "
        MySQLStr = MySQLStr & "tbl_ForecastOrderR3_Main_History.ROP, tbl_ForecastOrderR3_Main_History.InshuranceLVL "
        MySQLStr = MySQLStr & "FROM tbl_ForecastOrderR3_Main_History WITH(NOLOCK) INNER JOIN "
        MySQLStr = MySQLStr & "(SELECT Code, WarNo, MAX(Date) AS Expr1 "
        MySQLStr = MySQLStr & "FROM (SELECT tbl_ForecastOrderR3_Main_History_2.Code, tbl_ForecastOrderR3_Main_History_2.WarNo, "
        MySQLStr = MySQLStr & "tbl_ForecastOrderR3_Main_History_2.Date "
        MySQLStr = MySQLStr & "FROM tbl_ForecastOrderR3_Main_History AS tbl_ForecastOrderR3_Main_History_2 WITH(NOLOCK) LEFT OUTER JOIN "
        MySQLStr = MySQLStr & "(SELECT Code, WarNo, MAX(Date) AS Expr1 "
        MySQLStr = MySQLStr & "FROM tbl_ForecastOrderR3_Main_History AS tbl_ForecastOrderR3_Main_History_1 WITH(NOLOCK) "
        MySQLStr = MySQLStr & "WHERE (WarNo = N'" & WHCode & "') "
        MySQLStr = MySQLStr & "GROUP BY Code, WarNo) AS View_2_1_1 ON tbl_ForecastOrderR3_Main_History_2.Code = View_2_1_1.Code AND "
        MySQLStr = MySQLStr & "tbl_ForecastOrderR3_Main_History_2.WarNo = View_2_1_1.WarNo AND "
        MySQLStr = MySQLStr & "tbl_ForecastOrderR3_Main_History_2.Date = View_2_1_1.Expr1 "
        MySQLStr = MySQLStr & "WHERE (tbl_ForecastOrderR3_Main_History_2.WarNo = N'" & WHCode & "') AND (View_2_1_1.Expr1 IS NULL)) AS View_3 "
        MySQLStr = MySQLStr & "GROUP BY Code, WarNo) AS View_4 ON tbl_ForecastOrderR3_Main_History.Code = View_4.Code AND "
        MySQLStr = MySQLStr & "tbl_ForecastOrderR3_Main_History.WarNo = View_4.WarNo AND tbl_ForecastOrderR3_Main_History.Date = View_4.Expr1) AS View_1 ON "
        MySQLStr = MySQLStr & "tbl_ForecastOrderR3_Main.Code = View_1.Code "
        MySQLStr = MySQLStr & "WHERE (tbl_ForecastOrderR3_Main.WarNo = N'" & WHCode & "') AND (tbl_ForecastOrderR3_Main.WHass = - 1) "
        MySQLStr = MySQLStr & "ORDER BY tbl_ForecastOrderR3_Main.Code) AS View_2_1 LEFT OUTER JOIN "
        MySQLStr = MySQLStr & "(SELECT TOP (100) PERCENT View_8.Code, View_9.Name, View_9.ABC, View_9.XYZ, View_9.LT, View_9.OI, "
        MySQLStr = MySQLStr & "ROUND(CONVERT(float, View_8.MGZ), 3) AS HMGZ, "
        MySQLStr = MySQLStr & "ROUND(CONVERT(float, ISNULL(View_7.MGZ, View_8.MGZ)), 3) AS HMGZ_OLD, View_5.MGZ, View_5.MGZ_OLD, "
        MySQLStr = MySQLStr & "ROUND(Convert(float, View_8.ROP), 3)  AS HROP, ROUND(CONVERT(float, ISNULL(View_7.ROP, View_8.ROP)), 3) AS HROP_OLD, "
        MySQLStr = MySQLStr & "View_5.ROP, View_5.ROP_OLD, ROUND(CONVERT(float, View_8.IshuranceLVL), 3) AS HIshuranceLVL, "
        MySQLStr = MySQLStr & "ROUND(CONVERT(float, ISNULL(View_7.IshuranceLVL, View_8.IshuranceLVL)), 3) AS HIshuranceLVL_OLD, "
        MySQLStr = MySQLStr & "View_5.InshuranceLVL, View_5.InshuranceLVL_OLD, 1 AS IsCustom "
        MySQLStr = MySQLStr & "FROM (SELECT Code, Name, ABC, XYZ, LT, OI "
        MySQLStr = MySQLStr & "FROM tbl_ForecastOrderR3_Main AS tbl_ForecastOrderR3_Main_2 WITH(NOLOCK) "
        MySQLStr = MySQLStr & "WHERE (WarNo = N'" & WHCode & "') AND (WHass = - 1)) AS View_9 RIGHT OUTER JOIN "
        MySQLStr = MySQLStr & "(SELECT Code, MGZ, ROP, IshuranceLVL "
        MySQLStr = MySQLStr & "FROM tbl_ForecastOrderR2_CustomMGZROPINS WITH(NOLOCK) "
        MySQLStr = MySQLStr & "WHERE (WH = N'" & WHCode & "')) AS View_8 ON View_9.Code = View_8.Code LEFT OUTER JOIN "
        MySQLStr = MySQLStr & "(SELECT tbl_ForecastOrderR3_Main_1.Code, tbl_ForecastOrderR3_Main_1.Name, tbl_ForecastOrderR3_Main_1.ABC, "
        MySQLStr = MySQLStr & "tbl_ForecastOrderR3_Main_1.XYZ, tbl_ForecastOrderR3_Main_1.LT, tbl_ForecastOrderR3_Main_1.OI, "
        MySQLStr = MySQLStr & "ROUND(CONVERT(float, tbl_ForecastOrderR3_Main_1.MGZ), 3) AS MGZ, "
        MySQLStr = MySQLStr & "ROUND(CONVERT(float, ISNULL(View_1_1.MGZ, tbl_ForecastOrderR3_Main_1.MGZ)), 3) "
        MySQLStr = MySQLStr & "AS MGZ_OLD, ROUND(CONVERT(float, tbl_ForecastOrderR3_Main_1.ROP), 3) AS ROP, "
        MySQLStr = MySQLStr & "ROUND(CONVERT(float, ISNULL(View_1_1.ROP, tbl_ForecastOrderR3_Main_1.ROP)), 3) AS ROP_OLD, "
        MySQLStr = MySQLStr & "ROUND(CONVERT(float, tbl_ForecastOrderR3_Main_1.InshuranceLVL), 3) "
        MySQLStr = MySQLStr & "AS InshuranceLVL, ROUND(CONVERT(float, ISNULL(View_1_1.InshuranceLVL, tbl_ForecastOrderR3_Main_1.InshuranceLVL)), 3) "
        MySQLStr = MySQLStr & "AS InshuranceLVL_OLD "
        MySQLStr = MySQLStr & "FROM tbl_ForecastOrderR3_Main AS tbl_ForecastOrderR3_Main_1 WITH(NOLOCK) LEFT OUTER JOIN "
        MySQLStr = MySQLStr & "(SELECT tbl_ForecastOrderR3_Main_History_3.Code, tbl_ForecastOrderR3_Main_History_3.MGZ, "
        MySQLStr = MySQLStr & "tbl_ForecastOrderR3_Main_History_3.ROP, tbl_ForecastOrderR3_Main_History_3.InshuranceLVL "
        MySQLStr = MySQLStr & "FROM  tbl_ForecastOrderR3_Main_History AS tbl_ForecastOrderR3_Main_History_3 WITH(NOLOCK) INNER JOIN "
        MySQLStr = MySQLStr & "(SELECT Code, WarNo, MAX(Date) AS Expr1 "
        MySQLStr = MySQLStr & "FROM (SELECT tbl_ForecastOrderR3_Main_History_2.Code, tbl_ForecastOrderR3_Main_History_2.WarNo, "
        MySQLStr = MySQLStr & "tbl_ForecastOrderR3_Main_History_2.Date "
        MySQLStr = MySQLStr & "FROM tbl_ForecastOrderR3_Main_History AS tbl_ForecastOrderR3_Main_History_2 WITH(NOLOCK) LEFT OUTER JOIN "
        MySQLStr = MySQLStr & "(SELECT Code, WarNo, MAX(Date) AS Expr1 "
        MySQLStr = MySQLStr & "FROM tbl_ForecastOrderR3_Main_History AS tbl_ForecastOrderR3_Main_History_1 WITH(NOLOCK) "
        MySQLStr = MySQLStr & "WHERE (WarNo = N'" & WHCode & "') "
        MySQLStr = MySQLStr & "GROUP BY Code, WarNo) AS View_2 ON "
        MySQLStr = MySQLStr & "tbl_ForecastOrderR3_Main_History_2.Code = View_2.Code AND "
        MySQLStr = MySQLStr & "tbl_ForecastOrderR3_Main_History_2.WarNo = View_2.WarNo AND "
        MySQLStr = MySQLStr & "tbl_ForecastOrderR3_Main_History_2.Date = View_2.Expr1 "
        MySQLStr = MySQLStr & "WHERE (tbl_ForecastOrderR3_Main_History_2.WarNo = N'" & WHCode & "') AND (View_2.Expr1 IS NULL)) "
        MySQLStr = MySQLStr & "AS View_3_1 "
        MySQLStr = MySQLStr & "GROUP BY Code, WarNo) AS View_4_1 ON tbl_ForecastOrderR3_Main_History_3.Code = View_4_1.Code AND "
        MySQLStr = MySQLStr & "tbl_ForecastOrderR3_Main_History_3.WarNo = View_4_1.WarNo AND "
        MySQLStr = MySQLStr & "tbl_ForecastOrderR3_Main_History_3.Date = View_4_1.Expr1) AS View_1_1 ON "
        MySQLStr = MySQLStr & "tbl_ForecastOrderR3_Main_1.Code = View_1_1.Code "
        MySQLStr = MySQLStr & "WHERE (tbl_ForecastOrderR3_Main_1.WarNo = N'" & WHCode & "') AND (tbl_ForecastOrderR3_Main_1.WHass = - 1)) AS View_5 ON "
        MySQLStr = MySQLStr & "View_8.Code = View_5.Code LEFT OUTER JOIN "
        MySQLStr = MySQLStr & "(SELECT tbl_ForecastOrderR2_CustomMGZROPINS_History.Code, tbl_ForecastOrderR2_CustomMGZROPINS_History.MGZ, "
        MySQLStr = MySQLStr & "tbl_ForecastOrderR2_CustomMGZROPINS_History.ROP, tbl_ForecastOrderR2_CustomMGZROPINS_History.IshuranceLVL "
        MySQLStr = MySQLStr & "FROM tbl_ForecastOrderR2_CustomMGZROPINS_History WITH(NOLOCK) INNER JOIN "
        MySQLStr = MySQLStr & "(SELECT Code, MAX(DateFrom) AS DateFrom "
        MySQLStr = MySQLStr & "FROM tbl_ForecastOrderR2_CustomMGZROPINS_History AS tbl_ForecastOrderR2_CustomMGZROPINS_History_1 WITH(NOLOCK) "
        MySQLStr = MySQLStr & "WHERE (DateTo <> CONVERT(datetime, '31/12/9999', 103)) AND (WH = N'" & WHCode & "') "
        MySQLStr = MySQLStr & "GROUP BY Code) AS View_6 ON tbl_ForecastOrderR2_CustomMGZROPINS_History.Code = View_6.Code AND "
        MySQLStr = MySQLStr & "tbl_ForecastOrderR2_CustomMGZROPINS_History.DateFrom = View_6.DateFrom "
        MySQLStr = MySQLStr & "WHERE (tbl_ForecastOrderR2_CustomMGZROPINS_History.WH = N'" & WHCode & "') "
        MySQLStr = MySQLStr & "GROUP BY tbl_ForecastOrderR2_CustomMGZROPINS_History.Code, tbl_ForecastOrderR2_CustomMGZROPINS_History.MGZ, "
        MySQLStr = MySQLStr & "tbl_ForecastOrderR2_CustomMGZROPINS_History.ROP, tbl_ForecastOrderR2_CustomMGZROPINS_History.IshuranceLVL) AS View_7 ON "
        MySQLStr = MySQLStr & "View_8.Code = View_7.Code  ORDER BY View_8.Code) AS View_3_2 ON View_2_1.Code = View_3_2.Code "
        MySQLStr = MySQLStr & "ORDER BY View_2_1.Code "


        InitMyConn(False)
        InitMyRec(False, MySQLStr)
        If (Declarations.MyRec.EOF = True And Declarations.MyRec.BOF = True) Then
        Else
            Declarations.MyRec.MoveFirst()
            While Declarations.MyRec.EOF <> True
                MyWRKBook.ActiveSheet.Range("B" & StrNum) = "'" & Declarations.MyRec.Fields("Code").Value
                MyWRKBook.ActiveSheet.Range("C" & StrNum) = Declarations.MyRec.Fields("Name").Value
                MyWRKBook.ActiveSheet.Range("D" & StrNum) = Declarations.MyRec.Fields("ABC").Value
                MyWRKBook.ActiveSheet.Range("E" & StrNum) = Declarations.MyRec.Fields("XYZ").Value
                MyWRKBook.ActiveSheet.Range("F" & StrNum) = Declarations.MyRec.Fields("LT").Value
                MyWRKBook.ActiveSheet.Range("G" & StrNum) = Declarations.MyRec.Fields("OI").Value
                MyWRKBook.ActiveSheet.Range("H" & StrNum).NumberFormat = "#" & MySep & "##0" & MyDig & "00"
                MyWRKBook.ActiveSheet.Range("H" & StrNum) = Declarations.MyRec.Fields("MGZ_OLD").Value
                MyWRKBook.ActiveSheet.Range("I" & StrNum).NumberFormat = "#" & MySep & "##0" & MyDig & "00"
                MyWRKBook.ActiveSheet.Range("I" & StrNum) = Declarations.MyRec.Fields("MGZ").Value
                MyWRKBook.ActiveSheet.Range("J" & StrNum).NumberFormat = "#" & MySep & "##0" & MyDig & "00"
                MyWRKBook.ActiveSheet.Range("J" & StrNum) = Declarations.MyRec.Fields("ROP_OLD").Value
                MyWRKBook.ActiveSheet.Range("K" & StrNum).NumberFormat = "#" & MySep & "##0" & MyDig & "00"
                MyWRKBook.ActiveSheet.Range("K" & StrNum) = Declarations.MyRec.Fields("ROP").Value
                MyWRKBook.ActiveSheet.Range("L" & StrNum).NumberFormat = "#" & MySep & "##0" & MyDig & "00"
                MyWRKBook.ActiveSheet.Range("L" & StrNum) = Declarations.MyRec.Fields("InshuranceLVL").Value
                MyWRKBook.ActiveSheet.Range("M" & StrNum).NumberFormat = "#" & MySep & "##0" & MyDig & "00"
                MyWRKBook.ActiveSheet.Range("M" & StrNum) = Declarations.MyRec.Fields("InshuranceLVL").Value
                If (Declarations.MyRec.Fields("IsCustom").Value = 1) Then
                    MyWRKBook.ActiveSheet.Range("N" & StrNum).NumberFormat = "#" & MySep & "##0" & MyDig & "00"
                    MyWRKBook.ActiveSheet.Range("N" & StrNum) = Declarations.MyRec.Fields("HMGZ_OLD").Value
                    MyWRKBook.ActiveSheet.Range("O" & StrNum).NumberFormat = "#" & MySep & "##0" & MyDig & "00"
                    MyWRKBook.ActiveSheet.Range("O" & StrNum) = Declarations.MyRec.Fields("HMGZ").Value
                    MyWRKBook.ActiveSheet.Range("P" & StrNum).NumberFormat = "#" & MySep & "##0" & MyDig & "00"
                    MyWRKBook.ActiveSheet.Range("P" & StrNum) = Declarations.MyRec.Fields("HROP_OLD").Value
                    MyWRKBook.ActiveSheet.Range("Q" & StrNum).NumberFormat = "#" & MySep & "##0" & MyDig & "00"
                    MyWRKBook.ActiveSheet.Range("Q" & StrNum) = Declarations.MyRec.Fields("HROP").Value
                    MyWRKBook.ActiveSheet.Range("R" & StrNum).NumberFormat = "#" & MySep & "##0" & MyDig & "00"
                    MyWRKBook.ActiveSheet.Range("R" & StrNum) = Declarations.MyRec.Fields("HIshuranceLVL_OLD").Value
                    MyWRKBook.ActiveSheet.Range("S" & StrNum).NumberFormat = "#" & MySep & "##0" & MyDig & "00"
                    MyWRKBook.ActiveSheet.Range("S" & StrNum) = Declarations.MyRec.Fields("HIshuranceLVL").Value
                End If

                StrNum = StrNum + 1
                Declarations.MyRec.MoveNext()
            End While
            trycloseMyRec()
        End If

        StrNum = StrNum + 2
        Return StrNum
    End Function

    Private Sub Button6_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button6.Click
        '////////////////////////////////////////////////////////////////////////////////////////
        '//
        '// ��������� �������� �� Excel ����� �� ������ ������ 
        '//
        '////////////////////////////////////////////////////////////////////////////////////////
        Dim MyRez As MsgBoxResult
        Dim MyTxt As String

        MyTxt = "��� ������� ������ ��� ���������� ����������� ���� Excel, � ������� � ������ C1 ������� ����� ������ (� �������������� 0). " & Chr(13) & Chr(10)
        MyTxt = MyTxt & "������ � ����� Excel ������ ���������� � 5 ������, � ������� 'B'. ������ ������ ���� ��������� ��� ���������. " & Chr(13) & Chr(10)
        MyTxt = MyTxt & "� ������� 'B' ������ ������������� ���� ������� � ��������������� ������ " & Chr(13) & Chr(10)
        MyTxt = MyTxt & "� �������� 'C', 'D' � 'E' ������ ������������� ����� ���������� ������� ��������. ��� ������� ������ ���� ���������." & Chr(13) & Chr(10)
        MyTxt = MyTxt & "���, ROP � ������� ���������� ������." & Chr(13) & Chr(10)
        MyTxt = MyTxt & "� ��� ���� �������������� ���� Excel � �� ������ ������ ������?" & Chr(13) & Chr(10)
        MyRez = MsgBox(MyTxt, MsgBoxStyle.OkCancel, "��������!")
        If (MyRez = MsgBoxResult.Ok) Then
            ImportDataFromExcel()
            '---������ ������� �� ������� ���, ROP � ��������� ����� ��������� �������������
            BuildAutoItemList()
            '---������ ������� �� ������� ���, ROP � ��������� ����� ������������ �������
            BuildManualItemList()
            CheckButtons()
            SetWindowPos(Me.Handle.ToInt32, -2, 0, 0, 0, 0, &H3)
            Me.Cursor = Cursors.Default
        Else

        End If

    End Sub

    Private Sub ImportDataFromExcel()
        '////////////////////////////////////////////////////////////////////////////////////////
        '//
        '// �������� �� Excel ����� �� ������ ������ 
        '//
        '////////////////////////////////////////////////////////////////////////////////////////
        Dim appXLSRC As Object
        Dim MyWH As String
        Dim MyCode As String
        Dim MyMGZ As Double
        Dim MyROP As Double
        Dim MyInsLVL As Double
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
            MyWH = appXLSRC.Worksheets(1).Range("C1").Value

            '---��������� ��� � Excel ���������� ��� ������
            If MyWH = Nothing Then
                MsgBox("� ������������� ����� Excel � ������ 'C1' �� ���������� ��� ������ � Scala ", MsgBoxStyle.Critical, "��������!")
                appXLSRC.DisplayAlerts = 0
                appXLSRC.Workbooks.Close()
                appXLSRC.DisplayAlerts = 1
                appXLSRC.Quit()
                appXLSRC = Nothing
                Exit Sub
            End If
            '---��������� ��� ���� ����� ���� � Scala
            MySQLStr = "SELECT COUNT(SC23001) AS CC "
            MySQLStr = MySQLStr & "FROM SC230300 WITH(NOLOCK) "
            MySQLStr = MySQLStr & "WHERE (SC23006 = N'1') "
            MySQLStr = MySQLStr & "AND (SC23001 = N'" & MyWH & "')"
            InitMyConn(False)
            InitMyRec(False, MySQLStr)
            If (Declarations.MyRec.Fields("CC").Value = 0) Then
                MsgBox("� ������������� ����� Excel � ������ 'C1' ���������� �������� ��� ������ � Scala ", MsgBoxStyle.Critical, "��������!")
                appXLSRC.DisplayAlerts = 0
                appXLSRC.Workbooks.Close()
                appXLSRC.DisplayAlerts = 1
                appXLSRC.Quit()
                appXLSRC = Nothing
                trycloseMyRec()
                Exit Sub
            End If
            trycloseMyRec()

            '---��������� ��� ��� ���� ����������� �� ����� ������ � �������
            MySQLStr = "Update tbl_ForecastOrderR2_CustomMGZROPINS_History "
            MySQLStr = MySQLStr & "SET DateTo = GETDATE() "
            MySQLStr = MySQLStr & "WHERE (DateTo = Convert(datetime,'31/12/9999',103)) "
            MySQLStr = MySQLStr & "AND (WH = N'" & ComboBox1.SelectedValue & "') "
            InitMyConn(False)
            Declarations.MyConn.Execute(MySQLStr)

            '---�������� �������� �� ������� �������
            MySQLStr = "DELETE FROM tbl_ForecastOrderR2_CustomMGZROPINS "
            MySQLStr = MySQLStr & "WHERE (WH = N'" & ComboBox1.SelectedValue & "')"
            InitMyConn(False)
            Declarations.MyConn.Execute(MySQLStr)


            StrCnt = 5
            While Not appXLSRC.Worksheets(1).Range("B" & StrCnt).Value = Nothing
                MyCode = appXLSRC.Worksheets(1).Range("B" & StrCnt).Value.ToString
                If (appXLSRC.Worksheets(1).Range("C" & StrCnt).Value = Nothing And Not TypeOf (appXLSRC.Worksheets(1).Range("C" & StrCnt).Value) Is Double) Then
                    MsgBox("������ �" & StrCnt & " �������� ��� ������ ���� ���������.", MsgBoxStyle.Critical, "��������!")
                Else
                    If (Not TypeOf (appXLSRC.Worksheets(1).Range("C" & StrCnt).Value) Is Double) Then
                        MsgBox("������ �" & StrCnt & " �������� ��� ������ ���� ��������� �������� ���������.", MsgBoxStyle.Critical, "��������!")
                    Else
                        MyMGZ = appXLSRC.Worksheets(1).Range("C" & StrCnt).Value
                        If (appXLSRC.Worksheets(1).Range("D" & StrCnt).Value = Nothing And Not TypeOf (appXLSRC.Worksheets(1).Range("D" & StrCnt).Value) Is Double) Then
                            MsgBox("������ D" & StrCnt & " �������� ROP ������ ���� ���������.", MsgBoxStyle.Critical, "��������!")
                        Else
                            If (Not TypeOf (appXLSRC.Worksheets(1).Range("D" & StrCnt).Value) Is Double) Then
                                MsgBox("������ D" & StrCnt & " �������� ROP ������ ���� ��������� �������� ���������.", MsgBoxStyle.Critical, "��������!")
                            Else
                                MyROP = appXLSRC.Worksheets(1).Range("D" & StrCnt).Value
                                If (appXLSRC.Worksheets(1).Range("E" & StrCnt).Value = Nothing And Not TypeOf (appXLSRC.Worksheets(1).Range("E" & StrCnt).Value) Is Double) Then
                                    MsgBox("������ E" & StrCnt & " �������� ���������� ������ ������ ���� ���������.", MsgBoxStyle.Critical, "��������!")
                                Else
                                    If (Not TypeOf (appXLSRC.Worksheets(1).Range("E" & StrCnt).Value) Is Double) Then
                                        MsgBox("������ E" & StrCnt & " �������� ���������� ������ ������ ���� ��������� �������� ���������.", MsgBoxStyle.Critical, "��������!")
                                    Else
                                        MyInsLVL = appXLSRC.Worksheets(1).Range("E" & StrCnt).Value
                                        '---��� �������� - ���� �� ����� ��� � ������ ��������� �� ������� ������
                                        MySQLStr = "SELECT COUNT(Code) AS CC "
                                        MySQLStr = MySQLStr & "FROM tbl_ForecastOrderR3_Main WITH(NOLOCK) "
                                        MySQLStr = MySQLStr & "WHERE (WarNo = N'" & MyWH & "') "
                                        MySQLStr = MySQLStr & "AND (WHass = - 1) "
                                        MySQLStr = MySQLStr & "AND (Code = N'" & MyCode & "')"
                                        InitMyConn(False)
                                        InitMyRec(False, MySQLStr)
                                        If (Declarations.MyRec.Fields("CC").Value = 0) Then
                                            MsgBox("������ B" & StrCnt & " ��� ������ " & MyCode & "�� �������� ��������� �� ������ " & MyWH & ".", MsgBoxStyle.Critical, "��������!")
                                        Else
                                            UpdateDBInfo(MyWH, MyCode, MyMGZ, MyROP, MyInsLVL)
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
            MsgBox("������ ������ � ������������ ������� �������� ���, ROP � ������ ���������� ������ ����������.", MsgBoxStyle.OkOnly, "��������!")
        End If
    End Sub

    Private Sub UpdateDBInfo(ByVal MyWH As String, ByVal MyCode As String, ByVal MyMGZ As Double, ByVal MyROP As Double, ByVal MyInsLVL As Double)
        '////////////////////////////////////////////////////////////////////////////////////////
        '//
        '// ��������� (����������) ���������� - custom ���, ROP, ��������� �����
        '//
        '////////////////////////////////////////////////////////////////////////////////////////
        Dim MySQLStr As String

        '---������� � ������� �������
        MySQLStr = "INSERT INTO tbl_ForecastOrderR2_CustomMGZROPINS "
        MySQLStr = MySQLStr & "(ID, Code, WH, MGZ, ROP, IshuranceLVL) "
        MySQLStr = MySQLStr & "VALUES (NEWID(), "
        MySQLStr = MySQLStr & "N'" & MyCode & "', "
        MySQLStr = MySQLStr & "N'" & MyWH & "', "
        MySQLStr = MySQLStr & Replace(CStr(MyMGZ), ",", ".") & ", "
        MySQLStr = MySQLStr & Replace(CStr(MyROP), ",", ".") & ", "
        MySQLStr = MySQLStr & Replace(CStr(MyInsLVL), ",", ".") & ") "
        InitMyConn(False)
        Declarations.MyConn.Execute(MySQLStr)

        '---������� � �������
        MySQLStr = "INSERT INTO tbl_ForecastOrderR2_CustomMGZROPINS_History "
        MySQLStr = MySQLStr & "(ID, Code, WH, MGZ, ROP, IshuranceLVL, UserID, DateFrom, DateTo) "
        MySQLStr = MySQLStr & "VALUES (NEWID(), "
        MySQLStr = MySQLStr & "N'" & MyCode & "', "
        MySQLStr = MySQLStr & "N'" & MyWH & "', "
        MySQLStr = MySQLStr & Replace(CStr(MyMGZ), ",", ".") & ", "
        MySQLStr = MySQLStr & Replace(CStr(MyROP), ",", ".") & ", "
        MySQLStr = MySQLStr & Replace(CStr(MyInsLVL), ",", ".") & ", "
        MySQLStr = MySQLStr & " N'" & Declarations.UserCode & "', "
        MySQLStr = MySQLStr & "GETDATE(), Convert(datetime,'31/12/9999',103)) "
        InitMyConn(False)
        Declarations.MyConn.Execute(MySQLStr)

    End Sub

    Private Sub Button10_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button10.Click
        '////////////////////////////////////////////////////////////////////////////////////////
        '//
        '// ������ ������� �������������� �������� ���, ROP, ���������� ������
        '// ������� ���������� ��������� ���������� ������������
        '//
        '////////////////////////////////////////////////////////////////////////////////////////

        ReCalculate_Partial()

    End Sub

    Private Sub ReCalculate_Partial()
        '////////////////////////////////////////////////////////////////////////////////////////
        '//
        '// ��������� ������� �������������� �������� ���, ROP, ���������� ������
        '// ������� ���������� ��������� ���������� ������������
        '//
        '////////////////////////////////////////////////////////////////////////////////////////
        Dim MySQLStr As String

        MySQLStr = "Exec dbo.spp_ForecastOrderR3_Main_Incremental "
        InitMyConn(False)
        Declarations.MyConn.Execute(MySQLStr)

        '---������ ������� �� ������� ���, ROP � ��������� ����� ��������� �������������
        BuildAutoItemList()

        '---������ ������� �� ������� ���, ROP � ��������� ����� ������������ �������
        BuildManualItemList()

        CheckButtons()
    End Sub

    Private Sub Button11_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button11.Click
        '////////////////////////////////////////////////////////////////////////////////
        '//
        '// ����� ������� ������ �� ���� � ������ ������  � ������ ����
        '//
        '////////////////////////////////////////////////////////////////////////////////

        If Trim(TextBox4.Text) = "" Then
            MsgBox("���������� ������ �������� ������", MsgBoxStyle.OkOnly, "��������!")
            TextBox4.Select()
        Else
            Windows.Forms.Cursor.Current = Cursors.WaitCursor
            For i As Integer = 0 To DataGridView2.Rows.Count - 1
                If StrComp(UCase(Trim(TextBox4.Text)), Microsoft.VisualBasic.Left(UCase(Trim(DataGridView2.Item(0, i).Value.ToString)), Len(UCase(Trim(TextBox4.Text)))), 1) = 0 Then
                    DataGridView2.CurrentCell = DataGridView2.Item(0, i)
                    Windows.Forms.Cursor.Current = Cursors.Default
                    Exit Sub
                End If
            Next i
            Exit Sub
            Windows.Forms.Cursor.Current = Cursors.Default
        End If
    End Sub

    Private Sub Button14_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button14.Click
        '////////////////////////////////////////////////////////////////////////////////
        '//
        '// ����� ������� ����������� �� �������� ������ � ������ ����
        '//
        '////////////////////////////////////////////////////////////////////////////////

        If Trim(TextBox4.Text) = "" And Trim(TextBox1.Text) = "" Then
            MsgBox("���������� ������ �������� ������", MsgBoxStyle.OkOnly, "��������!")
            TextBox4.Select()
        Else
            Windows.Forms.Cursor.Current = Cursors.WaitCursor
            For i As Integer = 0 To DataGridView2.Rows.Count - 1
                If InStr(UCase(Trim(DataGridView2.Item(0, i).Value.ToString)), UCase(Trim(TextBox4.Text))) <> 0 And InStr(UCase(Trim(DataGridView2.Item(0, i).Value.ToString)), UCase(Trim(TextBox1.Text))) <> 0 _
                    Or InStr(UCase(Trim(DataGridView2.Item(1, i).Value.ToString)), UCase(Trim(TextBox4.Text))) <> 0 And InStr(UCase(Trim(DataGridView2.Item(1, i).Value.ToString)), UCase(Trim(TextBox1.Text))) <> 0 Then
                    DataGridView2.CurrentCell = DataGridView2.Item(0, i)
                    Windows.Forms.Cursor.Current = Cursors.Default
                    Exit Sub
                End If
            Next
            Windows.Forms.Cursor.Current = Cursors.Default
        End If
    End Sub

    Private Sub Button13_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button13.Click
        '////////////////////////////////////////////////////////////////////////////////
        '//
        '// ����� ���������� ����������� �� �������� ������ � ������ ����
        '//
        '////////////////////////////////////////////////////////////////////////////////
        Dim MyRez As Object

        If Trim(TextBox4.Text) = "" And Trim(TextBox1.Text) = "" Then
            MsgBox("���������� ������ �������� ������", MsgBoxStyle.OkOnly, "��������!")
            TextBox4.Select()
        Else
            Windows.Forms.Cursor.Current = Cursors.WaitCursor
            For i As Integer = DataGridView2.CurrentCellAddress.Y + 1 To DataGridView2.Rows.Count
                If i = DataGridView2.Rows.Count Then
                    MyRez = MsgBox("����� ����� �� ����� ������. ������ �������?", MsgBoxStyle.YesNo, "��������!")
                    If MyRez = 6 Then
                        i = 0
                    Else
                        Windows.Forms.Cursor.Current = Cursors.Default
                        Exit Sub
                    End If
                End If
                If DataGridView2.Rows.Count = 0 Then
                Else
                    If InStr(UCase(Trim(DataGridView2.Item(0, i).Value.ToString)), UCase(Trim(TextBox4.Text))) <> 0 And InStr(UCase(Trim(DataGridView2.Item(0, i).Value.ToString)), UCase(Trim(TextBox1.Text))) <> 0 _
                        Or InStr(UCase(Trim(DataGridView2.Item(1, i).Value.ToString)), UCase(Trim(TextBox4.Text))) <> 0 And InStr(UCase(Trim(DataGridView2.Item(1, i).Value.ToString)), UCase(Trim(TextBox1.Text))) <> 0 Then
                        DataGridView2.CurrentCell = DataGridView2.Item(0, i)
                        Windows.Forms.Cursor.Current = Cursors.Default
                        Exit Sub
                    End If
                End If
            Next i
            Windows.Forms.Cursor.Current = Cursors.Default
        End If
    End Sub

    Private Sub Button12_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button12.Click
        '////////////////////////////////////////////////////////////////////////////////
        '//
        '// ����� ���� ���������� �� �������� ������� � ��������� ����
        '//
        '////////////////////////////////////////////////////////////////////////////////

        If Trim(TextBox4.Text) = "" And Trim(TextBox1.Text) = "" Then
            MsgBox("���������� ������ �������� ������", MsgBoxStyle.OkOnly, "��������!")
            TextBox4.Select()
        Else
            MyItemSelectList2 = New ItemSelectList2
            MyItemSelectList2.ShowDialog()
        End If
    End Sub
End Class




