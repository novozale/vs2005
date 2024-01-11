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

    Public LoadFlag As Integer

    Private Sub MainForm_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        '/////////////////////////////////////////////////////////////////////////////////////
        '//
        '// ��� ������� ���������� ��������� - ���, ��������, ������������ � �.�.
        '// ����� ���� ������� ������ ����������� ������� ������������
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

        '---����� ������ � ����
        LoadFlag = 0
        '---������ �������
        BuildWHList()
        LoadFlag = 1

        '---�������, �� ������� ��������
        ComboBox2.SelectedText = "�� ���� ��������"
        
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

        If Declarations.MyWorkLevel = 0 Then          '---�������� �� ������ ��������
            MySQLStr = "SELECT tbl_ForecastOrderR4_Main_DC.Code, "
            MySQLStr = MySQLStr & "tbl_ForecastOrderR4_Main_DC.Name, "
            MySQLStr = MySQLStr & "tbl_ForecastOrderR4_Main_DC.DC, "
            MySQLStr = MySQLStr & "tbl_ForecastOrderR4_Main_DC.ABC, "
            MySQLStr = MySQLStr & "tbl_ForecastOrderR4_Main_DC.XYZ, "
            MySQLStr = MySQLStr & "tbl_ForecastOrderR4_Main_DC.LT, "
            MySQLStr = MySQLStr & "tbl_ForecastOrderR4_Main_DC.OI, "
            MySQLStr = MySQLStr & "tbl_ForecastOrderR4_Main_DC.MGZ, "
            MySQLStr = MySQLStr & "tbl_ForecastOrderR4_Main_DC.ROP, "
            MySQLStr = MySQLStr & "tbl_ForecastOrderR4_Main_DC.InshuranceLVL "
            MySQLStr = MySQLStr & "FROM tbl_ForecastOrderR4_Main_DC LEFT OUTER JOIN "
            MySQLStr = MySQLStr & "(SELECT Code FROM  dbo.tbl_ForecastOrderR4_CustomMGZROPINS_DC) AS View_1 ON tbl_ForecastOrderR4_Main_DC.Code = View_1.Code "
            MySQLStr = MySQLStr & "WHERE (View_1.Code IS NULL) AND "
            MySQLStr = MySQLStr & "(tbl_ForecastOrderR4_Main_DC.WHass = 1) "
            MySQLStr = MySQLStr & "Order BY tbl_ForecastOrderR4_Main_DC.Code "
        Else
            MySQLStr = "SELECT View_2.Code, "
            MySQLStr = MySQLStr & "View_2.Name, "
            MySQLStr = MySQLStr & "View_2.DC, "
            MySQLStr = MySQLStr & "View_2.ABC, "
            MySQLStr = MySQLStr & "View_2.XYZ, "
            MySQLStr = MySQLStr & "View_2.LT, "
            MySQLStr = MySQLStr & "View_2.OI, "
            MySQLStr = MySQLStr & "View_2.MGZ, "
            MySQLStr = MySQLStr & "View_2.ROP, "
            MySQLStr = MySQLStr & "View_2.InshuranceLVL "
            MySQLStr = MySQLStr & "FROM (SELECT Code, WH "
            MySQLStr = MySQLStr & "FROM tbl_ForecastOrderR4_CustomMGZROPINS_RWH "
            MySQLStr = MySQLStr & "WHERE (WH = N'" & ComboBox1.SelectedValue & "')) AS View_1 RIGHT OUTER JOIN "
            MySQLStr = MySQLStr & "(SELECT Code, Name, DC, ABC, XYZ, LT, OI, MGZ, ROP, InshuranceLVL, WarNo "
            MySQLStr = MySQLStr & "FROM tbl_ForecastOrderR4_Main_RWH "
            MySQLStr = MySQLStr & "WHERE (WarNo = N'" & ComboBox1.SelectedValue & "') AND ("
            MySQLStr = MySQLStr & "WHass = 1) AND "
            MySQLStr = MySQLStr & "(DC <> WarNo)) AS View_2 ON View_1.Code = View_2.Code AND View_1.WH = View_2.WarNo "
            MySQLStr = MySQLStr & "WHERE (View_1.Code Is NULL) AND (View_1.WH Is NULL) "
            MySQLStr = MySQLStr & "Order BY View_2.Code "
        End If
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
            DataGridView1.Columns(2).HeaderText = "DC"
            DataGridView1.Columns(2).Width = 50
            DataGridView1.Columns(3).HeaderText = "ABC"
            DataGridView1.Columns(3).Width = 50
            DataGridView1.Columns(4).HeaderText = "XYZ"
            DataGridView1.Columns(4).Width = 50
            DataGridView1.Columns(5).HeaderText = "LT"
            DataGridView1.Columns(5).Width = 50
            DataGridView1.Columns(6).HeaderText = "OI"
            DataGridView1.Columns(6).Width = 50
            DataGridView1.Columns(7).HeaderText = "���"
            DataGridView1.Columns(7).Width = 60
            DataGridView1.Columns(8).HeaderText = "ROP"
            DataGridView1.Columns(8).Width = 60
            DataGridView1.Columns(9).HeaderText = "����� �������"
            DataGridView1.Columns(9).Width = 60

            DataGridView1.SelectionMode = DataGridViewSelectionMode.FullRowSelect

            DataGridView1.Columns(7).DefaultCellStyle.Format = "### ##0.00"
            DataGridView1.Columns(8).DefaultCellStyle.Format = "### ##0.00"
            DataGridView1.Columns(9).DefaultCellStyle.Format = "### ##0.00"

        Catch ex As Exception
            MsgBox(ex.ToString)
        End Try
    End Sub

    Private Sub BuildManualItemList()
        '/////////////////////////////////////////////////////////////////////////////////////
        '//
        '// ����� � DataGridView2 ������ ������� �� ������� ���, ROP � ��������� ����� ������������ �������
        '// 
        '/////////////////////////////////////////////////////////////////////////////////////
        Dim MySQLStr As String                        '������� ������
        Dim MyAdapter As SqlClient.SqlDataAdapter     '
        Dim MyDs As New DataSet                       '

        If Declarations.MyWorkLevel = 0 Then          '---�������� �� ������ ��������
            MySQLStr = "SELECT tbl_ForecastOrderR4_Main_DC.Code, "
            MySQLStr = MySQLStr & "tbl_ForecastOrderR4_Main_DC.Name, "
            MySQLStr = MySQLStr & "tbl_ForecastOrderR4_Main_DC.DC, "
            MySQLStr = MySQLStr & "tbl_ForecastOrderR4_Main_DC.ABC, "
            MySQLStr = MySQLStr & "tbl_ForecastOrderR4_Main_DC.XYZ, "
            MySQLStr = MySQLStr & "tbl_ForecastOrderR4_Main_DC.LT, "
            MySQLStr = MySQLStr & "tbl_ForecastOrderR4_Main_DC.OI, "
            MySQLStr = MySQLStr & "tbl_ForecastOrderR4_Main_DC.MGZ AS AMGZ, "
            MySQLStr = MySQLStr & "tbl_ForecastOrderR4_Main_DC.ROP AS AROP, "
            MySQLStr = MySQLStr & "tbl_ForecastOrderR4_Main_DC.InshuranceLVL AS AIshuranceLVL, "
            MySQLStr = MySQLStr & "tbl_ForecastOrderR4_CustomMGZROPINS_DC.MGZ, "
            MySQLStr = MySQLStr & "tbl_ForecastOrderR4_CustomMGZROPINS_DC.ROP, "
            MySQLStr = MySQLStr & "tbl_ForecastOrderR4_CustomMGZROPINS_DC.IshuranceLVL, "
            MySQLStr = MySQLStr & "tbl_ForecastOrderR4_CustomMGZROPINS_DC.DueDate "
            MySQLStr = MySQLStr & "FROM  tbl_ForecastOrderR4_Main_DC INNER JOIN "
            MySQLStr = MySQLStr & "tbl_ForecastOrderR4_CustomMGZROPINS_DC ON tbl_ForecastOrderR4_Main_DC.Code = tbl_ForecastOrderR4_CustomMGZROPINS_DC.Code "
            MySQLStr = MySQLStr & "Order By tbl_ForecastOrderR4_Main_DC.Code "
        Else
            MySQLStr = "SELECT tbl_ForecastOrderR4_Main_RWH.Code, "
            MySQLStr = MySQLStr & "tbl_ForecastOrderR4_Main_RWH.Name, "
            MySQLStr = MySQLStr & "tbl_ForecastOrderR4_Main_RWH.DC, "
            MySQLStr = MySQLStr & "tbl_ForecastOrderR4_Main_RWH.ABC, "
            MySQLStr = MySQLStr & "tbl_ForecastOrderR4_Main_RWH.XYZ, "
            MySQLStr = MySQLStr & "tbl_ForecastOrderR4_Main_RWH.LT, "
            MySQLStr = MySQLStr & "tbl_ForecastOrderR4_Main_RWH.OI, "
            MySQLStr = MySQLStr & "tbl_ForecastOrderR4_Main_RWH.MGZ AS AMGZ, "
            MySQLStr = MySQLStr & "tbl_ForecastOrderR4_Main_RWH.ROP AS AROP, "
            MySQLStr = MySQLStr & "tbl_ForecastOrderR4_Main_RWH.InshuranceLVL AS AIshuranceLVL, "
            MySQLStr = MySQLStr & "tbl_ForecastOrderR4_CustomMGZROPINS_RWH.MGZ, "
            MySQLStr = MySQLStr & "tbl_ForecastOrderR4_CustomMGZROPINS_RWH.ROP, "
            MySQLStr = MySQLStr & "tbl_ForecastOrderR4_CustomMGZROPINS_RWH.IshuranceLVL, "
            MySQLStr = MySQLStr & "tbl_ForecastOrderR4_CustomMGZROPINS_RWH.DueDate "
            MySQLStr = MySQLStr & "FROM  tbl_ForecastOrderR4_Main_RWH INNER JOIN "
            MySQLStr = MySQLStr & "tbl_ForecastOrderR4_CustomMGZROPINS_RWH ON tbl_ForecastOrderR4_Main_RWH.Code = tbl_ForecastOrderR4_CustomMGZROPINS_RWH.Code AND "
            MySQLStr = MySQLStr & "tbl_ForecastOrderR4_Main_RWH.WarNo = tbl_ForecastOrderR4_CustomMGZROPINS_RWH.WH "
            MySQLStr = MySQLStr & "WHERE (tbl_ForecastOrderR4_Main_RWH.WarNo = N'" & ComboBox1.SelectedValue & "') "
            MySQLStr = MySQLStr & "Order By tbl_ForecastOrderR4_Main_RWH.Code "
        End If

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
            DataGridView2.Columns(2).HeaderText = "DC"
            DataGridView2.Columns(2).Width = 50
            DataGridView2.Columns(3).HeaderText = "ABC"
            DataGridView2.Columns(3).Width = 50
            DataGridView2.Columns(4).HeaderText = "XYZ"
            DataGridView2.Columns(4).Width = 50
            DataGridView2.Columns(5).HeaderText = "LT"
            DataGridView2.Columns(5).Width = 50
            DataGridView2.Columns(6).HeaderText = "OI"
            DataGridView2.Columns(6).Width = 50
            DataGridView2.Columns(7).HeaderText = "���� ���"
            DataGridView2.Columns(7).Width = 60
            DataGridView2.Columns(8).HeaderText = "���� ROP"
            DataGridView2.Columns(8).Width = 60
            DataGridView2.Columns(9).HeaderText = "���� ����� �������"
            DataGridView2.Columns(9).Width = 60
            DataGridView2.Columns(10).HeaderText = "���� ���"
            DataGridView2.Columns(10).Width = 60
            DataGridView2.Columns(11).HeaderText = "���� ROP"
            DataGridView2.Columns(11).Width = 60
            DataGridView2.Columns(12).HeaderText = "���� ����� �������"
            DataGridView2.Columns(12).Width = 60
            DataGridView2.Columns(13).HeaderText = "�� ����"
            DataGridView2.Columns(13).Width = 100


            DataGridView2.SelectionMode = DataGridViewSelectionMode.FullRowSelect

            DataGridView2.Columns(7).DefaultCellStyle.Format = "### ##0.00"
            DataGridView2.Columns(8).DefaultCellStyle.Format = "### ##0.00"
            DataGridView2.Columns(9).DefaultCellStyle.Format = "### ##0.00"
            DataGridView2.Columns(10).DefaultCellStyle.Format = "### ##0.00"
            DataGridView2.Columns(10).DefaultCellStyle.BackColor = Color.LightBlue
            DataGridView2.Columns(11).DefaultCellStyle.Format = "### ##0.00"
            DataGridView2.Columns(11).DefaultCellStyle.BackColor = Color.LightBlue
            DataGridView2.Columns(12).DefaultCellStyle.Format = "### ##0.00"
            DataGridView2.Columns(12).DefaultCellStyle.BackColor = Color.LightBlue
            DataGridView2.Columns(13).DefaultCellStyle.Format = "d"

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

        If DataGridView1.SelectedRows.Count = 0 Then
            Button2.Enabled = False
            Button5.Enabled = False
        Else
            Button2.Enabled = True
            Button5.Enabled = True
        End If
        If DataGridView2.SelectedRows.Count = 0 Then
            Button3.Enabled = False
            Button4.Enabled = False
        Else
            Button3.Enabled = True
            Button4.Enabled = True
        End If

        If Declarations.MyWorkLevel = 0 Then            '---�������� �� ������ ��������
            Button15.Enabled = True
            Button6.Enabled = False
        Else
            Button15.Enabled = False
            Button6.Enabled = True
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

    Private Sub DataGridView1_CellFormatting(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellFormattingEventArgs) Handles DataGridView1.CellFormatting
        '////////////////////////////////////////////////////////////////////////////////
        '//
        '// ��������� �������, � �������� ����������
        '//
        '////////////////////////////////////////////////////////////////////////////////
        Dim row As DataGridViewRow

        row = DataGridView1.Rows(e.RowIndex)
        If Declarations.MyWorkLevel = 0 Then          '---�������� �� ������ ��������
            If row.Cells(5).Value = 0 Or row.Cells(6).Value = 0 Or row.Cells(7).Value = 0 Or row.Cells(8).Value = 0 Or row.Cells(9).Value = 0 Then
                row.DefaultCellStyle.BackColor = Color.Yellow
            End If
        Else
            If row.Cells(5).Value = 0 Or row.Cells(7).Value = 0 Or row.Cells(8).Value = 0 Or row.Cells(9).Value = 0 Then
                row.DefaultCellStyle.BackColor = Color.Yellow
            End If
        End If
    End Sub

    Private Sub ComboBox1_SelectedValueChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles ComboBox1.SelectedValueChanged
        '////////////////////////////////////////////////////////////////////////////////
        '//
        '// ����� ������ � ComboBox1 - ������������� ������
        '//
        '////////////////////////////////////////////////////////////////////////////////

        If LoadFlag = 1 Then
            Label8.Text = "�� ������ ������ - " & ComboBox1.Text
            Label9.Text = "�� ������ ������ - " & ComboBox1.Text

            '---������ ������� �� ������� ���, ROP � ��������� ����� ��������� �������������
            BuildAutoItemList()

            '---������ ������� �� ������� ���, ROP � ��������� ����� ������������ �������
            BuildManualItemList()

            CheckButtons()
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
        '// �������� � Excel �������� ���, ROP � ���������� ������ (� ����������� �� ������ - �� ���� ������� ��� �� �����������)
        '//
        '////////////////////////////////////////////////////////////////////////////////

        Me.Cursor = Cursors.WaitCursor
        If Declarations.MyWorkLevel = 0 Then            '---�������� �� ������ ��������
            If My.Settings.UseOffice = "LibreOffice" Then
                UploadItemInfoDC_LO()
            Else
                UploadItemInfoDC()
            End If
        Else                                            '---�������� �� ������ ���������� ������
            If My.Settings.UseOffice = "LibreOffice" Then
                UploadItemInfo_LO()
            Else
                UploadItemInfo()
            End If
        End If
        Me.Cursor = Cursors.Default
    End Sub

    Private Sub AddItemCustomInfo()
        '////////////////////////////////////////////////////////////////////////////////////////
        '//
        '// ��������� ���������� ������ �������� ���, ROP � ���������� ������
        '//
        '////////////////////////////////////////////////////////////////////////////////////////
        Dim MySQLStr As String                        '������� ������
        Dim MyCode As String                          '��� ������

        Declarations.MySuccess = False
        MyAddCustom = New AddCustom
        MyAddCustom.ShowDialog()
        If Declarations.MySuccess = False Then
            Exit Sub
        Else '---���������� ������ �������� ���, ROP � ���������� ������
            MyCode = Trim(DataGridView1.SelectedRows.Item(0).Cells(0).Value.ToString())
            If Declarations.MyWorkLevel = 0 Then            '---�������� �� ������ ��������
                '---��������� ����� �������� � ������� �������
                MySQLStr = "INSERT INTO tbl_ForecastOrderR4_CustomMGZROPINS_DC "
                MySQLStr = MySQLStr & "(ID, Code, MGZ, ROP, IshuranceLVL, DueDate) "
                MySQLStr = MySQLStr & "VALUES (NEWID(), "
                MySQLStr = MySQLStr & "N'" & Trim(DataGridView1.SelectedRows.Item(0).Cells(0).Value.ToString()) & "', "
                MySQLStr = MySQLStr & Replace(CStr(Declarations.MyMGZ), ",", ".") & ", "
                MySQLStr = MySQLStr & Replace(CStr(Declarations.MyROP), ",", ".") & ", "
                MySQLStr = MySQLStr & Replace(CStr(Declarations.MyInsuranceLVL), ",", ".") & ", "
                MySQLStr = MySQLStr & "CONVERT(DATETIME, '" & FormatDateTime(Declarations.MyDueDate, DateFormat.ShortDate) & "', 103)) "
                InitMyConn(False)
                Declarations.MyConn.Execute(MySQLStr)
            Else                                            '---�������� �� ������ ���������� ������
                '---��������� ����� �������� � ������� �������
                MySQLStr = "INSERT INTO tbl_ForecastOrderR4_CustomMGZROPINS_RWH "
                MySQLStr = MySQLStr & "(ID, Code, WH, MGZ, ROP, IshuranceLVL, DueDate) "
                MySQLStr = MySQLStr & "VALUES (NEWID(), "
                MySQLStr = MySQLStr & "N'" & Trim(DataGridView1.SelectedRows.Item(0).Cells(0).Value.ToString()) & "', "
                MySQLStr = MySQLStr & "N'" & ComboBox1.SelectedValue & "', "
                MySQLStr = MySQLStr & Replace(CStr(Declarations.MyMGZ), ",", ".") & ", "
                MySQLStr = MySQLStr & Replace(CStr(Declarations.MyROP), ",", ".") & ", "
                MySQLStr = MySQLStr & Replace(CStr(Declarations.MyInsuranceLVL), ",", ".") & ", "
                MySQLStr = MySQLStr & "CONVERT(DATETIME, '" & FormatDateTime(Declarations.MyDueDate, DateFormat.ShortDate) & "', 103)) "
                InitMyConn(False)
                Declarations.MyConn.Execute(MySQLStr)
            End If

            '---������ ������� �� ������� ���, ROP � ��������� ����� ��������� �������������
            BuildAutoItemList()
            '---������ ������� �� ������� ���, ROP � ��������� ����� ������������ �������
            BuildManualItemList()
            '---������� ������� ������� ���������
            For i As Integer = 0 To DataGridView2.Rows.Count - 1
                If Trim(DataGridView2.Item(0, i).Value.ToString) = MyCode Then
                    DataGridView2.CurrentCell = DataGridView2.Item(2, i)
                End If
            Next
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

        If Declarations.MyWorkLevel = 0 Then            '---�������� �� ������ ��������
            MySQLStr = "DELETE FROM tbl_ForecastOrderR4_CustomMGZROPINS_DC "
            MySQLStr = MySQLStr & "WHERE (Code = N'" & Trim(DataGridView2.SelectedRows.Item(0).Cells(0).Value.ToString()) & "') "
            InitMyConn(False)
            Declarations.MyConn.Execute(MySQLStr)
        Else                                            '---�������� �� ������ ���������� ������
            '---�������� �������� �� ������� �������
            MySQLStr = "DELETE FROM tbl_ForecastOrderR4_CustomMGZROPINS_RWH "
            MySQLStr = MySQLStr & "WHERE (Code = N'" & Trim(DataGridView2.SelectedRows.Item(0).Cells(0).Value.ToString()) & "') "
            MySQLStr = MySQLStr & "AND (WH = N'" & ComboBox1.SelectedValue & "')"
            InitMyConn(False)
            Declarations.MyConn.Execute(MySQLStr)
        End If

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
        Dim MyCode As String                          '��� ������

        Declarations.MySuccess = False
        MyEditCustom = New EditCustom
        MyEditCustom.ShowDialog()
        If Declarations.MySuccess = False Then
            Exit Sub
        Else '---��������� ������ �������� ���, ROP � ���������� ������
            MyCode = Trim(DataGridView2.SelectedRows.Item(0).Cells(0).Value.ToString())
            If Declarations.MyWorkLevel = 0 Then            '---�������� �� ������ ��������
                MySQLStr = "UPDATE tbl_ForecastOrderR4_CustomMGZROPINS_DC "
                MySQLStr = MySQLStr & "SET MGZ = " & Replace(CStr(Declarations.MyMGZ), ",", ".") & ", "
                MySQLStr = MySQLStr & "ROP = " & Replace(CStr(Declarations.MyROP), ",", ".") & ", "
                MySQLStr = MySQLStr & "IshuranceLVL = " & Replace(CStr(Declarations.MyInsuranceLVL), ",", ".") & ", "
                MySQLStr = MySQLStr & "DueDate = CONVERT(DATETIME, '" & FormatDateTime(Declarations.MyDueDate, DateFormat.ShortDate) & "', 103) "
                MySQLStr = MySQLStr & "WHERE (Code = N'" & Trim(DataGridView2.SelectedRows.Item(0).Cells(0).Value.ToString()) & "') "
                InitMyConn(False)
                Declarations.MyConn.Execute(MySQLStr)
            Else                                            '---�������� �� ������ ���������� ������
                '---���������� ��������
                MySQLStr = "UPDATE tbl_ForecastOrderR4_CustomMGZROPINS_RWH "
                MySQLStr = MySQLStr & "SET MGZ = " & Replace(CStr(Declarations.MyMGZ), ",", ".") & ", "
                MySQLStr = MySQLStr & "ROP = " & Replace(CStr(Declarations.MyROP), ",", ".") & ", "
                MySQLStr = MySQLStr & "IshuranceLVL = " & Replace(CStr(Declarations.MyInsuranceLVL), ",", ".") & ", "
                MySQLStr = MySQLStr & "DueDate = CONVERT(DATETIME, '" & FormatDateTime(Declarations.MyDueDate, DateFormat.ShortDate) & "', 103) "
                MySQLStr = MySQLStr & "WHERE (Code = N'" & Trim(DataGridView2.SelectedRows.Item(0).Cells(0).Value.ToString()) & "') "
                MySQLStr = MySQLStr & "AND (WH = N'" & ComboBox1.SelectedValue & "') "
                InitMyConn(False)
                Declarations.MyConn.Execute(MySQLStr)
            End If

            '---������ ������� �� ������� ���, ROP � ��������� ����� ��������� �������������
            BuildAutoItemList()
            '---������ ������� �� ������� ���, ROP � ��������� ����� ������������ �������
            BuildManualItemList()
            '---������� ������� ������� ���������
            For i As Integer = 0 To DataGridView2.Rows.Count - 1
                If Trim(DataGridView2.Item(0, i).Value.ToString) = MyCode Then
                    DataGridView2.CurrentCell = DataGridView2.Item(2, i)
                End If
            Next
            CheckButtons()
        End If
    End Sub

    Private Sub UploadItemInfo()
        '////////////////////////////////////////////////////////////////////////////////////////
        '//
        '// ��������� �������� � Excel �������� ���, ROP � ���������� ������ �� ��������� �������
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

    Private Sub UploadItemInfo_LO()
        '////////////////////////////////////////////////////////////////////////////////////////
        '//
        '// ��������� �������� � LibreOffice �������� ���, ROP � ���������� ������ �� ��������� �������
        '//
        '////////////////////////////////////////////////////////////////////////////////////////
        Dim WHList(,) As String
        Dim MySQLStr As String
        Dim i As Integer
        Dim oServiceManager As Object
        Dim oDispatcher As Object
        Dim oDesktop As Object
        Dim oWorkBook As Object
        Dim oSheet As Object
        Dim StrNum As Double      '����� ������

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

        LOSetNotation(0)
        oServiceManager = CreateObject("com.sun.star.ServiceManager")
        oDesktop = oServiceManager.createInstance("com.sun.star.frame.Desktop")
        oDispatcher = oServiceManager.createInstance("com.sun.star.frame.DispatchHelper")
        Dim arg(1)
        arg(0) = mAkePropertyValue("Hidden", True, oServiceManager)
        arg(1) = mAkePropertyValue("MacroExecutionMode", 4, oServiceManager)
        oWorkBook = oDesktop.loadComponentFromURL("private:factory/scalc", "_blank", 0, arg)
        oSheet = oWorkBook.getSheets().getByIndex(0)

        UploadCommonHeader_LO(oSheet, oServiceManager, oWorkBook, oDispatcher)

        StrNum = 5
        For i = 0 To WHList.GetUpperBound(1)
            StrNum = UploadWHHeader_LO(oSheet, oServiceManager, oWorkBook, oDispatcher, WHList(0, i), WHList(1, i), StrNum)
            StrNum = UploadWHRows_LO(oSheet, oServiceManager, oWorkBook, oDispatcher, WHList(0, i), StrNum)
        Next

        Dim args() As Object
        ReDim args(0)
        args(0) = oServiceManager.Bridge_GetStruct("com.sun.star.beans.PropertyValue")
        args(0).Name = "ToPoint"
        args(0).Value = "$A$1"
        Dim oFrame As Object
        oFrame = oWorkBook.getCurrentController.getFrame
        oDispatcher.executeDispatch(oFrame, ".uno:GoToCell", "", 0, args)

        oWorkBook.CurrentController.Frame.ContainerWindow.Visible = True
        oWorkBook.CurrentController.Frame.ContainerWindow.toFront()
    End Sub

    Private Sub UploadItemInfoDC()
        '////////////////////////////////////////////////////////////////////////////////////////
        '//
        '// ��������� �������� � Excel �������� ���, ROP � ���������� ������ �� ���� �������
        '//
        '////////////////////////////////////////////////////////////////////////////////////////
        Dim StrNum As Double      '����� ������
        Dim MyObj As Object       'Excel
        Dim MyWRKBook As Object   '�����

        MyObj = CreateObject("Excel.Application")
        MyObj.SheetsInNewWorkbook = 1
        MyWRKBook = MyObj.Workbooks.Add

        UploadCommonHeaderDC(MyWRKBook)
        StrNum = 5
        StrNum = UploadDCHeader(MyWRKBook, StrNum)
        StrNum = UploadDCRows(MyWRKBook, StrNum)

        MyWRKBook.ActiveSheet.Range("A1").Select()
        MyObj.Application.Visible = True
        MyObj = Nothing
    End Sub

    Private Sub UploadItemInfoDC_LO()
        '////////////////////////////////////////////////////////////////////////////////////////
        '//
        '// ��������� �������� � LibreOffice �������� ���, ROP � ���������� ������ �� ���� �������
        '//
        '////////////////////////////////////////////////////////////////////////////////////////
        Dim StrNum As Double      '����� ������
        Dim oServiceManager As Object
        Dim oDispatcher As Object
        Dim oDesktop As Object
        Dim oWorkBook As Object
        Dim oSheet As Object

        LOSetNotation(0)
        oServiceManager = CreateObject("com.sun.star.ServiceManager")
        oDesktop = oServiceManager.createInstance("com.sun.star.frame.Desktop")
        oDispatcher = oServiceManager.createInstance("com.sun.star.frame.DispatchHelper")
        Dim arg(1)
        arg(0) = mAkePropertyValue("Hidden", True, oServiceManager)
        arg(1) = mAkePropertyValue("MacroExecutionMode", 4, oServiceManager)
        oWorkBook = oDesktop.loadComponentFromURL("private:factory/scalc", "_blank", 0, arg)
        oSheet = oWorkBook.getSheets().getByIndex(0)

        UploadCommonHeaderDC_LO(oSheet, oServiceManager, oWorkBook, oDispatcher)

        StrNum = 5
        StrNum = UploadDCHeader_LO(oSheet, oServiceManager, oWorkBook, oDispatcher, StrNum)
        StrNum = UploadDCRows_LO(oSheet, oServiceManager, oWorkBook, oDispatcher, StrNum)

        Dim args() As Object
        ReDim args(0)
        args(0) = oServiceManager.Bridge_GetStruct("com.sun.star.beans.PropertyValue")
        args(0).Name = "ToPoint"
        args(0).Value = "$A$1"
        Dim oFrame As Object
        oFrame = oWorkBook.getCurrentController.getFrame
        oDispatcher.executeDispatch(oFrame, ".uno:GoToCell", "", 0, args)

        oWorkBook.CurrentController.Frame.ContainerWindow.Visible = True
        oWorkBook.CurrentController.Frame.ContainerWindow.toFront()
    End Sub

    Private Function UploadCommonHeader(ByVal MyWRKBook As Object)
        '////////////////////////////////////////////////////////////////////////////////////////
        '//
        '// ��������� �������� � Excel ������ ��������� 
        '//
        '////////////////////////////////////////////////////////////////////////////////////////

        MyWRKBook.ActiveSheet.Range("B1") = "���������� � ���, ROP � ������ ���������� ������ "
        MyWRKBook.ActiveSheet.Range("B2") = "���������� ������������ �� ��������� ������� �� " & Now
        MyWRKBook.ActiveSheet.Range("B1:B2").Select()
        MyWRKBook.ActiveSheet.Range("B1:B2").Font.Bold = True

        '--- � ������� �����
        MyWRKBook.ActiveSheet.Columns("A:O").ColumnWidth = 10
        MyWRKBook.ActiveSheet.Columns("B:B").ColumnWidth = 17
        MyWRKBook.ActiveSheet.Columns("C:C").ColumnWidth = 30
    End Function

    Private Function UploadCommonHeader_LO(ByRef oSheet, ByRef oServiceManager, ByRef oWorkBook, ByRef oDispatcher)
        '////////////////////////////////////////////////////////////////////////////////////////
        '//
        '// ��������� �������� � LibreOffice ������ ��������� 
        '//
        '////////////////////////////////////////////////////////////////////////////////////////

        oSheet.getCellRangeByName("B1").String = "���������� � ���, ROP � ������ ���������� ������"
        oSheet.getCellRangeByName("B2").String = "���������� ������������ �� ��������� ������� �� " & Now

        '--- ������� �����
        oSheet.getColumns().getByName("B").Width = 3800
        oSheet.getColumns().getByName("C").Width = 11000

        '--- �����
        Dim args() As Object
        ReDim args(0)
        args(0) = oServiceManager.Bridge_GetStruct("com.sun.star.beans.PropertyValue")
        args(0).Name = "ToPoint"
        args(0).Value = "$B$1:$B$2"
        Dim oFrame As Object
        oFrame = oWorkBook.getCurrentController.getFrame
        oDispatcher.executeDispatch(oFrame, ".uno:GoToCell", "", 0, args)

        Dim args1() As Object
        ReDim args1(5)
        args1(0) = oServiceManager.Bridge_GetStruct("com.sun.star.beans.PropertyValue")
        args1(0).Name = "CharFontName.StyleName"
        args1(0).Value = "�������"
        args1(1) = oServiceManager.Bridge_GetStruct("com.sun.star.beans.PropertyValue")
        args1(1).Name = "CharFontName.Pitch"
        args1(1).Value = 2
        args1(2) = oServiceManager.Bridge_GetStruct("com.sun.star.beans.PropertyValue")
        args1(2).Name = "CharFontName.CharSet"
        args1(2).Value = 0
        args1(3) = oServiceManager.Bridge_GetStruct("com.sun.star.beans.PropertyValue")
        args1(3).Name = "CharFontName.Family"
        args1(3).Value = 5
        args1(4) = oServiceManager.Bridge_GetStruct("com.sun.star.beans.PropertyValue")
        args1(4).Name = "CharFontName.FamilyName"
        args1(4).Value = "Calibri"
        oDispatcher.executeDispatch(oFrame, ".uno:CharFontName", "", 0, args1)

        Dim args2() As Object
        ReDim args2(0)
        args2(0) = oServiceManager.Bridge_GetStruct("com.sun.star.beans.PropertyValue")
        args2(0).Name = "Bold"
        args2(0).Value = True
        oDispatcher.executeDispatch(oFrame, ".uno:Bold", "", 0, args2)
    End Function

    Private Function UploadCommonHeaderDC(ByVal MyWRKBook As Object)
        '////////////////////////////////////////////////////////////////////////////////////////
        '//
        '// ��������� �������� � Excel ������ ��������� ��� DC
        '//
        '////////////////////////////////////////////////////////////////////////////////////////

        MyWRKBook.ActiveSheet.Range("B1") = "���������� � ���, ROP � ������ ���������� ������ "
        MyWRKBook.ActiveSheet.Range("B2") = "���������� ������������ �� ���� ������� �� " & Now
        MyWRKBook.ActiveSheet.Range("B1:B2").Select()
        MyWRKBook.ActiveSheet.Range("B1:B2").Font.Bold = True

        '--- � ������� �����
        MyWRKBook.ActiveSheet.Columns("A:N").ColumnWidth = 10
        MyWRKBook.ActiveSheet.Columns("A:A").ColumnWidth = 17
        MyWRKBook.ActiveSheet.Columns("B:B").ColumnWidth = 30
    End Function

    Private Function UploadCommonHeaderDC_LO(ByRef oSheet, ByRef oServiceManager, ByRef oWorkBook, ByRef oDispatcher)
        '////////////////////////////////////////////////////////////////////////////////////////
        '//
        '// ��������� �������� � LibreOffice ������ ��������� ��� DC
        '//
        '////////////////////////////////////////////////////////////////////////////////////////

        oSheet.getCellRangeByName("B1").String = "���������� � ���, ROP � ������ ���������� ������"
        oSheet.getCellRangeByName("B2").String = "���������� ������������ �� ���� ������� �� " & Now

        '--- ������� �����
        oSheet.getColumns().getByName("A").Width = 3800
        oSheet.getColumns().getByName("B").Width = 11000

        '--- �����
        Dim args() As Object
        ReDim args(0)
        args(0) = oServiceManager.Bridge_GetStruct("com.sun.star.beans.PropertyValue")
        args(0).Name = "ToPoint"
        args(0).Value = "$B$1:$B$2"
        Dim oFrame As Object
        oFrame = oWorkBook.getCurrentController.getFrame
        oDispatcher.executeDispatch(oFrame, ".uno:GoToCell", "", 0, args)

        Dim args1() As Object
        ReDim args1(5)
        args1(0) = oServiceManager.Bridge_GetStruct("com.sun.star.beans.PropertyValue")
        args1(0).Name = "CharFontName.StyleName"
        args1(0).Value = "�������"
        args1(1) = oServiceManager.Bridge_GetStruct("com.sun.star.beans.PropertyValue")
        args1(1).Name = "CharFontName.Pitch"
        args1(1).Value = 2
        args1(2) = oServiceManager.Bridge_GetStruct("com.sun.star.beans.PropertyValue")
        args1(2).Name = "CharFontName.CharSet"
        args1(2).Value = 0
        args1(3) = oServiceManager.Bridge_GetStruct("com.sun.star.beans.PropertyValue")
        args1(3).Name = "CharFontName.Family"
        args1(3).Value = 5
        args1(4) = oServiceManager.Bridge_GetStruct("com.sun.star.beans.PropertyValue")
        args1(4).Name = "CharFontName.FamilyName"
        args1(4).Value = "Calibri"
        oDispatcher.executeDispatch(oFrame, ".uno:CharFontName", "", 0, args1)

        Dim args2() As Object
        ReDim args2(0)
        args2(0) = oServiceManager.Bridge_GetStruct("com.sun.star.beans.PropertyValue")
        args2(0).Name = "Bold"
        args2(0).Value = True
        oDispatcher.executeDispatch(oFrame, ".uno:Bold", "", 0, args2)
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
        MyWRKBook.ActiveSheet.Range("D" & StrNum) = "DC"
        MyWRKBook.ActiveSheet.Range("E" & StrNum) = "ABC"
        MyWRKBook.ActiveSheet.Range("F" & StrNum) = "XYZ"
        MyWRKBook.ActiveSheet.Range("G" & StrNum) = "����� ��������"
        MyWRKBook.ActiveSheet.Range("H" & StrNum) = "����� ����� ��������"
        MyWRKBook.ActiveSheet.Range("I" & StrNum) = "���� ���"
        MyWRKBook.ActiveSheet.Range("J" & StrNum) = "���� ROP"
        MyWRKBook.ActiveSheet.Range("K" & StrNum) = "���� ����� �������"
        MyWRKBook.ActiveSheet.Range("L" & StrNum) = "������ ���"
        MyWRKBook.ActiveSheet.Range("M" & StrNum) = "������ ROP"
        MyWRKBook.ActiveSheet.Range("N" & StrNum) = "������ ����� �������"
        MyWRKBook.ActiveSheet.Range("O" & StrNum) = "��������� �� ����"

        MyWRKBook.ActiveSheet.Range("B" & StrNum & ":O" & StrNum).Select()
        MyWRKBook.ActiveSheet.Range("B" & StrNum & ":O" & StrNum).HorizontalAlignment = -4108
        MyWRKBook.ActiveSheet.Range("B" & StrNum & ":O" & StrNum).VerticalAlignment = -4108
        MyWRKBook.ActiveSheet.Range("B" & StrNum & ":O" & StrNum).WrapText = True
        MyWRKBook.ActiveSheet.Range("B" & StrNum & ":O" & StrNum).Borders(5).LineStyle = -4142
        MyWRKBook.ActiveSheet.Range("B" & StrNum & ":O" & StrNum).Borders(6).LineStyle = -4142
        With MyWRKBook.ActiveSheet.Range("B" & StrNum & ":O" & StrNum).Borders(7)
            .LineStyle = 1
            .Weight = 4
            .ColorIndex = -4105
        End With
        With MyWRKBook.ActiveSheet.Range("B" & StrNum & ":O" & StrNum).Borders(8)
            .LineStyle = 1
            .Weight = 4
            .ColorIndex = -4105
        End With
        With MyWRKBook.ActiveSheet.Range("B" & StrNum & ":O" & StrNum).Borders(9)
            .LineStyle = 1
            .Weight = 4
            .ColorIndex = -4105
        End With
        With MyWRKBook.ActiveSheet.Range("B" & StrNum & ":O" & StrNum).Borders(10)
            .LineStyle = 1
            .Weight = 4
            .ColorIndex = -4105
        End With
        With MyWRKBook.ActiveSheet.Range("B" & StrNum & ":O" & StrNum).Borders(11)
            .LineStyle = 1
            .Weight = 4
            .ColorIndex = -4105
        End With
        With MyWRKBook.ActiveSheet.Range("B" & StrNum & ":O" & StrNum).Interior
            .ColorIndex = 35
            .Pattern = 1
            .PatternColorIndex = -4105
        End With

        StrNum = StrNum + 1
        Return StrNum
    End Function

    Private Function UploadWHHeader_LO(ByRef oSheet As Object, ByRef oServiceManager As Object, ByRef oWorkBook As Object, _
        ByRef oDispatcher As Object, ByVal WHCode As String, ByVal WHName As String, ByVal StrNum As Double) As Double
        '////////////////////////////////////////////////////////////////////////////////////////
        '//
        '// ��������� �������� � LibreOffice ��������� �� ������ ������ 
        '//
        '////////////////////////////////////////////////////////////////////////////////////////

        oSheet.getCellRangeByName("A" & StrNum).String = WHCode
        oSheet.getCellRangeByName("B" & StrNum).String = WHName

        oSheet.getCellRangeByName("A" & StrNum & ":B" & StrNum).CellBackColor = 16775598
        oSheet.getCellRangeByName("A" & StrNum & ":B" & StrNum).VertJustify = 2
        oSheet.getCellRangeByName("A" & StrNum & ":B" & StrNum).HoriJustify = 2

        Dim LineFormat As Object
        LineFormat = oServiceManager.Bridge_GetStruct("com.sun.star.table.BorderLine2")
        LineFormat.LineStyle = 0
        LineFormat.LineWidth = 70
        oSheet.getCellRangeByName("A" & StrNum & ":B" & StrNum).TopBorder = LineFormat
        oSheet.getCellRangeByName("A" & StrNum & ":B" & StrNum).RightBorder = LineFormat
        oSheet.getCellRangeByName("A" & StrNum & ":B" & StrNum).LeftBorder = LineFormat
        oSheet.getCellRangeByName("A" & StrNum & ":B" & StrNum).BottomBorder = LineFormat

        Dim args() As Object
        ReDim args(0)
        args(0) = oServiceManager.Bridge_GetStruct("com.sun.star.beans.PropertyValue")
        args(0).Name = "ToPoint"
        args(0).Value = "A" & StrNum & ":B" & StrNum
        Dim oFrame As Object
        oFrame = oWorkBook.getCurrentController.getFrame
        oDispatcher.executeDispatch(oFrame, ".uno:GoToCell", "", 0, args)

        Dim args1() As Object
        ReDim args1(5)
        args1(0) = oServiceManager.Bridge_GetStruct("com.sun.star.beans.PropertyValue")
        args1(0).Name = "CharFontName.StyleName"
        args1(0).Value = "�������"
        args1(1) = oServiceManager.Bridge_GetStruct("com.sun.star.beans.PropertyValue")
        args1(1).Name = "CharFontName.Pitch"
        args1(1).Value = 2
        args1(2) = oServiceManager.Bridge_GetStruct("com.sun.star.beans.PropertyValue")
        args1(2).Name = "CharFontName.CharSet"
        args1(2).Value = 0
        args1(3) = oServiceManager.Bridge_GetStruct("com.sun.star.beans.PropertyValue")
        args1(3).Name = "CharFontName.Family"
        args1(3).Value = 5
        args1(4) = oServiceManager.Bridge_GetStruct("com.sun.star.beans.PropertyValue")
        args1(4).Name = "CharFontName.FamilyName"
        args1(4).Value = "Calibri"
        oDispatcher.executeDispatch(oFrame, ".uno:CharFontName", "", 0, args1)

        Dim args2() As Object
        ReDim args2(0)
        args2(0) = oServiceManager.Bridge_GetStruct("com.sun.star.beans.PropertyValue")
        args2(0).Name = "Bold"
        args2(0).Value = True
        oDispatcher.executeDispatch(oFrame, ".uno:Bold", "", 0, args2)

        Dim args3() As Object
        ReDim args3(0)
        args3(0) = oServiceManager.Bridge_GetStruct("com.sun.star.beans.PropertyValue")
        args3(0).Name = "WrapText"
        args3(0).Value = True
        oDispatcher.executeDispatch(oFrame, ".uno:WrapText", "", 0, args3)

        StrNum = StrNum + 2
        '---� ��������� ��� �����
        oSheet.getCellRangeByName("B" & StrNum).String = "ID"
        oSheet.getCellRangeByName("C" & StrNum).String = "�������� ������"
        oSheet.getCellRangeByName("D" & StrNum).String = "DC"
        oSheet.getCellRangeByName("E" & StrNum).String = "ABC"
        oSheet.getCellRangeByName("F" & StrNum).String = "XYZ"
        oSheet.getCellRangeByName("G" & StrNum).String = "����� ��������"
        oSheet.getCellRangeByName("H" & StrNum).String = "����� ����� ��������"
        oSheet.getCellRangeByName("I" & StrNum).String = "���� ���"
        oSheet.getCellRangeByName("J" & StrNum).String = "���� ROP"
        oSheet.getCellRangeByName("K" & StrNum).String = "���� ����� �������"
        oSheet.getCellRangeByName("L" & StrNum).String = "������ ���"
        oSheet.getCellRangeByName("M" & StrNum).String = "������ ROP"
        oSheet.getCellRangeByName("N" & StrNum).String = "������ ����� �������"
        oSheet.getCellRangeByName("O" & StrNum).String = "��������� �� ����"

        oSheet.getCellRangeByName("B" & StrNum & ":O" & StrNum).CellBackColor = 12510163
        oSheet.getCellRangeByName("B" & StrNum & ":O" & StrNum).VertJustify = 2
        oSheet.getCellRangeByName("B" & StrNum & ":O" & StrNum).HoriJustify = 2

        LineFormat.LineStyle = 0
        LineFormat.LineWidth = 70
        oSheet.getCellRangeByName("B" & StrNum & ":O" & StrNum).TopBorder = LineFormat
        oSheet.getCellRangeByName("B" & StrNum & ":O" & StrNum).RightBorder = LineFormat
        oSheet.getCellRangeByName("B" & StrNum & ":O" & StrNum).LeftBorder = LineFormat
        oSheet.getCellRangeByName("B" & StrNum & ":O" & StrNum).BottomBorder = LineFormat

        ReDim args(0)
        args(0) = oServiceManager.Bridge_GetStruct("com.sun.star.beans.PropertyValue")
        args(0).Name = "ToPoint"
        args(0).Value = "B" & StrNum & ":O" & StrNum
        oDispatcher.executeDispatch(oFrame, ".uno:GoToCell", "", 0, args)

        ReDim args1(5)
        args1(0) = oServiceManager.Bridge_GetStruct("com.sun.star.beans.PropertyValue")
        args1(0).Name = "CharFontName.StyleName"
        args1(0).Value = "�������"
        args1(1) = oServiceManager.Bridge_GetStruct("com.sun.star.beans.PropertyValue")
        args1(1).Name = "CharFontName.Pitch"
        args1(1).Value = 2
        args1(2) = oServiceManager.Bridge_GetStruct("com.sun.star.beans.PropertyValue")
        args1(2).Name = "CharFontName.CharSet"
        args1(2).Value = 0
        args1(3) = oServiceManager.Bridge_GetStruct("com.sun.star.beans.PropertyValue")
        args1(3).Name = "CharFontName.Family"
        args1(3).Value = 5
        args1(4) = oServiceManager.Bridge_GetStruct("com.sun.star.beans.PropertyValue")
        args1(4).Name = "CharFontName.FamilyName"
        args1(4).Value = "Calibri"
        oDispatcher.executeDispatch(oFrame, ".uno:CharFontName", "", 0, args1)

        ReDim args2(0)
        args2(0) = oServiceManager.Bridge_GetStruct("com.sun.star.beans.PropertyValue")
        args2(0).Name = "Bold"
        args2(0).Value = True
        oDispatcher.executeDispatch(oFrame, ".uno:Bold", "", 0, args2)

        ReDim args3(0)
        args3(0) = oServiceManager.Bridge_GetStruct("com.sun.star.beans.PropertyValue")
        args3(0).Name = "WrapText"
        args3(0).Value = True
        oDispatcher.executeDispatch(oFrame, ".uno:WrapText", "", 0, args2)


        StrNum = StrNum + 1
        Return StrNum
    End Function

    Private Function UploadDCHeader(ByVal MyWRKBook As Object, ByVal StrNum As Double) As Double
        '////////////////////////////////////////////////////////////////////////////////////////
        '//
        '// ��������� �������� � Excel ��������� �� ���� ������� 
        '//
        '////////////////////////////////////////////////////////////////////////////////////////

        '---��������� ��� �����
        MyWRKBook.ActiveSheet.Range("A" & StrNum) = "ID"
        MyWRKBook.ActiveSheet.Range("B" & StrNum) = "�������� ������"
        MyWRKBook.ActiveSheet.Range("C" & StrNum) = "DC"
        MyWRKBook.ActiveSheet.Range("D" & StrNum) = "ABC"
        MyWRKBook.ActiveSheet.Range("E" & StrNum) = "XYZ"
        MyWRKBook.ActiveSheet.Range("F" & StrNum) = "����� ��������"
        MyWRKBook.ActiveSheet.Range("G" & StrNum) = "����� ����� ��������"
        MyWRKBook.ActiveSheet.Range("H" & StrNum) = "���� ���"
        MyWRKBook.ActiveSheet.Range("I" & StrNum) = "���� ROP"
        MyWRKBook.ActiveSheet.Range("J" & StrNum) = "���� ����� �������"
        MyWRKBook.ActiveSheet.Range("K" & StrNum) = "������ ���"
        MyWRKBook.ActiveSheet.Range("L" & StrNum) = "������ ROP"
        MyWRKBook.ActiveSheet.Range("M" & StrNum) = "������ ����� �������"
        MyWRKBook.ActiveSheet.Range("N" & StrNum) = "��������� �� ����"

        MyWRKBook.ActiveSheet.Range("A" & StrNum & ":N" & StrNum).Select()
        MyWRKBook.ActiveSheet.Range("A" & StrNum & ":N" & StrNum).HorizontalAlignment = -4108
        MyWRKBook.ActiveSheet.Range("A" & StrNum & ":N" & StrNum).VerticalAlignment = -4108
        MyWRKBook.ActiveSheet.Range("A" & StrNum & ":N" & StrNum).WrapText = True
        MyWRKBook.ActiveSheet.Range("A" & StrNum & ":N" & StrNum).Borders(5).LineStyle = -4142
        MyWRKBook.ActiveSheet.Range("A" & StrNum & ":N" & StrNum).Borders(6).LineStyle = -4142
        With MyWRKBook.ActiveSheet.Range("A" & StrNum & ":N" & StrNum).Borders(7)
            .LineStyle = 1
            .Weight = 4
            .ColorIndex = -4105
        End With
        With MyWRKBook.ActiveSheet.Range("A" & StrNum & ":N" & StrNum).Borders(8)
            .LineStyle = 1
            .Weight = 4
            .ColorIndex = -4105
        End With
        With MyWRKBook.ActiveSheet.Range("A" & StrNum & ":N" & StrNum).Borders(9)
            .LineStyle = 1
            .Weight = 4
            .ColorIndex = -4105
        End With
        With MyWRKBook.ActiveSheet.Range("A" & StrNum & ":N" & StrNum).Borders(10)
            .LineStyle = 1
            .Weight = 4
            .ColorIndex = -4105
        End With
        With MyWRKBook.ActiveSheet.Range("A" & StrNum & ":N" & StrNum).Borders(11)
            .LineStyle = 1
            .Weight = 4
            .ColorIndex = -4105
        End With
        With MyWRKBook.ActiveSheet.Range("A" & StrNum & ":N" & StrNum).Interior
            .ColorIndex = 35
            .Pattern = 1
            .PatternColorIndex = -4105
        End With

        StrNum = StrNum + 1
        Return StrNum
    End Function

    Private Function UploadDCHeader_LO(ByRef oSheet As Object, ByRef oServiceManager As Object, ByRef oWorkBook As Object, _
        ByRef oDispatcher As Object, ByVal StrNum As Double) As Double
        '////////////////////////////////////////////////////////////////////////////////////////
        '//
        '// ��������� �������� � Excel ��������� �� ���� ������� 
        '//
        '////////////////////////////////////////////////////////////////////////////////////////

        oSheet.getCellRangeByName("A" & StrNum).String = "ID"
        oSheet.getCellRangeByName("B" & StrNum).String = "�������� ������"
        oSheet.getCellRangeByName("C" & StrNum).String = "DC"
        oSheet.getCellRangeByName("D" & StrNum).String = "ABC"
        oSheet.getCellRangeByName("E" & StrNum).String = "XYZ"
        oSheet.getCellRangeByName("F" & StrNum).String = "����� ��������"
        oSheet.getCellRangeByName("G" & StrNum).String = "����� ����� ��������"
        oSheet.getCellRangeByName("H" & StrNum).String = "���� ���"
        oSheet.getCellRangeByName("I" & StrNum).String = "���� ROP"
        oSheet.getCellRangeByName("J" & StrNum).String = "���� ����� �������"
        oSheet.getCellRangeByName("K" & StrNum).String = "������ ���"
        oSheet.getCellRangeByName("L" & StrNum).String = "������ ROP"
        oSheet.getCellRangeByName("M" & StrNum).String = "������ ����� �������"
        oSheet.getCellRangeByName("N" & StrNum).String = "��������� �� ����"

        oSheet.getCellRangeByName("A" & StrNum & ":N" & StrNum).CellBackColor = 12510163
        oSheet.getCellRangeByName("A" & StrNum & ":N" & StrNum).VertJustify = 2
        oSheet.getCellRangeByName("A" & StrNum & ":N" & StrNum).HoriJustify = 2

        Dim LineFormat As Object
        LineFormat = oServiceManager.Bridge_GetStruct("com.sun.star.table.BorderLine2")
        LineFormat.LineStyle = 0
        LineFormat.LineWidth = 70
        oSheet.getCellRangeByName("A" & StrNum & ":N" & StrNum).TopBorder = LineFormat
        oSheet.getCellRangeByName("A" & StrNum & ":N" & StrNum).RightBorder = LineFormat
        oSheet.getCellRangeByName("A" & StrNum & ":N" & StrNum).LeftBorder = LineFormat
        oSheet.getCellRangeByName("A" & StrNum & ":N" & StrNum).BottomBorder = LineFormat

        Dim args() As Object
        ReDim args(0)
        args(0) = oServiceManager.Bridge_GetStruct("com.sun.star.beans.PropertyValue")
        args(0).Name = "ToPoint"
        args(0).Value = "A" & StrNum & ":N" & StrNum
        Dim oFrame As Object
        oFrame = oWorkBook.getCurrentController.getFrame
        oDispatcher.executeDispatch(oFrame, ".uno:GoToCell", "", 0, args)

        Dim args1() As Object
        ReDim args1(5)
        args1(0) = oServiceManager.Bridge_GetStruct("com.sun.star.beans.PropertyValue")
        args1(0).Name = "CharFontName.StyleName"
        args1(0).Value = "�������"
        args1(1) = oServiceManager.Bridge_GetStruct("com.sun.star.beans.PropertyValue")
        args1(1).Name = "CharFontName.Pitch"
        args1(1).Value = 2
        args1(2) = oServiceManager.Bridge_GetStruct("com.sun.star.beans.PropertyValue")
        args1(2).Name = "CharFontName.CharSet"
        args1(2).Value = 0
        args1(3) = oServiceManager.Bridge_GetStruct("com.sun.star.beans.PropertyValue")
        args1(3).Name = "CharFontName.Family"
        args1(3).Value = 5
        args1(4) = oServiceManager.Bridge_GetStruct("com.sun.star.beans.PropertyValue")
        args1(4).Name = "CharFontName.FamilyName"
        args1(4).Value = "Calibri"
        oDispatcher.executeDispatch(oFrame, ".uno:CharFontName", "", 0, args1)

        Dim args2() As Object
        ReDim args2(0)
        args2(0) = oServiceManager.Bridge_GetStruct("com.sun.star.beans.PropertyValue")
        args2(0).Name = "Bold"
        args2(0).Value = True
        oDispatcher.executeDispatch(oFrame, ".uno:Bold", "", 0, args2)

        Dim args3() As Object
        ReDim args3(0)
        args3(0) = oServiceManager.Bridge_GetStruct("com.sun.star.beans.PropertyValue")
        args3(0).Name = "WrapText"
        args3(0).Value = True
        oDispatcher.executeDispatch(oFrame, ".uno:WrapText", "", 0, args2)

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


        MySQLStr = "SELECT View_2.Code, View_2.Name, View_2.DC, View_2.ABC, View_2.XYZ, View_2.LT, View_2.OI, View_2.MGZ AS AMGZ, View_2.ROP AS AROP, "
        MySQLStr = MySQLStr & "View_2.InshuranceLVL AS AInshuranceLVL, CASE WHEN View_1.MGZ IS NULL THEN '' ELSE CONVERT(nvarchar(30), View_1.MGZ) END AS MGZ, "
        MySQLStr = MySQLStr & "CASE WHEN View_1.ROP IS NULL THEN '' ELSE CONVERT(nvarchar(30), View_1.ROP) END AS ROP, "
        MySQLStr = MySQLStr & "CASE WHEN View_1.IshuranceLVL IS NULL THEN '' ELSE CONVERT(nvarchar(30), View_1.IshuranceLVL) END AS IshuranceLVL, "
        MySQLStr = MySQLStr & "CASE WHEN View_1.DueDate IS NULL THEN '' ELSE CONVERT(nvarchar(30), View_1.DueDate, 103) END AS DueDate "
        MySQLStr = MySQLStr & "FROM (SELECT Code, Name, DC, ABC, XYZ, LT, OI, MGZ, ROP, InshuranceLVL "
        MySQLStr = MySQLStr & "FROM tbl_ForecastOrderR4_Main_RWH "
        MySQLStr = MySQLStr & "WHERE (WarNo = N'" & WHCode & "') AND "
        MySQLStr = MySQLStr & "(WHass = 1) AND (DC <> WarNo)) AS View_2 LEFT OUTER JOIN "
        MySQLStr = MySQLStr & "(SELECT Code, MGZ, ROP, IshuranceLVL, DueDate "
        MySQLStr = MySQLStr & "FROM tbl_ForecastOrderR4_CustomMGZROPINS_RWH "
        MySQLStr = MySQLStr & "WHERE (WH = N'" & WHCode & "')) "
        MySQLStr = MySQLStr & "AS View_1 ON View_2.Code = View_1.Code "
        MySQLStr = MySQLStr & "ORDER BY View_2.Code "

        InitMyConn(False)
        InitMyRec(False, MySQLStr)
        If (Declarations.MyRec.EOF = True And Declarations.MyRec.BOF = True) Then
        Else
            Declarations.MyRec.MoveFirst()
            While Declarations.MyRec.EOF <> True
                MyWRKBook.ActiveSheet.Range("B" & StrNum) = "'" & Declarations.MyRec.Fields("Code").Value
                MyWRKBook.ActiveSheet.Range("C" & StrNum) = Declarations.MyRec.Fields("Name").Value
                MyWRKBook.ActiveSheet.Range("D" & StrNum).NumberFormat = "@"
                MyWRKBook.ActiveSheet.Range("D" & StrNum) = Declarations.MyRec.Fields("DC").Value
                MyWRKBook.ActiveSheet.Range("E" & StrNum) = Declarations.MyRec.Fields("ABC").Value
                MyWRKBook.ActiveSheet.Range("F" & StrNum) = Declarations.MyRec.Fields("XYZ").Value
                MyWRKBook.ActiveSheet.Range("G" & StrNum) = Declarations.MyRec.Fields("LT").Value
                MyWRKBook.ActiveSheet.Range("H" & StrNum) = Declarations.MyRec.Fields("OI").Value
                MyWRKBook.ActiveSheet.Range("I" & StrNum).NumberFormatLocal = "#" & MySep & "##0" & MyDig & "00"
                If Declarations.MyRec.Fields("AMGZ").Value = 0 Then
                    MyWRKBook.ActiveSheet.Range("I" & StrNum).Interior.Color = 65535
                End If
                MyWRKBook.ActiveSheet.Range("I" & StrNum) = Declarations.MyRec.Fields("AMGZ").Value
                MyWRKBook.ActiveSheet.Range("J" & StrNum).NumberFormatLocal = "#" & MySep & "##0" & MyDig & "00"
                If Declarations.MyRec.Fields("AROP").Value = 0 Then
                    MyWRKBook.ActiveSheet.Range("J" & StrNum).Interior.Color = 65535
                End If
                MyWRKBook.ActiveSheet.Range("J" & StrNum) = Declarations.MyRec.Fields("AROP").Value
                MyWRKBook.ActiveSheet.Range("K" & StrNum).NumberFormatLocal = "#" & MySep & "##0" & MyDig & "00"
                If Declarations.MyRec.Fields("AInshuranceLVL").Value = 0 Then
                    MyWRKBook.ActiveSheet.Range("K" & StrNum).Interior.Color = 65535
                End If
                MyWRKBook.ActiveSheet.Range("K" & StrNum) = Declarations.MyRec.Fields("AInshuranceLVL").Value
                MyWRKBook.ActiveSheet.Range("L" & StrNum).NumberFormatLocal = "#" & MySep & "##0" & MyDig & "00"
                MyWRKBook.ActiveSheet.Range("L" & StrNum) = Declarations.MyRec.Fields("MGZ").Value
                MyWRKBook.ActiveSheet.Range("M" & StrNum).NumberFormatLocal = "#" & MySep & "##0" & MyDig & "00"
                MyWRKBook.ActiveSheet.Range("M" & StrNum) = Declarations.MyRec.Fields("ROP").Value
                MyWRKBook.ActiveSheet.Range("N" & StrNum).NumberFormatLocal = "#" & MySep & "##0" & MyDig & "00"
                MyWRKBook.ActiveSheet.Range("N" & StrNum) = Declarations.MyRec.Fields("IshuranceLVL").Value
                MyWRKBook.ActiveSheet.Range("O" & StrNum) = Declarations.MyRec.Fields("DueDate").Value


                StrNum = StrNum + 1
                Declarations.MyRec.MoveNext()
            End While
            trycloseMyRec()
        End If

        StrNum = StrNum + 2
        Return StrNum
    End Function

    Private Function UploadWHRows_LO(ByRef oSheet As Object, ByRef oServiceManager As Object, ByRef oWorkBook As Object, _
        ByRef oDispatcher As Object, ByVal WHCode As String, ByVal StrNum As Double) As Double
        '////////////////////////////////////////////////////////////////////////////////////////
        '//
        '// ��������� �������� � LibreOffice ����� �� ������ ������ 
        '//
        '////////////////////////////////////////////////////////////////////////////////////////
        Dim MySQLStr As String
        Dim StartStr As Double
        Dim NullDate As DateTime

        NullDate = New DateTime(1900, 1, 1)
        StartStr = StrNum
        MySQLStr = "SELECT View_2.Code, View_2.Name, View_2.DC, View_2.ABC, View_2.XYZ, View_2.LT, View_2.OI, View_2.MGZ AS AMGZ, View_2.ROP AS AROP, "
        MySQLStr = MySQLStr & "View_2.InshuranceLVL AS AInshuranceLVL, ISNULL(View_1.MGZ, 0) AS MGZ, "
        MySQLStr = MySQLStr & "ISNULL(View_1.ROP, 0) AS ROP, "
        MySQLStr = MySQLStr & "ISNULL(View_1.IshuranceLVL, 0) AS IshuranceLVL, "
        MySQLStr = MySQLStr & "ISNULL(View_1.DueDate, Convert(datetime, '01/01/1900', 103)) AS DueDate "
        MySQLStr = MySQLStr & "FROM (SELECT Code, Name, DC, ABC, XYZ, LT, OI, MGZ, ROP, InshuranceLVL "
        MySQLStr = MySQLStr & "FROM tbl_ForecastOrderR4_Main_RWH "
        MySQLStr = MySQLStr & "WHERE (WarNo = N'" & WHCode & "') AND "
        MySQLStr = MySQLStr & "(WHass = 1) AND (DC <> WarNo)) AS View_2 LEFT OUTER JOIN "
        MySQLStr = MySQLStr & "(SELECT Code, MGZ, ROP, IshuranceLVL, DueDate "
        MySQLStr = MySQLStr & "FROM tbl_ForecastOrderR4_CustomMGZROPINS_RWH "
        MySQLStr = MySQLStr & "WHERE (WH = N'" & WHCode & "')) "
        MySQLStr = MySQLStr & "AS View_1 ON View_2.Code = View_1.Code "
        MySQLStr = MySQLStr & "ORDER BY View_2.Code "

        InitMyConn(False)
        InitMyRec(False, MySQLStr)
        If (Declarations.MyRec.EOF = True And Declarations.MyRec.BOF = True) Then
        Else
            Declarations.MyRec.MoveFirst()
            While Declarations.MyRec.EOF <> True
                oSheet.getCellRangeByName("B" & StrNum).String = Declarations.MyRec.Fields("Code").Value
                oSheet.getCellRangeByName("C" & StrNum).String = Declarations.MyRec.Fields("Name").Value
                oSheet.getCellRangeByName("D" & StrNum).String = Declarations.MyRec.Fields("DC").Value
                oSheet.getCellRangeByName("E" & StrNum).String = Declarations.MyRec.Fields("ABC").Value
                oSheet.getCellRangeByName("F" & StrNum).String = Declarations.MyRec.Fields("XYZ").Value
                oSheet.getCellRangeByName("G" & StrNum).Value = Declarations.MyRec.Fields("LT").Value
                oSheet.getCellRangeByName("H" & StrNum).Value = Declarations.MyRec.Fields("OI").Value
                oSheet.getCellRangeByName("I" & StrNum).Value = Declarations.MyRec.Fields("AMGZ").Value
                If Declarations.MyRec.Fields("AMGZ").Value = 0 Then
                    oSheet.getCellRangeByName("I" & StrNum).CellBackColor = 16775598
                End If
                oSheet.getCellRangeByName("J" & StrNum).Value = Declarations.MyRec.Fields("AROP").Value
                If Declarations.MyRec.Fields("AROP").Value = 0 Then
                    oSheet.getCellRangeByName("J" & StrNum).CellBackColor = 16775598
                End If
                oSheet.getCellRangeByName("K" & StrNum).Value = Declarations.MyRec.Fields("AInshuranceLVL").Value
                If Declarations.MyRec.Fields("AInshuranceLVL").Value = 0 Then
                    oSheet.getCellRangeByName("K" & StrNum).CellBackColor = 16775598
                End If
                If Declarations.MyRec.Fields("MGZ").Value <> 0 Then
                    oSheet.getCellRangeByName("L" & StrNum).Value = Declarations.MyRec.Fields("MGZ").Value
                End If
                If Declarations.MyRec.Fields("ROP").Value <> 0 Then
                    oSheet.getCellRangeByName("M" & StrNum).Value = Declarations.MyRec.Fields("ROP").Value
                End If
                If Declarations.MyRec.Fields("IshuranceLVL").Value <> 0 Then
                    oSheet.getCellRangeByName("N" & StrNum).Value = Declarations.MyRec.Fields("IshuranceLVL").Value
                End If
                If Declarations.MyRec.Fields("DueDate").Value <> NullDate Then
                    oSheet.getCellRangeByName("O" & StrNum).Value = Declarations.MyRec.Fields("DueDate").Value
                End If

                StrNum = StrNum + 1
                Declarations.MyRec.MoveNext()
            End While
            trycloseMyRec()

            '----�����
            Dim args() As Object
            ReDim args(0)
            args(0) = oServiceManager.Bridge_GetStruct("com.sun.star.beans.PropertyValue")
            args(0).Name = "ToPoint"
            args(0).Value = "B" & StartStr & ":O" & StrNum
            Dim oFrame As Object
            oFrame = oWorkBook.getCurrentController.getFrame
            oDispatcher.executeDispatch(oFrame, ".uno:GoToCell", "", 0, args)

            Dim args1() As Object
            ReDim args1(5)
            args1(0) = oServiceManager.Bridge_GetStruct("com.sun.star.beans.PropertyValue")
            args1(0).Name = "CharFontName.StyleName"
            args1(0).Value = "�������"
            args1(1) = oServiceManager.Bridge_GetStruct("com.sun.star.beans.PropertyValue")
            args1(1).Name = "CharFontName.Pitch"
            args1(1).Value = 2
            args1(2) = oServiceManager.Bridge_GetStruct("com.sun.star.beans.PropertyValue")
            args1(2).Name = "CharFontName.CharSet"
            args1(2).Value = 0
            args1(3) = oServiceManager.Bridge_GetStruct("com.sun.star.beans.PropertyValue")
            args1(3).Name = "CharFontName.Family"
            args1(3).Value = 5
            args1(4) = oServiceManager.Bridge_GetStruct("com.sun.star.beans.PropertyValue")
            args1(4).Name = "CharFontName.FamilyName"
            args1(4).Value = "Calibri"
            oDispatcher.executeDispatch(oFrame, ".uno:CharFontName", "", 0, args1)

            '-----������
            args(0).Value = "B" & StartStr & ":N" & StrNum
            oDispatcher.executeDispatch(oFrame, ".uno:GoToCell", "", 0, args)

            Dim args2() As Object
            ReDim args2(0)
            args2(0) = oServiceManager.Bridge_GetStruct("com.sun.star.beans.PropertyValue")
            args2(0).Name = "NumberFormatValue"
            args2(0).Value = 4
            oDispatcher.executeDispatch(oFrame, ".uno:NumberFormatValue", "", 0, args2)

            args(0).Value = "O" & StartStr & ":O" & StrNum
            oDispatcher.executeDispatch(oFrame, ".uno:GoToCell", "", 0, args)

            args2(0).Name = "NumberFormatValue"
            args2(0).Value = 36
            oDispatcher.executeDispatch(oFrame, ".uno:NumberFormatValue", "", 0, args2)
        End If

        StrNum = StrNum + 2
        Return StrNum
    End Function

    Private Function UploadDCRows(ByVal MyWRKBook As Object, ByVal StrNum As Double) As Double
        '////////////////////////////////////////////////////////////////////////////////////////
        '//
        '// ��������� �������� � Excel ����� �� ���� ������� 
        '//
        '////////////////////////////////////////////////////////////////////////////////////////
        Dim MySQLStr As String
        Dim aa As New System.Globalization.NumberFormatInfo
        Dim MySep As String
        Dim MyDig As String

        MySep = aa.CurrentInfo.NumberGroupSeparator
        MyDig = aa.CurrentInfo.NumberDecimalSeparator


        MySQLStr = "SELECT View_2.Code, View_2.Name, View_2.DC, View_2.ABC, View_2.XYZ, View_2.LT, View_2.OI, View_2.AMGZ, View_2.AROP, View_2.AInshuranceLVL, "
        MySQLStr = MySQLStr & "CASE WHEN tbl_ForecastOrderR4_CustomMGZROPINS_DC.MGZ IS NULL THEN '' ELSE CONVERT(nvarchar(30), tbl_ForecastOrderR4_CustomMGZROPINS_DC.MGZ) END AS MGZ, "
        MySQLStr = MySQLStr & "CASE WHEN tbl_ForecastOrderR4_CustomMGZROPINS_DC.ROP IS NULL THEN '' ELSE CONVERT(nvarchar(30), tbl_ForecastOrderR4_CustomMGZROPINS_DC.ROP) END AS ROP, "
        MySQLStr = MySQLStr & "CASE WHEN tbl_ForecastOrderR4_CustomMGZROPINS_DC.IshuranceLVL IS NULL THEN '' ELSE CONVERT(nvarchar(30), tbl_ForecastOrderR4_CustomMGZROPINS_DC.IshuranceLVL) END AS IshuranceLVL, "
        MySQLStr = MySQLStr & "CASE WHEN tbl_ForecastOrderR4_CustomMGZROPINS_DC.DueDate IS NULL THEN '' ELSE CONVERT(nvarchar(30), tbl_ForecastOrderR4_CustomMGZROPINS_DC.DueDate, 103) END AS DueDate "
        MySQLStr = MySQLStr & "FROM tbl_ForecastOrderR4_CustomMGZROPINS_DC RIGHT OUTER JOIN "
        MySQLStr = MySQLStr & "(SELECT Code, Name, DC, ABC, XYZ, LT, OI, MGZ AS AMGZ, ROP AS AROP, InshuranceLVL AS AInshuranceLVL "
        MySQLStr = MySQLStr & "FROM tbl_ForecastOrderR4_Main_DC "
        MySQLStr = MySQLStr & "WHERE (WHass = 1)) AS  View_2 ON tbl_ForecastOrderR4_CustomMGZROPINS_DC.Code = View_2.Code "
        MySQLStr = MySQLStr & "ORDER BY View_2.Code "

        InitMyConn(False)
        InitMyRec(False, MySQLStr)
        If (Declarations.MyRec.EOF = True And Declarations.MyRec.BOF = True) Then
        Else
            Declarations.MyRec.MoveFirst()
            While Declarations.MyRec.EOF <> True
                MyWRKBook.ActiveSheet.Range("A" & StrNum) = "'" & Declarations.MyRec.Fields("Code").Value
                MyWRKBook.ActiveSheet.Range("B" & StrNum) = Declarations.MyRec.Fields("Name").Value
                MyWRKBook.ActiveSheet.Range("C" & StrNum).NumberFormat = "@"
                MyWRKBook.ActiveSheet.Range("C" & StrNum) = Declarations.MyRec.Fields("DC").Value
                MyWRKBook.ActiveSheet.Range("D" & StrNum) = Declarations.MyRec.Fields("ABC").Value
                MyWRKBook.ActiveSheet.Range("E" & StrNum) = Declarations.MyRec.Fields("XYZ").Value
                MyWRKBook.ActiveSheet.Range("F" & StrNum) = Declarations.MyRec.Fields("LT").Value
                MyWRKBook.ActiveSheet.Range("G" & StrNum) = Declarations.MyRec.Fields("OI").Value
                MyWRKBook.ActiveSheet.Range("H" & StrNum).NumberFormatLocal = "#" & MySep & "##0" & MyDig & "00"
                If Declarations.MyRec.Fields("AMGZ").Value = 0 Then
                    MyWRKBook.ActiveSheet.Range("H" & StrNum).Interior.Color = 65535
                End If
                MyWRKBook.ActiveSheet.Range("H" & StrNum) = Declarations.MyRec.Fields("AMGZ").Value
                MyWRKBook.ActiveSheet.Range("I" & StrNum).NumberFormatLocal = "#" & MySep & "##0" & MyDig & "00"
                If Declarations.MyRec.Fields("AROP").Value = 0 Then
                    MyWRKBook.ActiveSheet.Range("I" & StrNum).Interior.Color = 65535
                End If
                MyWRKBook.ActiveSheet.Range("I" & StrNum) = Declarations.MyRec.Fields("AROP").Value
                MyWRKBook.ActiveSheet.Range("J" & StrNum).NumberFormatLocal = "#" & MySep & "##0" & MyDig & "00"
                If Declarations.MyRec.Fields("AInshuranceLVL").Value = 0 Then
                    MyWRKBook.ActiveSheet.Range("J" & StrNum).Interior.Color = 65535
                End If
                MyWRKBook.ActiveSheet.Range("J" & StrNum) = Declarations.MyRec.Fields("AInshuranceLVL").Value
                MyWRKBook.ActiveSheet.Range("K" & StrNum).NumberFormatLocal = "#" & MySep & "##0" & MyDig & "00"
                MyWRKBook.ActiveSheet.Range("K" & StrNum) = Declarations.MyRec.Fields("MGZ").Value
                MyWRKBook.ActiveSheet.Range("L" & StrNum).NumberFormatLocal = "#" & MySep & "##0" & MyDig & "00"
                MyWRKBook.ActiveSheet.Range("L" & StrNum) = Declarations.MyRec.Fields("ROP").Value
                MyWRKBook.ActiveSheet.Range("M" & StrNum).NumberFormatLocal = "#" & MySep & "##0" & MyDig & "00"
                MyWRKBook.ActiveSheet.Range("M" & StrNum) = Declarations.MyRec.Fields("IshuranceLVL").Value
                MyWRKBook.ActiveSheet.Range("N" & StrNum) = Declarations.MyRec.Fields("DueDate").Value


                StrNum = StrNum + 1
                Declarations.MyRec.MoveNext()
            End While
            trycloseMyRec()
        End If

        StrNum = StrNum + 2
        Return StrNum
    End Function

    Private Function UploadDCRows_LO(ByRef oSheet As Object, ByRef oServiceManager As Object, ByRef oWorkBook As Object, _
        ByRef oDispatcher As Object, ByVal StrNum As Double) As Double
        '////////////////////////////////////////////////////////////////////////////////////////
        '//
        '// ��������� �������� � LibreOffice ����� �� ���� ������� 
        '//
        '////////////////////////////////////////////////////////////////////////////////////////
        Dim MySQLStr As String
        Dim StartStr As Double
        Dim NullDate As DateTime

        NullDate = New DateTime(1900, 1, 1)
        StartStr = StrNum
        MySQLStr = "SELECT View_2.Code, View_2.Name, View_2.DC, View_2.ABC, View_2.XYZ, View_2.LT, View_2.OI, View_2.AMGZ, View_2.AROP, View_2.AInshuranceLVL, "
        MySQLStr = MySQLStr & "ISNULL(tbl_ForecastOrderR4_CustomMGZROPINS_DC.MGZ, 0) AS MGZ, "
        MySQLStr = MySQLStr & "ISNULL(tbl_ForecastOrderR4_CustomMGZROPINS_DC.ROP, 0) AS ROP, "
        MySQLStr = MySQLStr & "ISNULL(tbl_ForecastOrderR4_CustomMGZROPINS_DC.IshuranceLVL, 0) AS IshuranceLVL, "
        MySQLStr = MySQLStr & "ISNULL(tbl_ForecastOrderR4_CustomMGZROPINS_DC.DueDate, Convert(datetime, '01/01/1900', 103)) AS DueDate "
        MySQLStr = MySQLStr & "FROM tbl_ForecastOrderR4_CustomMGZROPINS_DC RIGHT OUTER JOIN "
        MySQLStr = MySQLStr & "(SELECT Code, Name, DC, ABC, XYZ, LT, OI, MGZ AS AMGZ, ROP AS AROP, InshuranceLVL AS AInshuranceLVL "
        MySQLStr = MySQLStr & "FROM tbl_ForecastOrderR4_Main_DC "
        MySQLStr = MySQLStr & "WHERE (WHass = 1)) AS  View_2 ON tbl_ForecastOrderR4_CustomMGZROPINS_DC.Code = View_2.Code "
        MySQLStr = MySQLStr & "ORDER BY View_2.Code "

        InitMyConn(False)
        InitMyRec(False, MySQLStr)
        If (Declarations.MyRec.EOF = True And Declarations.MyRec.BOF = True) Then
        Else
            Declarations.MyRec.MoveFirst()
            While Declarations.MyRec.EOF <> True
                oSheet.getCellRangeByName("A" & StrNum).String = Declarations.MyRec.Fields("Code").Value
                oSheet.getCellRangeByName("B" & StrNum).String = Declarations.MyRec.Fields("Name").Value
                oSheet.getCellRangeByName("C" & StrNum).String = Declarations.MyRec.Fields("DC").Value
                oSheet.getCellRangeByName("D" & StrNum).String = Declarations.MyRec.Fields("ABC").Value
                oSheet.getCellRangeByName("E" & StrNum).String = Declarations.MyRec.Fields("XYZ").Value
                oSheet.getCellRangeByName("F" & StrNum).Value = Declarations.MyRec.Fields("LT").Value
                oSheet.getCellRangeByName("G" & StrNum).Value = Declarations.MyRec.Fields("OI").Value
                oSheet.getCellRangeByName("H" & StrNum).Value = Declarations.MyRec.Fields("AMGZ").Value
                If Declarations.MyRec.Fields("AMGZ").Value = 0 Then
                    oSheet.getCellRangeByName("H" & StrNum).CellBackColor = 16775598
                End If
                oSheet.getCellRangeByName("I" & StrNum).Value = Declarations.MyRec.Fields("AROP").Value
                If Declarations.MyRec.Fields("AROP").Value = 0 Then
                    oSheet.getCellRangeByName("I" & StrNum).CellBackColor = 16775598
                End If
                oSheet.getCellRangeByName("J" & StrNum).Value = Declarations.MyRec.Fields("AInshuranceLVL").Value
                If Declarations.MyRec.Fields("AInshuranceLVL").Value = 0 Then
                    oSheet.getCellRangeByName("J" & StrNum).CellBackColor = 16775598
                End If
                If Declarations.MyRec.Fields("MGZ").Value <> 0 Then
                    oSheet.getCellRangeByName("K" & StrNum).Value = Declarations.MyRec.Fields("MGZ").Value
                End If
                If Declarations.MyRec.Fields("ROP").Value <> 0 Then
                    oSheet.getCellRangeByName("L" & StrNum).Value = Declarations.MyRec.Fields("ROP").Value
                End If
                If Declarations.MyRec.Fields("IshuranceLVL").Value <> 0 Then
                    oSheet.getCellRangeByName("M" & StrNum).Value = Declarations.MyRec.Fields("IshuranceLVL").Value
                End If
                If Declarations.MyRec.Fields("DueDate").Value <> NullDate Then
                    oSheet.getCellRangeByName("N" & StrNum).Value = Declarations.MyRec.Fields("DueDate").Value
                End If

                StrNum = StrNum + 1
                Declarations.MyRec.MoveNext()
            End While
            trycloseMyRec()

            '----�����
            Dim args() As Object
            ReDim args(0)
            args(0) = oServiceManager.Bridge_GetStruct("com.sun.star.beans.PropertyValue")
            args(0).Name = "ToPoint"
            args(0).Value = "A" & StartStr & ":N" & StrNum
            Dim oFrame As Object
            oFrame = oWorkBook.getCurrentController.getFrame
            oDispatcher.executeDispatch(oFrame, ".uno:GoToCell", "", 0, args)

            Dim args1() As Object
            ReDim args1(5)
            args1(0) = oServiceManager.Bridge_GetStruct("com.sun.star.beans.PropertyValue")
            args1(0).Name = "CharFontName.StyleName"
            args1(0).Value = "�������"
            args1(1) = oServiceManager.Bridge_GetStruct("com.sun.star.beans.PropertyValue")
            args1(1).Name = "CharFontName.Pitch"
            args1(1).Value = 2
            args1(2) = oServiceManager.Bridge_GetStruct("com.sun.star.beans.PropertyValue")
            args1(2).Name = "CharFontName.CharSet"
            args1(2).Value = 0
            args1(3) = oServiceManager.Bridge_GetStruct("com.sun.star.beans.PropertyValue")
            args1(3).Name = "CharFontName.Family"
            args1(3).Value = 5
            args1(4) = oServiceManager.Bridge_GetStruct("com.sun.star.beans.PropertyValue")
            args1(4).Name = "CharFontName.FamilyName"
            args1(4).Value = "Calibri"
            oDispatcher.executeDispatch(oFrame, ".uno:CharFontName", "", 0, args1)

            '-----������
            args(0).Value = "A" & StartStr & ":M" & StrNum
            oDispatcher.executeDispatch(oFrame, ".uno:GoToCell", "", 0, args)

            Dim args2() As Object
            ReDim args2(0)
            args2(0) = oServiceManager.Bridge_GetStruct("com.sun.star.beans.PropertyValue")
            args2(0).Name = "NumberFormatValue"
            args2(0).Value = 4
            oDispatcher.executeDispatch(oFrame, ".uno:NumberFormatValue", "", 0, args2)

            args(0).Value = "N" & StartStr & ":N" & StrNum
            oDispatcher.executeDispatch(oFrame, ".uno:GoToCell", "", 0, args)

            args2(0).Name = "NumberFormatValue"
            args2(0).Value = 36
            oDispatcher.executeDispatch(oFrame, ".uno:NumberFormatValue", "", 0, args2)
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

        MyTxt = "��� ������� ������ �� ������ ������ ��� ���������� ����������� ���� Excel, � ������� � ������ B1 ������� ����� ������ (� �������������� 0). " & Chr(13) & Chr(10)
        MyTxt = MyTxt & "������ � ����� Excel ������ ���������� � 5 ������, � ������� 'A'. ������ ������ ���� ��������� ��� ���������. " & Chr(13) & Chr(10)
        MyTxt = MyTxt & "� ������� 'A' ������ ������������� ���� ������� � ��������������� ������ " & Chr(13) & Chr(10)
        MyTxt = MyTxt & "� �������� 'B', 'C' � 'D' ������ ������������� ����� ���������� ������� ��������. " & Chr(13) & Chr(10)
        MyTxt = MyTxt & "���, ROP � ������� ���������� ������." & Chr(13) & Chr(10)
        MyTxt = MyTxt & "� ������� 'E' ������ ���� ������� ����, �� ������� ��������� ���������." & Chr(13) & Chr(10)
        MyTxt = MyTxt & "��� ������� ������ ���� ���������." & Chr(13) & Chr(10)
        MyTxt = MyTxt & "� ��� ���� �������������� ���� Excel � �� ������ ������ ������?" & Chr(13) & Chr(10)
        MyRez = MsgBox(MyTxt, MsgBoxStyle.OkCancel, "��������!")
        If (MyRez = MsgBoxResult.Ok) Then
            If My.Settings.UseOffice = "LibreOffice" Then
                ImportDataFromLO()
            Else
                ImportDataFromExcel()
            End If
            '---������ ������� �� ������� ���, ROP � ��������� ����� ��������� �������������
            BuildAutoItemList()
            '---������ ������� �� ������� ���, ROP � ��������� ����� ������������ �������
            BuildManualItemList()
            CheckButtons()
            SetWindowPos(Me.Handle.ToInt32, -2, 0, 0, 0, 0, &H3)
            Me.Cursor = Cursors.Default
            MsgBox("������ ������ �������� ���, ��� � ���������� ������ �� ��������� (������������) ������� ����������.", MsgBoxStyle.Information, "��������!")
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
        Dim MyDueDate As Date
        Dim MyMonthNum As Double       '--���-�� ������� ��� ������� MGZ, ROP �  �.�.

        OpenFileDialog1.ShowDialog()
        If (OpenFileDialog1.FileName = "") Then
        Else
            Me.Cursor = Cursors.WaitCursor
            Me.Refresh()
            System.Windows.Forms.Application.DoEvents()

            appXLSRC = CreateObject("Excel.Application")
            appXLSRC.Workbooks.Open(OpenFileDialog1.FileName)
            MyWH = appXLSRC.Worksheets(1).Range("B1").Value

            '---��������� ��� � Excel ���������� ��� ������
            If MyWH = Nothing Then
                MsgBox("� ������������� ����� Excel � ������ 'B1' �� ���������� ��� ������ � Scala ", MsgBoxStyle.Critical, "��������!")
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
                MsgBox("� ������������� ����� Excel � ������ 'B1' ���������� �������� ��� ������ � Scala ", MsgBoxStyle.Critical, "��������!")
                appXLSRC.DisplayAlerts = 0
                appXLSRC.Workbooks.Close()
                appXLSRC.DisplayAlerts = 1
                appXLSRC.Quit()
                appXLSRC = Nothing
                trycloseMyRec()
                Exit Sub
            End If
            trycloseMyRec()

            '---�������� �������� �� ������� �������
            MySQLStr = "DELETE FROM tbl_ForecastOrderR4_CustomMGZROPINS_RWH "
            MySQLStr = MySQLStr & "WHERE (WH = N'" & MyWH & "')"
            InitMyConn(False)
            Declarations.MyConn.Execute(MySQLStr)


            StrCnt = 5
            While Not appXLSRC.Worksheets(1).Range("A" & StrCnt).Value = Nothing
                MyCode = appXLSRC.Worksheets(1).Range("A" & StrCnt).Value.ToString
                Try
                    MyMGZ = appXLSRC.Worksheets(1).Range("B" & StrCnt).Value
                    Try
                        MyROP = appXLSRC.Worksheets(1).Range("C" & StrCnt).Value
                        Try
                            MyInsLVL = appXLSRC.Worksheets(1).Range("D" & StrCnt).Value
                            Try
                                MyDueDate = appXLSRC.Worksheets(1).Range("E" & StrCnt).Value
                                '---��������� ���� - ������ ���� ������ ������� � ������ ��� ������� + �������� �������
                                MySQLStr = "SELECT MonthNum "
                                MySQLStr = MySQLStr & "FROM tbl_ForecastOrder_PeriodQTY "
                                InitMyConn(False)
                                InitMyRec(False, MySQLStr)
                                If (Declarations.MyRec.EOF = True And Declarations.MyRec.BOF = True) Then
                                    MyMonthNum = 12
                                Else
                                    MyMonthNum = Declarations.MyRec.Fields("MonthNum").Value
                                End If
                                trycloseMyRec()

                                If MyDueDate < Now() Then
                                    MsgBox("������ � ������ " & StrCnt & " ������� ""E"". ���� ""��������� �� ����:"" ������ ���� ������ �������.", MsgBoxStyle.Critical, "��������!")
                                Else
                                    If MyDueDate > DateAdd(DateInterval.Month, MyMonthNum, Now()) Then
                                        MsgBox("������ � ������ " & StrCnt & " ������� ""E"". ���� ""��������� �� ����:"" �� ������ ��������� ������� ������ ��� �� " & CStr(MyMonthNum) & " �������.", MsgBoxStyle.Critical, "��������!")
                                    Else
                                        '---��� �������� - ���� �� ����� ��� � ������ ��������� �� ������� ������
                                        MySQLStr = "SELECT COUNT(*) AS CC "
                                        MySQLStr = MySQLStr & "FROM tbl_ForecastOrderR4_Main_RWH "
                                        MySQLStr = MySQLStr & "WHERE (Code = N'" & Trim(MyCode) & "') "
                                        MySQLStr = MySQLStr & "AND (WarNo = N'" & Trim(MyWH) & "') "
                                        MySQLStr = MySQLStr & "AND (DC <> WarNo) "
                                        MySQLStr = MySQLStr & "AND (WHass = 1) "
                                        InitMyConn(False)
                                        InitMyRec(False, MySQLStr)
                                        If Declarations.MyRec.Fields("CC").Value = 0 Then
                                            trycloseMyRec()
                                            MsgBox("������ � ������ " & StrCnt & ". ����� �� �������� ��������� �� ������ ������ ��� ��� ������� ������ ����� DC ����� ������, ��� �������� �� ��������� ������ ������, ��� ������������.", MsgBoxStyle.Critical, "��������!")
                                        Else
                                            trycloseMyRec()
                                            '---������� � ������� �������
                                            MySQLStr = "INSERT INTO tbl_ForecastOrderR4_CustomMGZROPINS_RWH "
                                            MySQLStr = MySQLStr & "(ID, Code, WH, MGZ, ROP, IshuranceLVL, DueDate) "
                                            MySQLStr = MySQLStr & "VALUES (NEWID(), "
                                            MySQLStr = MySQLStr & "N'" & Trim(MyCode) & "', "
                                            MySQLStr = MySQLStr & "N'" & Trim(MyWH) & "', "
                                            MySQLStr = MySQLStr & Replace(CStr(MyMGZ), ",", ".") & ", "
                                            MySQLStr = MySQLStr & Replace(CStr(MyROP), ",", ".") & ", "
                                            MySQLStr = MySQLStr & Replace(CStr(MyInsLVL), ",", ".") & ", "
                                            MySQLStr = MySQLStr & "CONVERT(DATETIME, '" & FormatDateTime(MyDueDate, DateFormat.ShortDate) & "', 103)) "
                                            InitMyConn(False)
                                            Declarations.MyConn.Execute(MySQLStr)
                                        End If
                                    End If
                                End If
                            Catch ex As Exception
                                MsgBox("������ � ������ " & StrCnt & " ������� ""E"". " & ex.Message, MsgBoxStyle.Critical, "��������!")
                            End Try
                        Catch ex As Exception
                            MsgBox("������ � ������ " & StrCnt & " ������� ""D"". " & ex.Message, MsgBoxStyle.Critical, "��������!")
                        End Try
                    Catch ex As Exception
                        MsgBox("������ � ������ " & StrCnt & " ������� ""C"". " & ex.Message, MsgBoxStyle.Critical, "��������!")
                    End Try
                Catch ex As Exception
                    MsgBox("������ � ������ " & StrCnt & " ������� ""B"". " & ex.Message, MsgBoxStyle.Critical, "��������!")
                End Try
                StrCnt = StrCnt + 1
            End While
            ComboBox1.SelectedValue = MyWH
            appXLSRC.DisplayAlerts = 0
            appXLSRC.Workbooks.Close()
            appXLSRC.DisplayAlerts = 1
            appXLSRC.Quit()
            appXLSRC = Nothing
        End If
    End Sub

    Private Sub ImportDataFromLO()
        '////////////////////////////////////////////////////////////////////////////////////////
        '//
        '// �������� �� LibreOffice ����� �� ������ ������ 
        '//
        '////////////////////////////////////////////////////////////////////////////////////////
        Dim MyWH As String
        Dim MyCode As String
        Dim MyMGZ As Double
        Dim MyROP As Double
        Dim MyInsLVL As Double
        Dim StrCnt As Double
        Dim MySQLStr As String
        Dim MyDueDate As Date
        Dim MyMonthNum As Double       '--���-�� ������� ��� ������� MGZ, ROP �  �.�.
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
            Try
                LOSetNotation(0)
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

                '---��������� ��� � Excel ���������� ��� ������
                MyWH = oSheet.getCellRangeByName("B1").String
                If MyWH.Equals("") Then
                    MsgBox("� ������������� ����� Excel � ������ 'B1' �� ���������� ��� ������ � Scala ", MsgBoxStyle.Critical, "��������!")
                    oWorkBook.Close(True)
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
                    MsgBox("� ������������� ����� Excel � ������ 'B1' ���������� �������� ��� ������ � Scala ", MsgBoxStyle.Critical, "��������!")
                    oWorkBook.Close(True)
                    trycloseMyRec()
                    Exit Sub
                End If
                trycloseMyRec()

                '---�������� �������� �� ������� �������
                MySQLStr = "DELETE FROM tbl_ForecastOrderR4_CustomMGZROPINS_RWH "
                MySQLStr = MySQLStr & "WHERE (WH = N'" & MyWH & "')"
                InitMyConn(False)
                Declarations.MyConn.Execute(MySQLStr)

                StrCnt = 5
                StrCnt = 5
                While Not oSheet.getCellRangeByName("A" & StrCnt).String.Equals("")
                    MyCode = oSheet.getCellRangeByName("A" & StrCnt).String
                    Try
                        MyMGZ = oSheet.getCellRangeByName("B" & StrCnt).value
                        Try
                            MyROP = oSheet.getCellRangeByName("C" & StrCnt).value
                            Try
                                MyInsLVL = oSheet.getCellRangeByName("D" & StrCnt).value
                                Try
                                    MyDueDate = oSheet.getCellRangeByName("E" & StrCnt).String
                                    '---��������� ���� - ������ ���� ������ ������� � ������ ��� ������� + �������� �������
                                    MySQLStr = "SELECT MonthNum "
                                    MySQLStr = MySQLStr & "FROM tbl_ForecastOrder_PeriodQTY "
                                    InitMyConn(False)
                                    InitMyRec(False, MySQLStr)
                                    If (Declarations.MyRec.EOF = True And Declarations.MyRec.BOF = True) Then
                                        MyMonthNum = 12
                                    Else
                                        MyMonthNum = Declarations.MyRec.Fields("MonthNum").Value
                                    End If
                                    trycloseMyRec()

                                    If MyDueDate < Now() Then
                                        MsgBox("������ � ������ " & StrCnt & " ������� ""E"". ���� ""��������� �� ����:"" ������ ���� ������ �������.", MsgBoxStyle.Critical, "��������!")
                                    Else
                                        If MyDueDate > DateAdd(DateInterval.Month, MyMonthNum, Now()) Then
                                            MsgBox("������ � ������ " & StrCnt & " ������� ""E"". ���� ""��������� �� ����:"" �� ������ ��������� ������� ������ ��� �� " & CStr(MyMonthNum) & " �������.", MsgBoxStyle.Critical, "��������!")
                                        Else
                                            '---��� �������� - ���� �� ����� ��� � ������ ��������� �� ������� ������
                                            MySQLStr = "SELECT COUNT(*) AS CC "
                                            MySQLStr = MySQLStr & "FROM tbl_ForecastOrderR4_Main_RWH "
                                            MySQLStr = MySQLStr & "WHERE (Code = N'" & Trim(MyCode) & "') "
                                            MySQLStr = MySQLStr & "AND (WarNo = N'" & Trim(MyWH) & "') "
                                            MySQLStr = MySQLStr & "AND (DC <> WarNo) "
                                            MySQLStr = MySQLStr & "AND (WHass = 1) "
                                            InitMyConn(False)
                                            InitMyRec(False, MySQLStr)
                                            If Declarations.MyRec.Fields("CC").Value = 0 Then
                                                trycloseMyRec()
                                                MsgBox("������ � ������ " & StrCnt & ". ����� �� �������� ��������� �� ������ ������ ��� ��� ������� ������ ����� DC ����� ������, ��� �������� �� ��������� ������ ������, ��� ������������.", MsgBoxStyle.Critical, "��������!")
                                            Else
                                                trycloseMyRec()
                                                '---������� � ������� �������
                                                MySQLStr = "INSERT INTO tbl_ForecastOrderR4_CustomMGZROPINS_RWH "
                                                MySQLStr = MySQLStr & "(ID, Code, WH, MGZ, ROP, IshuranceLVL, DueDate) "
                                                MySQLStr = MySQLStr & "VALUES (NEWID(), "
                                                MySQLStr = MySQLStr & "N'" & Trim(MyCode) & "', "
                                                MySQLStr = MySQLStr & "N'" & Trim(MyWH) & "', "
                                                MySQLStr = MySQLStr & Replace(CStr(MyMGZ), ",", ".") & ", "
                                                MySQLStr = MySQLStr & Replace(CStr(MyROP), ",", ".") & ", "
                                                MySQLStr = MySQLStr & Replace(CStr(MyInsLVL), ",", ".") & ", "
                                                MySQLStr = MySQLStr & "CONVERT(DATETIME, '" & FormatDateTime(MyDueDate, DateFormat.ShortDate) & "', 103)) "
                                                InitMyConn(False)
                                                Declarations.MyConn.Execute(MySQLStr)
                                            End If
                                        End If
                                    End If
                                Catch ex As Exception
                                    MsgBox("������ � ������ " & StrCnt & " ������� ""E"". " & ex.Message, MsgBoxStyle.Critical, "��������!")
                                End Try
                            Catch ex As Exception
                                MsgBox("������ � ������ " & StrCnt & " ������� ""D"". " & ex.Message, MsgBoxStyle.Critical, "��������!")
                            End Try
                        Catch ex As Exception
                            MsgBox("������ � ������ " & StrCnt & " ������� ""C"". " & ex.Message, MsgBoxStyle.Critical, "��������!")
                        End Try
                    Catch ex As Exception
                        MsgBox("������ � ������ " & StrCnt & " ������� ""B"". " & ex.Message, MsgBoxStyle.Critical, "��������!")
                    End Try
                    StrCnt = StrCnt + 1
                End While


            Catch ex As Exception
                MsgBox("������ : " & ex.Message, MsgBoxStyle.Critical, "��������!")
            End Try
            Me.Cursor = Cursors.Default
            oWorkBook.Close(True)
            ComboBox1.SelectedValue = MyWH
            MsgBox("������ ������ ����������.", MsgBoxStyle.OkOnly, "��������!")
        End If
    End Sub

    Private Sub Button10_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button10.Click
        '////////////////////////////////////////////////////////////////////////////////////////
        '//
        '// ������ ������� �������������� �������� ���, ROP, ���������� ������
        '// ������� ���������� ��������� ���������� ������������
        '//
        '////////////////////////////////////////////////////////////////////////////////////////

        ReCalculate_All()

    End Sub

    Private Sub ReCalculate_All()
        '////////////////////////////////////////////////////////////////////////////////////////
        '//
        '// ��������� ������� �������������� �������� ���, ROP, ���������� ������
        '// �� ���� ������� � �� �������� ������ �������
        '//
        '////////////////////////////////////////////////////////////////////////////////////////
        Dim MySQLStr As String

        Me.Cursor = Cursors.WaitCursor
        MySQLStr = "Exec dbo.spp_ForecastOrderR4_Main_DC "
        InitMyConn(False)
        Declarations.MyConn.Execute(MySQLStr)

        MySQLStr = "Exec spp_ForecastOrderR4_Main_RWH "
        InitMyConn(False)
        Declarations.MyConn.Execute(MySQLStr)
        Me.Cursor = Cursors.Default

        '---������ ������� �� ������� ���, ROP � ��������� ����� ��������� �������������
        BuildAutoItemList()

        '---������ ������� �� ������� ���, ROP � ��������� ����� ������������ �������
        BuildManualItemList()

        CheckButtons()
        MsgBox("�������� ������� �������������� �������� ���, ROP � ���������� ������ ���������.", MsgBoxStyle.Information, "��������!")
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

    Private Sub ComboBox2_TextChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles ComboBox2.TextChanged
        '////////////////////////////////////////////////////////////////////////////////
        '//
        '// ����� ����� ������ ������ - ������������� ������
        '//
        '////////////////////////////////////////////////////////////////////////////////

        If ComboBox2.Text = "�� ���� ��������" Then
            Declarations.MyWorkLevel = 0
            ComboBox1.Enabled = False
            ComboBox1.Visible = False
            Label2.Visible = False
            Label8.Text = "���������� �� DC"
            Label8.ForeColor = Color.DarkGreen
            Label9.Text = "���������� �� DC"
            Label9.ForeColor = Color.DarkGreen
        Else
            Declarations.MyWorkLevel = 1
            ComboBox1.Enabled = True
            ComboBox1.Visible = True
            Label2.Visible = True
            Label8.Text = "�� ������ ������ - " & ComboBox1.Text
            Label8.ForeColor = Color.DarkBlue
            Label9.Text = "�� ������ ������ - " & ComboBox1.Text
            Label9.ForeColor = Color.DarkBlue
        End If

        '---������ ������� �� ������� ���, ROP � ��������� ����� ��������� �������������
        BuildAutoItemList()

        '---������ ������� �� ������� ���, ROP � ��������� ����� ������������ �������
        BuildManualItemList()

        CheckButtons()
    End Sub

    Private Sub DataGridView2_CellFormatting(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellFormattingEventArgs) Handles DataGridView2.CellFormatting
        '////////////////////////////////////////////////////////////////////////////////
        '//
        '// ��������� �������, � ���������� ����������
        '//
        '////////////////////////////////////////////////////////////////////////////////
        Dim row As DataGridViewRow

        row = DataGridView2.Rows(e.RowIndex)
        If row.Cells(7).Value <> 0 Or row.Cells(8).Value <> 0 Or row.Cells(9).Value <> 0 Then
            row.DefaultCellStyle.BackColor = Color.Yellow
        End If

    End Sub

    Private Sub Button15_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button15.Click
        '////////////////////////////////////////////////////////////////////////////////////////
        '//
        '// ��������� �������� �� Excel ����� �� ���� ������� 
        '//
        '////////////////////////////////////////////////////////////////////////////////////////
        Dim MyRez As MsgBoxResult
        Dim MyTxt As String

        MyTxt = "��� ������� ������ �� ���� ������� ��� ���������� ����������� ���� Excel. " & Chr(13) & Chr(10)
        MyTxt = MyTxt & "������ � ����� Excel ������ ���������� � 5 ������, � ������� 'A'. ������ ������ ���� ��������� ��� ���������. " & Chr(13) & Chr(10)
        MyTxt = MyTxt & "� ������� 'A' ������ ������������� ���� ������� � ��������������� ������ " & Chr(13) & Chr(10)
        MyTxt = MyTxt & "� �������� 'B', 'C' � 'D' ������ ������������� ����� ���������� ������� ��������. " & Chr(13) & Chr(10)
        MyTxt = MyTxt & "���, ROP � ������� ���������� ������." & Chr(13) & Chr(10)
        MyTxt = MyTxt & "� ������� 'E' ������ ���� ������� ����, �� ������� ��������� ���������." & Chr(13) & Chr(10)
        MyTxt = MyTxt & "��� ������� ������ ���� ���������." & Chr(13) & Chr(10)
        MyTxt = MyTxt & "� ��� ���� �������������� ���� Excel � �� ������ ������ ������?" & Chr(13) & Chr(10)
        MyRez = MsgBox(MyTxt, MsgBoxStyle.OkCancel, "��������!")
        If (MyRez = MsgBoxResult.Ok) Then
            If My.Settings.UseOffice = "LibreOffice" Then
                ImportDCDataFromLO()
            Else
                ImportDCDataFromExcel()
            End If
            '---������ ������� �� ������� ���, ROP � ��������� ����� ��������� �������������
            BuildAutoItemList()
            '---������ ������� �� ������� ���, ROP � ��������� ����� ������������ �������
            BuildManualItemList()
            CheckButtons()
            SetWindowPos(Me.Handle.ToInt32, -2, 0, 0, 0, 0, &H3)
            Me.Cursor = Cursors.Default
            MsgBox("������ ������ �������� ���, ��� � ���������� ������ �� ���� ������� ����������.", MsgBoxStyle.Information, "��������!")
        End If
    End Sub

    Private Sub ImportDCDataFromExcel()
        '////////////////////////////////////////////////////////////////////////////////////////
        '//
        '// �������� �� Excel ����� �� ���� ������� 
        '//
        '////////////////////////////////////////////////////////////////////////////////////////
        Dim appXLSRC As Object
        Dim MyCode As String
        Dim MyMGZ As Double
        Dim MyROP As Double
        Dim MyInsLVL As Double
        Dim StrCnt As Double
        Dim MySQLStr As String
        Dim MyDueDate As Date
        Dim MyMonthNum As Double       '--���-�� ������� ��� ������� MGZ, ROP �  �.�.

        OpenFileDialog1.ShowDialog()
        If (OpenFileDialog1.FileName = "") Then
        Else
            Me.Cursor = Cursors.WaitCursor
            Me.Refresh()
            System.Windows.Forms.Application.DoEvents()

            appXLSRC = CreateObject("Excel.Application")
            appXLSRC.Workbooks.Open(OpenFileDialog1.FileName)

            '---�������� �������� �� ������� �������
            MySQLStr = "DELETE FROM tbl_ForecastOrderR4_CustomMGZROPINS_DC "
            InitMyConn(False)
            Declarations.MyConn.Execute(MySQLStr)


            StrCnt = 5
            While Not appXLSRC.Worksheets(1).Range("A" & StrCnt).Value = Nothing
                MyCode = appXLSRC.Worksheets(1).Range("A" & StrCnt).Value.ToString
                Try
                    MyMGZ = appXLSRC.Worksheets(1).Range("B" & StrCnt).Value
                    Try
                        MyROP = appXLSRC.Worksheets(1).Range("C" & StrCnt).Value
                        Try
                            MyInsLVL = appXLSRC.Worksheets(1).Range("D" & StrCnt).Value
                            Try
                                MyDueDate = appXLSRC.Worksheets(1).Range("E" & StrCnt).Value
                                '---��������� ���� - ������ ���� ������ ������� � ������ ��� ������� + �������� �������
                                MySQLStr = "SELECT MonthNum "
                                MySQLStr = MySQLStr & "FROM tbl_ForecastOrder_PeriodQTY "
                                InitMyConn(False)
                                InitMyRec(False, MySQLStr)
                                If (Declarations.MyRec.EOF = True And Declarations.MyRec.BOF = True) Then
                                    MyMonthNum = 12
                                Else
                                    MyMonthNum = Declarations.MyRec.Fields("MonthNum").Value
                                End If
                                trycloseMyRec()

                                If MyDueDate < Now() Then
                                    MsgBox("������ � ������ " & StrCnt & " ������� ""E"". ���� ""��������� �� ����:"" ������ ���� ������ �������.", MsgBoxStyle.Critical, "��������!")
                                Else
                                    If MyDueDate > DateAdd(DateInterval.Month, MyMonthNum, Now()) Then
                                        MsgBox("������ � ������ " & StrCnt & " ������� ""E"". ���� ""��������� �� ����:"" �� ������ ��������� ������� ������ ��� �� " & CStr(MyMonthNum) & " �������.", MsgBoxStyle.Critical, "��������!")
                                    Else
                                        '---��� �������� - ���� �� ����� ��� � ������ ��������� 
                                        MySQLStr = "SELECT COUNT(*) AS CC "
                                        MySQLStr = MySQLStr & "FROM tbl_ForecastOrderR4_Main_DC "
                                        MySQLStr = MySQLStr & "WHERE (Code = N'" & Trim(MyCode) & "') "
                                        MySQLStr = MySQLStr & "AND (WHass = 1) "
                                        InitMyConn(False)
                                        InitMyRec(False, MySQLStr)
                                        If Declarations.MyRec.Fields("CC").Value = 0 Then
                                            trycloseMyRec()
                                            MsgBox("������ � ������ " & StrCnt & ". ����� �� �������� ���������.", MsgBoxStyle.Critical, "��������!")
                                        Else
                                            trycloseMyRec()
                                            '---������� � ������� �������
                                            MySQLStr = "INSERT INTO tbl_ForecastOrderR4_CustomMGZROPINS_DC "
                                            MySQLStr = MySQLStr & "(ID, Code, MGZ, ROP, IshuranceLVL, DueDate) "
                                            MySQLStr = MySQLStr & "VALUES (NEWID(), "
                                            MySQLStr = MySQLStr & "N'" & Trim(MyCode) & "', "
                                            MySQLStr = MySQLStr & Replace(CStr(MyMGZ), ",", ".") & ", "
                                            MySQLStr = MySQLStr & Replace(CStr(MyROP), ",", ".") & ", "
                                            MySQLStr = MySQLStr & Replace(CStr(MyInsLVL), ",", ".") & ", "
                                            MySQLStr = MySQLStr & "CONVERT(DATETIME, '" & FormatDateTime(MyDueDate, DateFormat.ShortDate) & "', 103)) "
                                            InitMyConn(False)
                                            Declarations.MyConn.Execute(MySQLStr)
                                        End If
                                    End If
                                End If
                            Catch ex As Exception
                                MsgBox("������ � ������ " & StrCnt & " ������� ""E"". " & ex.Message, MsgBoxStyle.Critical, "��������!")
                            End Try
                        Catch ex As Exception
                            MsgBox("������ � ������ " & StrCnt & " ������� ""D"". " & ex.Message, MsgBoxStyle.Critical, "��������!")
                        End Try
                    Catch ex As Exception
                        MsgBox("������ � ������ " & StrCnt & " ������� ""C"". " & ex.Message, MsgBoxStyle.Critical, "��������!")
                    End Try
                Catch ex As Exception
                    MsgBox("������ � ������ " & StrCnt & " ������� ""B"". " & ex.Message, MsgBoxStyle.Critical, "��������!")
                End Try
                StrCnt = StrCnt + 1
            End While
            appXLSRC.DisplayAlerts = 0
            appXLSRC.Workbooks.Close()
            appXLSRC.DisplayAlerts = 1
            appXLSRC.Quit()
            appXLSRC = Nothing
        End If
    End Sub

    Private Sub ImportDCDataFromLO()
        '////////////////////////////////////////////////////////////////////////////////////////
        '//
        '// �������� �� LibreOffice ����� �� ���� ������� 
        '//
        '////////////////////////////////////////////////////////////////////////////////////////
        Dim oServiceManager As Object
        Dim oDispatcher As Object
        Dim oDesktop As Object
        Dim oWorkBook As Object
        Dim oSheet As Object
        Dim oFileName As String
        Dim MyCode As String
        Dim MyMGZ As Double
        Dim MyROP As Double
        Dim MyInsLVL As Double
        Dim StrCnt As Double
        Dim MySQLStr As String
        Dim MyDueDate As Date
        Dim MyMonthNum As Double       '--���-�� ������� ��� ������� MGZ, ROP �  �.�.

        OpenFileDialog2.ShowDialog()
        If (OpenFileDialog2.FileName = "") Then
        Else
            Me.Cursor = Cursors.WaitCursor
            Me.Refresh()
            System.Windows.Forms.Application.DoEvents()
            Try
                LOSetNotation(0)
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

                '---�������� �������� �� ������� �������
                MySQLStr = "DELETE FROM tbl_ForecastOrderR4_CustomMGZROPINS_DC "
                InitMyConn(False)
                Declarations.MyConn.Execute(MySQLStr)
                StrCnt = 5
                While Not oSheet.getCellRangeByName("A" & StrCnt).String.Equals("")
                    MyCode = oSheet.getCellRangeByName("A" & StrCnt).String
                    Try
                        MyMGZ = oSheet.getCellRangeByName("B" & StrCnt).value
                        Try
                            MyROP = oSheet.getCellRangeByName("C" & StrCnt).value
                            Try
                                MyInsLVL = oSheet.getCellRangeByName("D" & StrCnt).value
                                Try
                                    MyDueDate = oSheet.getCellRangeByName("E" & StrCnt).String
                                    '---��������� ���� - ������ ���� ������ ������� � ������ ��� ������� + �������� �������
                                    MySQLStr = "SELECT MonthNum "
                                    MySQLStr = MySQLStr & "FROM tbl_ForecastOrder_PeriodQTY "
                                    InitMyConn(False)
                                    InitMyRec(False, MySQLStr)
                                    If (Declarations.MyRec.EOF = True And Declarations.MyRec.BOF = True) Then
                                        MyMonthNum = 12
                                    Else
                                        MyMonthNum = Declarations.MyRec.Fields("MonthNum").Value
                                    End If
                                    trycloseMyRec()

                                    If MyDueDate < Now() Then
                                        MsgBox("������ � ������ " & StrCnt & " ������� ""E"". ���� ""��������� �� ����:"" ������ ���� ������ �������.", MsgBoxStyle.Critical, "��������!")
                                    Else
                                        If MyDueDate > DateAdd(DateInterval.Month, MyMonthNum, Now()) Then
                                            MsgBox("������ � ������ " & StrCnt & " ������� ""E"". ���� ""��������� �� ����:"" �� ������ ��������� ������� ������ ��� �� " & CStr(MyMonthNum) & " �������.", MsgBoxStyle.Critical, "��������!")
                                        Else
                                            '---��� �������� - ���� �� ����� ��� � ������ ��������� 
                                            MySQLStr = "SELECT COUNT(*) AS CC "
                                            MySQLStr = MySQLStr & "FROM tbl_ForecastOrderR4_Main_DC "
                                            MySQLStr = MySQLStr & "WHERE (Code = N'" & Trim(MyCode) & "') "
                                            MySQLStr = MySQLStr & "AND (WHass = 1) "
                                            InitMyConn(False)
                                            InitMyRec(False, MySQLStr)
                                            If Declarations.MyRec.Fields("CC").Value = 0 Then
                                                trycloseMyRec()
                                                MsgBox("������ � ������ " & StrCnt & ". ����� �� �������� ���������.", MsgBoxStyle.Critical, "��������!")
                                            Else
                                                trycloseMyRec()
                                                '---������� � ������� �������
                                                MySQLStr = "INSERT INTO tbl_ForecastOrderR4_CustomMGZROPINS_DC "
                                                MySQLStr = MySQLStr & "(ID, Code, MGZ, ROP, IshuranceLVL, DueDate) "
                                                MySQLStr = MySQLStr & "VALUES (NEWID(), "
                                                MySQLStr = MySQLStr & "N'" & Trim(MyCode) & "', "
                                                MySQLStr = MySQLStr & Replace(CStr(MyMGZ), ",", ".") & ", "
                                                MySQLStr = MySQLStr & Replace(CStr(MyROP), ",", ".") & ", "
                                                MySQLStr = MySQLStr & Replace(CStr(MyInsLVL), ",", ".") & ", "
                                                MySQLStr = MySQLStr & "CONVERT(DATETIME, '" & FormatDateTime(MyDueDate, DateFormat.ShortDate) & "', 103)) "
                                                InitMyConn(False)
                                                Declarations.MyConn.Execute(MySQLStr)
                                            End If
                                        End If
                                    End If
                                Catch ex As Exception
                                    MsgBox("������ � ������ " & StrCnt & " ������� ""E"". " & ex.Message, MsgBoxStyle.Critical, "��������!")
                                End Try
                            Catch ex As Exception
                                MsgBox("������ � ������ " & StrCnt & " ������� ""D"". " & ex.Message, MsgBoxStyle.Critical, "��������!")
                            End Try
                        Catch ex As Exception
                            MsgBox("������ � ������ " & StrCnt & " ������� ""C"". " & ex.Message, MsgBoxStyle.Critical, "��������!")
                        End Try
                    Catch ex As Exception
                        MsgBox("������ � ������ " & StrCnt & " ������� ""B"". " & ex.Message, MsgBoxStyle.Critical, "��������!")
                    End Try
                    StrCnt = StrCnt + 1
                End While
            Catch ex As Exception
                MsgBox("������ : " & ex.Message, MsgBoxStyle.Critical, "��������!")
            End Try
            Me.Cursor = Cursors.Default
            oWorkBook.Close(True)
            MsgBox("������ ������ ����������.", MsgBoxStyle.OkOnly, "��������!")
        End If

    End Sub
End Class




