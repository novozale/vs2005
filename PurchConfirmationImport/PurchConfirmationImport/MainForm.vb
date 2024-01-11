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

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        '/////////////////////////////////////////////////////////////////////////////////////
        '//
        '// ����� �� ���������
        '// 
        '/////////////////////////////////////////////////////////////////////////////////////

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
    End Sub

    Private Sub Button2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button2.Click
        '////////////////////////////////////////////////////////////////////////////////////////
        '//
        '// ��������� �������� �� Excel ������������� �� �������� �� 1 ����������  
        '//
        '////////////////////////////////////////////////////////////////////////////////////////
        Label3.Text = ""
        Me.Refresh()
        System.Windows.Forms.Application.DoEvents()
        Button1.Enabled = False
        Button2.Enabled = False
        If My.Settings.UseOffice = "LibreOffice" Then
            ImportDataFromLO()
        Else
            ImportDataFromExcel()
        End If
        Button1.Enabled = True
        Button2.Enabled = True
        SetWindowPos(Me.Handle.ToInt32, -2, 0, 0, 0, 0, &H3)
        Me.Cursor = Cursors.Default
    End Sub

    Private Sub ImportDataFromExcel()
        '////////////////////////////////////////////////////////////////////////////////////////
        '//
        '// �������� �� Excel ������������� �� �������� �� 1 ����������  
        '//
        '////////////////////////////////////////////////////////////////////////////////////////
        Dim appXLSRC As Object
        Dim MySuppCode As String
        Dim MySQLStr As String                      'SQL ������
        Dim i As Double                             '������� �����
        Dim MyOrder As String                       '����� ������ �� �������
        Dim MySuppProductCode As String             '��� ������ ����������
        Dim MyProductCode As String                 '��� ������
        Dim MyConfDate As Date                      '�������������� ����
        Dim MyBackDate As Date                      '����������� ����
        Dim MyERRStr As String                      '��������� �� �������
        Dim MyOrderFlag As String                   '������� - ��������� ����� ������ ��� ���

        If OpenFileDialog1.ShowDialog() <> Windows.Forms.DialogResult.Cancel Then
            If (OpenFileDialog1.FileName = "") Then
            Else
                MyERRStr = ""

                Me.Cursor = Cursors.WaitCursor
                Me.Refresh()
                System.Windows.Forms.Application.DoEvents()

                appXLSRC = CreateObject("Excel.Application")
                appXLSRC.Workbooks.Open(OpenFileDialog1.FileName)

                '---���������, ��� ���������� ��� ����������
                MySuppCode = Trim(appXLSRC.Worksheets(1).Range("E1").Value)
                If MySuppCode = Nothing Then
                    MsgBox("� ������������� ����� Excel � ������ 'E1' �� ���������� ��� ���������� ", MsgBoxStyle.Critical, "��������!")
                    appXLSRC.DisplayAlerts = 0
                    appXLSRC.Workbooks.Close()
                    appXLSRC.DisplayAlerts = 1
                    appXLSRC.Quit()
                    appXLSRC = Nothing
                    Exit Sub
                End If

                '---��������� ��� ���� ��������� ���� � Scala
                MySQLStr = "SELECT COUNT(PL01001) AS CC "
                MySQLStr = MySQLStr & "FROM PL010300 "
                MySQLStr = MySQLStr & "WHERE (PL01001 = N'" & MySuppCode & "')"
                InitMyConn(False)
                InitMyRec(False, MySQLStr)
                If (Declarations.MyRec.Fields("CC").Value = 0) Then
                    MsgBox("� ������������� ����� Excel � ������ 'E1' ���������� �������� ��� ���������� � Scala ", MsgBoxStyle.Critical, "��������!")
                    appXLSRC.DisplayAlerts = 0
                    appXLSRC.Workbooks.Close()
                    appXLSRC.DisplayAlerts = 1
                    appXLSRC.Quit()
                    appXLSRC = Nothing
                    trycloseMyRec()
                    Exit Sub
                End If
                trycloseMyRec()

                i = 4
                While Not appXLSRC.Worksheets(1).Range("B" & i).Value = Nothing
                    Try
                        MyOrder = Microsoft.VisualBasic.Right("0000000000" & Trim(appXLSRC.Worksheets(1).Range("B" & CStr(i)).Value.ToString), 10)
                        '---��������� - ���� �� ����� ����� �� ������� �� ����� ���������� (����������)
                        MySQLStr = "SELECT COUNT(PC01001) AS CC "
                        MySQLStr = MySQLStr & "FROM PC010300 "
                        MySQLStr = MySQLStr & "WHERE (PC01001 = N'" & MyOrder & "') AND (PC01002 <> 2) "
                        InitMyConn(False)
                        InitMyRec(False, MySQLStr)
                        If (Declarations.MyRec.Fields("CC").Value = 0) Then
                            trycloseMyRec()
                            MsgBox("� ������������� ����� Excel � ������ 'B" & i & "' ���������� ����� ������ �� �������, �������� ��� � Scala ��� ������� ������ (2 ����) ", MsgBoxStyle.Critical, "��������!")
                        Else
                            trycloseMyRec()
                            If appXLSRC.Worksheets(1).Range("C" & i).Value <> Nothing Then
                                '---================================��������� ������������� ��� ����� ������
                                Try
                                    MyConfDate = CDate(appXLSRC.Worksheets(1).Range("E" & CStr(i)).Value.ToString)
                                    If appXLSRC.Worksheets(1).Range("F" & i).Value = Nothing Then
                                        MyBackDate = MyConfDate
                                    Else
                                        MyBackDate = CDate(appXLSRC.Worksheets(1).Range("F" & CStr(i)).Value.ToString)
                                    End If
                                    Try
                                        '---�� � ������� ���������� � Scala
                                        MySQLStr = "UPDATE PC030300 "
                                        MySQLStr = MySQLStr & "SET PC03016 = CONVERT(DATETIME, '" & Format(MyConfDate, "dd/MM/yyyy") & "', 103), "
                                        MySQLStr = MySQLStr & "PC03024 = CONVERT(DATETIME, '" & Format(MyConfDate, "dd/MM/yyyy") & "', 103), "
                                        MySQLStr = MySQLStr & "PC03031 = CONVERT(DATETIME, '" & Format(MyBackDate, "dd/MM/yyyy") & "', 103), "
                                        MySQLStr = MySQLStr & "PC03029 = N'1' "
                                        MySQLStr = MySQLStr & "WHERE (PC03001 = N'" & MyOrder & "') "
                                        Declarations.MyConn.Execute(MySQLStr)
                                    Catch
                                        MsgBox("� ������������� ����� Excel � ������ 'F" & i & "' ����������� �������� ���� ", MsgBoxStyle.Critical, "��������!")
                                    End Try
                                Catch
                                    MsgBox("� ������������� ����� Excel � ������ 'E" & i & "' ����������� �������� ���� ", MsgBoxStyle.Critical, "��������!")
                                End Try
                                '---===================================����� ��������� ������������� ��� ����� ������
                            Else
                                '---================================��������� ������������� ��� ������ ������ � ������
                                Try
                                    MySuppProductCode = appXLSRC.Worksheets(1).Range("D" & CStr(i)).Value.ToString
                                    '---��������� - ���� �� ����� ��� ������ ���������� � ����� ����������
                                    MySQLStr = "SELECT COUNT(*) AS CC "
                                    MySQLStr = MySQLStr & "FROM SC010300 "
                                    MySQLStr = MySQLStr & "WHERE (SC01060 = N'" & MySuppProductCode & "') AND (SC01058 = N'" & MySuppCode & "')"
                                    InitMyConn(False)
                                    InitMyRec(False, MySQLStr)
                                    If (Declarations.MyRec.Fields("CC").Value = 0) Then
                                        trycloseMyRec()
                                        '---��� ������ ���� � ������ ����������
                                        MyERRStr = MyERRStr & "������ " & i & " ��������� " & MySuppCode & " ��� ������ ���������� " & MySuppProductCode & " �� ������" & Microsoft.VisualBasic.Chr(13)
                                    Else
                                        trycloseMyRec()
                                        '---�������� ��� ��� ������
                                        MySQLStr = "Select SC01001 AS CC "
                                        MySQLStr = MySQLStr & "FROM SC010300 "
                                        MySQLStr = MySQLStr & "WHERE (SC01060 = N'" & MySuppProductCode & "') AND (SC01058 = N'" & MySuppCode & "')"
                                        InitMyConn(False)
                                        InitMyRec(False, MySQLStr)
                                        MyProductCode = Declarations.MyRec.Fields("CC").Value.ToString
                                        trycloseMyRec()
                                        '---��������� - ���� �� ����� ��� ������ � ���� ������ �� �������
                                        MySQLStr = "SELECT COUNT(*) AS CC "
                                        MySQLStr = MySQLStr & "FROM PC030300 "
                                        MySQLStr = MySQLStr & "WHERE (PC03001 = N'" & MyOrder & "') AND (PC03005 = N'" & MyProductCode & "') "
                                        InitMyConn(False)
                                        InitMyRec(False, MySQLStr)
                                        If (Declarations.MyRec.Fields("CC").Value = 0) Then
                                            trycloseMyRec()
                                            '--- ��� ������ ���� � ����� ������
                                            MyERRStr = MyERRStr & "������ " & i & " ��� ������ " & MyProductCode & " ��� ������ ���������� " & MySuppProductCode & " �� ������ � ������ �� ������� " & MyOrder & " " & Microsoft.VisualBasic.Chr(13)
                                        Else
                                            trycloseMyRec()
                                            Try
                                                MyConfDate = CDate(appXLSRC.Worksheets(1).Range("E" & CStr(i)).Value.ToString)
                                                If appXLSRC.Worksheets(1).Range("F" & i).Value = Nothing Then
                                                    MyBackDate = MyConfDate
                                                Else
                                                    MyBackDate = CDate(appXLSRC.Worksheets(1).Range("F" & CStr(i)).Value.ToString)
                                                End If
                                                Try
                                                    '---�� � ������� ���������� � Scala
                                                    MySQLStr = "UPDATE PC030300 "
                                                    MySQLStr = MySQLStr & "SET PC03016 = CONVERT(DATETIME, '" & Format(MyConfDate, "dd/MM/yyyy") & "', 103), "
                                                    MySQLStr = MySQLStr & "PC03024 = CONVERT(DATETIME, '" & Format(MyConfDate, "dd/MM/yyyy") & "', 103), "
                                                    MySQLStr = MySQLStr & "PC03031 = CONVERT(DATETIME, '" & Format(MyBackDate, "dd/MM/yyyy") & "', 103), "
                                                    MySQLStr = MySQLStr & "PC03029 = N'1' "
                                                    MySQLStr = MySQLStr & "WHERE (PC03001 = N'" & MyOrder & "') AND (PC03005 = N'" & MyProductCode & "') "
                                                    Declarations.MyConn.Execute(MySQLStr)
                                                Catch
                                                    MsgBox("� ������������� ����� Excel � ������ 'F" & i & "' ����������� �������� ���� ", MsgBoxStyle.Critical, "��������!")
                                                End Try
                                            Catch
                                                MsgBox("� ������������� ����� Excel � ������ 'E" & i & "' ����������� �������� ���� ", MsgBoxStyle.Critical, "��������!")
                                            End Try
                                        End If
                                    End If
                                Catch
                                    MsgBox("� ������������� ����� Excel � ������ 'D" & i & "' ���������� �������� ��� ������ ���������� ", MsgBoxStyle.Critical, "��������!")
                                End Try
                                '---===================================����� ��������� ������������� ��� ������ ������ � ������
                            End If
                        End If
                    Catch
                        MsgBox("� ������������� ����� Excel � ������ 'B" & i & "' ���������� �������� ����� ������ �� ������� ", MsgBoxStyle.Critical, "��������!")
                    End Try

                    Label3.Text = CStr(i - 3)
                    Me.Refresh()
                    System.Windows.Forms.Application.DoEvents()
                    i = i + 1
                End While

                appXLSRC.DisplayAlerts = 0
                appXLSRC.Workbooks.Close()
                appXLSRC.DisplayAlerts = 1
                appXLSRC.Quit()
                appXLSRC = Nothing
                If MyERRStr = "" Then '---��� ������
                    MsgBox("������ ������������� ������� �� ������� ����������", MsgBoxStyle.OkOnly, "��������!")
                Else
                    MyErrorForm = New ErrorForm
                    MyERRStr = "�� ����� ������� �������������� �������� ���� ������ " & Chr(13) & MyERRStr
                    MyErrorForm.MyMsg = MyERRStr
                    MyErrorForm.ShowDialog()
                End If
                    
            End If
        End If

    End Sub

    Private Sub ImportDataFromLO()
        '////////////////////////////////////////////////////////////////////////////////////////
        '//
        '// �������� �� Libre Office ������������� �� �������� �� 1 ����������  
        '//
        '////////////////////////////////////////////////////////////////////////////////////////
        Dim MyERRStr As String                      '��������� �� �������
        Dim MySuppCode As String
        Dim MySQLStr As String                      'SQL ������
        Dim i As Double                             '������� �����
        Dim MyOrder As String                       '����� ������ �� �������
        Dim MyConfDate As Date                      '�������������� ����
        Dim MyBackDate As Date                      '����������� ����
        Dim MySuppProductCode As String             '��� ������ ����������
        Dim MyProductCode As String                 '��� ������
        Dim oServiceManager As Object
        Dim oDispatcher As Object
        Dim oDesktop As Object
        Dim oWorkBook As Object
        Dim oSheet As Object
        Dim oFileName As String

        If OpenFileDialog2.ShowDialog() <> Windows.Forms.DialogResult.Cancel Then
            If (OpenFileDialog2.FileName = "") Then
            Else
                MyERRStr = ""

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

                    '---���������, ��� ���������� ��� ����������
                    MySuppCode = oSheet.getCellRangeByName("E1").String
                    If MySuppCode.Equals("") Then
                        MsgBox("� ������������� ����� Excel � ������ 'E1' �� ���������� ��� ���������� ", MsgBoxStyle.Critical, "��������!")
                        Me.Cursor = Cursors.Default
                        oWorkBook.Close(True)
                        Exit Sub
                    End If

                    '---��������� ��� ���� ��������� ���� � Scala
                    MySQLStr = "SELECT COUNT(PL01001) AS CC "
                    MySQLStr = MySQLStr & "FROM PL010300 WITH(NOLOCK) "
                    MySQLStr = MySQLStr & "WHERE (PL01001 = N'" & MySuppCode & "')"
                    InitMyConn(False)
                    InitMyRec(False, MySQLStr)
                    If (Declarations.MyRec.Fields("CC").Value = 0) Then
                        trycloseMyRec()
                        MsgBox("� ������������� ����� Excel � ������ 'E1' ���������� �������� ��� ���������� � Scala ", MsgBoxStyle.Critical, "��������!")
                        Me.Cursor = Cursors.Default
                        oWorkBook.Close(True)
                        Exit Sub
                    End If
                    trycloseMyRec()

                    i = 4
                    While Not oSheet.getCellRangeByName("B" & i).String.Equals("")
                        MyOrder = Microsoft.VisualBasic.Right("0000000000" & oSheet.getCellRangeByName("B" & i).String, 10)
                        '---��������� - ���� �� ����� ����� �� ������� �� ����� ���������� (����������)
                        MySQLStr = "SELECT COUNT(PC01001) AS CC "
                        MySQLStr = MySQLStr & "FROM PC010300 WITH(NOLOCK) "
                        MySQLStr = MySQLStr & "WHERE (PC01001 = N'" & MyOrder & "') " 'AND (PC01002 <> 2) "
                        InitMyConn(False)
                        InitMyRec(False, MySQLStr)
                        If (Declarations.MyRec.Fields("CC").Value = 0) Then
                            trycloseMyRec()
                            MsgBox("� ������������� ����� Excel � ������ 'B" & i & "' ���������� ����� ������ �� �������, �������� ��� � Scala ��� ������� ������ (2 ����) ", MsgBoxStyle.Critical, "��������!")
                        Else
                            trycloseMyRec()
                            If Not oSheet.getCellRangeByName("C" & i).String.Equals("") Then
                                '---================================��������� ������������� ��� ����� ������
                                Try
                                    MyConfDate = DateTime.FromOADate(oSheet.getCellRangeByName("E" & CStr(i)).Value)
                                    If oSheet.getCellRangeByName("F" & i).String.Equals("") Then
                                        MyBackDate = MyConfDate
                                    Else
                                        MyBackDate = DateTime.FromOADate(oSheet.getCellRangeByName("F" & CStr(i)).Value)
                                    End If
                                    Try
                                        '---�� � ������� ���������� � Scala
                                        MySQLStr = "UPDATE PC030300 "
                                        MySQLStr = MySQLStr & "SET PC03016 = CONVERT(DATETIME, '" & Format(MyConfDate, "dd/MM/yyyy") & "', 103), "
                                        MySQLStr = MySQLStr & "PC03024 = CONVERT(DATETIME, '" & Format(MyConfDate, "dd/MM/yyyy") & "', 103), "
                                        MySQLStr = MySQLStr & "PC03031 = CONVERT(DATETIME, '" & Format(MyBackDate, "dd/MM/yyyy") & "', 103), "
                                        MySQLStr = MySQLStr & "PC03029 = N'1' "
                                        MySQLStr = MySQLStr & "WHERE (PC03001 = N'" & MyOrder & "') "
                                        Declarations.MyConn.Execute(MySQLStr)

                                        MySQLStr = "UPDATE tbl_PurchaseWorkplace_ConsolidatedOrders "
                                        MySQLStr = MySQLStr & "SET ConfirmedDate = CONVERT(DATETIME, '" & Microsoft.VisualBasic.Right("00" & CStr(DatePart(DateInterval.Day, Now())), 2) & "/" & Microsoft.VisualBasic.Right("00" & CStr(DatePart(DateInterval.Month, Now())), 2) & "/" & CStr(DatePart(DateInterval.Year, Now())) & "', 103) "
                                        MySQLStr = MySQLStr & "FROM tbl_PurchaseWorkplace_ConsolidatedOrders INNER JOIN "
                                        MySQLStr = MySQLStr & "PC010300 ON tbl_PurchaseWorkplace_ConsolidatedOrders.ID = PC010300.PC01052 "
                                        MySQLStr = MySQLStr & "WHERE (PC010300.PC01001 = N'" & MyOrder & "') "
                                        Declarations.MyConn.Execute(MySQLStr)
                                    Catch ex As Exception
                                        MsgBox(ex.Message, MsgBoxStyle.OkOnly, "��������!")
                                    End Try
                                Catch ex As Exception
                                    MsgBox("� ������������� ����� Excel � ������ 'E" & i & "' ����������� �������� ���� ", MsgBoxStyle.Critical, "��������!")
                                End Try
                            Else
                                '---================================��������� ������������� ��� ������ ������ � ������
                                MySuppProductCode = oSheet.getCellRangeByName("D" & CStr(i)).String
                                '---��������� - ���� �� ����� ��� ������ ���������� � ����� ����������
                                MySQLStr = "SELECT COUNT(SC010300.SC01001) AS CC "
                                MySQLStr = MySQLStr & "FROM SC010300 WITH(NOLOCK) INNER JOIN "
                                MySQLStr = MySQLStr & "PC030300 ON SC010300.SC01001 = PC030300.PC03005 "
                                MySQLStr = MySQLStr & "WHERE (SC010300.SC01060 = N'" & MySuppProductCode & "') AND "
                                MySQLStr = MySQLStr & "(SC010300.SC01058 = N'" & MySuppCode & "') AND "
                                MySQLStr = MySQLStr & "(PC030300.PC03001 = N'" & MyOrder & "') "
                                InitMyConn(False)
                                InitMyRec(False, MySQLStr)
                                If (Declarations.MyRec.Fields("CC").Value = 0) Then
                                    trycloseMyRec()
                                    '---��� ������ ���� � ������ ����������
                                    MyERRStr = MyERRStr & "������ " & i & " ��������� " & MySuppCode & " ��� ������ ���������� " & MySuppProductCode & " �� ������" & Microsoft.VisualBasic.Chr(13)
                                Else
                                    trycloseMyRec()
                                    '---�������� ��� ��� ������
                                    MySQLStr = "SELECT SC010300.SC01001 AS CC "
                                    MySQLStr = MySQLStr & "FROM SC010300 WITH(NOLOCK) INNER JOIN "
                                    MySQLStr = MySQLStr & "PC030300 ON SC010300.SC01001 = PC030300.PC03005 "
                                    MySQLStr = MySQLStr & "WHERE (SC010300.SC01060 = N'" & MySuppProductCode & "') AND "
                                    MySQLStr = MySQLStr & "(SC010300.SC01058 = N'" & MySuppCode & "') AND "
                                    MySQLStr = MySQLStr & "(PC030300.PC03001 = N'" & MyOrder & "') "
                                    InitMyConn(False)
                                    InitMyRec(False, MySQLStr)
                                    MyProductCode = Declarations.MyRec.Fields("CC").Value.ToString
                                    trycloseMyRec()
                                    '---��������� - ���� �� ����� ��� ������ � ���� ������ �� �������
                                    MySQLStr = "SELECT COUNT(*) AS CC "
                                    MySQLStr = MySQLStr & "FROM PC030300 "
                                    MySQLStr = MySQLStr & "WHERE (PC03001 = N'" & MyOrder & "') AND (PC03005 = N'" & MyProductCode & "') "
                                    InitMyConn(False)
                                    InitMyRec(False, MySQLStr)
                                    If (Declarations.MyRec.Fields("CC").Value = 0) Then
                                        trycloseMyRec()
                                        '--- ��� ������ ���� � ����� ������
                                        MyERRStr = MyERRStr & "������ " & i & " ��� ������ " & MyProductCode & " ��� ������ ���������� " & MySuppProductCode & " �� ������ � ������ �� ������� " & MyOrder & " " & Microsoft.VisualBasic.Chr(13)
                                    Else
                                        trycloseMyRec()
                                        Try
                                            MyConfDate = DateTime.FromOADate(oSheet.getCellRangeByName("E" & CStr(i)).Value)
                                            If oSheet.getCellRangeByName("F" & i).String.Equals("") Then
                                                MyBackDate = MyConfDate
                                            Else
                                                MyBackDate = DateTime.FromOADate(oSheet.getCellRangeByName("F" & CStr(i)).Value)
                                            End If
                                            '---�� � ������� ���������� � Scala
                                            Try
                                                MySQLStr = "UPDATE PC030300 "
                                                MySQLStr = MySQLStr & "SET PC03016 = CONVERT(DATETIME, '" & Format(MyConfDate, "dd/MM/yyyy") & "', 103), "
                                                MySQLStr = MySQLStr & "PC03024 = CONVERT(DATETIME, '" & Format(MyConfDate, "dd/MM/yyyy") & "', 103), "
                                                MySQLStr = MySQLStr & "PC03031 = CONVERT(DATETIME, '" & Format(MyBackDate, "dd/MM/yyyy") & "', 103), "
                                                MySQLStr = MySQLStr & "PC03029 = N'1' "
                                                MySQLStr = MySQLStr & "WHERE (PC03001 = N'" & MyOrder & "') AND (PC03005 = N'" & MyProductCode & "') "
                                                Declarations.MyConn.Execute(MySQLStr)

                                                MySQLStr = "UPDATE tbl_PurchaseWorkplace_ConsolidatedOrders "
                                                MySQLStr = MySQLStr & "SET ConfirmedDate = CONVERT(DATETIME, '" & Microsoft.VisualBasic.Right("00" & CStr(DatePart(DateInterval.Day, Now())), 2) & "/" & Microsoft.VisualBasic.Right("00" & CStr(DatePart(DateInterval.Month, Now())), 2) & "/" & CStr(DatePart(DateInterval.Year, Now())) & "', 103) "
                                                MySQLStr = MySQLStr & "FROM tbl_PurchaseWorkplace_ConsolidatedOrders INNER JOIN "
                                                MySQLStr = MySQLStr & "PC010300 ON tbl_PurchaseWorkplace_ConsolidatedOrders.ID = PC010300.PC01052 "
                                                MySQLStr = MySQLStr & "WHERE (PC010300.PC01001 = N'" & MyOrder & "') "
                                                Declarations.MyConn.Execute(MySQLStr)
                                            Catch ex As Exception
                                                MsgBox(ex.Message, MsgBoxStyle.OkOnly, "��������!")
                                            End Try
                                        Catch
                                            MsgBox("� ������������� ����� Excel � ������ 'F" & i & "' ����������� �������� ���� ", MsgBoxStyle.Critical, "��������!")
                                        End Try
                                    End If
                                End If
                            End If
                        End If
                        Label3.Text = CStr(i - 3)
                        Me.Refresh()
                        System.Windows.Forms.Application.DoEvents()
                        i = i + 1
                    End While
                Catch ex As Exception
                    MsgBox("������ : " & ex.Message, MsgBoxStyle.Critical, "��������!")
                Finally
                    Try
                        oWorkBook.Close(True)
                    Catch ex As Exception
                    End Try
                    Declarations.MyConn.Close()
                    Declarations.MyConn = Nothing
                End Try
                Me.Cursor = Cursors.Default
                If MyERRStr = "" Then '---��� ������
                    MsgBox("������ ������������� ������� �� ������� ����������", MsgBoxStyle.OkOnly, "��������!")
                Else
                    MyErrorForm = New ErrorForm
                    MyERRStr = "�� ����� ������� �������������� �������� ���� ������ " & Chr(13) & MyERRStr
                    MyErrorForm.MyMsg = MyERRStr
                    MyErrorForm.ShowDialog()
                End If
            End If
        End If
    End Sub
End Class
