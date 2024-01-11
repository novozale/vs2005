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
        '// 
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
        '// ��������� �������� �� Excel ���������� �� �����������  
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
        '// �������� �� Excel ����� ����� �� ������� �� 1 ����������  
        '//
        '////////////////////////////////////////////////////////////////////////////////////////
        Dim MySQLStr As String                      'SQL ������
        Dim appXLSRC As Object
        Dim MyVersion As String                     '������ ���������
        Dim ScalaCode As String                     '��� ������ � Scala
        Dim Supplier As String                      '��� ����������
        Dim AltSupplier As String                   '��� �������������� ����������
        Dim i As Double                             '������� �����

        If OpenFileDialog1.ShowDialog() <> Windows.Forms.DialogResult.Cancel Then
            If (OpenFileDialog1.FileName = "") Then
            Else
                Me.Cursor = Cursors.WaitCursor
                Me.Refresh()
                System.Windows.Forms.Application.DoEvents()

                appXLSRC = CreateObject("Excel.Application")
                appXLSRC.Workbooks.Open(OpenFileDialog1.FileName)

                '---��������� ������ ����� Excel
                MyVersion = Trim(appXLSRC.Worksheets(1).Range("A1").Value)
                MySQLStr = "SELECT Version "
                MySQLStr = MySQLStr & "FROM tbl_VersionImportItemsFromExcel WITH(NOLOCK) "
                MySQLStr = MySQLStr & "WHERE (Name = N'������ �����������') "
                InitMyConn(False)
                InitMyRec(False, MySQLStr)
                If Declarations.MyRec.BOF = True And Declarations.MyRec.EOF = True Then
                    MsgBox("� Scala �� ����������� ������� ������ ����� Excel. ���������� � ��������������", vbCritical, "��������!")
                    appXLSRC.DisplayAlerts = 0
                    appXLSRC.Workbooks.Close()
                    appXLSRC.DisplayAlerts = 1
                    appXLSRC.Quit()
                    appXLSRC = Nothing
                    trycloseMyRec()
                    Exit Sub
                Else
                    If Trim(Declarations.MyRec.Fields("Version").Value) = MyVersion Then
                        trycloseMyRec()
                    Else
                        MsgBox("�� ��������� �������� � ������������ ������� ����� Excel. ���� �������� � ������� " & MyVersion & ".", vbCritical, "��������!")
                        appXLSRC.DisplayAlerts = 0
                        appXLSRC.Workbooks.Close()
                        appXLSRC.DisplayAlerts = 1
                        appXLSRC.Quit()
                        appXLSRC = Nothing
                        trycloseMyRec()
                        Exit Sub
                    End If
                End If

                i = 7
                '---������ ������
                While Not appXLSRC.Worksheets(1).Range("A" & CStr(i)).Value = Nothing
                    ScalaCode = Trim(appXLSRC.Worksheets(1).Range("A" & CStr(i)).Value)
                    '---���������, ��� ����� ��� ���� � Scala
                    MySQLStr = "SELECT COUNT(*) AS CC "
                    MySQLStr = MySQLStr & "FROM SC010300 WITH(NOLOCK) "
                    MySQLStr = MySQLStr & "WHERE (SC01001 = N'" & ScalaCode & "') "
                    InitMyConn(False)
                    InitMyRec(False, MySQLStr)
                    If (Declarations.MyRec.Fields("CC").Value = 0) Then
                        MsgBox("� ������������� ����� Excel � ������ ""A" & CStr(i) & """ ���������� �������� ��� ������ � Scala ", MsgBoxStyle.Critical, "��������!")
                        appXLSRC.DisplayAlerts = 0
                        appXLSRC.Workbooks.Close()
                        appXLSRC.DisplayAlerts = 1
                        appXLSRC.Quit()
                        appXLSRC = Nothing
                        trycloseMyRec()
                        Exit Sub
                    End If
                    trycloseMyRec()

                    If appXLSRC.Worksheets(1).Range("B" & CStr(i)).Value = Nothing Or Trim(appXLSRC.Worksheets(1).Range("B" & CStr(i)).Value.ToString) = "" Then
                        MsgBox("� ������������� ����� Excel � ������ ""B" & CStr(i) & """ �� ���������� ��� �������� ����������. ", MsgBoxStyle.Critical, "��������!")
                        appXLSRC.DisplayAlerts = 0
                        appXLSRC.Workbooks.Close()
                        appXLSRC.DisplayAlerts = 1
                        appXLSRC.Quit()
                        appXLSRC = Nothing
                        trycloseMyRec()
                        Exit Sub
                    Else
                        Supplier = Trim(appXLSRC.Worksheets(1).Range("B" & CStr(i)).Value.ToString)
                        '---���������, ��� ����� ��������� ���� � Scala
                        MySQLStr = "SELECT COUNT(*) AS CC "
                        MySQLStr = MySQLStr & "FROM PL010300 WITH(NOLOCK) "
                        MySQLStr = MySQLStr & "WHERE (PL01001 = N'" & Supplier & "') "
                        InitMyConn(False)
                        InitMyRec(False, MySQLStr)
                        If (Declarations.MyRec.Fields("CC").Value = 0) Then
                            MsgBox("� ������������� ����� Excel � ������ ""B" & CStr(i) & """ ���������� �������� ��� ���������� � Scala ", MsgBoxStyle.Critical, "��������!")
                            appXLSRC.DisplayAlerts = 0
                            appXLSRC.Workbooks.Close()
                            appXLSRC.DisplayAlerts = 1
                            appXLSRC.Quit()
                            appXLSRC = Nothing
                            trycloseMyRec()
                            Exit Sub
                        End If
                    End If
                    

                    If appXLSRC.Worksheets(1).Range("C" & CStr(i)).Value = Nothing Then
                        AltSupplier = ""
                    Else
                        AltSupplier = Trim(appXLSRC.Worksheets(1).Range("C" & CStr(i)).Value.ToString)
                        '---���������, ��� ����� ��������� ���� � Scala
                        MySQLStr = "SELECT COUNT(*) AS CC "
                        MySQLStr = MySQLStr & "FROM PL010300 WITH(NOLOCK) "
                        MySQLStr = MySQLStr & "WHERE (PL01001 = N'" & AltSupplier & "') "
                        InitMyConn(False)
                        InitMyRec(False, MySQLStr)
                        If (Declarations.MyRec.Fields("CC").Value = 0) Then
                            MsgBox("� ������������� ����� Excel � ������ ""C" & CStr(i) & """ ���������� �������� ��� ��������������� ���������� � Scala ", MsgBoxStyle.Critical, "��������!")
                            appXLSRC.DisplayAlerts = 0
                            appXLSRC.Workbooks.Close()
                            appXLSRC.DisplayAlerts = 1
                            appXLSRC.Quit()
                            appXLSRC = Nothing
                            trycloseMyRec()
                            Exit Sub
                        End If
                        trycloseMyRec()
                    End If


                    '---��������� ������------------------------------------------------------------------
                    '---�������� ���������
                    MySQLStr = "UPDATE SC010300 "
                    MySQLStr = MySQLStr & "SET SC01058 = N'" & Supplier & "' "
                    MySQLStr = MySQLStr & "WHERE (SC01001 = N'" & ScalaCode & "') "
                    InitMyConn(False)
                    Declarations.MyConn.Execute(MySQLStr)

                    '---�������������� ���������
                    MySQLStr = "UPDATE SC010300 "
                    MySQLStr = MySQLStr & "SET SC01059 = N'" & AltSupplier & "' "
                    MySQLStr = MySQLStr & "WHERE (SC01001 = N'" & ScalaCode & "') "
                    InitMyConn(False)
                    Declarations.MyConn.Execute(MySQLStr)

                    Label3.Text = CStr(i - 6)
                    Me.Refresh()
                    System.Windows.Forms.Application.DoEvents()
                    i = i + 1
                End While
                appXLSRC.DisplayAlerts = 0
                appXLSRC.Workbooks.Close()
                appXLSRC.DisplayAlerts = 1
                appXLSRC.Quit()
                appXLSRC = Nothing
                '---��������� �� ��������� ���������
                MsgBox("������ ���������� � ����������� ����������. ������ ROP, ���, ���, XYZ - ������ ����� ����������� �����.", MsgBoxStyle.OkOnly, "��������!")
            End If
        End If
    End Sub

    Private Sub ImportDataFromLO()
        '////////////////////////////////////////////////////////////////////////////////////////
        '//
        '// �������� �� LibreOffice ����� ����� �� ������� �� 1 ����������  
        '//
        '////////////////////////////////////////////////////////////////////////////////////////
        Dim oServiceManager As Object
        Dim oDispatcher As Object
        Dim oDesktop As Object
        Dim oWorkBook As Object
        Dim oSheet As Object
        Dim oFileName As String
        Dim MyVersion As String                       '������ ���������
        Dim MySQLStr As String                        'SQL ������
        Dim i As Integer                              '������� �����
        Dim ScalaCode As String                     '��� ������ � Scala
        Dim Supplier As String                      '��� ����������
        Dim AltSupplier As String                   '��� �������������� ����������

        If OpenFileDialog2.ShowDialog() <> Windows.Forms.DialogResult.Cancel Then
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

                '---��������� ������ ����� Excel
                MyVersion = oSheet.getCellRangeByName("A1").String
                If MyVersion = "" Then
                    MsgBox("� ������������� ����� Excel � ������ 'A1' �� ����������� ������ ����� Excel ", MsgBoxStyle.Critical, "��������!")
                    oWorkBook.Close(True)
                    Exit Sub
                Else
                    MySQLStr = "SELECT Version "
                    MySQLStr = MySQLStr & "FROM tbl_VersionImportItemsFromExcel "
                    MySQLStr = MySQLStr & "WHERE (Name = N'������ �����������') "
                    InitMyConn(False)
                    InitMyRec(False, MySQLStr)
                    If Declarations.MyRec.BOF = True And Declarations.MyRec.EOF = True Then
                        MsgBox("� Scala �� ����������� ������� ������ ����� Excel. ���������� � ��������������", vbCritical, "��������!")
                        trycloseMyRec()
                        oWorkBook.Close(True)
                        Exit Sub
                    Else
                        If Trim(Declarations.MyRec.Fields("Version").Value) = MyVersion Then
                            trycloseMyRec()
                        Else
                            MsgBox("�� ��������� �������� � ������������ ������� ����� Excel. ���� �������� � ������� " & Declarations.MyRec.Fields("Version").Value & ".", vbCritical, "��������!")
                            trycloseMyRec()
                            oWorkBook.Close(True)
                            Exit Sub
                        End If
                    End If
                End If

                '---������ ������
                i = 7
                While oSheet.getCellRangeByName("A" & CStr(i)).String.Equals("") = False
                    '---��� Scala
                    ScalaCode = Trim(oSheet.getCellRangeByName("A" & i).String)
                    '---���������, ��� ����� ��� ���� � Scala
                    MySQLStr = "SELECT COUNT(*) AS CC "
                    MySQLStr = MySQLStr & "FROM SC010300 WITH(NOLOCK) "
                    MySQLStr = MySQLStr & "WHERE (SC01001 = N'" & ScalaCode & "') "
                    InitMyConn(False)
                    InitMyRec(False, MySQLStr)
                    If (Declarations.MyRec.Fields("CC").Value = 0) Then
                        MsgBox("� ������������� ����� Excel � ������ ""A" & CStr(i) & """ ���������� �������� ��� ������ � Scala ", MsgBoxStyle.Critical, "��������!")
                        oWorkBook.Close(True)
                        trycloseMyRec()
                        Exit Sub
                    End If
                    trycloseMyRec()

                    '---���������
                    If Trim(oSheet.getCellRangeByName("B" & i).String).Equals("") Then
                        MsgBox("� ������������� ����� Excel � ������ ""B" & CStr(i) & """ �� ���������� ��� �������� ����������. ", MsgBoxStyle.Critical, "��������!")
                        oWorkBook.Close(True)
                        Exit Sub
                    End If
                    Supplier = Trim(oSheet.getCellRangeByName("B" & CStr(i)).String)
                    '---���������, ��� ����� ��������� ���� � Scala
                    MySQLStr = "SELECT COUNT(*) AS CC "
                    MySQLStr = MySQLStr & "FROM PL010300 WITH(NOLOCK) "
                    MySQLStr = MySQLStr & "WHERE (PL01001 = N'" & Supplier & "') "
                    InitMyConn(False)
                    InitMyRec(False, MySQLStr)
                    If (Declarations.MyRec.Fields("CC").Value = 0) Then
                        MsgBox("� ������������� ����� Excel � ������ ""B" & CStr(i) & """ ���������� �������� ��� ���������� � Scala ", MsgBoxStyle.Critical, "��������!")
                        oWorkBook.Close(True)
                        trycloseMyRec()
                        Exit Sub
                    End If

                    '---�������������� ���������
                    AltSupplier = Trim(oSheet.getCellRangeByName("C" & CStr(i)).String)
                    '---���������, ��� ����� ��������� ���� � Scala
                    If Not AltSupplier.Equals("") Then
                        MySQLStr = "SELECT COUNT(*) AS CC "
                        MySQLStr = MySQLStr & "FROM PL010300 WITH(NOLOCK) "
                        MySQLStr = MySQLStr & "WHERE (PL01001 = N'" & AltSupplier & "') "
                        InitMyConn(False)
                        InitMyRec(False, MySQLStr)
                        If (Declarations.MyRec.Fields("CC").Value = 0) Then
                            MsgBox("� ������������� ����� Excel � ������ ""C" & CStr(i) & """ ���������� �������� ��� ��������������� ���������� � Scala ", MsgBoxStyle.Critical, "��������!")
                            oWorkBook.Close(True)
                            trycloseMyRec()
                            Exit Sub
                        End If
                    End If

                    '---��������� ������------------------------------------------------------------------
                    '---�������� ���������
                    MySQLStr = "UPDATE SC010300 "
                    MySQLStr = MySQLStr & "SET SC01058 = N'" & Supplier & "' "
                    MySQLStr = MySQLStr & "WHERE (SC01001 = N'" & ScalaCode & "') "
                    InitMyConn(False)
                    Declarations.MyConn.Execute(MySQLStr)

                    '---�������������� ���������
                    MySQLStr = "UPDATE SC010300 "
                    MySQLStr = MySQLStr & "SET SC01059 = N'" & AltSupplier & "' "
                    MySQLStr = MySQLStr & "WHERE (SC01001 = N'" & ScalaCode & "') "
                    InitMyConn(False)
                    Declarations.MyConn.Execute(MySQLStr)


                    Label3.Text = CStr(i - 6)
                    Me.Refresh()
                    System.Windows.Forms.Application.DoEvents()
                    i = i + 1
                End While
                trycloseMyRec()
                oWorkBook.Close(True)
                MsgBox("������ ���������� � ����������� ����������. ������ ROP, ���, ���, XYZ - ������ ����� ����������� �����.", vbOKOnly, "��������!")
            End If
        End If
    End Sub
End Class
