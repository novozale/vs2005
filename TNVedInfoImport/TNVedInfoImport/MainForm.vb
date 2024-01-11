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
        '// �������� �� Excel ����� �� ���  
        '//
        '////////////////////////////////////////////////////////////////////////////////////////
        Dim MySQLStr As String                      'SQL ������
        Dim MyExcelCounter As Double                  '������� ����� Excel
        Dim appXLSRC As Object
        Dim ScalaCode As String
        Dim TNVedCode As String

        If OpenFileDialog1.ShowDialog() <> Windows.Forms.DialogResult.Cancel Then
            If (OpenFileDialog1.FileName = "") Then
            Else
                Me.Cursor = Cursors.WaitCursor
                Me.Refresh()
                System.Windows.Forms.Application.DoEvents()

                appXLSRC = CreateObject("Excel.Application")
                appXLSRC.Workbooks.Open(OpenFileDialog1.FileName)

                MyExcelCounter = 2
                While Not appXLSRC.Worksheets(1).Range("A" & CStr(MyExcelCounter)).Value = Nothing
                    '---��� Scala
                    ScalaCode = Trim(appXLSRC.Worksheets(1).Range("A" & CStr(MyExcelCounter)).Value)
                    '---���������, ��� ����� ��� ���� � Scala
                    MySQLStr = "SELECT COUNT(*) AS CC "
                    MySQLStr = MySQLStr & "FROM SC010300 WITH(NOLOCK) "
                    MySQLStr = MySQLStr & "WHERE (SC01001 = N'" & ScalaCode & "') "
                    InitMyConn(False)
                    InitMyRec(False, MySQLStr)
                    If (Declarations.MyRec.Fields("CC").Value = 0) Then
                        MsgBox("� ������������� ����� Excel � ������ ""A" & CStr(MyExcelCounter) & """ ���������� �������� ��� ������ � Scala ", MsgBoxStyle.Critical, "��������!")
                        appXLSRC.DisplayAlerts = 0
                        appXLSRC.Workbooks.Close()
                        appXLSRC.DisplayAlerts = 1
                        appXLSRC.Quit()
                        appXLSRC = Nothing
                        trycloseMyRec()
                        Exit Sub
                    End If
                    trycloseMyRec()

                    '---��� �� ���
                    If appXLSRC.Worksheets(1).Range("B" & CStr(MyExcelCounter)).Value = Nothing Or Trim(appXLSRC.Worksheets(1).Range("B" & CStr(MyExcelCounter)).Value.ToString) = "" Then
                        MsgBox("� ������������� ����� Excel � ������ ""B" & CStr(MyExcelCounter) & """ �� ���������� ��� �� ���. ", MsgBoxStyle.Critical, "��������!")
                        appXLSRC.DisplayAlerts = 0
                        appXLSRC.Workbooks.Close()
                        appXLSRC.DisplayAlerts = 1
                        appXLSRC.Quit()
                        appXLSRC = Nothing
                        Exit Sub
                    End If
                    TNVedCode = Trim(appXLSRC.Worksheets(1).Range("B" & CStr(MyExcelCounter)).Value)
                    '---���������, ��� ����� ��� �� ��� ���� � Scala
                    MySQLStr = "SELECT COUNT(*) AS CC "
                    MySQLStr = MySQLStr & "FROM SY240300 "
                    MySQLStr = MySQLStr & "WHERE (SY24001 = N'BN') AND "
                    MySQLStr = MySQLStr & "(SY24002 = N'" & Trim(TNVedCode) & "') "
                    InitMyConn(False)
                    InitMyRec(False, MySQLStr)
                    If Declarations.MyRec.Fields("CC").Value = 0 Then
                        MsgBox("��� �� ��� " & TNVedCode & " � ������ " & MyExcelCounter & " ������������ � ����������� Scala. ������ ��������� �� �����.", vbCritical, "��������!")
                        appXLSRC.DisplayAlerts = 0
                        appXLSRC.Workbooks.Close()
                        appXLSRC.DisplayAlerts = 1
                        appXLSRC.Quit()
                        appXLSRC = Nothing
                        trycloseMyRec()
                        Exit Sub
                    End If
                    '---��������� ���� �� ��� � Scala
                    MySQLStr = "UPDATE SC010300 "
                    MySQLStr = MySQLStr & "Set SC01100 = N'" & TNVedCode & "' "
                    MySQLStr = MySQLStr & "WHERE (SC01001 = N'" & Trim(ScalaCode) & "')"
                    InitMyConn(False)
                    Declarations.MyConn.Execute(MySQLStr)

                    Label3.Text = CStr(MyExcelCounter - 1)
                    Me.Refresh()
                    System.Windows.Forms.Application.DoEvents()
                    MyExcelCounter = MyExcelCounter + 1
                End While
                appXLSRC.DisplayAlerts = 0
                appXLSRC.Workbooks.Close()
                appXLSRC.DisplayAlerts = 1
                appXLSRC.Quit()
                appXLSRC = Nothing
                '---��������� �� ��������� ���������
                MsgBox("������ ���������� � ����� �� ��� ����������. ", MsgBoxStyle.OkOnly, "��������!")
            End If
        End If

    End Sub

    Private Sub ImportDataFromLO()
        '////////////////////////////////////////////////////////////////////////////////////////
        '//
        '// �������� �� LibreOffice ����� �� ���  
        '//
        '////////////////////////////////////////////////////////////////////////////////////////
        Dim MySQLStr As String                      'SQL ������
        Dim MyExcelCounter As Double                  '������� ����� Excel
        Dim ScalaCode As String
        Dim TNVedCode As String
        Dim oServiceManager As Object
        Dim oDispatcher As Object
        Dim oDesktop As Object
        Dim oWorkBook As Object
        Dim oSheet As Object
        Dim oFileName As String

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

                MyExcelCounter = 2
                While oSheet.getCellRangeByName("A" & MyExcelCounter).String.Equals("") = False
                    '---��� Scala
                    ScalaCode = Trim(oSheet.getCellRangeByName("A" & MyExcelCounter).String)
                    '---���������, ��� ����� ��� ���� � Scala
                    MySQLStr = "SELECT COUNT(*) AS CC "
                    MySQLStr = MySQLStr & "FROM SC010300 WITH(NOLOCK) "
                    MySQLStr = MySQLStr & "WHERE (SC01001 = N'" & ScalaCode & "') "
                    InitMyConn(False)
                    InitMyRec(False, MySQLStr)
                    If (Declarations.MyRec.Fields("CC").Value = 0) Then
                        MsgBox("� ������������� ����� Excel � ������ ""A" & CStr(MyExcelCounter) & """ ���������� �������� ��� ������ � Scala ", MsgBoxStyle.Critical, "��������!")
                        oWorkBook.Close(True)
                        trycloseMyRec()
                        Exit Sub
                    End If
                    trycloseMyRec()

                    '---��� �� ���
                    If Trim(oSheet.getCellRangeByName("B" & MyExcelCounter).String).Equals("") Then
                        MsgBox("� ������������� ����� Excel � ������ ""B" & CStr(MyExcelCounter) & """ �� ���������� ��� �� ���. ", MsgBoxStyle.Critical, "��������!")
                        oWorkBook.Close(True)
                        Exit Sub
                    End If
                    TNVedCode = Trim(oSheet.getCellRangeByName("B" & MyExcelCounter).String)
                    '---���������, ��� ����� ��� �� ��� ���� � Scala
                    MySQLStr = "SELECT COUNT(*) AS CC "
                    MySQLStr = MySQLStr & "FROM SY240300 "
                    MySQLStr = MySQLStr & "WHERE (SY24001 = N'BN') AND "
                    MySQLStr = MySQLStr & "(SY24002 = N'" & Trim(TNVedCode) & "') "
                    InitMyConn(False)
                    InitMyRec(False, MySQLStr)
                    If Declarations.MyRec.Fields("CC").Value = 0 Then
                        MsgBox("��� �� ��� " & TNVedCode & " � ������ " & MyExcelCounter & " ������������ � ����������� Scala. ������ ��������� �� �����.", vbCritical, "��������!")
                        oWorkBook.Close(True)
                        trycloseMyRec()
                        Exit Sub
                    End If

                    '---��������� ���� �� ��� � Scala
                    MySQLStr = "UPDATE SC010300 "
                    MySQLStr = MySQLStr & "Set SC01100 = N'" & TNVedCode & "' "
                    MySQLStr = MySQLStr & "WHERE (SC01001 = N'" & Trim(ScalaCode) & "')"
                    InitMyConn(False)
                    Declarations.MyConn.Execute(MySQLStr)

                    Label3.Text = CStr(MyExcelCounter - 1)
                    Me.Refresh()
                    System.Windows.Forms.Application.DoEvents()
                    MyExcelCounter = MyExcelCounter + 1
                End While
                oWorkBook.Close(True)
                MsgBox("������ ���������� � ����� �� ��� ����������.", vbOKOnly, "��������!")
            End If
        End If
    End Sub
End Class
