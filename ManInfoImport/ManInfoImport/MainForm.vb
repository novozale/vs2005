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

    Private Sub Button2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button2.Click
        '////////////////////////////////////////////////////////////////////////////////////////
        '//
        '// ��������� �������� ���������� �� �����������  
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

    Private Sub ImportDataFromExcel()
        '////////////////////////////////////////////////////////////////////////////////////////
        '//
        '// �������� �� Excel ���������� �� ��������������  
        '//
        '////////////////////////////////////////////////////////////////////////////////////////
        Dim MySQLStr As String                      'SQL ������
        Dim appXLSRC As Object
        Dim MyVersion As String                     '������ ���������
        Dim ScalaCode As String                     '��� ������ � Scala
        Dim ManCode As String                       '��� ������ �������������
        Dim Manufacturer As Integer                 '��� �������������
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
                MySQLStr = MySQLStr & "WHERE (Name = N'������ ����� �������������� � ��������������') "
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
                        MsgBox("� ������������� ����� Excel � ������ ""B" & CStr(i) & """ �� ���������� ��� ������ �������������. ", MsgBoxStyle.Critical, "��������!")
                        appXLSRC.DisplayAlerts = 0
                        appXLSRC.Workbooks.Close()
                        appXLSRC.DisplayAlerts = 1
                        appXLSRC.Quit()
                        appXLSRC = Nothing
                        trycloseMyRec()
                        Exit Sub
                    Else
                        ManCode = Trim(appXLSRC.Worksheets(1).Range("B" & CStr(i)).Value.ToString)
                    End If


                    If appXLSRC.Worksheets(1).Range("C" & CStr(i)).Value = Nothing Or Trim(appXLSRC.Worksheets(1).Range("C" & CStr(i)).Value.ToString) = "" Then
                        MsgBox("� ������������� ����� Excel � ������ ""C" & CStr(i) & """ �� ���������� ��� �������������. ", MsgBoxStyle.Critical, "��������!")
                        appXLSRC.DisplayAlerts = 0
                        appXLSRC.Workbooks.Close()
                        appXLSRC.DisplayAlerts = 1
                        appXLSRC.Quit()
                        appXLSRC = Nothing
                        trycloseMyRec()
                        Exit Sub
                    Else
                        Try
                            Manufacturer = Trim(appXLSRC.Worksheets(1).Range("C" & CStr(i)).Value.ToString)
                        Catch ex As Exception
                            MsgBox("� ������������� ����� Excel � ������ ""C" & CStr(i) & """ ��� ������������� � Scala ������ ���� ����� ������. ", MsgBoxStyle.Critical, "��������!")
                            appXLSRC.DisplayAlerts = 0
                            appXLSRC.Workbooks.Close()
                            appXLSRC.DisplayAlerts = 1
                            appXLSRC.Quit()
                            appXLSRC = Nothing
                            trycloseMyRec()
                            Exit Sub
                        End Try

                        '---���������, ��� ����� ������������� ���� � Scala
                        MySQLStr = "SELECT COUNT(*) AS CC "
                        MySQLStr = MySQLStr & "FROM tbl_Manufacturers "
                        MySQLStr = MySQLStr & "WHERE (ID = " & Manufacturer & ")"
                        InitMyConn(False)
                        InitMyRec(False, MySQLStr)
                        If (Declarations.MyRec.Fields("CC").Value = 0) Then
                            MsgBox("� ������������� ����� Excel � ������ ""C" & CStr(i) & """ ���������� �������� ��� ������������� � Scala ", MsgBoxStyle.Critical, "��������!")
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
                    '---��� ������ ������������� � ��� �������������
                    MySQLStr = "UPDATE tbl_ItemCard0300 "
                    MySQLStr = MySQLStr & "SET ManufacturerItemCode = N'" & ManCode & "', "
                    MySQLStr = MySQLStr & "Manufacturer = " & Manufacturer & " "
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
                MsgBox("������ ���������� � ����� ������� ������������� � �������������� ����������.", MsgBoxStyle.OkOnly, "��������!")
            End If
        End If
    End Sub

    Private Sub ImportDataFromLO()
        '////////////////////////////////////////////////////////////////////////////////////////
        '//
        '// �������� �� LibreOffice ���������� �� ��������������
        '//
        '////////////////////////////////////////////////////////////////////////////////////////
        Dim MySQLStr As String                      'SQL ������
        Dim MyVersion As String                     '������ ���������
        Dim ScalaCode As String                     '��� ������ � Scala
        Dim ManCode As String                       '��� ������ �������������
        Dim Manufacturer As Integer                 '��� �������������
        Dim i As Double                             '������� �����
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

                    '---��������� ������ ����� Excel
                    MyVersion = oSheet.getCellRangeByName("A1").String
                    MySQLStr = "SELECT Version "
                    MySQLStr = MySQLStr & "FROM tbl_VersionImportItemsFromExcel WITH(NOLOCK) "
                    MySQLStr = MySQLStr & "WHERE (Name = N'������ ����� �������������� � ��������������') "
                    InitMyConn(False)
                    InitMyRec(False, MySQLStr)
                    If Declarations.MyRec.BOF = True And Declarations.MyRec.EOF = True Then
                        MsgBox("� Scala �� ����������� ������� ������ ����� Excel. ���������� � ��������������", vbCritical, "��������!")
                        oWorkBook.Close(True)
                        trycloseMyRec()
                        Exit Sub
                    Else
                        If Trim(Declarations.MyRec.Fields("Version").Value) = MyVersion Then
                            trycloseMyRec()
                        Else
                            MsgBox("�� ��������� �������� � ������������ ������� ����� Excel. ���� �������� � ������� " & MyVersion & ".", vbCritical, "��������!")
                            oWorkBook.Close(True)
                            trycloseMyRec()
                            Exit Sub
                        End If
                    End If

                    i = 7
                    '---������ ������
                    While Not oSheet.getCellRangeByName("A" & i).String.Equals("")
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

                        '---��� ������ �������������
                        ManCode = Trim(oSheet.getCellRangeByName("B" & CStr(i)).String)
                        If ManCode.Equals("") Then
                            MsgBox("� ������������� ����� Excel � ������ ""B" & CStr(i) & """ �� ���������� ��� ������ �������������. ", MsgBoxStyle.Critical, "��������!")
                            oWorkBook.Close(True)
                            Exit Sub
                        End If

                        '---��� �������������
                        Try
                            Manufacturer = Trim(oSheet.getCellRangeByName("C" & CStr(i)).Value)
                        Catch ex As Exception
                            MsgBox("� ������������� ����� Excel � ������ ""C" & CStr(i) & """ ��� ������������� � Scala ������ ���� ����� ������. ", MsgBoxStyle.Critical, "��������!")
                            oWorkBook.Close(True)
                            Exit Sub
                        End Try

                        '---���������, ��� ����� ������������� ���� � Scala
                        MySQLStr = "SELECT COUNT(*) AS CC "
                        MySQLStr = MySQLStr & "FROM tbl_Manufacturers "
                        MySQLStr = MySQLStr & "WHERE (ID = " & Manufacturer & ")"
                        InitMyConn(False)
                        InitMyRec(False, MySQLStr)
                        If (Declarations.MyRec.Fields("CC").Value = 0) Then
                            MsgBox("� ������������� ����� Excel � ������ ""C" & CStr(i) & """ ���������� �������� ��� ������������� � Scala ", MsgBoxStyle.Critical, "��������!")
                            oWorkBook.Close(True)
                            trycloseMyRec()
                            Exit Sub
                        End If
                        trycloseMyRec()

                        '---��������� ������------------------------------------------------------------------
                        '---��� ������ ������������� � ��� �������������
                        MySQLStr = "UPDATE tbl_ItemCard0300 "
                        MySQLStr = MySQLStr & "SET ManufacturerItemCode = N'" & ManCode & "', "
                        MySQLStr = MySQLStr & "Manufacturer = " & Manufacturer & " "
                        MySQLStr = MySQLStr & "WHERE (SC01001 = N'" & ScalaCode & "') "
                        InitMyConn(False)
                        Declarations.MyConn.Execute(MySQLStr)

                        Label3.Text = CStr(i - 6)
                        Me.Refresh()
                        System.Windows.Forms.Application.DoEvents()
                        i = i + 1
                    End While
                Catch ex As Exception
                    MsgBox("������ : " & ex.Message, MsgBoxStyle.Critical, "��������!")
                End Try
                oWorkBook.Close(True)
                '---��������� �� ��������� ���������
                MsgBox("������ ���������� � ����� ������� ������������� � �������������� ����������.", MsgBoxStyle.OkOnly, "��������!")
            End If
        End If
    End Sub
End Class
