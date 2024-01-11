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
        '// выход из программы
        '// 
        '/////////////////////////////////////////////////////////////////////////////////////

        Application.Exit()
    End Sub

    Private Sub MainForm_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles Me.KeyDown
        '////////////////////////////////////////////////////////////////////////////////
        '//
        '// Запрет закрытия окна по ALT+F4
        '//
        '////////////////////////////////////////////////////////////////////////////////

        If e.KeyData = Keys.Alt + Keys.F4 Then
            e.Handled = True
        End If
    End Sub

    Private Sub MainForm_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        '/////////////////////////////////////////////////////////////////////////////////////
        '//
        '// При запуске определяем параметры - Год, компания, пользователь и т.д.
        '// 
        '/////////////////////////////////////////////////////////////////////////////////////

        '---параметры запуска
        Try
            Dim Scala As New SfwIII.Application

            Declarations.CompanyID = Scala.ActiveProcess.CommonVars.CompanyCode
            Declarations.Year = Mid(Scala.ActiveProcess.CommonVars.FiscalYear, 3)
            Declarations.UserCode = Scala.ActiveProcess.CommonVars.UserCode

        Catch
            MsgBox("Программа должна запускаться только из меню Scala", MsgBoxStyle.Critical, "Внимание!")
            Application.Exit()
        End Try
    End Sub

    Private Sub Button2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button2.Click
        '////////////////////////////////////////////////////////////////////////////////////////
        '//
        '// Процедура загрузки из Excel информации по поставщикам  
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
        '// Загрузка из Excel кодов ТН ВЭД  
        '//
        '////////////////////////////////////////////////////////////////////////////////////////
        Dim MySQLStr As String                      'SQL запрос
        Dim MyExcelCounter As Double                  'счетчик строк Excel
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
                    '---Код Scala
                    ScalaCode = Trim(appXLSRC.Worksheets(1).Range("A" & CStr(MyExcelCounter)).Value)
                    '---Проверяем, что такой код есть в Scala
                    MySQLStr = "SELECT COUNT(*) AS CC "
                    MySQLStr = MySQLStr & "FROM SC010300 WITH(NOLOCK) "
                    MySQLStr = MySQLStr & "WHERE (SC01001 = N'" & ScalaCode & "') "
                    InitMyConn(False)
                    InitMyRec(False, MySQLStr)
                    If (Declarations.MyRec.Fields("CC").Value = 0) Then
                        MsgBox("В импортируемом листе Excel в ячейке ""A" & CStr(MyExcelCounter) & """ проставлен неверный код запаса в Scala ", MsgBoxStyle.Critical, "Внимание!")
                        appXLSRC.DisplayAlerts = 0
                        appXLSRC.Workbooks.Close()
                        appXLSRC.DisplayAlerts = 1
                        appXLSRC.Quit()
                        appXLSRC = Nothing
                        trycloseMyRec()
                        Exit Sub
                    End If
                    trycloseMyRec()

                    '---Код ТН ВЭД
                    If appXLSRC.Worksheets(1).Range("B" & CStr(MyExcelCounter)).Value = Nothing Or Trim(appXLSRC.Worksheets(1).Range("B" & CStr(MyExcelCounter)).Value.ToString) = "" Then
                        MsgBox("В импортируемом листе Excel в ячейке ""B" & CStr(MyExcelCounter) & """ не проставлен код ТН ВЭД. ", MsgBoxStyle.Critical, "Внимание!")
                        appXLSRC.DisplayAlerts = 0
                        appXLSRC.Workbooks.Close()
                        appXLSRC.DisplayAlerts = 1
                        appXLSRC.Quit()
                        appXLSRC = Nothing
                        Exit Sub
                    End If
                    TNVedCode = Trim(appXLSRC.Worksheets(1).Range("B" & CStr(MyExcelCounter)).Value)
                    '---Проверяем, что такой код ТН ВЭД есть в Scala
                    MySQLStr = "SELECT COUNT(*) AS CC "
                    MySQLStr = MySQLStr & "FROM SY240300 "
                    MySQLStr = MySQLStr & "WHERE (SY24001 = N'BN') AND "
                    MySQLStr = MySQLStr & "(SY24002 = N'" & Trim(TNVedCode) & "') "
                    InitMyConn(False)
                    InitMyRec(False, MySQLStr)
                    If Declarations.MyRec.Fields("CC").Value = 0 Then
                        MsgBox("Код ТН ВЭД " & TNVedCode & " в строке " & MyExcelCounter & " оттсутствует в справочнике Scala. Данные обновлены не будут.", vbCritical, "Внимание!")
                        appXLSRC.DisplayAlerts = 0
                        appXLSRC.Workbooks.Close()
                        appXLSRC.DisplayAlerts = 1
                        appXLSRC.Quit()
                        appXLSRC = Nothing
                        trycloseMyRec()
                        Exit Sub
                    End If
                    '---занесение кода ТН ВЭД в Scala
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
                '---Сообщение об окончании процедуры
                MsgBox("Импорт информации о Кодах ТН ВЭД произведен. ", MsgBoxStyle.OkOnly, "Внимание!")
            End If
        End If

    End Sub

    Private Sub ImportDataFromLO()
        '////////////////////////////////////////////////////////////////////////////////////////
        '//
        '// Загрузка из LibreOffice кодов ТН ВЭД  
        '//
        '////////////////////////////////////////////////////////////////////////////////////////
        Dim MySQLStr As String                      'SQL запрос
        Dim MyExcelCounter As Double                  'счетчик строк Excel
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
                    '---Код Scala
                    ScalaCode = Trim(oSheet.getCellRangeByName("A" & MyExcelCounter).String)
                    '---Проверяем, что такой код есть в Scala
                    MySQLStr = "SELECT COUNT(*) AS CC "
                    MySQLStr = MySQLStr & "FROM SC010300 WITH(NOLOCK) "
                    MySQLStr = MySQLStr & "WHERE (SC01001 = N'" & ScalaCode & "') "
                    InitMyConn(False)
                    InitMyRec(False, MySQLStr)
                    If (Declarations.MyRec.Fields("CC").Value = 0) Then
                        MsgBox("В импортируемом листе Excel в ячейке ""A" & CStr(MyExcelCounter) & """ проставлен неверный код запаса в Scala ", MsgBoxStyle.Critical, "Внимание!")
                        oWorkBook.Close(True)
                        trycloseMyRec()
                        Exit Sub
                    End If
                    trycloseMyRec()

                    '---Код ТН ВЭД
                    If Trim(oSheet.getCellRangeByName("B" & MyExcelCounter).String).Equals("") Then
                        MsgBox("В импортируемом листе Excel в ячейке ""B" & CStr(MyExcelCounter) & """ не проставлен код ТН ВЭД. ", MsgBoxStyle.Critical, "Внимание!")
                        oWorkBook.Close(True)
                        Exit Sub
                    End If
                    TNVedCode = Trim(oSheet.getCellRangeByName("B" & MyExcelCounter).String)
                    '---Проверяем, что такой код ТН ВЭД есть в Scala
                    MySQLStr = "SELECT COUNT(*) AS CC "
                    MySQLStr = MySQLStr & "FROM SY240300 "
                    MySQLStr = MySQLStr & "WHERE (SY24001 = N'BN') AND "
                    MySQLStr = MySQLStr & "(SY24002 = N'" & Trim(TNVedCode) & "') "
                    InitMyConn(False)
                    InitMyRec(False, MySQLStr)
                    If Declarations.MyRec.Fields("CC").Value = 0 Then
                        MsgBox("Код ТН ВЭД " & TNVedCode & " в строке " & MyExcelCounter & " оттсутствует в справочнике Scala. Данные обновлены не будут.", vbCritical, "Внимание!")
                        oWorkBook.Close(True)
                        trycloseMyRec()
                        Exit Sub
                    End If

                    '---занесение кода ТН ВЭД в Scala
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
                MsgBox("Импорт информации о Кодах ТН ВЭД произведен.", vbOKOnly, "Внимание!")
            End If
        End If
    End Sub
End Class
