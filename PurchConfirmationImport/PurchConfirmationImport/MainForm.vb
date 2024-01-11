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
        '// после чего выводим прайс листы на доставку
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
        '// Процедура загрузки из Excel подтверждения на поставку от 1 поставщика  
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
        '// Загрузка из Excel подтверждения на поставку от 1 поставщика  
        '//
        '////////////////////////////////////////////////////////////////////////////////////////
        Dim appXLSRC As Object
        Dim MySuppCode As String
        Dim MySQLStr As String                      'SQL запрос
        Dim i As Double                             'счетчик строк
        Dim MyOrder As String                       'номер заказа на закупку
        Dim MySuppProductCode As String             'код товара поставщика
        Dim MyProductCode As String                 'код товара
        Dim MyConfDate As Date                      'подтвержденная дата
        Dim MyBackDate As Date                      'задолженная дата
        Dim MyERRStr As String                      'сообщения об ошибках
        Dim MyOrderFlag As String                   'Признак - прогрузка всего заказа или нет

        If OpenFileDialog1.ShowDialog() <> Windows.Forms.DialogResult.Cancel Then
            If (OpenFileDialog1.FileName = "") Then
            Else
                MyERRStr = ""

                Me.Cursor = Cursors.WaitCursor
                Me.Refresh()
                System.Windows.Forms.Application.DoEvents()

                appXLSRC = CreateObject("Excel.Application")
                appXLSRC.Workbooks.Open(OpenFileDialog1.FileName)

                '---Проверяем, что проставлен код поставщика
                MySuppCode = Trim(appXLSRC.Worksheets(1).Range("E1").Value)
                If MySuppCode = Nothing Then
                    MsgBox("В импортируемом листе Excel в ячейке 'E1' не проставлен код поставщика ", MsgBoxStyle.Critical, "Внимание!")
                    appXLSRC.DisplayAlerts = 0
                    appXLSRC.Workbooks.Close()
                    appXLSRC.DisplayAlerts = 1
                    appXLSRC.Quit()
                    appXLSRC = Nothing
                    Exit Sub
                End If

                '---проверяем что этот поставщик есть в Scala
                MySQLStr = "SELECT COUNT(PL01001) AS CC "
                MySQLStr = MySQLStr & "FROM PL010300 "
                MySQLStr = MySQLStr & "WHERE (PL01001 = N'" & MySuppCode & "')"
                InitMyConn(False)
                InitMyRec(False, MySQLStr)
                If (Declarations.MyRec.Fields("CC").Value = 0) Then
                    MsgBox("В импортируемом листе Excel в ячейке 'E1' проставлен неверный код поставщика в Scala ", MsgBoxStyle.Critical, "Внимание!")
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
                        '---проверяем - есть ли такой заказ на закупку от этого поставщика (незакрытый)
                        MySQLStr = "SELECT COUNT(PC01001) AS CC "
                        MySQLStr = MySQLStr & "FROM PC010300 "
                        MySQLStr = MySQLStr & "WHERE (PC01001 = N'" & MyOrder & "') AND (PC01002 <> 2) "
                        InitMyConn(False)
                        InitMyRec(False, MySQLStr)
                        If (Declarations.MyRec.Fields("CC").Value = 0) Then
                            trycloseMyRec()
                            MsgBox("В импортируемом листе Excel в ячейке 'B" & i & "' проставлен номер заказа на закупку, которого нет в Scala или который закрыт (2 типа) ", MsgBoxStyle.Critical, "Внимание!")
                        Else
                            trycloseMyRec()
                            If appXLSRC.Worksheets(1).Range("C" & i).Value <> Nothing Then
                                '---================================Прогрузка подтверждения для всего заказа
                                Try
                                    MyConfDate = CDate(appXLSRC.Worksheets(1).Range("E" & CStr(i)).Value.ToString)
                                    If appXLSRC.Worksheets(1).Range("F" & i).Value = Nothing Then
                                        MyBackDate = MyConfDate
                                    Else
                                        MyBackDate = CDate(appXLSRC.Worksheets(1).Range("F" & CStr(i)).Value.ToString)
                                    End If
                                    Try
                                        '---Ну и заносим обновления в Scala
                                        MySQLStr = "UPDATE PC030300 "
                                        MySQLStr = MySQLStr & "SET PC03016 = CONVERT(DATETIME, '" & Format(MyConfDate, "dd/MM/yyyy") & "', 103), "
                                        MySQLStr = MySQLStr & "PC03024 = CONVERT(DATETIME, '" & Format(MyConfDate, "dd/MM/yyyy") & "', 103), "
                                        MySQLStr = MySQLStr & "PC03031 = CONVERT(DATETIME, '" & Format(MyBackDate, "dd/MM/yyyy") & "', 103), "
                                        MySQLStr = MySQLStr & "PC03029 = N'1' "
                                        MySQLStr = MySQLStr & "WHERE (PC03001 = N'" & MyOrder & "') "
                                        Declarations.MyConn.Execute(MySQLStr)
                                    Catch
                                        MsgBox("В импортируемом листе Excel в ячейке 'F" & i & "' проставлена неверная дата ", MsgBoxStyle.Critical, "Внимание!")
                                    End Try
                                Catch
                                    MsgBox("В импортируемом листе Excel в ячейке 'E" & i & "' проставлена неверная дата ", MsgBoxStyle.Critical, "Внимание!")
                                End Try
                                '---===================================Конец прогрузки подтверждения для всего заказа
                            Else
                                '---================================Прогрузка подтверждения для одного запаса в заказе
                                Try
                                    MySuppProductCode = appXLSRC.Worksheets(1).Range("D" & CStr(i)).Value.ToString
                                    '---Проверяем - есть ли такой код товара поставщика у этого поставщика
                                    MySQLStr = "SELECT COUNT(*) AS CC "
                                    MySQLStr = MySQLStr & "FROM SC010300 "
                                    MySQLStr = MySQLStr & "WHERE (SC01060 = N'" & MySuppProductCode & "') AND (SC01058 = N'" & MySuppCode & "')"
                                    InitMyConn(False)
                                    InitMyRec(False, MySQLStr)
                                    If (Declarations.MyRec.Fields("CC").Value = 0) Then
                                        trycloseMyRec()
                                        '---Нет такого кода у такого поставщика
                                        MyERRStr = MyERRStr & "Строка " & i & " поставщик " & MySuppCode & " Код товара поставщика " & MySuppProductCode & " не найден" & Microsoft.VisualBasic.Chr(13)
                                    Else
                                        trycloseMyRec()
                                        '---Получаем наш код товара
                                        MySQLStr = "Select SC01001 AS CC "
                                        MySQLStr = MySQLStr & "FROM SC010300 "
                                        MySQLStr = MySQLStr & "WHERE (SC01060 = N'" & MySuppProductCode & "') AND (SC01058 = N'" & MySuppCode & "')"
                                        InitMyConn(False)
                                        InitMyRec(False, MySQLStr)
                                        MyProductCode = Declarations.MyRec.Fields("CC").Value.ToString
                                        trycloseMyRec()
                                        '---Проверяем - есть ли такой код товара в этом заказе на закупку
                                        MySQLStr = "SELECT COUNT(*) AS CC "
                                        MySQLStr = MySQLStr & "FROM PC030300 "
                                        MySQLStr = MySQLStr & "WHERE (PC03001 = N'" & MyOrder & "') AND (PC03005 = N'" & MyProductCode & "') "
                                        InitMyConn(False)
                                        InitMyRec(False, MySQLStr)
                                        If (Declarations.MyRec.Fields("CC").Value = 0) Then
                                            trycloseMyRec()
                                            '--- Нет такого кода в таком заказе
                                            MyERRStr = MyERRStr & "Строка " & i & " код товара " & MyProductCode & " Код товара поставщика " & MySuppProductCode & " не найден в заказе на закупку " & MyOrder & " " & Microsoft.VisualBasic.Chr(13)
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
                                                    '---Ну и заносим обновления в Scala
                                                    MySQLStr = "UPDATE PC030300 "
                                                    MySQLStr = MySQLStr & "SET PC03016 = CONVERT(DATETIME, '" & Format(MyConfDate, "dd/MM/yyyy") & "', 103), "
                                                    MySQLStr = MySQLStr & "PC03024 = CONVERT(DATETIME, '" & Format(MyConfDate, "dd/MM/yyyy") & "', 103), "
                                                    MySQLStr = MySQLStr & "PC03031 = CONVERT(DATETIME, '" & Format(MyBackDate, "dd/MM/yyyy") & "', 103), "
                                                    MySQLStr = MySQLStr & "PC03029 = N'1' "
                                                    MySQLStr = MySQLStr & "WHERE (PC03001 = N'" & MyOrder & "') AND (PC03005 = N'" & MyProductCode & "') "
                                                    Declarations.MyConn.Execute(MySQLStr)
                                                Catch
                                                    MsgBox("В импортируемом листе Excel в ячейке 'F" & i & "' проставлена неверная дата ", MsgBoxStyle.Critical, "Внимание!")
                                                End Try
                                            Catch
                                                MsgBox("В импортируемом листе Excel в ячейке 'E" & i & "' проставлена неверная дата ", MsgBoxStyle.Critical, "Внимание!")
                                            End Try
                                        End If
                                    End If
                                Catch
                                    MsgBox("В импортируемом листе Excel в ячейке 'D" & i & "' проставлен неверный код товара поставщика ", MsgBoxStyle.Critical, "Внимание!")
                                End Try
                                '---===================================Конец прогрузки подтверждения для одного запаса в заказе
                            End If
                        End If
                    Catch
                        MsgBox("В импортируемом листе Excel в ячейке 'B" & i & "' проставлен неверный номер заказа на закупку ", MsgBoxStyle.Critical, "Внимание!")
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
                If MyERRStr = "" Then '---нет ошибок
                    MsgBox("Импорт подтверждений заказов на закупку произведен", MsgBoxStyle.OkOnly, "Внимание!")
                Else
                    MyErrorForm = New ErrorForm
                    MyERRStr = "Во время импорта подтверждениия поставки были ошибки " & Chr(13) & MyERRStr
                    MyErrorForm.MyMsg = MyERRStr
                    MyErrorForm.ShowDialog()
                End If
                    
            End If
        End If

    End Sub

    Private Sub ImportDataFromLO()
        '////////////////////////////////////////////////////////////////////////////////////////
        '//
        '// Загрузка из Libre Office подтверждения на поставку от 1 поставщика  
        '//
        '////////////////////////////////////////////////////////////////////////////////////////
        Dim MyERRStr As String                      'сообщения об ошибках
        Dim MySuppCode As String
        Dim MySQLStr As String                      'SQL запрос
        Dim i As Double                             'счетчик строк
        Dim MyOrder As String                       'номер заказа на закупку
        Dim MyConfDate As Date                      'подтвержденная дата
        Dim MyBackDate As Date                      'задолженная дата
        Dim MySuppProductCode As String             'код товара поставщика
        Dim MyProductCode As String                 'код товара
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

                    '---Проверяем, что проставлен код поставщика
                    MySuppCode = oSheet.getCellRangeByName("E1").String
                    If MySuppCode.Equals("") Then
                        MsgBox("В импортируемом листе Excel в ячейке 'E1' не проставлен код поставщика ", MsgBoxStyle.Critical, "Внимание!")
                        Me.Cursor = Cursors.Default
                        oWorkBook.Close(True)
                        Exit Sub
                    End If

                    '---проверяем что этот поставщик есть в Scala
                    MySQLStr = "SELECT COUNT(PL01001) AS CC "
                    MySQLStr = MySQLStr & "FROM PL010300 WITH(NOLOCK) "
                    MySQLStr = MySQLStr & "WHERE (PL01001 = N'" & MySuppCode & "')"
                    InitMyConn(False)
                    InitMyRec(False, MySQLStr)
                    If (Declarations.MyRec.Fields("CC").Value = 0) Then
                        trycloseMyRec()
                        MsgBox("В импортируемом листе Excel в ячейке 'E1' проставлен неверный код поставщика в Scala ", MsgBoxStyle.Critical, "Внимание!")
                        Me.Cursor = Cursors.Default
                        oWorkBook.Close(True)
                        Exit Sub
                    End If
                    trycloseMyRec()

                    i = 4
                    While Not oSheet.getCellRangeByName("B" & i).String.Equals("")
                        MyOrder = Microsoft.VisualBasic.Right("0000000000" & oSheet.getCellRangeByName("B" & i).String, 10)
                        '---проверяем - есть ли такой заказ на закупку от этого поставщика (незакрытый)
                        MySQLStr = "SELECT COUNT(PC01001) AS CC "
                        MySQLStr = MySQLStr & "FROM PC010300 WITH(NOLOCK) "
                        MySQLStr = MySQLStr & "WHERE (PC01001 = N'" & MyOrder & "') " 'AND (PC01002 <> 2) "
                        InitMyConn(False)
                        InitMyRec(False, MySQLStr)
                        If (Declarations.MyRec.Fields("CC").Value = 0) Then
                            trycloseMyRec()
                            MsgBox("В импортируемом листе Excel в ячейке 'B" & i & "' проставлен номер заказа на закупку, которого нет в Scala или который закрыт (2 типа) ", MsgBoxStyle.Critical, "Внимание!")
                        Else
                            trycloseMyRec()
                            If Not oSheet.getCellRangeByName("C" & i).String.Equals("") Then
                                '---================================Прогрузка подтверждения для всего заказа
                                Try
                                    MyConfDate = DateTime.FromOADate(oSheet.getCellRangeByName("E" & CStr(i)).Value)
                                    If oSheet.getCellRangeByName("F" & i).String.Equals("") Then
                                        MyBackDate = MyConfDate
                                    Else
                                        MyBackDate = DateTime.FromOADate(oSheet.getCellRangeByName("F" & CStr(i)).Value)
                                    End If
                                    Try
                                        '---Ну и заносим обновления в Scala
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
                                        MsgBox(ex.Message, MsgBoxStyle.OkOnly, "Внимание!")
                                    End Try
                                Catch ex As Exception
                                    MsgBox("В импортируемом листе Excel в ячейке 'E" & i & "' проставлена неверная дата ", MsgBoxStyle.Critical, "Внимание!")
                                End Try
                            Else
                                '---================================Прогрузка подтверждения для одного запаса в заказе
                                MySuppProductCode = oSheet.getCellRangeByName("D" & CStr(i)).String
                                '---Проверяем - есть ли такой код товара поставщика у этого поставщика
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
                                    '---Нет такого кода у такого поставщика
                                    MyERRStr = MyERRStr & "Строка " & i & " поставщик " & MySuppCode & " Код товара поставщика " & MySuppProductCode & " не найден" & Microsoft.VisualBasic.Chr(13)
                                Else
                                    trycloseMyRec()
                                    '---Получаем наш код товара
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
                                    '---Проверяем - есть ли такой код товара в этом заказе на закупку
                                    MySQLStr = "SELECT COUNT(*) AS CC "
                                    MySQLStr = MySQLStr & "FROM PC030300 "
                                    MySQLStr = MySQLStr & "WHERE (PC03001 = N'" & MyOrder & "') AND (PC03005 = N'" & MyProductCode & "') "
                                    InitMyConn(False)
                                    InitMyRec(False, MySQLStr)
                                    If (Declarations.MyRec.Fields("CC").Value = 0) Then
                                        trycloseMyRec()
                                        '--- Нет такого кода в таком заказе
                                        MyERRStr = MyERRStr & "Строка " & i & " код товара " & MyProductCode & " Код товара поставщика " & MySuppProductCode & " не найден в заказе на закупку " & MyOrder & " " & Microsoft.VisualBasic.Chr(13)
                                    Else
                                        trycloseMyRec()
                                        Try
                                            MyConfDate = DateTime.FromOADate(oSheet.getCellRangeByName("E" & CStr(i)).Value)
                                            If oSheet.getCellRangeByName("F" & i).String.Equals("") Then
                                                MyBackDate = MyConfDate
                                            Else
                                                MyBackDate = DateTime.FromOADate(oSheet.getCellRangeByName("F" & CStr(i)).Value)
                                            End If
                                            '---Ну и заносим обновления в Scala
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
                                                MsgBox(ex.Message, MsgBoxStyle.OkOnly, "Внимание!")
                                            End Try
                                        Catch
                                            MsgBox("В импортируемом листе Excel в ячейке 'F" & i & "' проставлена неверная дата ", MsgBoxStyle.Critical, "Внимание!")
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
                    MsgBox("ошибка : " & ex.Message, MsgBoxStyle.Critical, "Внимание!")
                Finally
                    Try
                        oWorkBook.Close(True)
                    Catch ex As Exception
                    End Try
                    Declarations.MyConn.Close()
                    Declarations.MyConn = Nothing
                End Try
                Me.Cursor = Cursors.Default
                If MyERRStr = "" Then '---нет ошибок
                    MsgBox("Импорт подтверждений заказов на закупку произведен", MsgBoxStyle.OkOnly, "Внимание!")
                Else
                    MyErrorForm = New ErrorForm
                    MyERRStr = "Во время импорта подтверждениия поставки были ошибки " & Chr(13) & MyERRStr
                    MyErrorForm.MyMsg = MyERRStr
                    MyErrorForm.ShowDialog()
                End If
            End If
        End If
    End Sub
End Class
