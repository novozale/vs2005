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
        '// Процедура загрузки информации по поставщикам  
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

    Private Sub ImportDataFromExcel()
        '////////////////////////////////////////////////////////////////////////////////////////
        '//
        '// Загрузка из Excel информации по производителям  
        '//
        '////////////////////////////////////////////////////////////////////////////////////////
        Dim MySQLStr As String                      'SQL запрос
        Dim appXLSRC As Object
        Dim MyVersion As String                     'Версия документа
        Dim ScalaCode As String                     'код товара в Scala
        Dim ManCode As String                       'Код товара производителя
        Dim Manufacturer As Integer                 'Код производителя
        Dim i As Double                             'счетчик строк

        If OpenFileDialog1.ShowDialog() <> Windows.Forms.DialogResult.Cancel Then
            If (OpenFileDialog1.FileName = "") Then
            Else
                Me.Cursor = Cursors.WaitCursor
                Me.Refresh()
                System.Windows.Forms.Application.DoEvents()

                appXLSRC = CreateObject("Excel.Application")
                appXLSRC.Workbooks.Open(OpenFileDialog1.FileName)

                '---Проверяем версию листа Excel
                MyVersion = Trim(appXLSRC.Worksheets(1).Range("A1").Value)
                MySQLStr = "SELECT Version "
                MySQLStr = MySQLStr & "FROM tbl_VersionImportItemsFromExcel WITH(NOLOCK) "
                MySQLStr = MySQLStr & "WHERE (Name = N'Импорт кодов производителей и производителей') "
                InitMyConn(False)
                InitMyRec(False, MySQLStr)
                If Declarations.MyRec.BOF = True And Declarations.MyRec.EOF = True Then
                    MsgBox("В Scala не проставлена текущая версия листа Excel. Обратитесь к администратору", vbCritical, "Внимание!")
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
                        MsgBox("Вы пытаетесь работать с некорректной версией листа Excel. Надо работать с версией " & MyVersion & ".", vbCritical, "Внимание!")
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
                '---читаем данные
                While Not appXLSRC.Worksheets(1).Range("A" & CStr(i)).Value = Nothing
                    ScalaCode = Trim(appXLSRC.Worksheets(1).Range("A" & CStr(i)).Value)
                    '---проверяем, что такой код есть в Scala
                    MySQLStr = "SELECT COUNT(*) AS CC "
                    MySQLStr = MySQLStr & "FROM SC010300 WITH(NOLOCK) "
                    MySQLStr = MySQLStr & "WHERE (SC01001 = N'" & ScalaCode & "') "
                    InitMyConn(False)
                    InitMyRec(False, MySQLStr)
                    If (Declarations.MyRec.Fields("CC").Value = 0) Then
                        MsgBox("В импортируемом листе Excel в ячейке ""A" & CStr(i) & """ проставлен неверный код запаса в Scala ", MsgBoxStyle.Critical, "Внимание!")
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
                        MsgBox("В импортируемом листе Excel в ячейке ""B" & CStr(i) & """ не проставлен код товара производителя. ", MsgBoxStyle.Critical, "Внимание!")
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
                        MsgBox("В импортируемом листе Excel в ячейке ""C" & CStr(i) & """ не проставлен код производителя. ", MsgBoxStyle.Critical, "Внимание!")
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
                            MsgBox("В импортируемом листе Excel в ячейке ""C" & CStr(i) & """ Код производителя в Scala должен быть целым числом. ", MsgBoxStyle.Critical, "Внимание!")
                            appXLSRC.DisplayAlerts = 0
                            appXLSRC.Workbooks.Close()
                            appXLSRC.DisplayAlerts = 1
                            appXLSRC.Quit()
                            appXLSRC = Nothing
                            trycloseMyRec()
                            Exit Sub
                        End Try

                        '---проверяем, что такой Производитель есть в Scala
                        MySQLStr = "SELECT COUNT(*) AS CC "
                        MySQLStr = MySQLStr & "FROM tbl_Manufacturers "
                        MySQLStr = MySQLStr & "WHERE (ID = " & Manufacturer & ")"
                        InitMyConn(False)
                        InitMyRec(False, MySQLStr)
                        If (Declarations.MyRec.Fields("CC").Value = 0) Then
                            MsgBox("В импортируемом листе Excel в ячейке ""C" & CStr(i) & """ проставлен неверный код производителя в Scala ", MsgBoxStyle.Critical, "Внимание!")
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


                    '---Обновляем данные------------------------------------------------------------------
                    '---Код товара производителя и код производителя
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
                '---Сообщение об окончании процедуры
                MsgBox("Импорт информации о кодах товаров производителя и производителях произведен.", MsgBoxStyle.OkOnly, "Внимание!")
            End If
        End If
    End Sub

    Private Sub ImportDataFromLO()
        '////////////////////////////////////////////////////////////////////////////////////////
        '//
        '// Загрузка из LibreOffice информации по производителям
        '//
        '////////////////////////////////////////////////////////////////////////////////////////
        Dim MySQLStr As String                      'SQL запрос
        Dim MyVersion As String                     'Версия документа
        Dim ScalaCode As String                     'код товара в Scala
        Dim ManCode As String                       'Код товара производителя
        Dim Manufacturer As Integer                 'Код производителя
        Dim i As Double                             'счетчик строк
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

                    '---Проверяем версию листа Excel
                    MyVersion = oSheet.getCellRangeByName("A1").String
                    MySQLStr = "SELECT Version "
                    MySQLStr = MySQLStr & "FROM tbl_VersionImportItemsFromExcel WITH(NOLOCK) "
                    MySQLStr = MySQLStr & "WHERE (Name = N'Импорт кодов производителей и производителей') "
                    InitMyConn(False)
                    InitMyRec(False, MySQLStr)
                    If Declarations.MyRec.BOF = True And Declarations.MyRec.EOF = True Then
                        MsgBox("В Scala не проставлена текущая версия листа Excel. Обратитесь к администратору", vbCritical, "Внимание!")
                        oWorkBook.Close(True)
                        trycloseMyRec()
                        Exit Sub
                    Else
                        If Trim(Declarations.MyRec.Fields("Version").Value) = MyVersion Then
                            trycloseMyRec()
                        Else
                            MsgBox("Вы пытаетесь работать с некорректной версией листа Excel. Надо работать с версией " & MyVersion & ".", vbCritical, "Внимание!")
                            oWorkBook.Close(True)
                            trycloseMyRec()
                            Exit Sub
                        End If
                    End If

                    i = 7
                    '---читаем данные
                    While Not oSheet.getCellRangeByName("A" & i).String.Equals("")
                        ScalaCode = Trim(oSheet.getCellRangeByName("A" & i).String)
                        '---проверяем, что такой код есть в Scala
                        MySQLStr = "SELECT COUNT(*) AS CC "
                        MySQLStr = MySQLStr & "FROM SC010300 WITH(NOLOCK) "
                        MySQLStr = MySQLStr & "WHERE (SC01001 = N'" & ScalaCode & "') "
                        InitMyConn(False)
                        InitMyRec(False, MySQLStr)
                        If (Declarations.MyRec.Fields("CC").Value = 0) Then
                            MsgBox("В импортируемом листе Excel в ячейке ""A" & CStr(i) & """ проставлен неверный код запаса в Scala ", MsgBoxStyle.Critical, "Внимание!")
                            oWorkBook.Close(True)
                            trycloseMyRec()
                            Exit Sub
                        End If
                        trycloseMyRec()

                        '---Код товара производителя
                        ManCode = Trim(oSheet.getCellRangeByName("B" & CStr(i)).String)
                        If ManCode.Equals("") Then
                            MsgBox("В импортируемом листе Excel в ячейке ""B" & CStr(i) & """ не проставлен код товара производителя. ", MsgBoxStyle.Critical, "Внимание!")
                            oWorkBook.Close(True)
                            Exit Sub
                        End If

                        '---Код производителя
                        Try
                            Manufacturer = Trim(oSheet.getCellRangeByName("C" & CStr(i)).Value)
                        Catch ex As Exception
                            MsgBox("В импортируемом листе Excel в ячейке ""C" & CStr(i) & """ Код производителя в Scala должен быть целым числом. ", MsgBoxStyle.Critical, "Внимание!")
                            oWorkBook.Close(True)
                            Exit Sub
                        End Try

                        '---проверяем, что такой Производитель есть в Scala
                        MySQLStr = "SELECT COUNT(*) AS CC "
                        MySQLStr = MySQLStr & "FROM tbl_Manufacturers "
                        MySQLStr = MySQLStr & "WHERE (ID = " & Manufacturer & ")"
                        InitMyConn(False)
                        InitMyRec(False, MySQLStr)
                        If (Declarations.MyRec.Fields("CC").Value = 0) Then
                            MsgBox("В импортируемом листе Excel в ячейке ""C" & CStr(i) & """ проставлен неверный код производителя в Scala ", MsgBoxStyle.Critical, "Внимание!")
                            oWorkBook.Close(True)
                            trycloseMyRec()
                            Exit Sub
                        End If
                        trycloseMyRec()

                        '---Обновляем данные------------------------------------------------------------------
                        '---Код товара производителя и код производителя
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
                    MsgBox("ошибка : " & ex.Message, MsgBoxStyle.Critical, "Внимание!")
                End Try
                oWorkBook.Close(True)
                '---Сообщение об окончании процедуры
                MsgBox("Импорт информации о кодах товаров производителя и производителях произведен.", MsgBoxStyle.OkOnly, "Внимание!")
            End If
        End If
    End Sub
End Class
