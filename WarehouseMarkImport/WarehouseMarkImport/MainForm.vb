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
        '// Процедура загрузки из Excel   
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
        '// Загрузка из Excel   
        '//
        '////////////////////////////////////////////////////////////////////////////////////////
        Dim MySQLStr As String                      'SQL запрос
        Dim appXLSRC As Object
        Dim MyVersion As String                     'Версия документа
        Dim ScalaCode As String                     'код товара в Scala
        Dim WHMark As String                        'Признак складской - нескладской
        Dim i As Double                             'счетчик строк

        Dim cmd As New ADODB.Command                    'команда (spp процедура)
        Dim MyParam As ADODB.Parameter                  'передаваемый параметр номер 1
        Dim MyParam1 As ADODB.Parameter                 'передаваемый параметр номер 2

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
                MySQLStr = MySQLStr & "WHERE (Name = N'Признак складского ассортимента') "
                InitMyConn(False)
                InitMyRec(False, MySQLStr)
                If Declarations.MyRec.BOF = True And Declarations.MyRec.EOF = True Then
                    MsgBox("В Scala не проставлена текущая версия листа Excel. Обратитесь к администратору", vbCritical, "Внимание!")
                    trycloseMyRec()
                    Exit Sub
                Else
                    If Trim(Declarations.MyRec.Fields("Version").Value) = MyVersion Then
                        trycloseMyRec()
                    Else
                        MsgBox("Вы пытаетесь работать с некорректной версией листа Excel. Надо работать с версией " & MyVersion & ".", vbCritical, "Внимание!")
                        trycloseMyRec()
                        Exit Sub
                    End If
                End If


                '---читаем данные, запускаем хранимую процедуру
                i = 7

                '---Подготовка хранимой процедуры----------------------------------------------------------------
                cmd.ActiveConnection = Declarations.MyConn
                cmd.CommandText = "spp_ImportWHMark"
                cmd.CommandType = ADODB.CommandTypeEnum.adCmdStoredProc
                cmd.CommandTimeout = 300

                '----Создание параметров---------------------------------------------------
                '---Код товара в Scala
                MyParam = cmd.CreateParameter("@MyScalaCode", ADODB.DataTypeEnum.adBSTR, ADODB.ParameterDirectionEnum.adParamInput, 35)
                cmd.Parameters.Append(MyParam)
                '---Признак складской
                MyParam1 = cmd.CreateParameter("@MyWHMark", ADODB.DataTypeEnum.adBSTR, ADODB.ParameterDirectionEnum.adParamInput, 10)
                cmd.Parameters.Append(MyParam1)


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

                    WHMark = Trim(appXLSRC.Worksheets(1).Range("B" & CStr(i)).Value)
                    If CheckWHMark(WHMark) = False Then
                        MsgBox("В импортируемом листе Excel в ячейке ""B" & CStr(i) & """ проставлен неверный признак складского запаса. Можно использовать только 0 и 1, длина признака - 4 знака.", MsgBoxStyle.Critical, "Внимание!")
                        appXLSRC.DisplayAlerts = 0
                        appXLSRC.Workbooks.Close()
                        appXLSRC.DisplayAlerts = 1
                        appXLSRC.Quit()
                        appXLSRC = Nothing
                        Exit Sub
                    End If


                    '----значения параметров хранимой процедуры----------------------------------------------
                    MyParam.Value = ScalaCode
                    MyParam1.Value = WHMark

                    '---запуск хранимой процедуры------------------------------------------------------
                    cmd.Execute()


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
                MsgBox("Импорт признака складской произведен. Расчет ROP, МЖЗ, АБС, XYZ - анализ будут произведены ночью.", MsgBoxStyle.OkOnly, "Внимание!")
            End If
        End If
    End Sub

    Private Sub ImportDataFromLO()
        '////////////////////////////////////////////////////////////////////////////////////////
        '//
        '// Загрузка из LibreOffice   
        '//
        '////////////////////////////////////////////////////////////////////////////////////////
        Dim oServiceManager As Object
        Dim oDispatcher As Object
        Dim oDesktop As Object
        Dim oWorkBook As Object
        Dim oSheet As Object
        Dim oFileName As String
        Dim MyVersion As String                       'Версия документа
        Dim MySQLStr As String                        'SQL запрос
        Dim ScalaCode As String                     'код товара в Scala
        Dim WHMark As String                        'Признак складской - нескладской
        Dim i As Integer                              'счетчик строк

        Dim cmd As New ADODB.Command                    'команда (spp процедура)
        Dim MyParam As ADODB.Parameter                  'передаваемый параметр номер 1
        Dim MyParam1 As ADODB.Parameter                 'передаваемый параметр номер 2

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

                '---Проверяем версию листа Excel
                MyVersion = oSheet.getCellRangeByName("A1").String
                If MyVersion = "" Then
                    MsgBox("В импортируемом листе Excel в ячейке 'A1' не проставлена версия листа Excel ", MsgBoxStyle.Critical, "Внимание!")
                    oWorkBook.Close(True)
                    Exit Sub
                Else
                    MySQLStr = "SELECT Version "
                    MySQLStr = MySQLStr & "FROM tbl_VersionImportItemsFromExcel "
                    MySQLStr = MySQLStr & "WHERE (Name = N'Признак складского ассортимента') "
                    InitMyConn(False)
                    InitMyRec(False, MySQLStr)
                    If Declarations.MyRec.BOF = True And Declarations.MyRec.EOF = True Then
                        MsgBox("В Scala не проставлена текущая версия листа Excel. Обратитесь к администратору", vbCritical, "Внимание!")
                        trycloseMyRec()
                        oWorkBook.Close(True)
                        Exit Sub
                    Else
                        If Trim(Declarations.MyRec.Fields("Version").Value) = MyVersion Then
                            trycloseMyRec()
                        Else
                            MsgBox("Вы пытаетесь работать с некорректной версией листа Excel. Надо работать с версией " & Declarations.MyRec.Fields("Version").Value & ".", vbCritical, "Внимание!")
                            trycloseMyRec()
                            oWorkBook.Close(True)
                            Exit Sub
                        End If
                    End If
                End If

                '---читаем данные, запускаем хранимую процедуру
                i = 7

                '---Подготовка хранимой процедуры----------------------------------------------------------------
                cmd.ActiveConnection = Declarations.MyConn
                cmd.CommandText = "spp_ImportWHMark"
                cmd.CommandType = ADODB.CommandTypeEnum.adCmdStoredProc
                cmd.CommandTimeout = 300

                '----Создание параметров---------------------------------------------------
                '---Код товара в Scala
                MyParam = cmd.CreateParameter("@MyScalaCode", ADODB.DataTypeEnum.adBSTR, ADODB.ParameterDirectionEnum.adParamInput, 35)
                cmd.Parameters.Append(MyParam)
                '---Признак складской
                MyParam1 = cmd.CreateParameter("@MyWHMark", ADODB.DataTypeEnum.adBSTR, ADODB.ParameterDirectionEnum.adParamInput, 10)
                cmd.Parameters.Append(MyParam1)

                While oSheet.getCellRangeByName("A" & CStr(i)).String.Equals("") = False
                    '---Код Scala
                    ScalaCode = Trim(oSheet.getCellRangeByName("A" & i).String)
                    '---Проверяем, что такой код есть в Scala
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

                    '---Признак складской
                    WHMark = Trim(oSheet.getCellRangeByName("B" & i).String)
                    If CheckWHMark(WHMark) = False Then
                        MsgBox("В импортируемом листе Excel в ячейке ""B" & CStr(i) & """ проставлен неверный признак складского запаса. Можно использовать только 0 и 1, длина признака - 4 знака.", MsgBoxStyle.Critical, "Внимание!")
                        oWorkBook.Close(True)
                        trycloseMyRec()
                        Exit Sub
                    End If

                    '----значения параметров хранимой процедуры----------------------------------------------
                    MyParam.Value = ScalaCode
                    MyParam1.Value = WHMark

                    '---запуск хранимой процедуры------------------------------------------------------
                    cmd.Execute()

                    Label3.Text = CStr(i - 6)
                    Me.Refresh()
                    System.Windows.Forms.Application.DoEvents()
                    i = i + 1
                End While
                trycloseMyRec()
                oWorkBook.Close(True)
                MsgBox("Импорт признака складской произведен. Расчет ROP, МЖЗ, АБС, XYZ - анализ будут произведены ночью.", vbOKOnly, "Внимание!")
            End If
        End If
    End Sub

    Private Function CheckWHMark(ByVal WHMark As String) As Boolean
        '////////////////////////////////////////////////////////////////////////////////////////
        '//
        '// Проверка правильности заполнения признака складской
        '//
        '////////////////////////////////////////////////////////////////////////////////////////
        Dim i As Integer

        If Len(WHMark) <> 4 Then
            MsgBox("Длина признака складской должна быть 4 знака ", MsgBoxStyle.Critical, "Внимание!")
            CheckWHMark = False
            Exit Function
        End If

        For i = 0 To Len(WHMark) - 1
            If Mid(WHMark, i + 1, 1) <> "0" And Mid(WHMark, i + 1, 1) <> "1" Then
                MsgBox("В признаке складской должны использоваться только цифры 0 и 1 ", MsgBoxStyle.Critical, "Внимание!")
                CheckWHMark = False
                Exit Function
            End If
        Next
        CheckWHMark = True
    End Function
End Class